#!/usr/bin/env python3
"""
Proyecto Estadísticas de YouTube - Etapa III: iteración y mejora del modelo

Uso (después de haber ejecutado analisis.py):
    python iteracion.py --data data --out salida
    python iteracion.py --data data --out salida --rapido      # solo para verificar que corre

Cada iteración cambia UNA cosa respecto a un modelo de comparación y se contrasta con la prueba t remuestreada
corregida de Nadeau y Bengio (2003) sobre las mismas particiones. Genera salida/results_iii.json, figuras fig3_*.png,
modelo_final_iii.joblib y modelo_iii_metadata.json.
"""
import argparse, json, platform, datetime as dt
from pathlib import Path
import numpy as np
import pandas as pd
from scipy import stats
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import joblib, sklearn
import analisis as A
from analisis import (cargar, limpiar, construir, NUM_B, CAT_B, ETIQ, SEED, ct, cv_reg, corr_t, modelos_reg,
                      modelos_clf, savefig, C1, C2, C3)
from sklearn.base import BaseEstimator, clone
from sklearn.calibration import CalibratedClassifierCV
from sklearn.compose import ColumnTransformer
from sklearn.decomposition import TruncatedSVD
from sklearn.ensemble import HistGradientBoostingRegressor, HistGradientBoostingClassifier
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.feature_selection import SelectFromModel, SelectKBest, mutual_info_regression
from sklearn.impute import SimpleImputer
from sklearn.linear_model import RidgeCV, LassoCV, LogisticRegression, LinearRegression
from sklearn.metrics import (accuracy_score, balanced_accuracy_score, f1_score, roc_auc_score, confusion_matrix,
                             brier_score_loss, mean_squared_error)
from sklearn.model_selection import (KFold, RepeatedKFold, RepeatedStratifiedKFold, RandomizedSearchCV,
                                     train_test_split, cross_val_predict)
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

NREPS, NITER, CV_INT = 2, 20, 3        # --rapido los reduce
EXTRA = ["vader_max", "vader_min", "strong_pos", "strong_neg", "lex_div", "emoji_share"]
NUM_NCL = [c for c in NUM_B if c not in ("clike_mean", "clike_max", "sent_w")]   # sin likes de comentarios
NUM_X = NUM_B + EXTRA
ETIQ.update({"vader_max": "VADER máx. entre comentarios", "vader_min": "VADER mín. entre comentarios",
             "strong_pos": "% comentarios muy positivos", "strong_neg": "% comentarios muy negativos",
             "lex_div": "Diversidad léxica", "emoji_share": "% con emojis"})
GRUPOS = {
    "Sentimiento de los comentarios": ["sent_mean", "sent_std", "pos_share", "neg_share", "neu_share", "vader_mean", "vader_std"] + EXTRA[:4],
    "Likes de los comentarios": ["sent_w", "clike_mean", "clike_max"],
    "Volumen de comentarios": ["n_com"],
    "Estilo textual": ["len_mean", "words_mean", "q_share", "excl_share", "upper_share", "lex_div", "emoji_share"],
    "Metadatos del video": ["age_days", "title_len", "title_vader", "title_q", "Keyword"],
    "Texto (TF-IDF / LSA)": ["text"],
}


# ---------------------------------------------------------------------------------------------- utilidades
def rmse(y, p): return float(np.sqrt(mean_squared_error(y, p)))


def extras(c, D):
    c = c.copy()
    c["strong_pos"] = (c["vader"] >= .5).astype(float); c["strong_neg"] = (c["vader"] <= -.5).astype(float)
    c["emoji"] = c["Comment"].str.contains(r"[\U0001F300-\U0001FAFF\u2600-\u27BF]").astype(float)
    c["lex"] = c["Comment"].apply(lambda s: len(set(s.lower().split())) / max(len(s.split()), 1))
    E = c.groupby("Video ID").agg(vader_max=("vader", "max"), vader_min=("vader", "min"), strong_pos=("strong_pos", "mean"),
                                  strong_neg=("strong_neg", "mean"), lex_div=("lex", "mean"), emoji_share=("emoji", "mean"))
    return D.join(E)


class Apilado(BaseEstimator):
    """Apilamiento (Wolpert, 1992) propio: robusto con DataFrames y texto. meta=None -> promedio simple."""
    def __init__(self, bases, meta=None, tipo="reg", cv=3):
        self.bases, self.meta, self.tipo, self.cv = bases, meta, tipo, cv

    def _z(self, ests, X):
        f = (lambda e: e.predict(X)) if self.tipo == "reg" else (lambda e: e.predict_proba(X)[:, 1])
        return np.column_stack([f(e) for e in ests])

    def fit(self, X, y):
        self.fit_ = [clone(b).fit(X, y) for b in self.bases]
        if self.meta is not None:
            met = "predict" if self.tipo == "reg" else "predict_proba"
            Z = []
            for b in self.bases:
                p = cross_val_predict(clone(b), X, y, cv=self.cv, method=met)
                Z.append(p if self.tipo == "reg" else p[:, 1])
            self.meta_ = clone(self.meta).fit(np.column_stack(Z), y)
        self.classes_ = np.array([0, 1])
        return self

    def predict(self, X):
        Z = self._z(self.fit_, X)
        if self.tipo == "reg":
            return self.meta_.predict(Z) if self.meta is not None else Z.mean(axis=1)
        return (self.predict_proba(X)[:, 1] >= .5).astype(int)

    def predict_proba(self, X):
        Z = self._z(self.fit_, X)
        p = self.meta_.predict_proba(Z)[:, 1] if self.meta is not None else Z.mean(axis=1)
        return np.column_stack([1 - p, p])


def P(prep, m, sel=None):
    pasos = [("prep", prep)] + ([("sel", sel)] if sel is not None else []) + [("m", m)]
    return Pipeline(pasos)


def hgb(**kw):
    base = dict(learning_rate=.05, max_iter=A.MAXIT, max_depth=4, min_samples_leaf=15, l2_regularization=1.0)
    return HistGradientBoostingRegressor(random_state=SEED, **{**base, **kw})


def hgbc(**kw):
    base = dict(learning_rate=.05, max_iter=A.MAXIT, max_depth=4, min_samples_leaf=15, l2_regularization=1.0)
    return HistGradientBoostingClassifier(random_state=SEED, **{**base, **kw})


def ct_lsa(num, cat, k=50):
    return ColumnTransformer([
        ("num", Pipeline([("i", SimpleImputer(strategy="median")), ("s", StandardScaler())]), num),
        ("cat", OneHotEncoder(handle_unknown="ignore", min_frequency=5), cat),
        ("lsa", Pipeline([("tf", TfidfVectorizer(ngram_range=(1, 2), min_df=3, max_features=20000, sublinear_tf=True, stop_words="english")),
                          ("svd", TruncatedSVD(k, random_state=SEED))]), "text")])


def disponibles():
    H = {"python": platform.python_version(), "scikit-learn": sklearn.__version__}
    for mod in ("lightgbm", "xgboost"):
        try:
            m = __import__(mod); H[mod] = m.__version__
        except Exception:
            H[mod] = None
    return H


# ---------------------------------------------------------------------------------------------- registro de modelos
def registro_reg(H):
    base = modelos_reg(); AL = np.logspace(-2, 3, 12)
    R = []

    def add(id_, nombre, it, herr, est, cols, cfg=1, compl=1):
        R.append(dict(id=id_, nombre=nombre, it=it, herr=herr, est=est, cols=cols, cfg=cfg, compl=compl))
    e, n, c, k = base["Referencia (mediana)"]; add("R0", "Referencia (mediana)", "I0", "DummyRegressor", e, n + c, k, 0)
    e, n, c, k = base["Ridge – B (+metadatos)"]; add("R1", "Ridge – B (E-II)", "I0", "RidgeCV", e, n + c, k, 1)
    e, n, c, k = base["Grad. boosting – B"]; add("R2", "Grad. boosting – B (E-II, malla)", "I0", "GridSearchCV", e, n + c, k, 3)
    e, n, c, k = base["Híbrido B + TF-IDF + Ridge"]; add("R3", "Híbrido B + TF-IDF + Ridge (E-II)", "I0", "TF-IDF + RidgeCV", e, n + c + ["text"], k, 2)
    add("R4", "HGB fijo – B", "I1", "HistGradientBoosting", P(ct(NUM_B, CAT_B), hgb()), NUM_B + CAT_B, 1, 3)
    add("R5", "HGB fijo – B sin likes de comentarios", "I1", "Ablación de variables", P(ct(NUM_NCL, CAT_B), hgb()), NUM_NCL + CAT_B, 1, 3)
    add("R6", "Ridge – B sin likes de comentarios", "I1", "Ablación de variables", P(ct(NUM_NCL, CAT_B), RidgeCV(alphas=AL)), NUM_NCL + CAT_B, 12, 1)
    add("R7", "HGB – B + Lasso (SelectFromModel)", "I2", "LassoCV + SelectFromModel",
        P(ct(NUM_B, CAT_B), hgb(), SelectFromModel(LassoCV(cv=3, max_iter=5000, random_state=SEED))), NUM_B + CAT_B, 1, 3)
    mi = lambda X, y: mutual_info_regression(X, y, random_state=SEED)
    add("R8", "HGB – B + información mutua (k = 15)", "I2", "SelectKBest (información mutua)",
        P(ct(NUM_B, CAT_B), hgb(), SelectKBest(mi, k=15)), NUM_B + CAT_B, 1, 3)
    add("R9", "HGB fijo – B + variables nuevas", "I3", "Ingeniería de variables", P(ct(NUM_X, CAT_B), hgb()), NUM_X + CAT_B, 1, 3)
    add("R10", "Ridge – B + variables nuevas", "I3", "Ingeniería de variables", P(ct(NUM_X, CAT_B), RidgeCV(alphas=AL)), NUM_X + CAT_B, 12, 1)
    add("R11", "HGB fijo – B + nuevas + LSA (50)", "I3", "TruncatedSVD (LSA)", P(ct_lsa(NUM_X, CAT_B), hgb()), NUM_X + CAT_B + ["text"], 1, 4)
    dist = {"m__learning_rate": stats.loguniform(.01, .2), "m__max_depth": [2, 3, 4, 6, None], "m__max_leaf_nodes": stats.randint(8, 48),
            "m__min_samples_leaf": stats.randint(5, 40), "m__l2_regularization": stats.loguniform(1e-3, 10)}
    rs = RandomizedSearchCV(P(ct(NUM_X, CAT_B), HistGradientBoostingRegressor(max_iter=400, early_stopping=True, n_iter_no_change=20,
                            validation_fraction=.15, random_state=SEED)), dist, n_iter=NITER, cv=KFold(CV_INT, shuffle=True, random_state=SEED),
                            scoring="neg_root_mean_squared_error", random_state=SEED, n_jobs=-1)
    add("R12", "HGB afinado (búsqueda aleatoria) + nuevas", "I4", "RandomizedSearchCV", rs, NUM_X + CAT_B, NITER, 4)
    if H.get("lightgbm"):
        from lightgbm import LGBMRegressor
        add("R13", "LightGBM + nuevas", "I4", "LightGBM", P(ct(NUM_X, CAT_B), LGBMRegressor(
            n_estimators=400, learning_rate=.03, num_leaves=15, subsample=.8, subsample_freq=1, colsample_bytree=.8,
            random_state=SEED, verbose=-1, n_jobs=1)), NUM_X + CAT_B, 1, 4)
    if H.get("xgboost"):
        from xgboost import XGBRegressor
        add("R14", "XGBoost + nuevas", "I4", "XGBoost", P(ct(NUM_X, CAT_B), XGBRegressor(
            n_estimators=400, learning_rate=.03, max_depth=4, subsample=.8, colsample_bytree=.8, random_state=SEED, n_jobs=1, verbosity=0)),
            NUM_X + CAT_B, 1, 4)
    bases = [P(ct(NUM_X, CAT_B), hgb()), P(ct(NUM_X, CAT_B), RidgeCV(alphas=AL)), P(ct(NUM_X, CAT_B, True), RidgeCV(alphas=np.logspace(-1, 3.5, 12)))]
    cols = NUM_X + CAT_B + ["text"]
    add("R15", "Promedio (HGB + Ridge + Híbrido)", "I5", "Promedio de modelos", Apilado(bases, None, "reg"), cols, 3, 5)
    add("R16", "Apilado (HGB + Ridge + Híbrido → regresión no negativa)", "I5", "Apilamiento propio",
        Apilado(bases, LinearRegression(positive=True), "reg", CV_INT), cols, 3, 5)
    return R


def registro_clf(H):
    base = modelos_clf(); R = []

    def add(id_, nombre, it, herr, est, cols, cfg=1, compl=1):
        R.append(dict(id=id_, nombre=nombre, it=it, herr=herr, est=est, cols=cols, cfg=cfg, compl=compl))
    lr = lambda: LogisticRegression(max_iter=2000, C=.5)
    e, n, c, k = base["Referencia (clase mayoritaria)"]; add("C0", "Referencia (clase mayoritaria)", "I0", "DummyClassifier", e, n + c, k, 0)
    e, n, c, k = base["Regresión logística – B"]; add("C1", "Reg. logística – B (E-II)", "I0", "LogisticRegression", e, n + c, k, 1)
    e, n, c, k = base["Grad. boosting – B"]; add("C2", "Grad. boosting – B (E-II, malla)", "I0", "GridSearchCV", e, n + c, k, 3)
    add("C3", "Reg. logística – B sin likes de comentarios", "I1", "Ablación de variables", P(ct(NUM_NCL, CAT_B), lr()), NUM_NCL + CAT_B, 1, 1)
    add("C4", "HGB fijo – B + variables nuevas", "I3", "Ingeniería de variables", P(ct(NUM_X, CAT_B), hgbc()), NUM_X + CAT_B, 1, 3)
    dist = {"m__learning_rate": stats.loguniform(.01, .2), "m__max_depth": [2, 3, 4, 6, None], "m__max_leaf_nodes": stats.randint(8, 48),
            "m__min_samples_leaf": stats.randint(5, 40), "m__l2_regularization": stats.loguniform(1e-3, 10)}
    rs = RandomizedSearchCV(P(ct(NUM_X, CAT_B), HistGradientBoostingClassifier(max_iter=400, early_stopping=True, n_iter_no_change=20,
                            validation_fraction=.15, random_state=SEED)), dist, n_iter=NITER, cv=KFold(CV_INT, shuffle=True, random_state=SEED),
                            scoring="roc_auc", random_state=SEED, n_jobs=-1)
    add("C5", "HGB afinado (búsqueda aleatoria) + nuevas", "I4", "RandomizedSearchCV", rs, NUM_X + CAT_B, NITER, 4)
    add("C6", "HGB calibrado (isotónica) + nuevas", "I6", "CalibratedClassifierCV",
        CalibratedClassifierCV(P(ct(NUM_X, CAT_B), hgbc()), method="isotonic", cv=CV_INT), NUM_X + CAT_B, 1, 4)
    bases = [P(ct(NUM_X, CAT_B), lr()), P(ct(NUM_X, CAT_B), hgbc()), P(ct(text=True), lr())]
    add("C7", "Apilado (LR + HGB + TF-IDF LR)", "I5", "Apilamiento propio", Apilado(bases, LogisticRegression(max_iter=2000), "clf", CV_INT),
        NUM_X + CAT_B + ["text"], 3, 5)
    return R


# ---------------------------------------------------------------------------------------------- evaluación
def ece(y, p, bins=10):
    idx = np.clip(np.digitize(p, np.linspace(0, 1, bins + 1)) - 1, 0, bins - 1); e = 0.0
    for b in range(bins):
        m = idx == b
        if m.any(): e += m.mean() * abs(y[m].mean() - p[m].mean())
    return float(e)


def cv_clf3(est, X, y):
    cv = RepeatedStratifiedKFold(n_splits=5, n_repeats=A.NREPS, random_state=SEED)
    rows, oofp = [], np.zeros(len(y))
    for i, (tr, te) in enumerate(cv.split(X, y)):
        m = clone(est).fit(X.iloc[tr], y.iloc[tr]); pr = m.predict_proba(X.iloc[te])[:, 1]
        pred = (pr >= .5).astype(int); yt = y.iloc[te].values
        tn, fp, fn, tp = confusion_matrix(yt, pred, labels=[0, 1]).ravel()
        rows.append((accuracy_score(yt, pred), balanced_accuracy_score(yt, pred), f1_score(yt, pred, zero_division=0),
                     roc_auc_score(yt, pr), tp / max(tp + fp, 1), tp / max(tp + fn, 1), tn / max(tn + fp, 1),
                     brier_score_loss(yt, pr), ece(yt, pr)))
        if i < 5: oofp[te] = pr
    return np.array(rows), oofp, (len(tr), len(te))


def elegir(res, key, mayor=False):
    cand = {k: v for k, v in res.items() if "Referencia" not in k}
    mejor = (max if mayor else min)(cand, key=lambda k: cand[k][key])
    K = len(cand[mejor]["folds_" + key]); se = cand[mejor][key + "_sd"] / np.sqrt(K)
    lim = cand[mejor][key] - se if mayor else cand[mejor][key] + se
    ok = [k for k, v in cand.items() if (v[key] >= lim if mayor else v[key] <= lim)]
    elegido = min(ok, key=lambda k: (cand[k]["compl"], (-1 if mayor else 1) * cand[k][key]))
    return dict(mejor=mejor, candidatos=ok, elegido=elegido, umbral=float(lim), se=float(se))


def prueba(res, a, b, key, tam, menor_mejor=True):
    """b frente a a: delta = métrica(b) - métrica(a)."""
    fa, fb = np.array(res[a]["folds_" + key]), np.array(res[b]["folds_" + key])
    t, p = corr_t(fb - fa, *tam)
    return dict(a=a, b=b, delta=float((fb - fa).mean()), t=t, p=p)


def imp_perm(est, D, cols, y, grupos, rep=20):
    Xtr, Xte, ytr, yte = train_test_split(D[cols], y, test_size=.2, random_state=SEED)
    m = clone(est).fit(Xtr, ytr); base = rmse(yte, m.predict(Xte)); rng = np.random.default_rng(SEED); out = []
    for g, cs in grupos.items():
        cs = [c for c in cs if c in cols]
        if not cs: continue
        d = []
        for _ in range(rep):
            X2 = Xte.copy(); idx = rng.permutation(len(X2))
            for c in cs: X2[c] = Xte[c].values[idx]
            d.append(rmse(yte, m.predict(X2)) - base)
        out.append(dict(grupo=g, etiqueta=ETIQ.get(g, g), vars=cs, delta=float(np.mean(d)), sd=float(np.std(d))))
    return sorted(out, key=lambda r: -r["delta"]), base


def nombres_sel(est, D, cols, y):
    """Variables que conserva un selector (Lasso / información mutua) ajustado con todos los datos."""
    m = clone(est).fit(D[cols], y)
    nom = m.named_steps["prep"].get_feature_names_out(); mask = m.named_steps["sel"].get_support()
    sel = [str(n).split("__", 1)[-1] for n, k in zip(nom, mask) if k]
    return dict(n_entrada=int(len(nom)), n_sel=int(mask.sum()), sel=sel)


# ---------------------------------------------------------------------------------------------- principal
def main():
    global NITER, CV_INT
    ap = argparse.ArgumentParser(); ap.add_argument("--data", default="data"); ap.add_argument("--out", default="salida")
    ap.add_argument("--rapido", action="store_true"); a = ap.parse_args(); out = Path(a.out); out.mkdir(exist_ok=True, parents=True)
    A.NREPS = NREPS
    if a.rapido: A.NREPS, A.NTREE, A.MAXIT, NITER, CV_INT = 1, 40, 60, 4, 2
    H = disponibles(); print("Entorno:", H)
    v, c, vp, cp = cargar(a.data); v, c = limpiar(v, c); c, D = construir(v, c); D = extras(c, D)
    y = D["log_views"]; yc = D["alta"]
    R3 = dict(entorno=H, config=dict(repeticiones=A.NREPS, iter_busqueda=NITER, cv_interna=CV_INT, semilla=SEED, n_videos=int(len(D))))

    # ---- regresión
    reg, oofs, tam = {}, {}, None
    for r in registro_reg(H):
        f, oof, tam = cv_reg(r["est"], D[r["cols"]], y, nrep=A.NREPS)
        reg[r["nombre"]] = dict(id=r["id"], it=r["it"], herr=r["herr"], rmse=float(f[:, 0].mean()), rmse_sd=float(f[:, 0].std()),
                                mae=float(f[:, 1].mean()), r2=float(f[:, 2].mean()), r2_sd=float(f[:, 2].std()), rho=float(f[:, 3].mean()),
                                configs=r["cfg"], compl=r["compl"], n_vars=len(r["cols"]), cols=r["cols"], folds_rmse=f[:, 0].tolist())
        oofs[r["nombre"]] = oof
        print(f"  [{r['id']}] {r['nombre']}: RMSE={reg[r['nombre']]['rmse']:.3f} R2={reg[r['nombre']]['r2']:.3f}")
    por_id = {v_["id"]: k for k, v_ in reg.items()}
    R3["reg"] = dict(modelos=reg, n_modelos=len(reg), n_configs=int(sum(v_["configs"] for v_ in reg.values())))
    sel = elegir(reg, "rmse"); R3["seleccion"] = sel
    # comparaciones pareadas = evidencia de cada iteración
    pares = [("R4", "R5", "I1", "Quitar likes de comentarios (HGB)"), ("R1", "R6", "I1", "Quitar likes de comentarios (Ridge)"),
             ("R4", "R7", "I2", "Selección con Lasso"), ("R4", "R8", "I2", "Selección por información mutua"),
             ("R4", "R9", "I3", "Variables nuevas (HGB)"), ("R1", "R10", "I3", "Variables nuevas (Ridge)"),
             ("R9", "R11", "I3", "Añadir LSA al texto"), ("R9", "R12", "I4", "Afinamiento por búsqueda aleatoria"),
             ("R9", "R13", "I4", "LightGBM frente a HGB"), ("R9", "R14", "I4", "XGBoost frente a HGB"),
             ("R9", "R15", "I5", "Promedio frente a HGB individual"), ("R9", "R16", "I5", "Apilado frente a HGB individual"),
             ("R2", "R9", "I3", "HGB nuevo frente al de la etapa II")]
    P_ = []
    for ia, ib, it, desc in pares:
        if ia in por_id and ib in por_id:
            d = prueba(reg, por_id[ia], por_id[ib], "rmse", tam); d.update(it=it, desc=desc, id_a=ia, id_b=ib); P_.append(d)
    ref = por_id["R0"]; base_e2 = min((por_id[i] for i in ("R1", "R2", "R3")), key=lambda k: reg[k]["rmse"])
    ind = [k for k, v_ in reg.items() if v_["it"] in ("I0", "I1", "I2", "I3", "I4") and "Referencia" not in k]
    mej_ind = min(ind, key=lambda k: reg[k]["rmse"]); ens = [k for k, v_ in reg.items() if v_["it"] == "I5"]
    mej_ens = min(ens, key=lambda k: reg[k]["rmse"])
    R3["pares"] = P_
    R3["pruebas"] = dict(elegido_vs_ref=prueba(reg, ref, sel["elegido"], "rmse", tam),
                         elegido_vs_base=prueba(reg, base_e2, sel["elegido"], "rmse", tam), base_e2=base_e2,
                         mejor_individual=mej_ind, mejor_ensamble=mej_ens,
                         ensamble_vs_individual=prueba(reg, mej_ind, mej_ens, "rmse", tam))
    eleg = sel["elegido"]; est = next(r["est"] for r in registro_reg(H) if r["nombre"] == eleg); cols = reg[eleg]["cols"]
    # ---- selección de variables
    sv = {}
    for k, i in (("lasso", "R7"), ("mi", "R8")):
        est_s = next(r["est"] for r in registro_reg(H) if r["id"] == i)
        sv[k] = nombres_sel(est_s, D, NUM_B + CAT_B, y)
    R3["seleccion_vars"] = sv
    # ---- importancia (grupos y variables) del modelo elegido
    g, base_rmse = imp_perm(est, D, cols, y, GRUPOS)
    iv, _ = imp_perm(est, D, cols, y, {c_: [c_] for c_ in cols if c_ != "text"} | ({"text": ["text"]} if "text" in cols else {}))
    R3["importancia_grupos"] = g; R3["importancia_vars"] = iv[:15]; R3["rmse_test_importancia"] = base_rmse
    # ---- errores por categoría del elegido
    o = oofs[eleg]; res = y.values - o
    R3["por_keyword"] = [dict(keyword=k, n=int(len(gk)), rmse=float(np.sqrt(np.mean(res[D.index.get_indexer(gk.index)] ** 2))))
                         for k, gk in D.groupby("Keyword") if len(gk) >= 15]
    R3["heterocedasticidad_rho"] = float(stats.spearmanr(o, np.abs(res))[0])
    # ---- intervalos de predicción (regresión cuantílica; Koenker & Bassett, 1978)
    lo, hi = np.zeros(len(y)), np.zeros(len(y)); cov = []; anch = []
    for tr, te in RepeatedKFold(n_splits=5, n_repeats=1, random_state=SEED).split(D):
        q = {}
        for qq in (.1, .9):
            m = P(ct(NUM_X, CAT_B), hgb(loss="quantile", quantile=qq)).fit(D.iloc[tr][NUM_X + CAT_B], y.iloc[tr]); q[qq] = m.predict(D.iloc[te][NUM_X + CAT_B])
        l, h = np.minimum(q[.1], q[.9]), np.maximum(q[.1], q[.9]); lo[te], hi[te] = l, h
        cov.append(float(np.mean((y.iloc[te] >= l) & (y.iloc[te] <= h)))); anch.append(float(np.mean(h - l)))
    inside = (y.values >= lo) & (y.values <= hi)
    R3["intervalos"] = dict(nominal=.8, cobertura=float(np.mean(cov)), cobertura_sd=float(np.std(cov)), ancho=float(np.mean(anch)),
                            por_keyword=[dict(keyword=k, n=int(len(gk)), cobertura=float(inside[D.index.get_indexer(gk.index)].mean()))
                                         for k, gk in D.groupby("Keyword") if len(gk) >= 15])
    # ---- clasificación
    clf, oofc = {}, {}
    for r in registro_clf(H):
        f, oofp, tamc = cv_clf3(r["est"], D[r["cols"]], yc)
        clf[r["nombre"]] = dict(id=r["id"], it=r["it"], herr=r["herr"], acc=float(f[:, 0].mean()), bacc=float(f[:, 1].mean()), f1=float(f[:, 2].mean()),
                                auc=float(f[:, 3].mean()), auc_sd=float(f[:, 3].std()), prec=float(f[:, 4].mean()), sens=float(f[:, 5].mean()),
                                espec=float(f[:, 6].mean()), brier=float(f[:, 7].mean()), ece=float(f[:, 8].mean()), configs=r["cfg"], compl=r["compl"],
                                cols=r["cols"], folds_auc=f[:, 3].tolist())
        oofc[r["nombre"]] = oofp
        print(f"  [{r['id']}] {r['nombre']}: AUC={clf[r['nombre']]['auc']:.3f} Brier={clf[r['nombre']]['brier']:.3f}")
    selc = elegir(clf, "auc", mayor=True); ec = selc["elegido"]
    ids = {v_["id"]: k for k, v_ in clf.items()}
    cm = confusion_matrix(yc, (oofc[ec] >= .5).astype(int), labels=[0, 1])
    bins = np.linspace(0, 1, 11); idx = np.clip(np.digitize(oofc[ec], bins) - 1, 0, 9)
    cal = lambda p: [dict(pred=float(p[idx == b].mean()), obs=float(yc.values[idx == b].mean()), n=int((idx == b).sum())) for b in range(10) if (idx == b).any()]
    base_c = ids["C2"]; idxb = np.clip(np.digitize(oofc[base_c], bins) - 1, 0, 9)
    calb = [dict(pred=float(oofc[base_c][idxb == b].mean()), obs=float(yc.values[idxb == b].mean()), n=int((idxb == b).sum())) for b in range(10) if (idxb == b).any()]
    R3["clf"] = dict(modelos=clf, seleccion=selc, cm=cm.tolist(), calibracion=cal(oofc[ec]), calibracion_base=calb, base=base_c,
                     n_modelos=len(clf), n_configs=int(sum(v_["configs"] for v_ in clf.values())),
                     prueba_vs_base=prueba(clf, base_c, ec, "auc", tamc), prueba_vs_ref=prueba(clf, ids["C0"], ec, "auc", tamc))
    R3["clf"]["prueba_vs_base"]["delta_auc"] = R3["clf"]["prueba_vs_base"].pop("delta")
    R3["clf"]["prueba_vs_ref"]["delta_auc"] = R3["clf"]["prueba_vs_ref"].pop("delta")

    # ---- modelo final (regresión elegida)
    final = clone(est).fit(D[cols], y); desc = dict(modelo=eleg, tipo=type(final).__name__)
    if hasattr(final, "best_params_"): desc["hiperparametros"] = {k: (float(v_) if isinstance(v_, (float, np.floating)) else (int(v_) if isinstance(v_, (int, np.integer)) else str(v_))) for k, v_ in final.best_params_.items()}
    if isinstance(final, Apilado) and hasattr(final, "meta_") and hasattr(final.meta_, "coef_"):
        desc["pesos_apilado"] = [float(w) for w in final.meta_.coef_]; desc["intercepto"] = float(final.meta_.intercept_)
    R3["final"] = desc
    joblib.dump(final, out / "modelo_final_iii.joblib")
    json.dump(dict(modelo=eleg, entrada=cols, objetivo="log1p(Views)", rmse_cv=reg[eleg]["rmse"], r2_cv=reg[eleg]["r2"], fecha=str(dt.date.today()),
                   semilla=SEED, entorno=H, detalle=desc), open(out / "modelo_iii_metadata.json", "w", encoding="utf-8"), indent=2, ensure_ascii=False)

    # ---- figuras
    col_it = {"I0": C3, "I1": C2, "I2": "#8C564B", "I3": C1, "I4": "#2CA02C", "I5": "#9467BD", "I6": "#17BECF"}
    ordn = sorted(reg, key=lambda k: int(reg[k]["id"][1:]))
    fig, ax = plt.subplots(figsize=(6.4, 4.4))
    ax.barh(range(len(ordn)), [reg[k]["rmse"] for k in ordn], xerr=[reg[k]["rmse_sd"] for k in ordn], color=[col_it[reg[k]["it"]] for k in ordn], ecolor="k", capsize=2)
    ax.set_yticks(range(len(ordn))); ax.set_yticklabels([f"{reg[k]['id']} {k}"[:52] for k in ordn], fontsize=6.5); ax.invert_yaxis()
    ax.set_xlabel("RMSE en log(Views) (media ± DE entre particiones)"); ax.set_xlim(left=max(0, min(reg[k]["rmse"] for k in ordn) * .85))
    ax.axvline(sel["umbral"], color="k", ls="--", lw=.8)
    ax.legend(handles=[plt.Rectangle((0, 0), 1, 1, color=c_) for c_ in col_it.values()], labels=list(col_it), fontsize=6.5, frameon=False, loc="lower right")
    savefig("fig3_iteraciones.png", out)
    top = sorted((k for k in reg), key=lambda k: reg[k]["rmse"])[:7]
    if ref not in top: top.append(ref)
    fig, ax = plt.subplots(figsize=(6.2, 3.2))
    ax.boxplot([reg[k]["folds_rmse"] for k in top[::-1]], vert=False, patch_artist=True, boxprops=dict(facecolor="#DCE6F1"), medianprops=dict(color=C2))
    ax.set_yticks(range(1, len(top) + 1)); ax.set_yticklabels([k[:46] for k in top[::-1]], fontsize=7); ax.set_xlabel("RMSE por partición")
    savefig("fig3_boxplot.png", out)
    fig, ax = plt.subplots(figsize=(5.2, 2.8)); gg = g[::-1]
    ax.barh([r["etiqueta"] for r in gg], [r["delta"] for r in gg], xerr=[r["sd"] for r in gg], color=C1, ecolor=C3)
    ax.set_xlabel("Aumento del RMSE al permutar el grupo de variables"); savefig("fig3_grupos.png", out)
    fig, ax = plt.subplots(figsize=(3.8, 3.4))
    ax.plot([0, 1], [0, 1], color="k", lw=.8, ls="--")
    for datos, lab, col_ in ((R3["clf"]["calibracion_base"], "Base (etapa II)", C3), (R3["clf"]["calibracion"], "Elegido", C2)):
        ax.plot([d_["pred"] for d_ in datos], [d_["obs"] for d_ in datos], marker="o", ms=3, label=lab, color=col_)
    ax.set_xlabel("Probabilidad predicha"); ax.set_ylabel("Frecuencia observada"); ax.legend(frameon=False); savefig("fig3_calibracion.png", out)
    fig, ax = plt.subplots(figsize=(3.8, 3.4)); ax.scatter(y, o, s=6, alpha=.45, color=C1)
    lim = [min(y.min(), np.nanmin(o)), max(y.max(), np.nanmax(o))]; ax.plot(lim, lim, color=C2, lw=1)
    ax.set_xlabel("log(Views) real"); ax.set_ylabel("log(Views) predicho (fuera de muestra)"); savefig("fig3_real_vs_predicho.png", out)
    ik = R3["intervalos"]["por_keyword"]
    if ik:
        fig, ax = plt.subplots(figsize=(5.0, 2.8)); ax.bar([d_["keyword"] for d_ in ik], [d_["cobertura"] for d_ in ik], color=C1)
        ax.axhline(.8, color=C2, ls="--"); ax.set_ylabel("Cobertura empírica"); plt.setp(ax.get_xticklabels(), rotation=45, ha="right", fontsize=7)
        savefig("fig3_intervalos.png", out)
    json.dump(R3, open(out / "results_iii.json", "w", encoding="utf-8"), indent=1, ensure_ascii=False,
              default=lambda x: float(x) if isinstance(x, (np.floating, np.integer)) else str(x))
    print("Listo →", out / "results_iii.json")


if __name__ == "__main__":
    main()
