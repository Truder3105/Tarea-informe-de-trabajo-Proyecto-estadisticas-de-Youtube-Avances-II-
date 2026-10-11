#!/usr/bin/env python3
"""
Proyecto Estadísticas de YouTube - Pipeline de análisis (Etapas I y II)

Uso:
    python analisis.py --data data --out salida

Espera en --data los archivos de Kaggle (advaypatil/youtube-statistics):
    videos-stats.csv  (Title, Video ID, Published At, Keyword, Likes, Comments, Views)
    comments.csv      (Video ID, Comment, Likes, Sentiment)

Genera en --out: results.json, figuras fig_*.png y modelo_final.joblib.
Todas las cifras de los informes provienen de results.json (nada se escribe a mano).
"""
import argparse, hashlib, json, re, warnings, datetime as dt
from pathlib import Path
import numpy as np
import pandas as pd
from scipy import stats
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import joblib
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from sklearn.base import clone
from sklearn.compose import ColumnTransformer
from sklearn.dummy import DummyRegressor, DummyClassifier
from sklearn.ensemble import (RandomForestRegressor, RandomForestClassifier,
                              HistGradientBoostingRegressor, HistGradientBoostingClassifier)
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.impute import SimpleImputer
from sklearn.inspection import permutation_importance
from sklearn.linear_model import RidgeCV, LogisticRegression
from sklearn.metrics import (mean_squared_error, mean_absolute_error, r2_score, accuracy_score,
                             balanced_accuracy_score, f1_score, roc_auc_score, confusion_matrix,
                             cohen_kappa_score)
from sklearn.model_selection import (RepeatedKFold, RepeatedStratifiedKFold, GridSearchCV,
                                     KFold, StratifiedKFold, train_test_split)
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler

warnings.filterwarnings("ignore")
SEED = 42
NTREE, MAXIT, NREPS = 150, 250, 3   # --rapido los reduce (solo para pruebas)
RNG = np.random.default_rng(SEED)
plt.rcParams.update({"font.family": "serif", "font.size": 9, "axes.spines.top": False,
                     "axes.spines.right": False, "figure.dpi": 100})
C1, C2, C3 = "#2F5D8A", "#C9772B", "#7A7A7A"
VADER = SentimentIntensityAnalyzer()


# ----------------------------------------------------------------------------- utilidades
def leer(path):
    for enc in ("utf-8", "utf-8-sig", "latin-1"):
        try:
            return pd.read_csv(path, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise RuntimeError(f"No se pudo leer {path}")


def buscar(d, patron):
    for p in sorted(Path(d).glob("*.csv")):
        if re.search(patron, p.name, re.I):
            return p
    raise FileNotFoundError(f"No hay CSV que coincida con '{patron}' en {d}")


def bh(p):
    p = np.asarray(p, float); n = len(p); o = np.argsort(p)
    q = p[o] * n / np.arange(1, n + 1)
    q = np.minimum.accumulate(q[::-1])[::-1]
    r = np.empty(n); r[o] = np.minimum(q, 1)
    return r


def boot_rho(x, y, B=600):
    n = len(x); rs = []
    for _ in range(B):
        i = RNG.integers(0, n, n)
        rs.append(stats.spearmanr(x[i], y[i])[0])
    return [float(v) for v in np.nanpercentile(rs, [2.5, 97.5])]


def partial_spearman(x, y, z):
    rx, ry, rz = (stats.rankdata(a) for a in (x, y, z))
    ex = rx - np.polyval(np.polyfit(rz, rx, 1), rz)
    ey = ry - np.polyval(np.polyfit(rz, ry, 1), rz)
    r, p = stats.pearsonr(ex, ey)
    return float(r), float(p)


def savefig(name, out):
    plt.tight_layout(); plt.savefig(out / name, dpi=200); plt.close()


# ----------------------------------------------------------------------------- carga y calidad
def cargar(d):
    vp, cp = buscar(d, r"video.*stat"), buscar(d, r"comment")
    v, c = leer(vp), leer(cp)
    for df in (v, c):
        df.drop(columns=[x for x in df.columns if str(x).startswith("Unnamed")], inplace=True)
    return v, c, vp, cp


def calidad(v, c, vp, cp):
    Q = {}
    Q["archivos"] = {p.name: {"bytes": p.stat().st_size,
                              "sha256": hashlib.sha256(p.read_bytes()).hexdigest()} for p in (vp, cp)}
    Q["v_cols"] = list(v.columns); Q["c_cols"] = list(c.columns)
    Q["v_dtypes"] = {k: str(t) for k, t in v.dtypes.items()}
    Q["c_dtypes"] = {k: str(t) for k, t in c.dtypes.items()}
    for col in ("Likes", "Comments", "Views"):
        v[col] = pd.to_numeric(v[col], errors="coerce")
    c["Likes"] = pd.to_numeric(c["Likes"], errors="coerce")
    Q["v_n"], Q["c_n"] = int(len(v)), int(len(c))
    Q["v_missing"] = {k: int(x) for k, x in v.isna().sum().items()}
    Q["c_missing"] = {k: int(x) for k, x in c.isna().sum().items()}
    Q["v_dup_id"] = int(v["Video ID"].duplicated().sum())
    Q["c_dup_rows"] = int(c.duplicated(["Video ID", "Comment"]).sum())
    Q["neg_values"] = {k: int((v[k] < 0).sum()) for k in ("Likes", "Comments", "Views")}
    Q["likes_gt_views"] = int((v["Likes"] > v["Views"]).sum())
    Q["comments_gt_views"] = int((v["Comments"] > v["Views"]).sum())
    Q["orphan_comments"] = int((~c["Video ID"].isin(v["Video ID"])).sum())
    Q["videos_sin_comentarios"] = int((~v["Video ID"].isin(c["Video ID"])).sum())
    fechas = pd.to_datetime(v["Published At"], errors="coerce", utc=True)
    Q["fecha_invalida"] = int(fechas.isna().sum())
    Q["fecha_min"], Q["fecha_max"] = str(fechas.min().date()), str(fechas.max().date())
    Q["keywords"] = {k: int(x) for k, x in v["Keyword"].value_counts().items()}
    Q["sentiment_valores"] = {str(k): int(x) for k, x in c["Sentiment"].value_counts().sort_index().items()}
    cm = c["Comment"].fillna("").astype(str)   # pandas 3: astype(str) conserva NaN
    Q["c_vacios"] = int((c["Comment"].notna() & (cm.str.strip() == "")).sum())
    Q["c_len_mediana"] = float(cm.str.len().median()); Q["c_len_p95"] = float(cm.str.len().quantile(.95))
    Q["c_url"] = float(cm.str.contains(r"http|www\.", case=False).mean())
    Q["c_emoji"] = float(cm.str.contains(r"[\U0001F300-\U0001FAFF\u2600-\u27BF]").mean())
    Q["c_no_ascii"] = float((cm.apply(lambda s: sum(ord(ch) > 127 for ch in s) / max(len(s), 1)) > .3).mean())
    npv = c.groupby("Video ID").size()
    Q["com_por_video"] = {k: float(x) for k, x in npv.describe().items()}
    Q["cob_1"] = float(v["Video ID"].isin(npv.index).mean())
    Q["cob_5"] = float(v["Video ID"].map(npv).fillna(0).ge(5).mean())
    Q["cob_10"] = float(v["Video ID"].map(npv).fillna(0).ge(10).mean())
    tot = v.drop_duplicates("Video ID").set_index("Video ID")["Comments"]
    ratio = (npv / tot.reindex(npv.index)).replace([np.inf], np.nan).dropna()
    Q["frac_comentarios_observados_mediana"] = float(ratio.median()) if len(ratio) else None
    lv = np.log1p(v["Views"].dropna().clip(lower=0))
    q1, q3 = lv.quantile([.25, .75]); iqr = q3 - q1
    Q["outliers_iqr_logviews"] = int(((lv < q1 - 1.5 * iqr) | (lv > q3 + 1.5 * iqr)).sum())
    Q["describe"] = {k: {a: float(b) for a, b in v[k].describe().items()} for k in ("Views", "Likes", "Comments")}
    return Q


def limpiar(v, c):
    v = v.drop_duplicates("Video ID").copy()
    v = v[(v["Views"].notna()) & (v["Views"] > 0)]
    v["Likes"] = v["Likes"].where(v["Likes"] >= 0)
    v["Comments"] = v["Comments"].where(v["Comments"] >= 0)
    v["Published At"] = pd.to_datetime(v["Published At"], errors="coerce", utc=True)
    c = c.dropna(subset=["Comment", "Video ID"]).copy()
    c["Comment"] = c["Comment"].astype(str)
    c = c[c["Comment"].str.strip() != ""].drop_duplicates(["Video ID", "Comment"])
    c["Likes"] = c["Likes"].fillna(0).clip(lower=0)
    c = c[c["Sentiment"].notna()]
    c["Sentiment"] = c["Sentiment"].astype(int)
    c = c[c["Video ID"].isin(v["Video ID"])]
    return v, c


# ----------------------------------------------------------------------------- ingeniería de variables
def vader_c(s): return VADER.polarity_scores(s)["compound"]


def construir(v, c):
    c = c.copy()
    c["len"] = c["Comment"].str.len(); c["words"] = c["Comment"].str.split().str.len()
    c["vader"] = c["Comment"].apply(vader_c)
    c["q"] = c["Comment"].str.contains(r"\?").astype(float)
    c["ex"] = c["Comment"].str.contains("!").astype(float)
    c["up"] = c["Comment"].apply(lambda s: sum(ch.isupper() for ch in s) / max(sum(ch.isalpha() for ch in s), 1))
    c["pos"] = (c["Sentiment"] == 2).astype(float); c["neg"] = (c["Sentiment"] == 0).astype(float)
    c["neu"] = (c["Sentiment"] == 1).astype(float)
    w = c["Likes"] + 1
    c["sw"] = c["Sentiment"] * w
    g = c.groupby("Video ID")
    A = g.agg(n_com=("Comment", "size"), sent_mean=("Sentiment", "mean"), sent_std=("Sentiment", "std"),
              pos_share=("pos", "mean"), neg_share=("neg", "mean"), neu_share=("neu", "mean"),
              clike_mean=("Likes", "mean"), clike_max=("Likes", "max"), len_mean=("len", "mean"),
              words_mean=("words", "mean"), vader_mean=("vader", "mean"), vader_std=("vader", "std"),
              q_share=("q", "mean"), excl_share=("ex", "mean"), upper_share=("up", "mean"))
    A["sent_w"] = g["sw"].sum() / g.apply(lambda d: (d["Likes"] + 1).sum())
    A["text"] = g["Comment"].apply(" ".join)
    for col in ("n_com", "clike_mean", "clike_max", "len_mean", "words_mean"):
        A[col] = np.log1p(A[col])
    D = v.set_index("Video ID").join(A, how="inner")
    ref = D["Published At"].max()
    D["age_days"] = np.log1p((ref - D["Published At"]).dt.days.clip(lower=0))
    D["title_len"] = D["Title"].fillna("").astype(str).str.len()
    D["title_vader"] = D["Title"].fillna("").astype(str).apply(vader_c)
    D["title_q"] = D["Title"].fillna("").astype(str).str.contains(r"\?").astype(float)
    D["log_views"] = np.log1p(D["Views"]); D["log_likes"] = np.log1p(D["Likes"])
    D["like_rate"] = D["Likes"] / D["Views"]
    D["alta"] = (D["log_views"] > D["log_views"].median()).astype(int)
    for col in ("sent_std", "vader_std"):
        D[col] = D[col].fillna(0)
    return c, D


NUM_A = ["n_com", "sent_mean", "sent_std", "pos_share", "neg_share", "neu_share", "sent_w", "clike_mean",
         "clike_max", "len_mean", "words_mean", "vader_mean", "vader_std", "q_share", "excl_share", "upper_share"]
NUM_B = NUM_A + ["age_days", "title_len", "title_vader", "title_q"]
CAT_B = ["Keyword"]
ETIQ = {"n_com": "N.º de comentarios (log)", "sent_mean": "Sentimiento medio", "sent_std": "Dispersión del sentimiento",
        "pos_share": "% comentarios positivos", "neg_share": "% comentarios negativos", "neu_share": "% neutros",
        "sent_w": "Sentimiento ponderado por likes", "clike_mean": "Likes medios del comentario (log)",
        "clike_max": "Likes máx. del comentario (log)", "len_mean": "Longitud media (log)",
        "words_mean": "Palabras medias (log)", "vader_mean": "VADER compuesto medio", "vader_std": "VADER dispersión",
        "q_share": "% con '?'", "excl_share": "% con '!'", "upper_share": "% mayúsculas", "age_days": "Antigüedad (log días)",
        "title_len": "Longitud del título", "title_vader": "VADER del título", "title_q": "Título con '?'",
        "Keyword": "Categoría (Keyword)", "text": "Texto de comentarios (TF-IDF)"}


def ct(num=None, cat=None, text=False):
    t = []
    if num: t.append(("num", Pipeline([("i", SimpleImputer(strategy="median")), ("s", StandardScaler())]), num))
    if cat: t.append(("cat", OneHotEncoder(handle_unknown="ignore", min_frequency=5), cat))
    if text: t.append(("txt", TfidfVectorizer(ngram_range=(1, 2), min_df=3, max_features=20000,
                                              sublinear_tf=True, stop_words="english"), "text"))
    return ColumnTransformer(t)


def gs(p, grid, scoring):
    return GridSearchCV(p, grid, cv=KFold(3, shuffle=True, random_state=SEED), scoring=scoring, n_jobs=-1)


def modelos_reg():
    P = lambda prep, m: Pipeline([("prep", prep), ("m", m)])
    rf = lambda: RandomForestRegressor(n_estimators=NTREE, random_state=SEED, n_jobs=1)
    hg = lambda: HistGradientBoostingRegressor(max_iter=MAXIT, random_state=SEED)
    alphas = np.logspace(-2, 3, 12)
    rfg = {"m__max_depth": [None, 8], "m__min_samples_leaf": [2, 5]}
    hgg = {"m__learning_rate": [0.03, 0.1], "m__max_depth": [3, None]}
    S = "neg_root_mean_squared_error"
    return {
        "Referencia (mediana)": (P(ct(NUM_A), DummyRegressor(strategy="median")), NUM_A, [], 1),
        "Ridge – A (comentarios)": (P(ct(NUM_A), RidgeCV(alphas=alphas)), NUM_A, [], 12),
        "Bosque aleatorio – A": (gs(P(ct(NUM_A), rf()), rfg, S), NUM_A, [], 4),
        "Grad. boosting – A": (gs(P(ct(NUM_A), hg()), hgg, S), NUM_A, [], 4),
        "Ridge – B (+metadatos)": (P(ct(NUM_B, CAT_B), RidgeCV(alphas=alphas)), NUM_B, CAT_B, 12),
        "Bosque aleatorio – B": (gs(P(ct(NUM_B, CAT_B), rf()), rfg, S), NUM_B, CAT_B, 4),
        "Grad. boosting – B": (gs(P(ct(NUM_B, CAT_B), hg()), hgg, S), NUM_B, CAT_B, 4),
        "TF-IDF + Ridge (solo texto)": (P(ct(text=True), RidgeCV(alphas=alphas)), [], [], 12),
        "Híbrido B + TF-IDF + Ridge": (P(ct(NUM_B, CAT_B, True), RidgeCV(alphas=np.logspace(-1, 3.5, 12))), NUM_B, CAT_B, 12),
    }


def modelos_clf():
    P = lambda prep, m: Pipeline([("prep", prep), ("m", m)])
    rfg = {"m__max_depth": [None, 8], "m__min_samples_leaf": [2, 5]}
    hgg = {"m__learning_rate": [0.03, 0.1], "m__max_depth": [3, None]}
    lr = lambda: LogisticRegression(max_iter=2000, C=0.5)
    return {
        "Referencia (clase mayoritaria)": (P(ct(NUM_A), DummyClassifier(strategy="most_frequent")), NUM_A, [], 1),
        "Regresión logística – A": (P(ct(NUM_A), lr()), NUM_A, [], 1),
        "Regresión logística – B": (P(ct(NUM_B, CAT_B), lr()), NUM_B, CAT_B, 1),
        "Bosque aleatorio – B": (gs(P(ct(NUM_B, CAT_B), RandomForestClassifier(n_estimators=NTREE, random_state=SEED, n_jobs=1)),
                                    rfg, "roc_auc"), NUM_B, CAT_B, 4),
        "Grad. boosting – B": (gs(P(ct(NUM_B, CAT_B), HistGradientBoostingClassifier(max_iter=MAXIT, random_state=SEED)),
                                  hgg, "roc_auc"), NUM_B, CAT_B, 4),
        "TF-IDF + Reg. logística (texto)": (P(ct(text=True), lr()), [], [], 1),
    }


# ----------------------------------------------------------------------------- validación cruzada
def corr_t(d, ntr, nte):
    """Prueba t remuestreada corregida (Nadeau & Bengio, 2003)."""
    d = np.asarray(d, float); k = len(d)
    var = d.var(ddof=1)
    if var == 0: return 0.0, 1.0
    t = d.mean() / np.sqrt((1 / k + nte / ntr) * var)
    return float(t), float(2 * stats.t.sf(abs(t), k - 1))


def cv_reg(est, X, y, nrep=None):
    nrep = min(nrep or NREPS, NREPS)
    cv = RepeatedKFold(n_splits=5, n_repeats=nrep, random_state=SEED)
    rows, oof = [], np.full(len(y), np.nan)
    for i, (tr, te) in enumerate(cv.split(X)):
        m = clone(est).fit(X.iloc[tr], y.iloc[tr]); p = m.predict(X.iloc[te])
        if i < 5: oof[te] = p
        rows.append((np.sqrt(mean_squared_error(y.iloc[te], p)), mean_absolute_error(y.iloc[te], p),
                     r2_score(y.iloc[te], p), stats.spearmanr(y.iloc[te], p)[0] if np.std(p) > 0 else 0.0))
    return np.array(rows), oof, (len(tr), len(te))


def cv_clf(est, X, y, nrep=None):
    nrep = min(nrep or 2, NREPS)
    cv = RepeatedStratifiedKFold(n_splits=5, n_repeats=nrep, random_state=SEED)
    rows, oof, oofp = [], np.zeros(len(y), int), np.zeros(len(y))
    for i, (tr, te) in enumerate(cv.split(X, y)):
        m = clone(est).fit(X.iloc[tr], y.iloc[tr]); yt = y.iloc[te]
        pred = m.predict(X.iloc[te])
        pr = m.predict_proba(X.iloc[te])[:, 1]
        if i < 5: oof[te], oofp[te] = pred, pr
        tn, fp, fn, tp = confusion_matrix(yt, pred, labels=[0, 1]).ravel()
        rows.append((accuracy_score(yt, pred), balanced_accuracy_score(yt, pred), f1_score(yt, pred, zero_division=0),
                     roc_auc_score(yt, pr) if len(set(pr)) > 1 else 0.5,
                     tp / max(tp + fp, 1), tp / max(tp + fn, 1), tn / max(tn + fp, 1)))
    return np.array(rows), oof, (len(tr), len(te))


# ----------------------------------------------------------------------------- análisis estadístico (Etapa I)
def estadistica(c, D, out):
    S = {}
    y = D["log_views"].values
    S["n_videos"] = int(len(D)); S["n_comentarios"] = int(len(c))
    S["skew_views"] = float(stats.skew(D["Views"])); S["skew_logviews"] = float(stats.skew(y))
    S["kurt_logviews"] = float(stats.kurtosis(y))
    try:
        _, lam = stats.boxcox(D["Views"].values + 1); S["boxcox_lambda"] = float(lam)
    except Exception:
        S["boxcox_lambda"] = None
    # 1. Spearman con IC bootstrap y FDR
    rows = []
    for k in NUM_B:
        r, p = stats.spearmanr(D[k], y)
        lo, hi = boot_rho(D[k].values, y)
        pr, pp = partial_spearman(D[k].values, y, D["age_days"].values) if k != "age_days" else (np.nan, np.nan)
        rows.append(dict(var=k, etiqueta=ETIQ[k], rho=float(r), lo=lo, hi=hi,
                         p=float(p), rho_parcial=float(pr), p_parcial=float(pp)))
    q = bh([r["p"] for r in rows])
    for r, qq in zip(rows, q): r["q"] = float(qq)
    S["spearman"] = sorted(rows, key=lambda r: -abs(r["rho"]))
    # 2. Kruskal-Wallis views ~ Keyword
    grp = [g["log_views"].values for _, g in D.groupby("Keyword") if len(g) > 1]
    H, p = stats.kruskal(*grp)
    S["kruskal"] = dict(H=float(H), p=float(p), eps2=float(H / (len(D) - 1)), k=len(grp))
    S["kw_tabla"] = [dict(keyword=k, n=int(len(g)), mediana_views=float(g["Views"].median()),
                          q1=float(g["Views"].quantile(.25)), q3=float(g["Views"].quantile(.75)),
                          sent_medio=float(g["sent_mean"].mean())) for k, g in D.groupby("Keyword")]
    # 3. Chi-cuadrado Keyword x sentimiento dominante, V de Cramér
    dom = D[["pos_share", "neu_share", "neg_share"]].idxmax(axis=1)
    tab = pd.crosstab(D["Keyword"], dom)
    chi2, pc, dof, _ = stats.chi2_contingency(tab)
    S["chi"] = dict(chi2=float(chi2), p=float(pc), dof=int(dof),
                    V=float(np.sqrt(chi2 / (tab.values.sum() * (min(tab.shape) - 1)))))
    # 4. Paradoja de Simpson
    pooled = stats.spearmanr(D["sent_mean"], y)[0]
    within = []
    for k, g in D.groupby("Keyword"):
        if len(g) >= 20: within.append((k, len(g), float(stats.spearmanr(g["sent_mean"], g["log_views"])[0])))
    wavg = float(sum(n * r for _, n, r in within) / max(sum(n for _, n, _ in within), 1))
    S["simpson"] = dict(pooled=float(pooled), within_w=wavg, within=within,
                        invierte=bool(np.sign(pooled) != np.sign(wavg) and abs(wavg) > .02))
    # 5. Sesgo de selección: likes del comentario vs sentimiento (nivel comentario)
    kw = stats.kruskal(*[g["Likes"].values for _, g in c.groupby("Sentiment")])
    S["sesgo_sel"] = dict(H=float(kw[0]), p=float(kw[1]),
                          med={str(k): float(g["Likes"].median()) for k, g in c.groupby("Sentiment")},
                          mean={str(k): float(g["Likes"].mean()) for k, g in c.groupby("Sentiment")})
    # 6. Tasa de likes vs sentimiento
    ok = D["like_rate"].notna()
    r, p = stats.spearmanr(D.loc[ok, "sent_mean"], D.loc[ok, "like_rate"])
    S["like_rate"] = dict(rho=float(r), p=float(p), mediana=float(D.loc[ok, "like_rate"].median()))
    # 7. Concordancia etiqueta Kaggle vs VADER
    pv = np.where(c["vader"] >= .05, 2, np.where(c["vader"] <= -.05, 0, 1))
    cmx = confusion_matrix(c["Sentiment"], pv, labels=[0, 1, 2])
    S["vader"] = dict(kappa=float(cohen_kappa_score(c["Sentiment"], pv)), acc=float(accuracy_score(c["Sentiment"], pv)),
                      cm=cmx.tolist(), n=int(len(c)))
    # 8. Potencia
    n = len(D); S["potencia"] = dict(n=int(n), rho_min=float(np.tanh((1.96 + 0.84) / np.sqrt(n - 3))))
    S["muestras_por_param"] = float(n / (len(NUM_B) + len(D["Keyword"].unique())))
    # figuras de la etapa I
    fig, ax = plt.subplots(1, 3, figsize=(7.2, 2.4))
    for a, col, colr in zip(ax, ["Views", "Likes", "Comments"], [C1, C2, C3]):
        a.hist(np.log10(D[col].dropna() + 1), bins=30, color=colr, edgecolor="white", linewidth=.3)
        a.set_xlabel(f"log10({col}+1)"); a.set_ylabel("Videos")
    savefig("fig_distribuciones.png", out)
    fig, ax = plt.subplots(1, 2, figsize=(7.2, 2.6))
    vc = D["Keyword"].value_counts()
    ax[0].barh(vc.index[::-1][-15:], vc.values[::-1][-15:], color=C1); ax[0].set_xlabel("Videos")
    sc = c["Sentiment"].value_counts().sort_index()
    ax[1].bar([{0: "Negativo", 1: "Neutro", 2: "Positivo"}.get(i, str(i)) for i in sc.index], sc.values, color=[C2, C3, C1])
    ax[1].set_ylabel("Comentarios")
    savefig("fig_keyword_sentimiento.png", out)
    fig, ax = plt.subplots(figsize=(3.6, 2.6))
    ax.hist(np.expm1(D["n_com"]), bins=range(0, int(np.expm1(D["n_com"]).max()) + 2), color=C1, edgecolor="white", linewidth=.3)
    ax.set_xlabel("Comentarios disponibles por video"); ax.set_ylabel("Videos")
    savefig("fig_comentarios_por_video.png", out)
    sp = S["spearman"][:12][::-1]
    fig, ax = plt.subplots(figsize=(5.4, 3.4))
    ax.errorbar([r["rho"] for r in sp], range(len(sp)), xerr=[[r["rho"] - r["lo"] for r in sp], [r["hi"] - r["rho"] for r in sp]],
                fmt="o", color=C1, ecolor=C3, capsize=2)
    ax.set_yticks(range(len(sp))); ax.set_yticklabels([r["etiqueta"] for r in sp]); ax.axvline(0, color="k", lw=.6)
    ax.set_xlabel("ρ de Spearman con log(Views) (IC 95 % bootstrap)")
    savefig("fig_spearman.png", out)
    return S


def suficiencia(D, out):
    """Curva de aprendizaje: ¿mejoran las predicciones al aumentar n?"""
    X, y = D[NUM_B + CAT_B], D["log_views"]
    fr = [0.1, 0.2, 0.4, 0.6, 0.8, 1.0]
    res = {}
    for nombre, est in {"Ridge": Pipeline([("prep", ct(NUM_B, CAT_B)), ("m", RidgeCV(alphas=np.logspace(-2, 3, 12)))]),
                        "Grad. boosting": Pipeline([("prep", ct(NUM_B, CAT_B)),
                                                    ("m", HistGradientBoostingRegressor(max_iter=150, learning_rate=.05, random_state=SEED))])}.items():
        pts = []
        for f in fr:
            sc = []
            for tr, te in KFold(5, shuffle=True, random_state=SEED).split(X):
                sub = np.random.default_rng(SEED).choice(tr, max(int(len(tr) * f), 10), replace=False)
                m = clone(est).fit(X.iloc[sub], y.iloc[sub]); sc.append(r2_score(y.iloc[te], m.predict(X.iloc[te])))
            pts.append((int(len(sub)), float(np.mean(sc)), float(np.std(sc))))
        res[nombre] = pts
    fig, ax = plt.subplots(figsize=(4.4, 2.8))
    for (k, pts), col in zip(res.items(), (C1, C2)):
        ax.errorbar([p[0] for p in pts], [p[1] for p in pts], yerr=[p[2] for p in pts], marker="o", color=col, label=k, capsize=2)
    ax.set_xlabel("Videos usados para entrenar"); ax.set_ylabel("R² (validación)"); ax.legend(frameon=False)
    savefig("fig_curva_aprendizaje.png", out)
    last = res["Grad. boosting"]
    return dict(curvas=res, pendiente_final=float((last[-1][1] - last[-2][1]) / max(last[-1][0] - last[-2][0], 1) * 100))


# ----------------------------------------------------------------------------- modelado (Etapa II)
def modelado(D, out):
    M = {}
    y = D["log_views"]
    # ----- regresión
    reg, oofs, tests, ft = {}, {}, None, None
    for nombre, (est, num, cat, ncfg) in modelos_reg().items():
        cols = num + cat + (["text"] if ("TF-IDF" in nombre or "Híbrido" in nombre) else [])
        X = D[cols]
        folds, oof, (ntr, nte) = cv_reg(est, X, y)
        reg[nombre] = dict(rmse=float(folds[:, 0].mean()), rmse_sd=float(folds[:, 0].std()), mae=float(folds[:, 1].mean()),
                           r2=float(folds[:, 2].mean()), r2_sd=float(folds[:, 2].std()), rho=float(folds[:, 3].mean()),
                           configs=ncfg, n_vars=len(num) + len(cat), folds_rmse=folds[:, 0].tolist(), cols=cols)
        oofs[nombre] = oof; tests = (ntr, nte)
        print(f"  [reg] {nombre}: RMSE={reg[nombre]['rmse']:.3f} R2={reg[nombre]['r2']:.3f}")
    names = list(reg)
    best = min((n for n in names if "Referencia" not in n), key=lambda n: reg[n]["rmse"])
    ranking = sorted((n for n in names if n != "Referencia (mediana)"), key=lambda n: reg[n]["rmse"])
    second = ranking[1] if len(ranking) > 1 else best
    base = reg["Referencia (mediana)"]["folds_rmse"]
    pr = {}
    for nm, other in (("vs_referencia", base), ("vs_segundo", reg[second]["folds_rmse"])):
        d = np.array(other) - np.array(reg[best]["folds_rmse"])
        t, p = corr_t(d, *tests); pr[nm] = dict(t=t, p=p, delta_rmse=float(d.mean()))
    pr["segundo"] = second
    M["reg"] = dict(modelos=reg, mejor=best, pruebas=pr, n_configs=int(sum(r["configs"] for r in reg.values())),
                    n_modelos=len(reg), orden=ranking)
    # ----- por keyword (calidad de predicción por segmento)
    o = oofs[best]; res = y.values - o
    M["reg"]["por_keyword"] = [dict(keyword=k, n=int(len(g)), rmse=float(np.sqrt(np.mean(res[D.index.get_indexer(g.index)] ** 2))))
                               for k, g in D.groupby("Keyword") if len(g) >= 15]
    M["reg"]["heterocedasticidad_rho"] = float(stats.spearmanr(o, np.abs(res))[0])
    # ----- objetivo secundario: log(Likes)
    Dl = D[D["Likes"].notna()]
    sec = {}
    for nombre in ("Referencia (mediana)", "Ridge – B (+metadatos)", "Grad. boosting – B"):
        est, num, cat, _ = modelos_reg()[nombre]
        f, _, _ = cv_reg(est, Dl[num + cat], Dl["log_likes"], nrep=2)
        sec[nombre] = dict(rmse=float(f[:, 0].mean()), r2=float(f[:, 2].mean()), rho=float(f[:, 3].mean()))
    M["likes"] = sec
    # ----- clasificación
    yc, clf, coofs, cp = D["alta"], {}, {}, {}
    for nombre, (est, num, cat, ncfg) in modelos_clf().items():
        cols = num + cat + (["text"] if "TF-IDF" in nombre else [])
        folds, oof, (ntr, nte) = cv_clf(est, D[cols], yc)
        clf[nombre] = dict(acc=float(folds[:, 0].mean()), bacc=float(folds[:, 1].mean()), f1=float(folds[:, 2].mean()),
                           auc=float(folds[:, 3].mean()), auc_sd=float(folds[:, 3].std()), prec=float(folds[:, 4].mean()),
                           sens=float(folds[:, 5].mean()), espec=float(folds[:, 6].mean()), configs=ncfg, cols=cols,
                           folds_auc=folds[:, 3].tolist())
        coofs[nombre] = oof; print(f"  [clf] {nombre}: AUC={clf[nombre]['auc']:.3f} Acc={clf[nombre]['acc']:.3f}")
    bestc = max((n for n in clf if "Referencia" not in n), key=lambda n: clf[n]["auc"])
    cm = confusion_matrix(yc, coofs[bestc], labels=[0, 1])
    M["clf"] = dict(modelos=clf, mejor=bestc, cm=cm.tolist(), n_configs=int(sum(r["configs"] for r in clf.values())))
    d = np.array(clf["Referencia (clase mayoritaria)"]["folds_auc"]) * 0 + np.array(clf[bestc]["folds_auc"]) - 0.5
    t, p = corr_t(d, *tests); M["clf"]["prueba_vs_azar"] = dict(t=t, p=p)
    # ----- importancia por permutación (mejor regresor)
    est = modelos_reg()[best][0]; cols = reg[best]["cols"]
    Xtr, Xte, ytr, yte = train_test_split(D[cols], y, test_size=.2, random_state=SEED)
    m = clone(est).fit(Xtr, ytr)
    pi = permutation_importance(m, Xte, yte, n_repeats=25, random_state=SEED, scoring="neg_root_mean_squared_error")
    imp = sorted([dict(var=c_, etiqueta=ETIQ.get(c_, c_), delta=float(a), sd=float(b))
                  for c_, a, b in zip(cols, pi.importances_mean, pi.importances_std)], key=lambda r: -r["delta"])
    M["importancia"] = imp
    comentarios_only = all(c_ in NUM_A or c_ == "text" for c_ in cols)
    M["mejor_usa_metadatos"] = not comentarios_only
    # ----- figuras
    fig, ax = plt.subplots(figsize=(6.0, 3.2))
    ordn = sorted(reg, key=lambda n: reg[n]["rmse"])
    ax.boxplot([reg[n]["folds_rmse"] for n in ordn], vert=False, patch_artist=True,
               boxprops=dict(facecolor="#DCE6F1"), medianprops=dict(color=C2))
    ax.set_yticks(range(1, len(ordn) + 1)); ax.set_yticklabels(ordn)   # compatible con matplotlib 3.x
    ax.set_xlabel("RMSE en log(Views) por partición (menor es mejor)")
    savefig("fig_modelos_rmse.png", out)
    fig, ax = plt.subplots(figsize=(3.8, 3.4))
    ax.scatter(y, oofs[best], s=6, alpha=.45, color=C1)
    lim = [min(y.min(), np.nanmin(oofs[best])), max(y.max(), np.nanmax(oofs[best]))]
    ax.plot(lim, lim, color=C2, lw=1); ax.set_xlabel("log(Views) real"); ax.set_ylabel("log(Views) predicho (fuera de muestra)")
    savefig("fig_real_vs_predicho.png", out)
    t10 = imp[:10][::-1]
    fig, ax = plt.subplots(figsize=(5.0, 3.0))
    ax.barh([r["etiqueta"] for r in t10], [r["delta"] for r in t10], xerr=[r["sd"] for r in t10], color=C1, ecolor=C3)
    ax.set_xlabel("Aumento del RMSE al permutar la variable")
    savefig("fig_importancia.png", out)
    fig, ax = plt.subplots(figsize=(3.0, 2.8))
    ax.imshow(cm, cmap="Blues")
    for (i, j), val in np.ndenumerate(cm): ax.text(j, i, int(val), ha="center", va="center")
    ax.set_xticks([0, 1]); ax.set_yticks([0, 1]); ax.set_xticklabels(["Baja", "Alta"]); ax.set_yticklabels(["Baja", "Alta"])
    ax.set_xlabel("Predicción"); ax.set_ylabel("Real")
    savefig("fig_matriz_confusion.png", out)
    # ----- artefacto final (MLOps: modelo versionado + metadatos)
    final = clone(est).fit(D[cols], y)
    joblib.dump(final, out / "modelo_final.joblib")
    json.dump(dict(modelo=best, entrada=cols, objetivo="log1p(Views)", fecha=str(dt.date.today()),
                   rmse_cv=reg[best]["rmse"], r2_cv=reg[best]["r2"], semilla=SEED),
              open(out / "modelo_metadata.json", "w", encoding="utf-8"), indent=2, ensure_ascii=False)
    return M


# ----------------------------------------------------------------------------- principal
def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--data", default="data"); ap.add_argument("--out", default="salida"); ap.add_argument("--rapido", action="store_true")
    a = ap.parse_args(); out = Path(a.out); out.mkdir(exist_ok=True, parents=True)
    global NTREE, MAXIT, NREPS
    if a.rapido: NTREE, MAXIT, NREPS = 40, 60, 1
    v, c, vp, cp = cargar(a.data)
    R = {"calidad": calidad(v, c, vp, cp)}
    v, c = limpiar(v, c)
    c, D = construir(v, c)
    R["limpieza"] = dict(videos_finales=int(len(D)), comentarios_finales=int(len(c)),
                         videos_descartados=R["calidad"]["v_n"] - int(len(D)))
    print("Etapa I: estadística…"); R["stats"] = estadistica(c, D, out)
    print("Etapa I: suficiencia…"); R["suficiencia"] = suficiencia(D, out)
    print("Etapa II: modelado…"); R["modelado"] = modelado(D, out)
    json.dump(R, open(out / "results.json", "w", encoding="utf-8"), indent=1, ensure_ascii=False,
              default=lambda o: float(o) if isinstance(o, (np.floating, np.integer)) else str(o))
    print("Listo →", out / "results.json")


if __name__ == "__main__":
    main()