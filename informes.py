#!/usr/bin/env python3
"""
Genera los informes Avances I y Avances II (APA 7) a partir de Plantilla_APA7.docx y salida/results.json.

Uso:  python informes.py --plantilla Plantilla_APA7.docx --salida salida
Edite el bloque CONFIG con los datos del grupo antes de ejecutar.
"""
import argparse, json, re, copy
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ============================== CONFIG (EDITAR) ==============================
CONFIG = dict(
    autores="Julian Esteban Ballesteros Ortiz, Juan Diego Walteros Cortes",
    facultad="Inteligencia Artificial, Universidad de Cundinamarca",
    curso="CAD2202023206 – Procesamiento de Lenguaje Natural - EIAIIPA2026_2",
    profesor="Docente, Paul Alexander Diaz Montaña",
    fecha_1="[Fecha de entrega Semana 7]",
    fecha_2="domingo, 11 de octubre de 2026",
    fecha_3="domingo, 18 de octubre de 2026",
)
# =============================================================================

REFS = [
    "Bergstra, J., & Bengio, Y. (2012). Random search for hyper-parameter optimization. *Journal of Machine Learning Research, 13*, 281–305.",
    "Brier, G. W. (1950). Verification of forecasts expressed in terms of probability. *Monthly Weather Review, 78*(1), 1–3. https://doi.org/10.1175/1520-0493(1950)078<0001:VOFEIT>2.0.CO;2",
    "Chen, T., & Guestrin, C. (2016). XGBoost: A scalable tree boosting system. En *Proceedings of the 22nd ACM SIGKDD International Conference on Knowledge Discovery and Data Mining* (pp. 785–794). ACM. https://doi.org/10.1145/2939672.2939785",
    "Deerwester, S., Dumais, S. T., Furnas, G. W., Landauer, T. K., & Harshman, R. (1990). Indexing by latent semantic analysis. *Journal of the American Society for Information Science, 41*(6), 391–407.",
    "Guyon, I., & Elisseeff, A. (2003). An introduction to variable and feature selection. *Journal of Machine Learning Research, 3*, 1157–1182.",
    "Hastie, T., Tibshirani, R., & Friedman, J. (2009). *The elements of statistical learning: Data mining, inference, and prediction* (2.ª ed.). Springer. https://doi.org/10.1007/978-0-387-84858-7",
    "Ke, G., Meng, Q., Finley, T., Wang, T., Chen, W., Ma, W., Ye, Q., & Liu, T.-Y. (2017). LightGBM: A highly efficient gradient boosting decision tree. *Advances in Neural Information Processing Systems, 30*, 3146–3154.",
    "Koenker, R., & Bassett, G. (1978). Regression quantiles. *Econometrica, 46*(1), 33–50. https://doi.org/10.2307/1913643",
    "Niculescu-Mizil, A., & Caruana, R. (2005). Predicting good probabilities with supervised learning. En *Proceedings of the 22nd International Conference on Machine Learning* (pp. 625–632). ACM. https://doi.org/10.1145/1102351.1102430",
    "Tibshirani, R. (1996). Regression shrinkage and selection via the lasso. *Journal of the Royal Statistical Society: Series B, 58*(1), 267–288. https://doi.org/10.1111/j.2517-6161.1996.tb02080.x",
    "Wolpert, D. H. (1992). Stacked generalization. *Neural Networks, 5*(2), 241–259. https://doi.org/10.1016/S0893-6080(05)80023-1",
    "Alemán Viteri, S. B. (2021). Análisis de sentimientos para Twitter con Vader y TextBlob. *Revista Odigos, 2*(3), 9–25. https://doi.org/10.35290/ro.v2n3.2021.494",
    "Barrios Arce, J. (2019, 26 de julio). La matriz de confusión y sus métricas. *Health Big Data*. https://www.juanbarrios.com/la-matriz-de-confusion-y-sus-metricas/",
    "Benjamini, Y., & Hochberg, Y. (1995). Controlling the false discovery rate: A practical and powerful approach to multiple testing. *Journal of the Royal Statistical Society: Series B, 57*(1), 289–300. https://doi.org/10.1111/j.2517-6161.1995.tb02031.x",
    "Box, G. E. P., & Cox, D. R. (1964). An analysis of transformations. *Journal of the Royal Statistical Society: Series B, 26*(2), 211–252. https://doi.org/10.1111/j.2517-6161.1964.tb00553.x",
    "Breiman, L. (2001). Random forests. *Machine Learning, 45*(1), 5–32. https://doi.org/10.1023/A:1010933404324",
    "Cohen, J. (1960). A coefficient of agreement for nominal scales. *Educational and Psychological Measurement, 20*(1), 37–46. https://doi.org/10.1177/001316446002000104",
    "Efron, B. (1979). Bootstrap methods: Another look at the jackknife. *The Annals of Statistics, 7*(1), 1–26. https://doi.org/10.1214/aos/1176344552",
    "Friedman, J. H. (2001). Greedy function approximation: A gradient boosting machine. *The Annals of Statistics, 29*(5), 1189–1232. https://doi.org/10.1214/aos/1013203451",
    "Google. (s. f.). *CommentThreads: list* [Documentación de YouTube Data API v3]. https://developers.google.com/youtube/v3/docs/commentThreads/list",
    "Hoerl, A. E., & Kennard, R. W. (1970). Ridge regression: Biased estimation for nonorthogonal problems. *Technometrics, 12*(1), 55–67. https://doi.org/10.1080/00401706.1970.10488634",
    "Hutto, C., & Gilbert, E. (2014). VADER: A parsimonious rule-based model for sentiment analysis of social media text. *Proceedings of the International AAAI Conference on Web and Social Media, 8*(1), 216–225. https://doi.org/10.1609/icwsm.v8i1.14550",
    "Kruskal, W. H., & Wallis, W. A. (1952). Use of ranks in one-criterion variance analysis. *Journal of the American Statistical Association, 47*(260), 583–621. https://doi.org/10.1080/01621459.1952.10483441",
    "Landis, J. R., & Koch, G. G. (1977). The measurement of observer agreement for categorical data. *Biometrics, 33*(1), 159–174. https://doi.org/10.2307/2529310",
    "Nadeau, C., & Bengio, Y. (2003). Inference for the generalization error. *Machine Learning, 52*(3), 239–281. https://doi.org/10.1023/A:1024068626366",
    "Patil, A. (s. f.). *YouTube statistics* [Conjunto de datos]. Kaggle. https://www.kaggle.com/datasets/advaypatil/youtube-statistics",
    "Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., Grisel, O., Blondel, M., Prettenhofer, P., Weiss, R., Dubourg, V., Vanderplas, J., Passos, A., Cournapeau, D., Brucher, M., Perrot, M., & Duchesnay, É. (2011). Scikit-learn: Machine learning in Python. *Journal of Machine Learning Research, 12*, 2825–2830.",
    "Salton, G., & Buckley, C. (1988). Term-weighting approaches in automatic text retrieval. *Information Processing & Management, 24*(5), 513–523. https://doi.org/10.1016/0306-4573(88)90021-0",
    "Sculley, D., Holt, G., Golovin, D., Davydov, E., Phillips, T., Ebner, D., Chaudhary, V., Young, M., Crespo, J.-F., & Dennison, D. (2015). Hidden technical debt in machine learning systems. *Advances in Neural Information Processing Systems, 28*, 2503–2511.",
    "Simpson, E. H. (1951). The interpretation of interaction in contingency tables. *Journal of the Royal Statistical Society: Series B, 13*(2), 238–241. https://doi.org/10.1111/j.2517-6161.1951.tb00088.x",
    "Wang, R. Y., & Strong, D. M. (1996). Beyond accuracy: What data quality means to data consumers. *Journal of Management Information Systems, 12*(4), 5–33. https://doi.org/10.1080/07421222.1996.11518099",
]

# ------------------------------------------------------------------ formato numérico (es)
def n(x, d=2):
    s = f"{x:,.{d}f}"
    return s.replace(",", "\u00a0").replace(".", ",")

def r_(x, d=2):
    """Estadístico acotado en [-1,1] o [0,1]: sin cero inicial (APA 7)."""
    s = f"{abs(x):.{d}f}".lstrip("0").replace(".", ",") if abs(x) < 1 else f"{abs(x):.{d}f}".replace(".", ",")
    return ("−" if x < 0 else "") + (s if s.startswith(",") or abs(x) >= 1 else "0" + s)

def pct(x, d=1): return n(100 * x, d) + " %"

def pv(p):
    return "< ,001" if p < .001 else "= " + f"{p:.3f}".lstrip("0").replace(".", ",")

def ent(x): return n(x, 0)

# ------------------------------------------------------------------ primitivas docx
FONT = "Times New Roman"

def _fmt(run, bold=False, italic=False, size=12):
    run.font.name = FONT; run.font.size = Pt(size); run.bold = bold; run.italic = italic
    rpr = run._r.get_or_add_rPr(); rf = rpr.find(qn("w:rFonts"))
    if rf is None: rf = OxmlElement("w:rFonts"); rpr.insert(0, rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"): rf.set(qn(a), FONT)
    run.font.color.rgb = RGBColor(0, 0, 0)

def _runs(p, text, bold=False, italic=False, size=12):
    for tok in re.split(r"(\*\*.+?\*\*|\*.+?\*)", text):
        if not tok: continue
        if tok.startswith("**"): _fmt(p.add_run(tok[2:-2]), True, italic, size)
        elif tok.startswith("*"): _fmt(p.add_run(tok[1:-1]), bold, True, size)
        else: _fmt(p.add_run(tok), bold, italic, size)

def _pf(p, line=2.0, first=None, left=None, hanging=None, before=0, after=0, align=None, keep=False):
    f = p.paragraph_format
    f.line_spacing = line; f.space_before = Pt(before); f.space_after = Pt(after)
    if first is not None: f.first_line_indent = Inches(first)
    if left is not None: f.left_indent = Inches(left)
    if hanging is not None: f.first_line_indent = Inches(-hanging)
    if align is not None: p.alignment = align
    if keep: f.keep_with_next = True


class Informe:
    def __init__(self, plantilla, abreviado):
        self.doc = Document(plantilla)
        body = self.doc.element.body
        for el in list(body):
            if el.tag != qn("w:sectPr"): body.remove(el)
        hp = self.doc.sections[0].header.paragraphs[0]
        hp.runs[0].text = abreviado.upper()
        self.nt = self.nf = 0

    # --- bloques
    def p(self, text, first=0.5, **kw):
        para = self.doc.add_paragraph(); _runs(para, text); _pf(para, first=first, **kw); return para

    @staticmethod
    def _sty(para, sid):
        pPr = para._p.get_or_add_pPr(); ps = OxmlElement("w:pStyle"); ps.set(qn("w:val"), sid); pPr.insert(0, ps)

    def h1(self, t):
        para = self.doc.add_paragraph(); self._sty(para, "Heading1"); _runs(para, t, bold=True)
        _pf(para, before=12, align=WD_ALIGN_PARAGRAPH.CENTER, first=0, keep=True)

    def h2(self, t):
        para = self.doc.add_paragraph(); self._sty(para, "Heading2"); _runs(para, t, bold=True)
        _pf(para, before=12, align=WD_ALIGN_PARAGRAPH.LEFT, first=0, keep=True)

    def h3(self, titulo, texto):
        para = self.doc.add_paragraph(); _runs(para, titulo.rstrip(".") + ".", bold=True, italic=True)
        _runs(para, " " + texto); _pf(para, first=0.5, before=12)

    def bloque(self, t):
        para = self.doc.add_paragraph(); _runs(para, t); _pf(para, first=0, left=0.5)

    def salto(self):
        para = self.doc.add_paragraph(); para.paragraph_format.page_break_before = True

    def _caption(self, kind, titulo):
        if kind == "Tabla": self.nt += 1; k = self.nt
        else: self.nf += 1; k = self.nf
        a = self.doc.add_paragraph(); _runs(a, f"{kind} {k}", bold=True); _pf(a, first=0, before=12, keep=True)
        b = self.doc.add_paragraph(); _runs(b, titulo, italic=True); _pf(b, first=0, keep=True)
        return k

    def nota(self, t):
        para = self.doc.add_paragraph(); _runs(para, "*Nota.* " + t, size=11); _pf(para, line=1.0, first=0, after=10)

    def figura(self, titulo, ruta, ancho=5.8, nota=None):
        k = self._caption("Figura", titulo)
        para = self.doc.add_paragraph(); para.add_run().add_picture(str(ruta), width=Inches(ancho))
        _pf(para, line=1.0, first=0, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
        if nota: self.nota(nota)
        else: self.doc.paragraphs[-1].paragraph_format.space_after = Pt(10)
        return k

    def tabla(self, titulo, cab, filas, anchos=None, nota=None, size=10, alin=None):
        kt = self._caption("Tabla", titulo)
        t = self.doc.add_table(rows=1 + len(filas), cols=len(cab)); t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        tblPr = t._tbl.tblPr
        for tag in ("w:tblBorders",):
            for e in tblPr.findall(qn(tag)): tblPr.remove(e)
        anchos = anchos or [6.5 / len(cab)] * len(cab)
        for j, w_ in enumerate(anchos): t.columns[j].width = Inches(w_)
        for i, row in enumerate([cab] + filas):
            for j, val in enumerate(row):
                cell = t.cell(i, j); cell.width = Inches(anchos[j])
                para = cell.paragraphs[0]; _runs(para, str(val), bold=(i == 0), size=size)
                _pf(para, line=1.0, first=0, before=2, after=2, keep=(i < len(filas)),
                    align=(WD_ALIGN_PARAGRAPH.LEFT if (j == 0 or (alin and alin[j] == "l")) else WD_ALIGN_PARAGRAPH.CENTER))
                tcPr = cell._tc.get_or_add_tcPr(); b = OxmlElement("w:tcBorders")
                edges = []
                if i == 0: edges += ["top", "bottom"]
                if i == len(filas): edges += ["bottom"]
                for e in edges:
                    el = OxmlElement(f"w:{e}"); el.set(qn("w:val"), "single"); el.set(qn("w:sz"), "8"); el.set(qn("w:color"), "000000")
                    b.append(el)
                tcPr.append(b)
        if nota: self.nota(nota)
        else:
            sp = self.doc.add_paragraph(); _pf(sp, line=1.0, first=0, after=6)
        return kt

    def portada(self, titulo, fecha):
        sp = self.doc.add_paragraph(); sp.paragraph_format.space_before = Pt(120)
        for i, (t, b) in enumerate([(titulo, True), (CONFIG["autores"], False), (CONFIG["facultad"], False),
                                    (CONFIG["curso"], False), (CONFIG["profesor"], False), (fecha, False)]):
            para = self.doc.add_paragraph(); _runs(para, t, bold=b); _pf(para, first=0, align=WD_ALIGN_PARAGRAPH.CENTER)

    def resumen(self, texto, clave):
        self.salto()
        a = self.doc.add_paragraph(); _runs(a, "Resumen", bold=True); _pf(a, first=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        self.p(texto, first=0)
        b = self.doc.add_paragraph(); _runs(b, "*Palabras clave:* " + clave); _pf(b, first=0.5, before=12)
        self.nw = len(texto.split())

    def referencias(self, citadas=None):
        self.salto()
        a = self.doc.add_paragraph(); _runs(a, "Referencias", bold=True); _pf(a, first=0, align=WD_ALIGN_PARAGRAPH.CENTER)
        for r in sorted(REFS, key=lambda s: s.lower()):
            if citadas is None or any(c in r for c in citadas):
                para = self.doc.add_paragraph(); _runs(para, r); _pf(para, first=None, left=0.5, hanging=0.5)

    def guardar(self, ruta): self.doc.save(ruta)


def cargar(sal):
    return json.load(open(Path(sal) / "results.json", encoding="utf-8"))

def nivel_r2(r2):
    return "muy limitada" if r2 < .10 else "modesta" if r2 < .30 else "moderada" if r2 < .50 else "alta"

def nivel_kappa(k):
    return ("pobre" if k < .21 else "débil" if k < .41 else "moderada" if k < .61 else "sustancial" if k < .81 else "casi perfecta")

TXT_KW = "Keyword"


# =============================================================================== AVANCES I
def build_I(R, sal, plantilla, out):
    q, S, Su, L = R["calidad"], R["stats"], R["suficiencia"], R["limpieza"]
    sal = Path(sal); nv, nc = L["videos_finales"], L["comentarios_finales"]
    top = S["spearman"][0]; sig = [r for r in S["spearman"] if r["q"] < .05]
    sigc = [r for r in S["spearman"] if r["q"] < .05 and r["var"] in
            ("n_com", "sent_mean", "pos_share", "neg_share", "neu_share", "sent_w", "clike_mean", "clike_max", "len_mean",
             "words_mean", "vader_mean", "vader_std", "q_share", "excl_share", "upper_share", "sent_std")]
    kap = S["vader"]["kappa"]; gan = Su["pendiente_final"]
    rep = Informe(plantilla, "Estadísticas de YouTube: Avances I")
    rep.portada("Proyecto Estadísticas de YouTube – Avances I: Calidad, Acceso y Análisis Exploratorio de los Datos", CONFIG["fecha_1"])
    resumen = (f"Este informe documenta la primera etapa del proyecto de procesamiento de lenguaje natural para estimar el desempeño de un video "
               f"de YouTube a partir de sus comentarios. Se trabajó con el conjunto *YouTube Statistics* de Kaggle, que tras la depuración quedó en "
               f"{ent(nv)} videos y {ent(nc)} comentarios. Se describen el origen y el procedimiento reproducible de acceso, y se evalúa la calidad "
               f"con indicadores de completitud, unicidad, validez, consistencia y cobertura. La etiqueta de sentimiento se contrastó con VADER "
               f"(κ de Cohen = {r_(kap)}, concordancia {nivel_kappa(kap)}). La suficiencia se examinó con una curva de aprendizaje y un cálculo de potencia: "
               f"con {ent(nv)} videos solo se detectan correlaciones de |ρ| ≥ {r_(S['potencia']['rho_min'])}. Se definieron como objetivos log(Views), "
               f"log(Likes) y una clase binaria de alta popularidad, y se aplicaron análisis no triviales: correlación de Spearman con intervalos bootstrap y "
               f"control de la tasa de falsos descubrimientos, correlación parcial por antigüedad, Kruskal-Wallis, V de Cramér, paradoja de Simpson y sesgo de "
               f"selección. Se proponen mejoras de recolección mediante la API oficial de YouTube y se plantean las decisiones para la etapa de modelado.")
    rep.resumen(resumen, "análisis de sentimientos, calidad de datos, YouTube, NLP, MLOps, estadística no paramétrica")
    rep.salto()
    t = rep.doc.add_paragraph(); _runs(t, "Proyecto Estadísticas de YouTube – Avances I: Calidad, Acceso y Análisis Exploratorio de los Datos", bold=True)
    _pf(t, first=0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # ---------------- Introducción
    rep.h1("Introducción")
    rep.p("El proyecto busca construir un modelo de procesamiento de lenguaje natural (NLP) que estime el desempeño de un video en YouTube a partir de los "
          "comentarios de su audiencia, y que culmine como un servicio consumible mediante una interfaz o una API, siguiendo prácticas de MLOps: manejo "
          "controlado de los datos, modelo alojado en la nube y capacidad de monitoreo. Sculley et al. (2015) advierten que buena parte de la deuda técnica "
          "de los sistemas de aprendizaje automático nace en las dependencias de datos; por ello esta primera etapa concentra el esfuerzo en los datos y no en el modelo.")
    rep.p("El análisis de sentimientos es una tarea consolidada para cuantificar opiniones en redes sociales. Alemán Viteri (2021), por ejemplo, comparó las "
          "librerías VADER y TextBlob sobre tuits y evaluó su desempeño con matriz de confusión, exactitud, precisión, sensibilidad y especificidad, métricas que "
          "se reutilizan en este proyecto. Ese trabajo también ilustra un riesgo metodológico que aquí se busca evitar: evaluar con una muestra pequeña, etiquetada "
          "por una sola persona y sin intervalos de incertidumbre. Este informe responde, de forma explícita, a cuatro preguntas: ¿cómo se accede a los datos?, "
          "¿cuál es su calidad y son suficientes para un resultado óptimo?, ¿qué variables pueden predecirse? y ¿qué análisis estadísticos no triviales pueden extraerse?")
    rep.p("El documento se organiza así: primero se describe el acceso y la estructura de los datos; después se evalúa su calidad y suficiencia y se proponen "
          "mejoras; luego se identifican variables y predicciones viables; finalmente se presentan los análisis estadísticos no triviales y las conclusiones. "
          "Todas las cifras provienen de un *pipeline* reproducible en Python (semilla fija y huellas SHA-256 de los archivos), de modo que cualquier resultado puede regenerarse.")

    # ---------------- Acceso y descripción
    rep.h1("Acceso y descripción de los datos")
    rep.h2("Origen y procedencia")
    rep.p("Los datos provienen del conjunto *YouTube Statistics* publicado en Kaggle por Patil (s. f.), que contiene dos archivos relacionados por el identificador "
          "del video: *videos-stats.csv*, con estadísticas agregadas por video, y *comments.csv*, con comentarios y una etiqueta de sentimiento. La tarea del curso permite "
          "usar este conjunto o construir uno propio; en esta etapa se parte del conjunto público por su reproducibilidad y se diseña, en paralelo, la ampliación propia "
          "descrita más adelante. Como la ficha pública no documenta un protocolo de muestreo aleatorio, este informe no asume que los videos representen a toda la plataforma: "
          "las conclusiones se limitan a la población de videos obtenidos por las palabras clave del conjunto.")
    rep.h2("Procedimiento de acceso y trazabilidad")
    rep.p("El acceso se realiza descargando el conjunto desde la página de Kaggle (opcionalmente con la herramienta de línea de comandos *kaggle datasets download -d "
          "advaypatil/youtube-statistics*) y colocando los dos CSV en una carpeta de datos de solo lectura. Para garantizar trazabilidad, el *pipeline* calcula la huella "
          "SHA-256 de cada archivo (Tabla 1), de modo que cualquier cambio en los datos de origen sea detectable y quede registrado en el informe, práctica básica de versionado de datos en MLOps.")
    rows = [[k, ent(v["bytes"]), v["sha256"][:16] + "…"] for k, v in q["archivos"].items()]
    rep.tabla("Archivos de origen y huellas de integridad", ["Archivo", "Tamaño (bytes)", "SHA-256 (primeros 16 caracteres)"], rows, [2.4, 1.6, 2.5],
              nota="La huella completa se conserva en results.json. Fuente: elaboración propia a partir de Patil (s. f.).")
    rep.h2("Estructura del conjunto")
    rep.p(f"Tras leer los archivos se obtuvieron {ent(q['v_n'])} filas en la tabla de videos y {ent(q['c_n'])} filas en la de comentarios. La Tabla 2 resume el diccionario de datos. "
          f"Los videos cubren publicaciones entre {q['fecha_min']} y {q['fecha_max']} y se reparten en {len(q['keywords'])} categorías (*Keyword*). La variable *Sentiment* "
          f"toma los valores {', '.join(sorted(q['sentiment_valores']))} (0 = negativo, 1 = neutro, 2 = positivo).")
    rep.tabla("Diccionario de datos", ["Tabla", "Variable", "Tipo", "Descripción"],
              [["Videos", "Title", "Texto", "Título del video"], ["Videos", "Video ID", "Texto (clave)", "Identificador único; llave de unión"],
               ["Videos", "Published At", "Fecha", "Fecha de publicación"], ["Videos", "Keyword", "Categórica", "Palabra clave/categoría temática"],
               ["Videos", "Likes", "Numérica", "Me gusta del video"], ["Videos", "Comments", "Numérica", "Total de comentarios del video"],
               ["Videos", "Views", "Numérica", "Reproducciones (variable objetivo principal)"], ["Comentarios", "Video ID", "Texto (clave)", "Video al que pertenece"],
               ["Comentarios", "Comment", "Texto", "Contenido del comentario"], ["Comentarios", "Likes", "Numérica", "Me gusta del comentario"],
               ["Comentarios", "Sentiment", "Ordinal (0–2)", "Etiqueta de sentimiento provista por el autor"]],
              [1.1, 1.2, 1.2, 3.0], alin=["l", "l", "l", "l"], nota="Fuente: Patil (s. f.).")
    rep.h2("Estrategias avanzadas de recolección")
    rep.p("Para superar las limitaciones del conjunto público se diseñó una recolección propia mediante la API oficial *YouTube Data API v3*. Los métodos *videos.list* "
          "(estadísticas del video) y *commentThreads.list* (hilos de comentarios, con paginación y orden por relevancia o por fecha) permiten reconstruir ambas tablas con la misma "
          "estructura (Google, s. f.). La estrategia incorpora cinco decisiones: (a) muestreo estratificado por categoría y por cuartiles de reproducciones, para no sobrerrepresentar "
          "videos virales; (b) captura de comentarios por dos órdenes (relevancia y fecha) para medir el sesgo de selección de los comentarios destacados; (c) registro de la fecha y hora "
          "de captura y de dos o más instantáneas por video, para modelar la evolución temporal; (d) control de cuota diaria con reintentos y almacenamiento incremental en formato "
          "columnar (Parquet) con esquema validado; y (e) cumplimiento de los términos de servicio de la plataforma y anonimización de autores de comentarios.")

    # ---------------- Calidad
    rep.h1("Calidad de los datos")
    rep.p("La calidad se evaluó con las dimensiones de completitud, unicidad, validez, consistencia y cobertura, derivadas del marco de Wang y Strong (1996), que entiende la "
          "calidad como adecuación al uso por parte del consumidor de los datos. La Tabla 3 reúne los indicadores calculados sobre los archivos originales, antes de la limpieza.")
    mv = sum(q["v_missing"].get(k, 0) for k in ("Views", "Likes", "Comments")); mc = q["c_missing"].get("Comment", 0)
    ok = lambda c: "Adecuada" if c else "Requiere tratamiento"
    fo = q["frac_comentarios_observados_mediana"]
    rows = [
        ["Completitud", "Faltantes en Views/Likes/Comments (videos)", ent(mv), ok(mv == 0)],
        ["Completitud", "Comentarios vacíos o nulos", ent(mc + q["c_vacios"]), ok(mc + q["c_vacios"] == 0)],
        ["Unicidad", "Video ID duplicados", ent(q["v_dup_id"]), ok(q["v_dup_id"] == 0)],
        ["Unicidad", "Comentarios duplicados (mismo video y texto)", ent(q["c_dup_rows"]), ok(q["c_dup_rows"] == 0)],
        ["Validez", "Valores negativos en Views/Likes/Comments", ent(sum(q["neg_values"].values())), ok(sum(q["neg_values"].values()) == 0)],
        ["Validez", "Fechas no interpretables", ent(q["fecha_invalida"]), ok(q["fecha_invalida"] == 0)],
        ["Consistencia", "Videos con Likes > Views", ent(q["likes_gt_views"]), ok(q["likes_gt_views"] == 0)],
        ["Consistencia", "Comentarios cuyo video no está en la tabla de videos", ent(q["orphan_comments"]), ok(q["orphan_comments"] == 0)],
        ["Cobertura", "Videos sin ningún comentario", ent(q["videos_sin_comentarios"]), ok(q["videos_sin_comentarios"] == 0)],
        ["Cobertura", "Videos con ≥ 5 comentarios / ≥ 10", f"{pct(q['cob_5'])} / {pct(q['cob_10'])}", "Ver suficiencia"],
        ["Representatividad", "Mediana de la fracción de comentarios observados", pct(fo) if fo is not None else "n. d.", "Muestra de los más relevantes"],
        ["Ruido textual", "Comentarios con >30 % de caracteres no ASCII", pct(q["c_no_ascii"]), ok(q["c_no_ascii"] < .05)],
        ["Ruido textual", "Comentarios con emojis / con URL", f"{pct(q['c_emoji'])} / {pct(q['c_url'])}", "Normalizar"],
        ["Valores extremos", "Videos atípicos en log(Views) (regla de Tukey)", ent(q["outliers_iqr_logviews"]), "Conservar y modelar en log"],
    ]
    rep.tabla("Indicadores de calidad de los datos originales", ["Dimensión", "Indicador", "Resultado", "Valoración"], rows, [1.2, 2.9, 1.1, 1.3], alin=["l", "l", "c", "l"],
              nota="Elaboración propia. «Requiere tratamiento» indica que el indicador es distinto de cero y el pipeline aplica la regla de depuración descrita en el texto.")
    rep.p(f"La depuración aplicada fue deliberadamente conservadora: se eliminaron videos duplicados y sin la variable objetivo (*Views* ausente o no positiva), valores negativos en "
          f"*Likes* y *Comments* se trataron como ausentes, se descartaron comentarios nulos, vacíos o duplicados y se conservaron únicamente comentarios cuyo video existe en la tabla de videos. "
          f"El resultado fueron {ent(nv)} videos y {ent(nc)} comentarios ({ent(L['videos_descartados'])} videos descartados). Los valores extremos no se eliminaron: en datos de "
          f"popularidad son información legítima, y se manejan con una transformación logarítmica.")
    rep.h2("Validación de la etiqueta de sentimiento")
    v = S["vader"]; cm = v["cm"]
    rep.p(f"Como el sentimiento será una variable predictora clave, se verificó su fiabilidad contrastándola con VADER (Hutto & Gilbert, 2014), un modelo basado en reglas y léxico "
          f"diseñado para texto de redes sociales, con umbrales estándar de ±0,05 sobre el puntaje compuesto. La concordancia fue de κ de Cohen = {r_(kap)} (Cohen, 1960), "
          f"clasificada como {nivel_kappa(kap)} según los rangos de Landis y Koch (1977), con una exactitud de {pct(v['acc'])} sobre {ent(v['n'])} comentarios (Tabla 4). "
          f"Esto implica que la etiqueta del conjunto contiene ruido respecto de un léxico independiente, o que ambas fuentes miden constructos distintos; por eso en el modelado se emplean "
          f"ambas señales y no solo una.")
    rep.tabla("Matriz de confusión: etiqueta del conjunto frente a VADER", ["Etiqueta del conjunto", "VADER negativo", "VADER neutro", "VADER positivo"],
              [["Negativo (0)"] + [ent(x) for x in cm[0]], ["Neutro (1)"] + [ent(x) for x in cm[1]], ["Positivo (2)"] + [ent(x) for x in cm[2]]],
              [1.8, 1.5, 1.5, 1.5], nota="Filas: etiqueta del conjunto; columnas: clase asignada por VADER. Elaboración propia.")
    rep.h2("Distribución de las variables")
    rep.figura("Distribución de reproducciones, likes y comentarios por video (escala logarítmica)", sal / "fig_distribuciones.png", 6.3,
               "Elaboración propia con los datos de Patil (s. f.).")
    lam = S["boxcox_lambda"]
    rep.p(f"Las tres métricas de popularidad tienen colas muy largas (Figura 1). La asimetría de *Views* pasa de {n(S['skew_views'])} en escala original a {n(S['skew_logviews'])} tras aplicar "
          f"logaritmo, con curtosis en exceso de {n(S['kurt_logviews'])}. {('El parámetro λ de Box-Cox estimado (' + n(lam) + ') ' + ('es cercano a cero, lo que respalda la transformación logarítmica (Box & Cox, 1964).' if abs(lam) < .3 else 'indica que una transformación más suave que el logaritmo también sería razonable (Box & Cox, 1964); se mantiene el logaritmo por su interpretabilidad.')) if lam is not None else ''} "
          f"Por tanto, el modelado se hace sobre log(1 + Views), lo que además estabiliza la varianza y hace que el error se interprete de forma multiplicativa.")
    rep.figura("Videos por categoría y distribución de la etiqueta de sentimiento", sal / "fig_keyword_sentimiento.png", 6.3, "Elaboración propia.")
    sv = q["sentiment_valores"]; tot = sum(sv.values())
    rep.p("La Figura 2 muestra el balance de categorías y de clases de sentimiento: " + "; ".join(f"la clase {k} concentra {pct(x / tot)}" for k, x in sv.items()) +
          ". Un desbalance marcado afecta métricas como la exactitud; por ello se complementará con exactitud balanceada, F1 y área bajo la curva ROC en la etapa de modelado.")
    rep.h2("Sesgos y limitaciones identificados")
    sb = S["sesgo_sel"]
    rep.h3("Sesgo de selección de comentarios", f"El conjunto contiene, por video, una muestra de comentarios relevantes (mediana de {n(q['com_por_video']['50%'], 0)} por video; "
           f"Figura 3), no la totalidad. En el nivel de comentario, los *likes* difieren entre clases de sentimiento (Kruskal-Wallis *H* = {n(sb['H'])}, *p* {pv(sb['p'])}); "
           f"las medianas son {', '.join('clase ' + k + ': ' + n(x, 1) for k, x in sb['med'].items())}. Es decir, los comentarios visibles no son una muestra aleatoria de la conversación.")
    rep.figura("Comentarios disponibles por video", sal / "fig_comentarios_por_video.png", 3.6, "Elaboración propia.")
    rep.h3("Confusión por antigüedad", "Un video más antiguo acumula más reproducciones y también más comentarios; sin control, la antigüedad produce correlaciones espurias. "
           "Por eso todas las asociaciones se reportan también como correlaciones parciales controlando la antigüedad.")
    rep.h3("Instantánea única e idioma", "Los datos son una fotografía en un único momento, lo que impide separar el efecto del tiempo del efecto del contenido, y el modelo léxico de VADER está "
           f"calibrado para inglés; {pct(q['c_no_ascii'])} de los comentarios tiene más de 30 % de caracteres no ASCII, indicio de texto en otros alfabetos o con emojis intensivos.")

    # ---------------- Suficiencia
    rep.h1("Suficiencia de los datos y propuestas de mejora")
    P = S["potencia"]
    rep.p(f"Se evaluó la suficiencia con tres criterios. Primero, el tamaño muestral relativo al número de parámetros: hay {n(S['muestras_por_param'], 1)} videos por parámetro candidato "
          f"(variables numéricas más niveles de categoría), por encima de la regla práctica de 10 a 20 observaciones por parámetro para modelos lineales. Segundo, la potencia estadística: con "
          f"*n* = {ent(P['n'])} videos y α = ,05 se alcanza 80 % de potencia solo para correlaciones de |ρ| ≥ {r_(P['rho_min'])}; efectos menores no serán detectables con estos datos. "
          f"Tercero, una curva de aprendizaje (Figura 4), que mide el R² de validación cruzada al entrenar con fracciones crecientes de los datos.")
    cu = Su["curvas"]["Grad. boosting"]
    rep.figura("Curva de aprendizaje para predecir log(Views) con variables de comentarios, categoría y antigüedad", sal / "fig_curva_aprendizaje.png", 4.4,
               "Barras de error: desviación estándar en validación cruzada de 5 particiones. Elaboración propia.")
    rep.p(f"Con {ent(cu[0][0])} videos el R² del modelo de refuerzo del gradiente es {r_(cu[0][1])}, y con {ent(cu[-1][0])} llega a {r_(cu[-1][1])}. La pendiente del último tramo equivale a "
          f"{'un aumento de ' + n(gan, 3) + ' puntos de R² por cada 100 videos adicionales' if gan > 0 else 'una meseta'}, por lo que se concluye que "
          f"{'los datos todavía no alcanzan una meseta y duplicar la muestra aportaría mejoras apreciables' if gan > 0.5 else 'el rendimiento se aproxima a una meseta con este tamaño; más videos aportarían mejoras marginales y conviene enriquecer la información de cada video (más comentarios y metadatos) en lugar de solo ampliar el número de videos'}. "
          f"En síntesis, los datos son suficientes para un prototipo y para análisis exploratorio, pero no para garantizar un resultado óptimo; la Tabla 5 prioriza las mejoras.")
    rep.tabla("Limitaciones detectadas y mejoras propuestas", ["Limitación", "Mejora propuesta", "Justificación", "Esfuerzo"],
              [["Solo comentarios destacados", "Capturar comentarios por relevancia y por fecha", "Mide y corrige el sesgo de selección", "Medio"],
               ["Pocos comentarios por video", "Paginar hasta ≥ 100 hilos por video con la API", "Reduce la varianza de los agregados de sentimiento", "Medio"],
               ["Etiqueta de sentimiento ruidosa", "Doble anotación manual de una submuestra y κ entre anotadores", "Cuantifica el error de la etiqueta", "Alto"],
               ["Instantánea única", "Dos o más capturas por video con fecha registrada", "Permite estudiar crecimiento y separar antigüedad", "Medio"],
               ["Idioma mixto", "Detectar idioma y usar léxico/modelo multilingüe", "VADER está calibrado para inglés", "Bajo"],
               ["Cobertura temática limitada", "Muestreo estratificado por categoría y cuartil de vistas", "Evita sesgo hacia videos virales", "Medio"],
               ["Sin contexto del canal", "Agregar suscriptores y duración del video", "Variables de confusión de primer orden", "Bajo"]],
              [1.45, 2.0, 2.1, 0.95], alin=["l", "l", "l", "c"], nota="Elaboración propia.")

    # ---------------- Variables
    rep.h1("Identificación de variables y predicciones")
    rep.h2("Variables candidatas")
    rep.p("Los comentarios se agregan al nivel de video para construir tres grupos de predictores (Tabla 6). El grupo A usa solo señales de la audiencia; el grupo B añade metadatos del "
          "video. Se descartan *Likes* y *Comments* del video como predictores de *Views* por circularidad: son consecuencias simultáneas de la audiencia, no antecedentes.")
    rep.tabla("Variables predictoras y objetivos", ["Grupo", "Variables", "Rol", "Riesgo principal"],
              [["A. Sentimiento", "Media, desviación y proporciones positiva/negativa/neutra de *Sentiment*; sentimiento ponderado por likes; VADER medio y dispersión", "Predictor", "Ruido de etiqueta"],
               ["A. Interacción", "Número de comentarios (log); likes medios y máximos del comentario (log)", "Predictor", "Sesgo de selección"],
               ["A. Estilo textual", "Longitud, palabras, proporción de «?», «!» y mayúsculas", "Predictor", "Idioma"],
               ["A. Texto", "TF-IDF de unigramas y bigramas de los comentarios (Salton & Buckley, 1988)", "Predictor", "Dimensionalidad"],
               ["B. Metadatos", "Categoría (*Keyword*), antigüedad (log), longitud y sentimiento del título", "Predictor", "Confusión por tiempo"],
               ["Objetivo 1", "log(1 + Views)", "Regresión", "Asimetría (resuelta con log)"],
               ["Objetivo 2", "log(1 + Likes) y tasa de likes (Likes/Views)", "Regresión", "Dependencia con Views"],
               ["Objetivo 3", "Alta vs. baja popularidad (sobre/bajo la mediana de log Views)", "Clasificación", "Umbral arbitrario"]],
              [1.2, 3.0, 0.9, 1.4], alin=["l", "l", "c", "l"], nota="Elaboración propia.")
    rep.h2("Viabilidad de las predicciones")
    rep.p(f"La predicción de log(Views) es viable pero acotada. La señal individual más fuerte es *{top['etiqueta']}* (ρ = {r_(top['rho'])}), y {len(sig)} de {len(S['spearman'])} variables "
          f"mantienen significación tras controlar la tasa de falsos descubrimientos. Además, se espera una capacidad explicativa limitada porque las reproducciones dependen de factores ajenos a los "
          f"comentarios (canal, algoritmo de recomendación, promoción). Conviene precisar el alcance del producto: dado que los comentarios aparecen después de que el video recibe vistas, el modelo no "
          f"«pronostica» el futuro de forma causal, sino que estima el desempeño esperado dada la reacción de la audiencia, útil para diagnóstico y como señal temprana con los primeros comentarios. "
          f"Para la tasa de likes, la correlación con el sentimiento medio es ρ = {r_(S['like_rate']['rho'])} (*p* {pv(S['like_rate']['p'])}), lo que la convierte en un objetivo secundario razonable.")

    # ---------------- Estadística
    rep.h1("Análisis estadísticos no triviales")
    rep.p("Dado que las variables de popularidad no son normales, se emplearon métodos no paramétricos y de remuestreo. Para controlar el error por comparaciones múltiples se aplicó el procedimiento de "
          "Benjamini y Hochberg (1995), que acota la tasa esperada de falsos descubrimientos, y los intervalos de confianza se estimaron por bootstrap (Efron, 1979) con 600 remuestreos.")
    rep.h2("Asociación entre variables de comentarios y popularidad")
    rows = [[r["etiqueta"], r_(r["rho"]), f"[{r_(r['lo'])}; {r_(r['hi'])}]", r_(r["rho_parcial"]) if r["var"] != "age_days" else "—", r_(r["q"], 3)] for r in S["spearman"][:10]]
    rep.tabla("Correlaciones de Spearman con log(Views)", ["Variable", "ρ", "IC 95 % (bootstrap)", "ρ parcial (antigüedad)", "q (FDR)"], rows, [2.5, 0.7, 1.4, 1.1, 0.8],
              nota="q: valor p ajustado por Benjamini-Hochberg sobre las 20 variables. ρ parcial: control lineal por rangos de la antigüedad. Se muestran las 10 mayores en valor absoluto.")
    rep.figura("Correlación de Spearman con log(Views) e intervalos de confianza", sal / "fig_spearman.png", 5.2, "Se muestran las 12 variables con mayor |ρ|. Elaboración propia.")
    rep.p(f"Al descontar la antigüedad, {'la señal de ' + top['etiqueta'] + ' pasa de ρ = ' + r_(top['rho']) + ' a ρ parcial = ' + r_(top['rho_parcial']) if top['var'] != 'age_days' else 'la antigüedad aparece como la variable más asociada, lo que confirma su papel de confusor'}. "
          f"De las variables derivadas de comentarios, {len(sigc)} conservan significación tras el ajuste por falsos descubrimientos.")
    K, C = S["kruskal"], S["chi"]
    rep.h3("Diferencias entre categorías", f"Una prueba de Kruskal-Wallis (Kruskal & Wallis, 1952) indica {'diferencias' if K['p'] < .05 else 'que no hay diferencias claras'} en log(Views) entre las {K['k']} categorías "
           f"(*H* = {n(K['H'])}, *p* {pv(K['p'])}, ε² = {r_(K['eps2'], 3)}), un tamaño de efecto {'pequeño' if K['eps2'] < .06 else 'moderado' if K['eps2'] < .14 else 'grande'}. "
           f"Además, la asociación entre categoría y sentimiento dominante es {'significativa' if C['p'] < .05 else 'no significativa'} (χ²({C['dof']}) = {n(C['chi2'])}, *p* {pv(C['p'])}, *V* de Cramér = {r_(C['V'])}).")
    Sm = S["simpson"]
    rep.h3("Paradoja de Simpson", f"La correlación global entre sentimiento medio y log(Views) es ρ = {r_(Sm['pooled'])}, mientras que el promedio ponderado dentro de cada categoría (≥ 20 videos) es ρ = {r_(Sm['within_w'])}. "
           f"{'El signo se invierte, por lo que la relación global es en parte un artefacto de la mezcla de categorías (Simpson, 1951); la categoría debe incluirse como covariable.' if Sm['invierte'] else 'El signo se mantiene, de modo que la relación no es un artefacto de la mezcla de categorías (Simpson, 1951), aunque su magnitud cambia al condicionar.'}")
    rep.h3("Potencia y tamaño del efecto", f"Con *n* = {ent(P['n'])}, el efecto mínimo detectable es |ρ| = {r_(P['rho_min'])}; los resultados no significativos deben leerse como ausencia de evidencia suficiente y no como ausencia de efecto.")
    rep.h2("Plan para la etapa de modelado")
    rep.p("A partir de estos hallazgos se decide: (a) usar log(Views) como objetivo principal; (b) incluir categoría y antigüedad como covariables para evitar la confusión descrita; (c) comparar "
          "modelos lineales regularizados (Hoerl & Kennard, 1970), bosques aleatorios (Breiman, 2001) y refuerzo del gradiente (Friedman, 2001) frente a una referencia ingenua; (d) validar con "
          "validación cruzada repetida y pruebas de comparación corregidas; y (e) evaluar el clasificador con exactitud, precisión, sensibilidad y especificidad (Alemán Viteri, 2021; Barrios Arce, 2019). "
          "Todo ello se implementa con scikit-learn (Pedregosa et al., 2011).")

    # ---------------- Conclusiones
    rep.h1("Conclusiones")
    rep.p(f"Los datos son accesibles de forma reproducible y su calidad estructural es {'buena' if q['v_dup_id'] == 0 and q['orphan_comments'] == 0 else 'aceptable tras la depuración'}, pero presentan limitaciones de fondo: "
          f"comentarios seleccionados por relevancia, etiqueta de sentimiento con concordancia {nivel_kappa(kap)} frente a VADER, instantánea única y una potencia que solo permite detectar efectos de |ρ| ≥ {r_(P['rho_min'])}. "
          f"Aun así, son suficientes para un primer modelo. Las variables viables son log(Views), log(Likes) y la clase de alta popularidad, y los análisis no triviales muestran que la antigüedad y la "
          f"categoría deben controlarse. La siguiente etapa se centrará en comparar modelos con validación rigurosa y en empaquetar el mejor como servicio monitoreado.")
    rep.referencias(["Alemán", "Barrios", "Benjamini", "Box", "Breiman", "Cohen", "Efron", "Friedman", "Google", "Hoerl", "Hutto", "Kruskal", "Landis", "Patil", "Pedregosa", "Salton", "Sculley", "Simpson", "Wang"])
    rep.guardar(out)
    return rep.nw


# =============================================================================== AVANCES II
def build_II(R, sal, plantilla, out):
    q, S, Su, L, M = R["calidad"], R["stats"], R["suficiencia"], R["limpieza"], R["modelado"]
    sal = Path(sal); nv, nc = L["videos_finales"], L["comentarios_finales"]
    reg, clf, P = M["reg"]["modelos"], M["clf"]["modelos"], M["reg"]["pruebas"]
    best, bestc = M["reg"]["mejor"], M["clf"]["mejor"]; B = reg[best]; BC = clf[bestc]; ref = reg["Referencia (mediana)"]
    mej = 100 * (1 - B["rmse"] / ref["rmse"])
    ra = min((k for k in reg if "– A" in k), key=lambda k: reg[k]["rmse"]); rb = min((k for k in reg if "– B" in k), key=lambda k: reg[k]["rmse"])
    rt = reg["TF-IDF + Ridge (solo texto)"]; rh = reg["Híbrido B + TF-IDF + Ridge"]
    imp = M["importancia"]; ntot = M["reg"]["n_modelos"] + len(clf); ncfg = M["reg"]["n_configs"] + M["clf"]["n_configs"]
    cmx = M["clf"]["cm"]; (tn, fp), (fn, tp) = cmx
    sens, esp, prec = tp / max(tp + fn, 1), tn / max(tn + fp, 1), tp / max(tp + fp, 1)
    rep = Informe(plantilla, "Estadísticas de YouTube: Avances II")
    rep.portada("Proyecto Estadísticas de YouTube – Avances II: Análisis Estadístico, Selección y Evaluación de Modelos", CONFIG["fecha_2"])
    resumen = (f"Este informe presenta la segunda etapa del proyecto: el análisis estadístico y el modelado para estimar el desempeño de videos de YouTube a partir de sus comentarios. "
               f"Con {ent(nv)} videos y {ent(nc)} comentarios se construyeron variables agregadas de sentimiento, interacción y estilo textual, además de TF-IDF y metadatos. Se evaluaron "
               f"{ntot} modelos y {ent(ncfg)} configuraciones, con validación cruzada repetida, prevención de fuga de información y pruebas de comparación corregidas. "
               f"El mejor regresor para log(Views) fue «{best}», con RMSE de {n(B['rmse'])} (R² = {r_(B['r2'])}, capacidad explicativa {nivel_r2(B['r2'])}), "
               f"{n(mej, 1)} % por debajo de la referencia. En clasificación de alta popularidad, «{bestc}» alcanzó un AUC de {r_(BC['auc'])}, sensibilidad de {r_(sens)} y "
               f"especificidad de {r_(esp)}. La importancia por permutación y la ablación indican qué variables aportan. Se concluye que un único modelo no es suficiente y se propone "
               f"una arquitectura con un regresor, un clasificador y un módulo de sentimiento, con monitoreo y despliegue como API.")
    rep.resumen(resumen, "modelado predictivo, validación cruzada, importancia de variables, análisis de sentimientos, MLOps")
    rep.salto()
    t = rep.doc.add_paragraph(); _runs(t, "Proyecto Estadísticas de YouTube – Avances II: Análisis Estadístico, Selección y Evaluación de Modelos", bold=True)
    _pf(t, first=0, align=WD_ALIGN_PARAGRAPH.CENTER)

    rep.h1("Introducción")
    rep.p("La primera etapa mostró que los datos son suficientes para un prototipo pero con limitaciones de fondo: comentarios seleccionados por relevancia, una etiqueta de sentimiento con concordancia "
          f"{nivel_kappa(S['vader']['kappa'])} frente a VADER y la necesidad de controlar la antigüedad y la categoría del video. Esta segunda etapa se centra en el modelo y responde a las preguntas de la "
          "asignación: ¿qué análisis estadísticos pueden hacerse con el conjunto?, ¿qué modelo se seleccionó y con qué desempeño y métricas?, ¿cuántos modelos se probaron y con qué herramientas?, "
          "¿es suficiente un solo modelo? y ¿qué variables se seleccionaron para predecir?")
    rep.p("El enfoque prioriza el rigor de la evaluación sobre el valor de una métrica aislada. Toda comparación se hace sobre predicciones fuera de muestra con validación cruzada repetida, "
          "el preprocesamiento se ajusta únicamente con los datos de entrenamiento de cada partición y las diferencias entre modelos se contrastan con una prueba estadística apropiada. "
          "El informe sigue la lógica del trabajo de Alemán Viteri (2021), que evaluó clasificadores de sentimiento con matriz de confusión, exactitud, precisión, sensibilidad y especificidad, y "
          "la amplía con intervalos de variabilidad, comparación contra una referencia ingenua y análisis de errores por segmento.")

    rep.h1("Análisis estadísticos realizados")
    rep.h2("Hallazgos de la etapa I que condicionan el modelado")
    top = S["spearman"][0]; Sm = S["simpson"]
    rep.p(f"Las variables de popularidad son fuertemente asimétricas (asimetría de *Views* = {n(S['skew_views'])}; en logaritmo = {n(S['skew_logviews'])}), por lo que el objetivo es log(1 + Views). "
          f"La señal individual más fuerte fue *{top['etiqueta']}* (ρ = {r_(top['rho'])}); la categoría explica diferencias de popularidad con un tamaño de efecto ε² = {r_(S['kruskal']['eps2'], 3)} y "
          f"{'la relación global entre sentimiento y vistas cambia de signo al condicionar por categoría (paradoja de Simpson), por lo que la categoría se incluye como covariable' if Sm['invierte'] else 'la relación entre sentimiento y vistas conserva su signo al condicionar por categoría'}. "
          f"Estos resultados justifican los grupos de variables A (comentarios) y B (comentarios más metadatos).")
    rep.h2("Análisis adicionales alineados con el conjunto de datos")
    rep.h3("Comparación estadística de modelos", f"Las diferencias de RMSE entre modelos se contrastaron con la prueba *t* remuestreada corregida de Nadeau y Bengio (2003), que corrige la subestimación de la varianza "
           f"causada por el solapamiento de los conjuntos de entrenamiento en validación cruzada. Frente a la referencia, el mejor modelo reduce el RMSE en {n(P['vs_referencia']['delta_rmse'], 3)} "
           f"(*t* = {n(P['vs_referencia']['t'])}, *p* {pv(P['vs_referencia']['p'])}); frente al segundo mejor («{P['segundo']}»), la diferencia es de {n(P['vs_segundo']['delta_rmse'], 3)} "
           f"(*t* = {n(P['vs_segundo']['t'])}, *p* {pv(P['vs_segundo']['p'])}).")
    rep.h3("Importancia por permutación", "La contribución de cada variable se estimó como el aumento del RMSE en un conjunto de prueba cuando sus valores se permutan al azar (25 repeticiones), "
           "medida agnóstica al modelo y menos sesgada que la importancia interna de los árboles ante variables de alta cardinalidad.")
    rep.h3("Ablación por grupos de variables", f"Se compararon cuatro configuraciones: solo comentarios numéricos (A), comentarios más metadatos (B), solo texto con TF-IDF y un híbrido B más TF-IDF. "
           f"Esto aísla cuánto aporta cada fuente de información al desempeño.")
    rep.h3("Diagnóstico de errores", f"Se examinó la heterocedasticidad (correlación de Spearman entre el valor predicho y el error absoluto: ρ = {r_(M['reg']['heterocedasticidad_rho'])}) y el RMSE por categoría, "
           f"para detectar segmentos donde el modelo no es confiable.")

    rep.h1("Diseño experimental y herramientas")
    rep.h2("Preprocesamiento")
    rep.p("Los comentarios se agregan por video. Se calcularon variables numéricas (proporciones y medias de sentimiento, sentimiento ponderado por likes, puntaje compuesto de VADER (Hutto & Gilbert, 2014), "
          "longitud, signos de interrogación y exclamación, mayúsculas) y se transformaron con log(1 + x) las variables de conteo y longitud. Para el texto se concatenaron los comentarios de cada video y se "
          "vectorizaron con TF-IDF (Salton & Buckley, 1988) sobre unigramas y bigramas, con frecuencia mínima de documento de 3, máximo de 20 000 términos, escala sublineal y eliminación de palabras vacías "
          "del inglés. Los valores faltantes numéricos se imputan con la mediana, las variables se estandarizan y la categoría se codifica en indicadores con agrupación de niveles poco frecuentes. "
          "Todos estos pasos están dentro de un *Pipeline* de scikit-learn (Pedregosa et al., 2011) que se ajusta solo con los datos de entrenamiento de cada partición, lo que evita la fuga de información.")
    rep.h2("Variables seleccionadas")
    rep.tabla("Conjuntos de variables evaluados", ["Conjunto", "Contenido", "N.º de variables"],
              [["A", "Sentimiento, interacción y estilo de los comentarios", str(len([1 for k in reg[ra]['cols']]))],
               ["B", "A + categoría, antigüedad, longitud, sentimiento e interrogación del título", str(reg[rb]['n_vars'])],
               ["Texto", "TF-IDF de unigramas y bigramas (hasta 20 000 términos)", "≤ 20 000"],
               ["Híbrido", "B + TF-IDF", f"{reg['Híbrido B + TF-IDF + Ridge']['n_vars']} + TF-IDF"]],
              [1.0, 4.2, 1.3], alin=["l", "l", "c"], nota="B incluye una variable categórica que se expande en indicadores. Elaboración propia.")
    rep.p("Los conjuntos se resumen en la Tabla 1. El objetivo principal es log(1 + Views). Se evaluaron además dos objetivos secundarios: log(1 + Likes) y la clase de alta popularidad (log Views por encima de la mediana). "
          "Se excluyeron *Likes* y *Comments* del video como predictores de *Views* por circularidad, y la antigüedad se incluye como covariable de control.")
    rep.h2("Protocolo de evaluación")
    rep.p("La regresión se evaluó con validación cruzada de 5 particiones repetida (RMSE, error absoluto medio, R² y correlación de Spearman entre valores reales y predichos); la clasificación, con "
          "validación estratificada repetida (exactitud, exactitud balanceada, F1, área bajo la curva ROC, precisión, sensibilidad y especificidad; véase Barrios Arce, 2019). Los hiperparámetros de "
          "bosques aleatorios (Breiman, 2001) y de refuerzo del gradiente (Friedman, 2001) se optimizaron con búsqueda en malla anidada de 3 particiones dentro de cada partición de entrenamiento, y la penalización "
          "de la regresión Ridge (Hoerl & Kennard, 1970) con validación cruzada interna. Las herramientas fueron Python, pandas, SciPy, scikit-learn, vaderSentiment, matplotlib y joblib, con semilla fija "
          "(42) y huellas SHA-256 de los datos.")

    rep.h1("Modelos evaluados")
    fam = lambda k: ("Referencia" if "Referencia" in k else "Lineal regularizado" if "Ridge" in k or "TF-IDF" in k or "Híbrido" in k else "Ensamble de árboles (bagging)" if "Bosque" in k else "Ensamble de árboles (boosting)")
    rows = [[k, fam(k), ", ".join(["A"] if "– A" in k else ["B"] if ("– B" in k) else ["Texto"] if k.startswith("TF-IDF") else ["B + texto"] if "Híbrido" in k else ["A"]), str(v["configs"])] for k, v in reg.items()]
    rep.p(f"Se probaron {M['reg']['n_modelos']} modelos de regresión y {len(clf)} de clasificación ({ntot} en total), con {ent(ncfg)} configuraciones de hiperparámetros evaluadas (Tabla 2). "
          "La inclusión de una referencia ingenua es esencial: un modelo solo es útil si supera de forma estadísticamente significativa a predecir siempre la mediana.")
    rep.tabla("Modelos de regresión evaluados", ["Modelo", "Familia", "Variables", "Configuraciones"], rows, [2.4, 2.1, 1.0, 1.0], alin=["l", "l", "c", "c"],
              nota="Configuraciones: combinaciones de hiperparámetros exploradas (malla) o valores de penalización evaluados (Ridge). Elaboración propia.")

    rep.h1("Resultados")
    rep.h2("Regresión sobre log(Views)")
    ordn = sorted(reg, key=lambda k: reg[k]["rmse"])
    rows = [[k, f"{n(reg[k]['rmse'])} ({n(reg[k]['rmse_sd'])})", n(reg[k]["mae"]), f"{r_(reg[k]['r2'])} ({r_(reg[k]['r2_sd'])})", r_(reg[k]["rho"])] for k in ordn]
    rep.tabla("Desempeño de los modelos de regresión (validación cruzada repetida)", ["Modelo", "RMSE (DE)", "MAE", "R² (DE)", "ρ de Spearman"], rows, [2.6, 1.1, 0.7, 1.1, 1.0],
              nota="DE: desviación estándar entre particiones. Errores en unidades de log(1 + Views). Ordenado de menor a mayor RMSE.")
    rep.figura("Distribución del RMSE por partición para cada modelo", sal / "fig_modelos_rmse.png", 6.0, "Elaboración propia.")
    rep.p(f"La Tabla 3 y la Figura 1 resumen el desempeño. El mejor modelo fue «{best}» con RMSE = {n(B['rmse'])}, MAE = {n(B['mae'])} y R² = {r_(B['r2'])}, una capacidad explicativa {nivel_r2(B['r2'])}. Como el RMSE se expresa en log, un error de {n(B['rmse'])} equivale "
          f"a un factor multiplicativo típico de {n(2.718281828 ** B['rmse'])} sobre las vistas. Mejora {n(mej, 1)} % al RMSE de la referencia ({n(ref['rmse'])}), diferencia "
          f"{'estadísticamente significativa' if P['vs_referencia']['p'] < .05 else 'que no alcanza significación estadística'} (*p* {pv(P['vs_referencia']['p'])}). Frente al segundo mejor modelo la diferencia "
          f"{'es significativa' if P['vs_segundo']['p'] < .05 else 'no es significativa'} (*p* {pv(P['vs_segundo']['p'])}), "
          f"{'por lo que la elección está respaldada estadísticamente' if P['vs_segundo']['p'] < .05 else 'lo que indica que ambos son estadísticamente equivalentes y la elección puede basarse en simplicidad e interpretabilidad'}.")
    rep.h2("Aporte de cada fuente de información")
    gb = 100 * (1 - reg[rb]["rmse"] / reg[ra]["rmse"])
    rep.p(f"El mejor modelo con solo comentarios (A) fue «{ra}» (RMSE = {n(reg[ra]['rmse'])}), mientras que con metadatos (B) fue «{rb}» (RMSE = {n(reg[rb]['rmse'])}); añadir categoría, antigüedad y "
          f"título {'reduce' if gb > 0 else 'no reduce'} el error en {n(abs(gb), 1)} %. El modelo de solo texto obtuvo RMSE = {n(rt['rmse'])} (R² = {r_(rt['r2'])}) y el híbrido, RMSE = {n(rh['rmse'])} "
          f"(R² = {r_(rh['r2'])}). {'La información textual no aporta de forma apreciable sobre las variables numéricas, lo que sugiere que el contenido léxico de los comentarios es poco predictivo del volumen de vistas.' if rt['rmse'] >= reg[rb]['rmse'] else 'El texto aporta señal propia, lo que justifica conservar la representación TF-IDF.'}")
    rep.h2("Calibración y errores por segmento")
    rep.figura("Valores reales y predichos de log(Views) fuera de muestra", sal / "fig_real_vs_predicho.png", 3.8, f"Modelo: {best}. La línea es la identidad. Elaboración propia.")
    pk = sorted(M["reg"]["por_keyword"], key=lambda r: -r["rmse"])
    rep.p(f"La Figura 2 compara valores reales y predichos. La correlación entre el valor predicho y el error absoluto es ρ = {r_(M['reg']['heterocedasticidad_rho'])}, "
          f"{'lo que indica que el error crece con la magnitud de la predicción y que los intervalos de predicción deben ampliarse en los extremos' if abs(M['reg']['heterocedasticidad_rho']) > .15 else 'sin evidencia fuerte de heterocedasticidad'}. "
          + (f"Por categoría, el RMSE oscila entre {n(min(r['rmse'] for r in pk))} y {n(max(r['rmse'] for r in pk))}; la categoría con mayor error es *{pk[0]['keyword']}* (*n* = {pk[0]['n']}), donde las predicciones deben usarse con más cautela." if pk else ""))
    rep.h2("Objetivo secundario: log(Likes)")
    LK = M["likes"]
    rep.tabla("Desempeño para predecir log(1 + Likes)", ["Modelo", "RMSE", "R²", "ρ de Spearman"], [[k, n(v["rmse"]), r_(v["r2"]), r_(v["rho"])] for k, v in LK.items()], [2.9, 1.2, 1.2, 1.2],
              nota="Validación cruzada repetida. Elaboración propia.")
    bl = min((k for k in LK if "Referencia" not in k), key=lambda k: LK[k]["rmse"])
    rep.p(f"Como muestra la Tabla 4, para los likes «{bl}» obtuvo R² = {r_(LK[bl]['r2'])}, {'superior' if LK[bl]['r2'] > B['r2'] else 'inferior'} al de las vistas ({r_(B['r2'])}). Que los likes sean {'más' if LK[bl]['r2'] > B['r2'] else 'menos'} predecibles es coherente con su relación directa con la reacción de la audiencia.")
    rep.h2("Clasificación de alta popularidad")
    rows = [[k, r_(v["acc"]), r_(v["bacc"]), r_(v["f1"]), f"{r_(v['auc'])} ({r_(v['auc_sd'])})", r_(v["prec"]), r_(v["sens"]), r_(v["espec"])] for k, v in sorted(clf.items(), key=lambda kv: -kv[1]["auc"])]
    rep.tabla("Desempeño de los clasificadores (validación cruzada estratificada repetida)", ["Modelo", "Exact.", "Exact. bal.", "F1", "AUC (DE)", "Prec.", "Sens.", "Espec."], rows,
              [1.9, 0.5, 0.65, 0.5, 0.85, 0.5, 0.5, 0.6], size=9, nota="Exact.: exactitud; bal.: balanceada; Prec.: precisión; Sens.: sensibilidad; Espec.: especificidad. Clase positiva: alta popularidad.")
    rep.figura("Matriz de confusión del mejor clasificador (predicciones fuera de muestra)", sal / "fig_matriz_confusion.png", 2.8, f"Modelo: {bestc}. Elaboración propia.")
    PQ = M["clf"]["prueba_vs_azar"]
    rep.p(f"Según la Tabla 5, el mejor clasificador, «{bestc}», alcanzó un AUC de {r_(BC['auc'])} y una exactitud de {r_(BC['acc'])}. La matriz de confusión (Figura 3) muestra {ent(tp)} verdaderos positivos, {ent(tn)} verdaderos negativos, "
          f"{ent(fp)} falsos positivos y {ent(fn)} falsos negativos, que equivalen a una precisión de {r_(prec)}, una sensibilidad de {r_(sens)} y una especificidad de {r_(esp)}. Su AUC "
          f"{'supera de forma significativa al azar (0,5)' if PQ['p'] < .05 else 'no supera de forma significativa al azar (0,5)'} (*t* = {n(PQ['t'])}, *p* {pv(PQ['p'])}). "
          f"Se reutilizan las mismas métricas que en Alemán Viteri (2021), pero los resultados no son comparables: allí se clasificaba el sentimiento de tuits y aquí la popularidad de videos.")
    rep.h2("Selección de variables e importancia")
    t10 = imp[:10]
    rep.figura("Importancia por permutación de las variables del mejor modelo", sal / "fig_importancia.png", 5.0, f"Aumento del RMSE al permutar cada variable. Modelo: {best}. Elaboración propia.")
    util = [r for r in imp if r["delta"] > r["sd"] * 2]
    rep.p(f"La Figura 4 presenta la importancia de las variables. Las más relevantes fueron {', '.join('*' + r['etiqueta'] + '*' for r in t10[:3])}. De {len(imp)} variables de entrada, {len(util)} "
          f"aportan una mejora superior a dos desviaciones estándar de su variabilidad; el resto puede descartarse en una versión simplificada del modelo sin pérdida apreciable, lo que reduce el costo de cálculo y de monitoreo.")

    rep.h1("¿Es suficiente un solo modelo?")
    rep.p("La evidencia indica que no. Primero, el problema tiene dos formulaciones distintas, regresión de la magnitud y clasificación de la categoría de popularidad, con métricas y usos diferentes. "
          f"Segundo, el desempeño varía por segmento (RMSE entre {n(min(r['rmse'] for r in pk)) if pk else 'n. d.'} y {n(max(r['rmse'] for r in pk)) if pk else 'n. d.'} según la categoría), por lo que un modelo global puede ser poco confiable en ciertos nichos. "
          "Tercero, la etiqueta de sentimiento del conjunto es ruidosa; mantener un módulo de sentimiento independiente (VADER y un clasificador entrenado) permite auditar y reemplazar esa señal. "
          "Cuarto, el equilibrio entre precisión e interpretabilidad favorece mantener un modelo lineal como alternativa explicable y un ensamble de árboles como candidato de mayor capacidad, bajo un esquema campeón-retador.")
    rep.p("Se propone por tanto una arquitectura de tres componentes: (a) un regresor de log(Views) con intervalo de predicción; (b) un clasificador de alta popularidad con umbral calibrado según el costo de "
          "los errores; y (c) un módulo de sentimiento que procesa los comentarios entrantes. Los tres se versionan y monitorean de forma independiente.")

    rep.h1("Modelo seleccionado y plan de despliegue")
    rep.p(f"Se selecciona «{best}» como modelo principal de regresión y «{bestc}» como clasificador. La elección se justifica por: (a) menor RMSE medio entre los {M['reg']['n_modelos']} modelos evaluados; "
          f"(b) diferencia {'significativa' if P['vs_referencia']['p'] < .05 else 'no significativa'} frente a la referencia; (c) validación sin fuga de información; y (d) variables disponibles en producción, ya que se calculan "
          "a partir de los comentarios y los metadatos públicos del video. El modelo final se reentrena con todos los datos y se guarda con *joblib* junto con un archivo de metadatos (variables de entrada, fecha, métricas de validación y semilla).")
    rep.p("Para cumplir las pautas de MLOps, la siguiente etapa empaquetará el modelo en una API (FastAPI) dentro de un contenedor, alojada en un servicio en la nube, con un registro de predicciones. El monitoreo "
          "incluirá el cambio en la distribución de las variables de entrada (índice de estabilidad poblacional), el error cuando se disponga del valor real de vistas y alertas de reentrenamiento. "
          "Estos mecanismos responden al riesgo de deuda técnica descrito por Sculley et al. (2015).")
    rep.h1("Limitaciones")
    rep.p(f"El modelo explica una fracción limitada de la variabilidad (R² = {r_(B['r2'])}); el volumen de vistas depende de factores externos a los comentarios. La evaluación usa una única fuente y una instantánea, "
          "por lo que no garantiza generalización a otras categorías, idiomas o fechas. La comparación de modelos con validación cruzada, aun corregida, es aproximada. Por último, la relación entre comentarios y vistas "
          "es asociativa y no debe interpretarse como causal.")
    rep.h1("Conclusiones")
    rep.p(f"Se evaluaron {ntot} modelos con validación rigurosa. El mejor regresor ({best}) reduce el error de la referencia en {n(mej, 1)} %, con una capacidad explicativa {nivel_r2(B['r2'])}; el mejor clasificador ({bestc}) logra un AUC de {r_(BC['auc'])}. "
          f"Los resultados son {'sólidos para un prototipo' if B['r2'] >= .2 else 'una base inicial que debe fortalecerse con más datos y variables'}, y las mejoras más prometedoras se encuentran en la calidad de la etiqueta de sentimiento, la captura de más comentarios por video y la "
          "incorporación de variables de contexto del canal. La siguiente semana se dedicará a los algoritmos del modelo y al despliegue monitoreado.")
    rep.referencias(["Alemán", "Barrios", "Breiman", "Friedman", "Hoerl", "Hutto", "Nadeau", "Pedregosa", "Salton", "Sculley"])
    rep.guardar(out)
    return rep.nw




# =============================================================================== AVANCES III
COMPL = {0: "Nula", 1: "Baja (lineal)", 2: "Baja (lineal + texto)", 3: "Media (árboles)", 4: "Media-alta", 5: "Alta (ensamble)"}
DESC = {
    "R1": "una regresión Ridge (penalización L2; Hoerl & Kennard, 1970) sobre variables estandarizadas y la categoría codificada en indicadores",
    "R2": "un modelo de refuerzo del gradiente por histogramas (Friedman, 2001) con hiperparámetros elegidos por búsqueda en malla",
    "R3": "una regresión Ridge sobre las variables numéricas, la categoría y la representación TF-IDF del texto",
    "R4": "un modelo de refuerzo del gradiente por histogramas con hiperparámetros fijos y regularización L2",
    "R5": "un modelo de refuerzo del gradiente que excluye por diseño las variables basadas en los likes de los comentarios",
    "R6": "una regresión Ridge que excluye por diseño las variables basadas en los likes de los comentarios",
    "R7": "un refuerzo del gradiente precedido de una selección de variables con Lasso (Tibshirani, 1996)",
    "R8": "un refuerzo del gradiente precedido de una selección de las 15 columnas con mayor información mutua",
    "R9": "un refuerzo del gradiente con el conjunto ampliado de variables derivadas de los comentarios",
    "R10": "una regresión Ridge con el conjunto ampliado de variables derivadas de los comentarios",
    "R11": "un refuerzo del gradiente que añade 50 componentes de análisis semántico latente del texto (Deerwester et al., 1990)",
    "R12": "un refuerzo del gradiente cuyos hiperparámetros se eligieron por búsqueda aleatoria (Bergstra & Bengio, 2012)",
    "R13": "un modelo LightGBM (Ke et al., 2017)", "R14": "un modelo XGBoost (Chen & Guestrin, 2016)",
    "R15": "el promedio simple de un refuerzo del gradiente, una regresión Ridge y el modelo híbrido con TF-IDF",
    "R16": "un apilamiento (Wolpert, 1992) de un refuerzo del gradiente, una regresión Ridge y el modelo híbrido, combinados con una regresión no negativa",
}


def build_III(R2, R3, sal, plantilla, out):
    sal = Path(sal); M2 = R2["modelado"]; L = R2["limpieza"]; nv, nc = L["videos_finales"], L["comentarios_finales"]
    reg, clf, sel, sc = R3["reg"]["modelos"], R3["clf"]["modelos"], R3["seleccion"], R3["clf"]["seleccion"]
    pid = {v["id"]: k for k, v in reg.items()}; pares = {(d["id_a"], d["id_b"]): d for d in R3["pares"]}
    eleg, ec = sel["elegido"], sc["elegido"]; E, EC = reg[eleg], clf[ec]; ref = reg[pid["R0"]]
    base = R3["pruebas"]["base_e2"]; B = reg[base]; PR = R3["pruebas"]; H = R3["entorno"]; rps = R3["config"]["repeticiones"]
    n3 = len(reg) + len(clf); n2 = M2["reg"]["n_modelos"] + len(M2["clf"]["modelos"]); ncfg = R3["reg"]["n_configs"] + R3["clf"]["n_configs"]
    mej = 100 * (1 - E["rmse"] / ref["rmse"])
    vd = lambda d: ("mejora significativa" if d["delta"] < 0 else "empeoramiento significativo") if d["p"] < .05 else "sin diferencia significativa"
    fp = lambda d: f"ΔRMSE = {n(d['delta'], 3)}; *t* = {n(d['t'])}, *p* {pv(d['p'])}"
    G = lambda a, b: pares.get((a, b))
    frase = lambda a, b: (f"{fp(G(a, b))}, es decir, {vd(G(a, b))}" if G(a, b) else "no evaluada en esta ejecución")
    REDUC = {"R5", "R6", "R7", "R8"}

    def decision(d):
        sig = d["p"] < .05
        if d["id_b"] in REDUC:
            return "Se conservan las variables" if (sig and d["delta"] > 0) else ("Se adopta la reducción" if sig else "Equivalente: se admite simplificar")
        return "Se adopta" if (sig and d["delta"] < 0) else ("Se descarta (empeora)" if sig else "Se descarta (sin mejora)")
    likes_importan = any(G(a, b) and G(a, b)["p"] < .05 and G(a, b)["delta"] > 0 for a, b in (("R4", "R5"), ("R1", "R6")))
    ens = PR["ensamble_vs_individual"]; ens_gana = ens["p"] < .05 and ens["delta"] < 0
    mejor_ind, mejor_ens = PR["mejor_individual"], PR["mejor_ensamble"]
    CV, IV = R3["clf"], R3["intervalos"]; cb = clf[CV["base"]]; c6 = clf[next(k for k, v in clf.items() if v["id"] == "C6")]

    rep = Informe(plantilla, "Estadísticas de YouTube: Avances III")
    rep.portada('Tarea Informe de trabajo “Proyecto estadísticas de Youtube Avances III”', CONFIG["fecha_3"])
    resumen = (f"Este informe documenta la tercera etapa del proyecto, centrada en la iteración controlada del modelo que estima el desempeño de un video de YouTube a partir de sus "
               f"comentarios. Partiendo de la mejor solución de la etapa II, se ejecutaron seis iteraciones (diagnóstico del efecto de los likes de los comentarios, selección de variables, "
               f"nuevas variables y análisis semántico latente, afinamiento y modelos alternativos, ensambles y calibración), cada una contrastada con la prueba t remuestreada corregida y con "
               f"la regla de una desviación estándar. En total se evaluaron {n3} modelos adicionales y {ent(ncfg)} configuraciones. El modelo seleccionado para log(Views) fue «{eleg}», con RMSE de "
               f"{n(E['rmse'])} y R² de {r_(E['r2'])}, {n(mej, 1)} % por debajo de la referencia; para la clase de alta popularidad se eligió «{ec}», con AUC de {r_(EC['auc'])} y puntaje de Brier de "
               f"{r_(EC['brier'], 3)}. Se discuten las variables elegidas mediante importancia por permutación agrupada y se concluye que, para el producto completo, un solo modelo no es suficiente: "
               f"se requieren un regresor, un clasificador calibrado y un modelo de intervalos, con cobertura empírica de {pct(IV['cobertura'])} frente al 80 % nominal.")
    rep.resumen(resumen, "iteración de modelos, selección de variables, ensambles, calibración, regresión cuantílica, MLOps")
    rep.salto()
    t = rep.doc.add_paragraph(); _runs(t, "Tarea Informe de trabajo “Proyecto estadísticas de Youtube Avances III”", bold=True); _pf(t, first=0, align=WD_ALIGN_PARAGRAPH.CENTER)

    # ------------------------------------------------------------------ Introducción
    rep.h1("Introducción")
    rep.p(f"La etapa II dejó como mejor regresor a «{M2['reg']['mejor']}», con una capacidad explicativa que el propio informe calificó como {nivel_r2(M2['reg']['modelos'][M2['reg']['mejor']]['r2'])} "
          f"(R² = {r_(M2['reg']['modelos'][M2['reg']['mejor']]['r2'])}), y a «{M2['clf']['mejor']}» como mejor clasificador de alta popularidad (AUC = {r_(M2['clf']['modelos'][M2['clf']['mejor']]['auc'])}). "
          f"Esos resultados abren preguntas que la tercera etapa debe resolver: si la señal proviene de la reacción del público o del tamaño de la audiencia, qué variables son realmente necesarias, si "
          f"existen mejoras reales frente a la solución inicial y si un solo modelo basta para el producto final.")
    rep.p("Esta etapa trata el desarrollo del modelo como un proceso iterativo y medible, no como una búsqueda del mayor número en una tabla. Cada iteración parte de una hipótesis, modifica una sola cosa "
          "respecto de un modelo de comparación, se evalúa con las mismas particiones de validación y se acepta o descarta con un criterio estadístico explícito. Así, cada mejora queda justificada y "
          "cada descarte también. El informe responde de forma directa a las preguntas de la asignación: ¿cuál fue el modelo seleccionado?, ¿cuál fue su desempeño y bajo qué métricas?, "
          "¿cuántos modelos se probaron?, ¿qué herramientas se usaron para determinarlo?, ¿es suficiente un solo modelo? y ¿qué variables se seleccionaron para predecir?")
    rep.p(f"El documento presenta primero el marco de la iteración y las herramientas probadas, luego las seis iteraciones con su evidencia, los resultados comparativos de regresión y clasificación, "
          f"la selección y descripción del modelo, las variables elegidas, la discusión sobre la suficiencia de un solo modelo, las amenazas a la validez y el plan de la semana 10. Se trabaja con "
          f"{ent(nv)} videos y {ent(nc)} comentarios depurados del conjunto de Patil (s. f.), los mismos de la etapa II.")

    # ------------------------------------------------------------------ Marco
    rep.h1("Marco metodológico de la iteración")
    rep.h2("Principios de diseño")
    rep.h3("Una sola variación por iteración", "Para atribuir un cambio de desempeño a una causa concreta, cada iteración se compara contra un modelo idéntico salvo por esa modificación (ablación controlada).")
    rep.h3("Mismas particiones", f"Todos los modelos de esta etapa se evalúan con validación cruzada de 5 particiones repetida {rps} {'vez' if rps == 1 else 'veces'} y semilla fija, de modo que las diferencias se calculan por pares de particiones y no sobre promedios aislados.")
    rep.h3("Criterio estadístico de decisión", "Una diferencia se considera real solo si la prueba *t* remuestreada corregida de Nadeau y Bengio (2003) arroja *p* < ,05. Cuando no hay diferencia significativa, se prefiere la opción más simple.")
    rep.h3("Regla de una desviación estándar", "El modelo final se elige entre los que quedan dentro de un error estándar del mejor RMSE, escogiendo el de menor complejidad (Hastie et al., 2009). Este criterio reduce el riesgo de elegir un modelo complejo por una ventaja que puede ser ruido.")
    rep.h3("Sin fuga de información", "Imputación, escalado, codificación, TF-IDF, selección de variables y reducción dimensional están dentro de *pipelines* de scikit-learn (Pedregosa et al., 2011) y se ajustan solo con los datos de entrenamiento de cada partición.")
    rep.h2("Herramientas probadas")
    lg = H.get("lightgbm"); xg = H.get("xgboost")
    rows = [
        ["*Pipeline* y *ColumnTransformer* (scikit-learn)", "Evitar fuga y asegurar reproducibilidad", "Todas", "Adoptada como base"],
        ["GridSearchCV", "Afinamiento por malla (etapa II)", "I0", "Sustituida por búsqueda aleatoria por costo"],
        ["RandomizedSearchCV (Bergstra & Bengio, 2012)", "Afinamiento de hiperparámetros", "I4", vd(G("R9", "R12")).capitalize() if G("R9", "R12") else "—"],
        ["LassoCV + SelectFromModel (Tibshirani, 1996)", "Selección de variables", "I2", (f"Conserva {R3['seleccion_vars']['lasso']['n_sel']} de {R3['seleccion_vars']['lasso']['n_entrada']} columnas; {vd(G('R4', 'R7'))}" if G("R4", "R7") else "—")],
        ["SelectKBest con información mutua (Guyon & Elisseeff, 2003)", "Selección de variables", "I2", (f"Conserva {R3['seleccion_vars']['mi']['n_sel']} columnas; {vd(G('R4', 'R8'))}" if G("R4", "R8") else "—")],
        ["TruncatedSVD / LSA (Deerwester et al., 1990)", "Representación semántica del texto", "I3", vd(G("R9", "R11")).capitalize() if G("R9", "R11") else "—"],
        ["Importancia por permutación (Breiman, 2001), individual y agrupada", "Interpretar y seleccionar variables", "I1–I6", "Adoptada"],
        ["Promedio y apilamiento (Wolpert, 1992), implementación propia", "Combinar modelos", "I5", f"Mejor ensamble frente al mejor individual: {vd(ens)}"],
        ["LightGBM (Ke et al., 2017)", "Boosting alternativo", "I4", (vd(G("R9", "R13")).capitalize() if G("R9", "R13") else f"No disponible en el entorno ({'no instalado' if not lg else lg})")],
        ["XGBoost (Chen & Guestrin, 2016)", "Boosting alternativo", "I4", (vd(G("R9", "R14")).capitalize() if G("R9", "R14") else f"No disponible en el entorno ({'no instalado' if not xg else xg})")],
        ["Regresión cuantílica (Koenker & Bassett, 1978)", "Intervalos de predicción", "I6", f"Cobertura {pct(IV['cobertura'])} (nominal 80 %)"],
        ["CalibratedClassifierCV; Brier (1950) y ECE", "Calibrar probabilidades", "I6", f"Brier {r_(cb['brier'], 3)} → {r_(c6['brier'], 3)} con calibración isotónica (Niculescu-Mizil & Caruana, 2005)"],
        ["Prueba *t* corregida y regla de 1 error estándar", "Decidir entre modelos", "Todas", "Adoptada"],
    ]
    rep.p(f"La Tabla {rep.nt + 1} lista las herramientas evaluadas, el propósito de cada una, la iteración donde se usó y el resultado, de modo que la elección final no dependa de una sola técnica. "
          f"Las librerías opcionales (LightGBM y XGBoost) se incorporan solo si están instaladas, y el informe registra cuáles se ejecutaron.")
    rep.tabla("Herramientas probadas y resultado de su evaluación", ["Herramienta", "Propósito", "Iteración", "Resultado"], rows, [2.3, 1.7, 0.8, 1.7], size=9, alin=["l", "l", "c", "l"],
              nota=f"Entorno: Python {H['python']}, scikit-learn {H['scikit-learn']}. El resultado resume la comparación pareada de cada herramienta frente a su modelo de comparación.")

    # ------------------------------------------------------------------ Iteraciones
    rep.h1("Iteraciones del modelo")
    filas = [[d["it"], d["desc"], n(d["delta"], 3), ("< ,001" if d["p"] < .001 else r_(d["p"], 3)), decision(d)] for d in R3["pares"]]
    ti = rep.nt + 1; fi = rep.nf + 1
    rep.p(f"La Tabla {ti} resume la evidencia de cada iteración y la Figura {fi} muestra el RMSE de todos los modelos probados; la línea punteada es el umbral de una desviación estándar sobre el mejor valor. "
          f"Una diferencia negativa indica que el modelo nuevo reduce el error respecto de su comparación.")
    rep.tabla("Resumen de iteraciones: cambio evaluado, diferencia de RMSE y decisión", ["It.", "Cambio evaluado", "ΔRMSE", "*p*", "Decisión"], filas, [0.4, 2.6, 0.8, 0.7, 2.0], size=9, alin=["c", "l", "c", "c", "l"],
              nota="ΔRMSE = RMSE del modelo nuevo menos RMSE del modelo de comparación (unidades de log(1 + Views)). *p*: prueba *t* remuestreada corregida.")
    rep.figura("RMSE de todos los modelos evaluados por iteración", sal / "fig3_iteraciones.png", 6.2,
               "Colores: iteración (I0 = línea base; I1 = diagnóstico de likes; I2 = selección; I3 = variables nuevas; I4 = afinamiento; I5 = ensambles). Barras de error: desviación estándar entre particiones.")

    rep.h2("Iteración 0: línea base de la etapa II")
    rep.p(f"Se reevaluaron los modelos clave de la etapa II con el protocolo actual. Ridge – B (Hoerl & Kennard, 1970) obtuvo RMSE = {n(reg[pid['R1']]['rmse'])}, el refuerzo del gradiente con malla (Friedman, 2001), {n(reg[pid['R2']]['rmse'])}, y el híbrido con TF-IDF, "
          f"{n(reg[pid['R3']]['rmse'])}, frente a {n(ref['rmse'])} de la referencia. El mejor de ellos, «{base}» (RMSE = {n(B['rmse'])}; R² = {r_(B['r2'])}), es la base contra la que se mide toda mejora. "
          f"En la etapa II se habían evaluado {n2} modelos, de modo que, con los {n3} de esta etapa, el proyecto acumula {n2 + n3} modelos probados.")
    rep.h2("Iteración 1: ¿cuánto de la señal proviene del tamaño de la audiencia?")
    rep.p(f"Hipótesis: los likes de los comentarios más destacados crecen con la audiencia, por lo que las variables derivadas de ellos (likes medios, likes máximos y sentimiento ponderado) podrían estar actuando como un "
          f"indicador indirecto de las vistas y no como una reacción independiente del público. Cambio: se eliminaron esas tres variables y se compararon los mismos modelos. Resultado: con refuerzo del gradiente, "
          f"{frase('R4', 'R5')}; con Ridge, {frase('R1', 'R6')}. En clasificación, la regresión logística sin esas variables obtuvo AUC = {r_(clf[next(k for k, v in clf.items() if v['id'] == 'C3')]['auc'])} frente a "
          f"{r_(clf[next(k for k, v in clf.items() if v['id'] == 'C1')]['auc'])} con ellas. "
          + ("Conclusión: las variables basadas en likes sí aportan información, pero el modelo conserva una parte sustancial de su capacidad sin ellas, de modo que la señal no se reduce a ese proxy; se mantienen en el modelo y se declara la dependencia como limitación."
             if likes_importan else
             "Conclusión: quitar esas variables no empeora de forma significativa el desempeño, es decir, el modelo no depende de ese posible proxy de audiencia, lo que refuerza la validez del resultado y permite un modelo más sencillo."))
    rep.h2("Iteración 2: selección de variables")
    sv = R3["seleccion_vars"]
    rep.p(f"Hipótesis: parte de las {len(reg[pid['R4']]['cols'])} variables de entrada es redundante. Cambio: se probaron dos selectores incorporados en el *pipeline*: Lasso, que conserva {sv['lasso']['n_sel']} de "
          f"{sv['lasso']['n_entrada']} columnas tras la codificación, y la información mutua con las 15 columnas de mayor puntaje. Resultado: Lasso, {frase('R4', 'R7')}; información mutua, {frase('R4', 'R8')}. "
          f"Entre las columnas que conserva Lasso están {', '.join('*' + ETIQ_X(x) + '*' for x in sv['lasso']['sel'][:6])}. "
          "Cuando la reducción no empeora el error, se obtiene un modelo más simple, barato de monitorear y menos propenso a sobreajuste (Guyon & Elisseeff, 2003).")
    rep.h2("Iteración 3: ingeniería de variables y análisis semántico latente")
    rep.p(f"Hipótesis: la distribución del sentimiento dentro de los comentarios contiene más información que su promedio. Cambio: se añadieron seis variables (VADER máximo y mínimo, proporciones de comentarios muy positivos y "
          f"muy negativos, diversidad léxica y proporción de emojis) y, por separado, 50 componentes de análisis semántico latente del texto. Resultado: con refuerzo del gradiente, {frase('R4', 'R9')}; con Ridge, "
          f"{frase('R1', 'R10')}; la adición del texto latente, {frase('R9', 'R11')}. Frente al refuerzo del gradiente de la etapa II, el modelo ampliado muestra {frase('R2', 'R9')}.")
    rep.h2("Iteración 4: afinamiento de hiperparámetros y boosting alternativos")
    rep.p(f"Hipótesis: los hiperparámetros fijos limitan el modelo. Cambio: búsqueda aleatoria de {R3['config']['iter_busqueda']} configuraciones sobre tasa de aprendizaje, profundidad, hojas máximas, "
          f"mínimo de muestras por hoja y regularización L2, con validación interna de {R3['config']['cv_interna']} particiones y parada temprana. Resultado: {frase('R9', 'R12')}. "
          f"LightGBM: {frase('R9', 'R13') if G('R9', 'R13') else 'no se ejecutó porque la librería no estaba instalada en este entorno'}; XGBoost: {frase('R9', 'R14') if G('R9', 'R14') else 'no se ejecutó porque la librería no estaba instalada en este entorno'}.")
    rep.h2("Iteración 5: ensambles")
    rep.p(f"Hipótesis: modelos de familias distintas cometen errores distintos y combinarlos reduce el error. Cambio: promedio simple y apilamiento de un refuerzo del gradiente, una regresión Ridge y el híbrido con TF-IDF; "
          f"en el apilamiento, un metamodelo de regresión no negativa aprende los pesos con predicciones fuera de muestra. Resultado: promedio, {frase('R9', 'R15')}; apilamiento, {frase('R9', 'R16')}. "
          f"El mejor ensamble («{mejor_ens}») frente al mejor modelo individual («{mejor_ind}») muestra {fp(ens)}, es decir, {vd(ens)}.")
    rep.h2("Iteración 6: incertidumbre y calibración")
    rep.p(f"Hipótesis: un producto útil no solo entrega un valor sino cuánta confianza merece. Cambio: se entrenaron modelos de regresión cuantílica (percentiles 10 y 90; Koenker & Bassett, 1978) para construir intervalos "
          f"de predicción del 80 % y se calibraron las probabilidades del clasificador con regresión isotónica. Resultado: la cobertura empírica de los intervalos fue {pct(IV['cobertura'])} (desviación estándar entre "
          f"particiones {pct(IV['cobertura_sd'])}) con un ancho medio de {n(IV['ancho'])} unidades logarítmicas, es decir, un factor multiplicativo de {n(2.718281828 ** IV['ancho'])}; "
          f"{'la cobertura está cerca del valor nominal' if abs(IV['cobertura'] - .8) < .05 else 'la cobertura se aleja del valor nominal, por lo que los intervalos deben recalibrarse antes de usarse en producción'}. "
          f"En clasificación, el puntaje de Brier (Brier, 1950) pasó de {r_(cb['brier'], 3)} a {r_(c6['brier'], 3)} y el error de calibración esperado, de {r_(cb['ece'], 3)} a {r_(c6['ece'], 3)} con la calibración isotónica.")

    # ------------------------------------------------------------------ Resultados
    rep.h1("Resultados comparativos")
    rep.h2("Regresión sobre log(Views)")
    ids = sorted(reg, key=lambda k: int(reg[k]["id"][1:]))
    rows = [[reg[k]["id"], k, reg[k]["it"], f"{n(reg[k]['rmse'])} ({n(reg[k]['rmse_sd'])})", n(reg[k]["mae"]), r_(reg[k]["r2"]), r_(reg[k]["rho"]), str(reg[k]["configs"])] for k in ids]
    tm = rep.nt + 1
    rep.p(f"La Tabla {tm} reúne las métricas de los {len(reg)} modelos de regresión de esta etapa: RMSE, error absoluto medio (MAE), R², correlación de Spearman entre valores reales y predichos y número de "
          f"configuraciones exploradas. El menor RMSE lo obtuvo «{sel['mejor']}» ({n(reg[sel['mejor']]['rmse'])}); sin embargo, {len(sel['candidatos'])} modelos quedan dentro de un error estándar ({n(sel['se'], 3)}) de ese valor, "
          f"por lo que son estadísticamente indistinguibles en la práctica.")
    rep.tabla("Desempeño de los modelos de regresión (validación cruzada repetida)", ["ID", "Modelo", "It.", "RMSE (DE)", "MAE", "R²", "ρ", "Config."], rows, [0.4, 2.7, 0.4, 0.95, 0.5, 0.5, 0.5, 0.55], size=8,
              alin=["c", "l", "c", "c", "c", "c", "c", "c"], nota="RMSE y MAE en unidades de log(1 + Views). DE: desviación estándar entre particiones. Config.: configuraciones exploradas.")
    rep.figura("Distribución del RMSE por partición de los modelos con mejor desempeño y la referencia", sal / "fig3_boxplot.png", 5.8, "Elaboración propia.")
    rep.p(f"El modelo seleccionado reduce el RMSE de la referencia en {n(mej, 1)} % ({fp(PR['elegido_vs_ref'])}) y, frente a la base de la etapa II, muestra {fp(PR['elegido_vs_base'])}, es decir, {vd(PR['elegido_vs_base'])}.")
    rep.h2("Clasificación de alta popularidad y calibración")
    crow = [[v["id"], k, r_(v["acc"]), r_(v["f1"]), f"{r_(v['auc'])} ({r_(v['auc_sd'])})", r_(v["prec"]), r_(v["sens"]), r_(v["espec"]), r_(v["brier"], 3), r_(v["ece"], 3)]
            for k, v in sorted(clf.items(), key=lambda kv: int(kv[1]["id"][1:]))]
    tc = rep.nt + 1
    rep.p(f"La Tabla {tc} presenta los {len(clf)} clasificadores. Además de la exactitud, el F1, el AUC, la precisión, la sensibilidad y la especificidad (Barrios Arce, 2019; Alemán Viteri, 2021), se reportan el puntaje de Brier y "
          f"el error de calibración esperado (ECE), que miden si las probabilidades son confiables y no solo si la clase es correcta.")
    rep.tabla("Desempeño de los clasificadores de alta popularidad", ["ID", "Modelo", "Exact.", "F1", "AUC (DE)", "Prec.", "Sens.", "Espec.", "Brier", "ECE"], crow, [0.35, 2.0, 0.5, 0.45, 0.8, 0.45, 0.45, 0.5, 0.5, 0.5], size=8,
              alin=["c", "l", "c", "c", "c", "c", "c", "c", "c", "c"], nota="Clase positiva: alta popularidad (sobre la mediana de log Views). Exact.: exactitud; Prec.: precisión; Sens.: sensibilidad; Espec.: especificidad; ECE: error de calibración esperado.")
    (tn, fp_), (fn, tp) = CV["cm"]
    rep.p(f"El clasificador elegido por la regla de una desviación estándar fue «{ec}» (AUC = {r_(EC['auc'])}; exactitud = {r_(EC['acc'])}; sensibilidad = {r_(EC['sens'])}; especificidad = {r_(EC['espec'])}). "
          f"Frente al modelo base de la etapa II, la diferencia de AUC es {n(CV['prueba_vs_base']['delta_auc'], 3)} (*p* {pv(CV['prueba_vs_base']['p'])}). "
          f"Su matriz de confusión fuera de muestra registra {ent(tp)} verdaderos positivos, {ent(tn)} verdaderos negativos, {ent(fp_)} falsos positivos y {ent(fn)} falsos negativos. "
          f"La Figura {rep.nf + 1} compara la calibración del modelo base y la del elegido.")
    rep.figura("Curva de calibración del clasificador base y del seleccionado", sal / "fig3_calibracion.png", 3.8, "La diagonal representa una calibración perfecta. Elaboración propia.")

    # ------------------------------------------------------------------ Selección y descripción
    rep.h1("Selección y descripción del modelo")
    cand = sorted(sel["candidatos"], key=lambda k: (reg[k]["compl"], reg[k]["rmse"]))
    rows = [[reg[k]["id"], k, f"{n(reg[k]['rmse'])}", COMPL[reg[k]["compl"]], "Sí" if k == eleg else ""] for k in cand]
    rep.p(f"El modelo se eligió con un procedimiento definido de antemano: (a) se identificó el mejor RMSE medio ({n(reg[sel['mejor']]['rmse'])}, «{sel['mejor']}»); (b) se calculó el umbral de un error estándar "
          f"({n(sel['umbral'])}); (c) se listaron los modelos por debajo del umbral (Tabla {rep.nt + 1}); y (d) se escogió el de menor complejidad. Este procedimiento evita premiar diferencias que son ruido de validación.")
    rep.tabla("Candidatos dentro de un error estándar del mejor RMSE", ["ID", "Modelo", "RMSE", "Complejidad", "Elegido"], rows, [0.5, 3.4, 0.7, 1.4, 0.6], size=9, alin=["c", "l", "c", "l", "c"],
              nota=f"Umbral: {n(sel['umbral'])} = mejor RMSE + 1 error estándar.")
    fin = R3["final"]
    descr = DESC.get(E["id"], "un modelo de regresión")
    extra = ""
    if "hiperparametros" in fin: extra += " Hiperparámetros elegidos: " + ", ".join(f"{k.replace('m__', '')} = {n(v, 3) if isinstance(v, float) else v}" for k, v in fin["hiperparametros"].items()) + "."
    if "pesos_apilado" in fin: extra += " Pesos del apilamiento (HGB, Ridge, híbrido): " + ", ".join(n(w, 3) for w in fin["pesos_apilado"]) + "."
    rep.p(f"El modelo seleccionado para estimar log(Views) es «{eleg}» ({E['id']}), {descr}. Utiliza {E['n_vars']} variables de entrada, entra a un *pipeline* con imputación por la mediana, estandarización "
          f"y codificación de la categoría, y se evaluó con RMSE = {n(E['rmse'])}, MAE = {n(E['mae'])}, R² = {r_(E['r2'])} y ρ = {r_(E['rho'])}.{extra}")
    rep.p(f"La elección se justifica por cuatro razones: desempeño estadísticamente equivalente al mejor modelo, menor complejidad entre los equivalentes, ausencia de fuga de información y disponibilidad en producción de todas sus "
          f"variables (se calculan a partir de los comentarios y metadatos públicos). Para clasificar la alta popularidad se selecciona «{ec}». El modelo final se reentrenó con todos los datos y se guardó junto con un "
          f"archivo de metadatos (variables de entrada, fecha, métricas y versiones de las librerías) para su auditoría y despliegue.")

    # ------------------------------------------------------------------ Variables
    rep.h1("Variables seleccionadas para predecir")
    g = R3["importancia_grupos"]
    rows = [[r["etiqueta"], ", ".join(ETIQ_X(v_) for v_ in r["vars"][:5]) + ("…" if len(r["vars"]) > 5 else ""), f"{n(r['delta'], 3)} ({n(r['sd'], 3)})"] for r in g]
    rep.p(f"El modelo elegido usa las variables de entrada que muestra la Tabla {rep.nt + 1}, organizadas en grupos. Para valorar cada grupo se permutaron conjuntamente sus columnas en un conjunto de prueba y se midió el aumento del "
          f"RMSE (importancia por permutación agrupada), lo que evita repartir de forma engañosa la importancia entre variables correlacionadas, como el sentimiento medio y las proporciones de clases.")
    rep.tabla("Grupos de variables del modelo y su importancia por permutación", ["Grupo", "Variables (primeras cinco)", "Aumento de RMSE (DE)"], rows, [2.0, 3.2, 1.3], size=9, alin=["l", "l", "c"],
              nota=f"RMSE de referencia en el conjunto de prueba: {n(R3['rmse_test_importancia'])}. 25 repeticiones de permutación por grupo; 20 en la medición individual.")
    rep.figura("Importancia por permutación agrupada", sal / "fig3_grupos.png", 5.0, "Elaboración propia.")
    iv = R3["importancia_vars"][:8]
    rep.p(f"El grupo con mayor efecto es *{g[0]['etiqueta']}* (aumento de RMSE = {n(g[0]['delta'], 3)}), seguido de *{g[1]['etiqueta']}*" + (f" y *{g[2]['etiqueta']}*" if len(g) > 2 else "") +
          f". En el nivel individual, las variables más relevantes son {', '.join('*' + ETIQ_X(r['grupo']) + '*' for r in iv[:4])}. "
          f"Las variables cuyo efecto no supera la variabilidad de la permutación se pueden retirar en versiones simplificadas sin pérdida apreciable, lo que se corrobora con la Iteración 2.")

    # ------------------------------------------------------------------ Suficiencia de un solo modelo
    rep.h1("¿Es suficiente un solo modelo?")
    pk = sorted(R3["por_keyword"], key=lambda r: -r["rmse"])
    rep.p(f"La respuesta tiene matices y se apoya en cuatro evidencias. Primera, para la estimación puntual de log(Views), el mejor ensamble frente al mejor modelo individual arroja {fp(ens)}: "
          f"{'el ensamble sí aporta una mejora real, aunque a costa de mayor complejidad, por lo que se justifica una arquitectura con más de un modelo' if ens_gana else 'no hay mejora significativa al combinar modelos, de modo que, para la estimación puntual, un solo modelo sencillo es suficiente y los ensambles no compensan su costo de mantenimiento'}.")
    rep.p(f"Segunda, el producto requiere tareas distintas: estimar la magnitud (regresión, «{eleg}»), clasificar la popularidad (clasificación, «{ec}») y expresar incertidumbre (regresión cuantílica, con cobertura de "
          f"{pct(IV['cobertura'])} frente al 80 % nominal). Un solo modelo no cubre estas salidas con métricas y calibración apropiadas para cada una.")
    tk = rep.nt + 1
    rep.p(f"Tercera, el error es heterogéneo entre categorías (Tabla {tk}): el RMSE va de {n(min(r['rmse'] for r in pk))} a {n(max(r['rmse'] for r in pk))}, lo que indica que un modelo global es menos confiable en ciertos nichos "
          f"y que conviene monitorear el error por segmento y evaluar modelos específicos si la brecha persiste. La correlación entre el valor predicho y el error absoluto es ρ = {r_(R3['heterocedasticidad_rho'])}"
          f"{', señal de heterocedasticidad que justifica los intervalos de predicción' if abs(R3['heterocedasticidad_rho']) > .15 else ''}.")
    rep.tabla("Error del modelo seleccionado por categoría (cinco mayores)", ["Categoría", "Videos", "RMSE"], [[r["keyword"], str(r["n"]), n(r["rmse"])] for r in pk[:5]], [3.0, 1.2, 1.2], alin=["l", "c", "c"],
              nota="Categorías con al menos 15 videos; predicciones fuera de muestra.")
    rep.p("Cuarta, un criterio de operación: en producción los modelos se degradan por cambios en los datos. Mantener un regresor principal, un clasificador calibrado y un modelo de intervalos, versionados y monitoreados "
          "por separado, permite detectar y corregir cada degradación sin reemplazar todo el sistema. En conclusión, un único modelo no es suficiente para el producto completo, aunque sí lo es, "
          f"{'solo con ciertas salvedades,' if ens_gana else 'con holgura,'} para la estimación puntual de la popularidad.")

    # ------------------------------------------------------------------ Amenazas y plan
    rep.h1("Amenazas a la validez y limitaciones")
    rep.p(f"Primero, la selección entre {len(reg)} modelos sobre las mismas particiones introduce un optimismo leve en el desempeño reportado; se mitigó con la regla de una desviación estándar y con pruebas "
          f"pareadas corregidas, pero la estimación definitiva requiere un conjunto de prueba retenido que no se haya usado en ninguna decisión. Segundo, " +
          ("las variables basadas en likes de comentarios aportan información y podrían reflejar en parte el tamaño de la audiencia (Iteración 1). " if likes_importan else "el modelo no depende de manera significativa de las variables basadas en likes de comentarios (Iteración 1). ") +
          "Tercero, los datos son una única instantánea y los comentarios son los más destacados, no una muestra aleatoria. Cuarto, la etiqueta de sentimiento del conjunto tiene concordancia limitada con VADER, que además está calibrado para inglés. "
          "Quinto, las relaciones son asociativas y no causales.")
    rep.h1("Plan para la semana 10")
    rep.p("La última etapa se enfocará en el producto: (a) evaluación final sobre un conjunto de prueba retenido; (b) empaquetado del modelo seleccionado en una API (FastAPI) dentro de un contenedor y despliegue en un servicio "
          "en la nube; (c) registro de predicciones y monitoreo del cambio en la distribución de las variables (índice de estabilidad poblacional), del error y de la cobertura de los intervalos; (d) una prueba en un "
          "entorno real con videos nuevos; y (e) el análisis de retos, cambios respecto del diseño inicial y lecciones aprendidas que pide la asignación.")
    rep.h1("Conclusiones")
    rep.p(f"La iteración controlada permitió justificar cada decisión con evidencia: se probaron {n3} modelos en esta etapa ({n2 + n3} en total) y {ent(ncfg)} configuraciones, con pruebas pareadas y una regla de parsimonia. "
          f"El modelo seleccionado para log(Views), «{eleg}», alcanza RMSE = {n(E['rmse'])} y R² = {r_(E['r2'])}, y para la alta popularidad se eligió «{ec}» con AUC = {r_(EC['auc'])}. "
          f"El diagnóstico de los likes de los comentarios, la selección de variables y la comparación de ensambles muestran qué componentes del modelo son realmente necesarios. Un solo modelo no es suficiente para "
          f"el producto completo, que debe ofrecer estimación, clasificación calibrada e incertidumbre; la semana 10 consolidará estos componentes en un servicio desplegado y monitoreado.")
    rep.referencias(["Alemán", "Barrios", "Bergstra", "Brier", "Breiman", "Chen, T.", "Deerwester", "Friedman", "Guyon", "Hastie", "Hoerl", "Ke, G.", "Koenker", "Nadeau", "Niculescu", "Patil", "Pedregosa", "Tibshirani", "Wolpert"])
    rep.guardar(out)
    return rep.nw


def ETIQ_X(v):
    """Etiqueta legible de una variable (acepta nombres de columnas codificadas, p. ej. Keyword_tech)."""
    et = {"sent_mean": "Sentimiento medio", "sent_std": "Dispersión del sentimiento", "pos_share": "% positivos", "neg_share": "% negativos", "neu_share": "% neutros",
          "sent_w": "Sentimiento ponderado por likes", "clike_mean": "Likes medios del comentario", "clike_max": "Likes máx. del comentario", "n_com": "N.º de comentarios",
          "len_mean": "Longitud media", "words_mean": "Palabras medias", "vader_mean": "VADER medio", "vader_std": "VADER dispersión", "q_share": "% con «?»",
          "excl_share": "% con «!»", "upper_share": "% mayúsculas", "age_days": "Antigüedad", "title_len": "Longitud del título", "title_vader": "VADER del título",
          "title_q": "Título con «?»", "Keyword": "Categoría", "text": "Texto (TF-IDF/LSA)", "vader_max": "VADER máximo", "vader_min": "VADER mínimo",
          "strong_pos": "% muy positivos", "strong_neg": "% muy negativos", "lex_div": "Diversidad léxica", "emoji_share": "% con emojis"}
    if v in et: return et[v]
    if v.startswith("Keyword_"): return "Categoría = " + v.split("_", 1)[1]
    return v


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--plantilla", default="Plantilla_APA7.docx"); ap.add_argument("--salida", default="salida")
    ap.add_argument("--informe", default="todos", choices=["1", "2", "3", "todos"])
    a = ap.parse_args(); sal = Path(a.salida)
    r12, r3 = sal / "results.json", sal / "results_iii.json"
    R = cargar(sal) if r12.exists() else None; R3 = json.load(open(r3, encoding="utf-8")) if r3.exists() else None
    hechos = []
    if a.informe in ("1", "todos") and R:
        o = sal / "Avances_I_YouTube_APA7.docx"; w = build_I(R, sal, a.plantilla, o); hechos.append((o, w))
    if a.informe in ("2", "todos") and R:
        o = sal / "Avances_II_YouTube_APA7.docx"; w = build_II(R, sal, a.plantilla, o); hechos.append((o, w))
    if a.informe in ("3", "todos") and R and R3:
        o = sal / "Avances_III_YouTube_APA7.docx"; w = build_III(R, R3, sal, a.plantilla, o); hechos.append((o, w))
    elif a.informe == "3":
        print("Falta results.json o results_iii.json: ejecute analisis.py e iteracion.py primero.")
    for o, w in hechos: print(f"{o.name}: resumen de {w} palabras (APA: 150–250)")


if __name__ == "__main__":
    main()