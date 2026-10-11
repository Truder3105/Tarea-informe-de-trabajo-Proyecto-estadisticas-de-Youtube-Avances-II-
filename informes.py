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
    facultad="Inteligencia Artificial, Uninversidad de Cundinamarca",
    curso="CAD2202023206: Procesamiento de Lenguaje Natural - EIAIIPA2026_2",
    profesor="PAUL ALEXANDER DIAZ MONTAÑA",
    fecha_1="domingo, 11 de octubre de 2026",
    fecha_2="domingo, 11 de octubre de 2026",
)
# =============================================================================

REFS = [
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
        self._caption("Figura", titulo)
        para = self.doc.add_paragraph(); para.add_run().add_picture(str(ruta), width=Inches(ancho))
        _pf(para, line=1.0, first=0, align=WD_ALIGN_PARAGRAPH.CENTER, after=4)
        if nota: self.nota(nota)
        else: self.doc.paragraphs[-1].paragraph_format.space_after = Pt(10)

    def tabla(self, titulo, cab, filas, anchos=None, nota=None, size=10, alin=None):
        self._caption("Tabla", titulo)
        t = self.doc.add_table(rows=1 + len(filas), cols=len(cab)); t.alignment = WD_TABLE_ALIGNMENT.CENTER
        t.autofit = False
        tblPr = t._tbl.tblPr
        for tag in ("w:tblBorders",):
            for e in tblPr.findall(qn(tag)): tblPr.remove(e)
        anchos = anchos or [6.5 / len(cab)] * len(cab)
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


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--plantilla", default="Plantilla_APA7.docx"); ap.add_argument("--salida", default="salida")
    a = ap.parse_args(); R = cargar(a.salida)
    o1, o2 = Path(a.salida) / "Avances_I_YouTube_APA7.docx", Path(a.salida) / "Avances_II_YouTube_APA7.docx"
    w1 = build_I(R, a.salida, a.plantilla, o1); w2 = build_II(R, a.salida, a.plantilla, o2)
    print(f"Resumen I: {w1} palabras; Resumen II: {w2} palabras (APA: 150–250)")
    print("Generados:", o1, o2)


if __name__ == "__main__":
    main()
