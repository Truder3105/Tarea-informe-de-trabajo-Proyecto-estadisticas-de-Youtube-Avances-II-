# Proyecto: Análisis y Predicción de Estadísticas de YouTube (Avances I, II y III)

Este repositorio ([Tarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-)) contiene el código, los datos y los artefactos del proyecto de analítica avanzada y aprendizaje automático sobre el conjunto de datos de YouTube (Patil, s. f.). El flujo de trabajo abarca desde la exploración y línea base (**Etapas I y II**) hasta el refinamiento iterativo, ingeniería de características, selección de variables, optimización de hiperparámetros, ensambles, calibración de probabilidades e intervalos de predicción (**[Etapa III / Commit e556a7b](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fcommit%2Fe556a7b560a3f51340a81e1e52518d0994141c56)**), junto con la generación automatizada de informes académicos en formato APA 7.

## Estructura del Repositorio

| Ruta / Archivo | Descripción | Etapa |
| :--- | :--- | :--- |
| [LEEME.md](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2FLEEME.md) | Guía rápida de ejecución del pipeline | General |
| [README.md](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2FREADME.md) | Documentación técnica principal del proyecto | General |
| [requirements.txt](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Frequirements.txt) | Dependencias base del entorno Python | General |
| [Plantilla\_APA7.docx](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2FPlantilla_APA7.docx) | Plantilla institucional base en Word con estilos APA 7 | General |
| [analisis.py](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fanalisis.py) | Pipeline de exploración, preprocesamiento, modelado base y gráficos | Etapas I y II |
| [iteracion.py](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fiteracion.py) | Pipeline de experimentación iterativa (`I0`–`I6`), ablación, selección 1-SE y calibración | Etapa III |
| [informes.py](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Finformes.py) | Generador automatizado de los informes de Avances I, II y III en `.docx` | Etapas I, II y III |
| [data/videos-stats.csv](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fdata%2Fvideos-stats.csv) | Métricas de videos de YouTube (1,881 registros brutos; 1,867 válidos) | Datos fuente |
| [data/comments.csv](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fdata%2Fcomments.csv) | Comentarios asociados a los videos (18,409 registros) | Datos fuente |
| [salida/modelo\_final.joblib](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Fmodelo_final.joblib) | Pipeline serializado del modelo seleccionado en la Etapa II | Etapa II |
| [salida/modelo\_metadata.json](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Fmodelo_metadata.json) | Metadatos e hiperparámetros del modelo de la Etapa II | Etapa II |
| [salida/results.json](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Fresults.json) | Resultados numéricos completos de las Etapas I y II | Etapas I y II |
| [salida/modelo\_final\_iii.joblib](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Fmodelo_final_iii.joblib) | Pipeline serializado del modelo ganador de la Etapa III (`XGBoost + nuevas`) | Etapa III |
| [salida/modelo\_iii\_metadata.json](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Fmodelo_iii_metadata.json) | Metadatos, predictores (27 variables) y métricas CV del modelo de la Etapa III | Etapa III |
| [salida/results\_iii.json](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Fresults_iii.json) | Resultados experimentales de las iteraciones `I0`–`I6` (17 regresores y 8 clasificadores) | Etapa III |
| [salida/Avances\_I\_YouTube\_APA7.docx](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2FAvances_I_YouTube_APA7.docx) | Informe académico del Avance I en formato Word (APA 7) | Etapa I |
| [salida/Avances\_II\_YouTube\_APA7.docx](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2FAvances_II_YouTube_APA7.docx) | Informe académico del Avance II en formato Word (APA 7) | Etapa II |
| [salida/Avances\_III\_YouTube\_APA7.docx](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2FAvances_III_YouTube_APA7.docx) | Informe académico del Avance III en formato Word (APA 7) | Etapa III |
| [salida/Tarea informe de trabajo Proyecto estadisticas de Youtube Avances II.pdf](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2FTarea%2520informe%2520de%2520trabajo%2520Proyecto%2520estadisticas%2520de%2520Youtube%2520Avances%2520II.pdf) | Versión final exportada en PDF del Avance II | Etapa II |
| [salida/Tarea Informe de trabajo Proyecto estadisticas de Youtube Avances III.pdf](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2FTarea%2520Informe%2520de%2520trabajo%2520Proyecto%2520estadisticas%2520de%2520Youtube%2520Avances%2520III.pdf) | Versión final exportada en PDF del Avance III | Etapa III |

### Figuras Generadas en `salida/`

* **Etapas I y II (`analisis.py`):** `fig_distribuciones.png`, `fig_spearman.png`, `fig_comentarios_por_video.png`, `fig_keyword_sentimiento.png`, `fig_modelos_rmse.png`, `fig_real_vs_predicho.png`, `fig_curva_aprendizaje.png`, `fig_matriz_confusion.png`, `fig_importancia.png`.
* **Etapa III (`iteracion.py`):**
  * [fig3\_iteraciones.png](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Ffig3_iteraciones.png): Evolución de desempeño (RMSE y AUC-ROC) a lo largo de las iteraciones `I0`–`I6` y umbral de selección 1-SE.
  * [fig3\_grupos.png](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Ffig3_grupos.png): Importancia por permutación agrupada por familias temáticas de variables.
  * [fig3\_real\_vs\_predicho.png](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Ffig3_real_vs_predicho.png): Diagnóstico de valores observados frente a predichos y análisis de residuos del modelo ganador.
  * [fig3\_boxplot.png](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Ffig3_boxplot.png): Distribución de errores y métricas por pliegue y categoría temática (`Keyword`).
  * [fig3\_intervalos.png](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Ffig3_intervalos.png): Cobertura empírica de los intervalos de predicción al 80% mediante regresión cuantílica.
  * [fig3\_calibracion.png](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Ffig3_calibracion.png): Curvas de confiabilidad (calibración isotónica) para el clasificador de popularidad.

## Preguntas de Investigación y Objetivos

### Pregunta Principal (Regresión)
> **¿En qué medida el volumen de interacción en comentarios, la polaridad afectiva (VADER), las características léxicas del título y el contenido temático (`Keyword`) permiten predecir el número de visualizaciones (`Views`) de un video de YouTube sin depender de métricas contemporáneas con fuga de información?**

* **Objetivo:** Construir y refinar iterativamente un modelo supervisado capaz de estimar las visualizaciones esperadas de un video e incorporar intervalos de predicción calibrados.
* **Variable Objetivo:** $\log_{10}(\text{Views} + 1)$ en la exploración inicial y $\ln(\text{Views} + 1)$ (`log1p(Views)`) en el pipeline de experimentación iterativa de la Etapa III.

### Pregunta Secundaria (Clasificación)
> **¿Es posible clasificar con alta precisión y probabilidades bien calibradas si un video superará la mediana de popularidad del dataset a partir de sus características de interacción en comentarios y contenido?**

* **Objetivo:** Desarrollar, ensamblar y calibrar un clasificador binario que categorice los videos en alta frente a baja popularidad.
* **Variable Objetivo:** `es_popular` ($1$ si $\text{Views} > \text{mediana}$, $0$ en caso contrario).

## Metodología y Evolución del Pipeline

### Etapas I y II: Línea Base y Diagnóstico (`analisis.py`)
1. **Procesamiento de Comentarios y Sentimiento:** Agrupación por video para calcular volumen de comentarios, likes en comentarios y puntajes de sentimiento con **VADER** (*Valence Aware Dictionary and sEntiment Reasoner*).
2. **Prevención de Fuga de Información (*Data Leakage*):** Separación entre escenarios con métricas contemporáneas (`Likes` del video) y escenarios realistas basados únicamente en señales de comentarios, antigüedad, texto del título y nicho temático (`Escenario B`).
3. **Esquema de Validación:** Validación cruzada estratificada y pipelines `ColumnTransformer` + `Pipeline` de Scikit-Learn.

### Etapa III: Experimentación Iterativa y Selección Estadística (`iteracion.py`)
El script [iteracion.py](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fiteracion.py) estructura el refinamiento sobre $N = 1,867$ videos válidos a través de **6 iteraciones formales (`I0`–`I6`)**, evaluando 17 modelos de regresión (87 configuraciones) y 8 modelos de clasificación:

| Iteración | Enfoque Metodológico | Modelos de Regresión (`R0`–`R16`) | Modelos de Clasificación (`C0`–`C7`) |
| :--- | :--- | :--- | :--- |
| **I0** | **Línea Base (Etapa II)** | `R0`: `DummyRegressor` (mediana)<br>`R1`: `RidgeCV`<br>`R2`: `GradientBoosting` (`GridSearchCV`)<br>`R3`: Híbrido `B + TF-IDF + RidgeCV` | `C0`: `DummyClassifier`<br>`C1`: `LogisticRegression`<br>`C2`: `GradientBoosting` (`GridSearchCV`) |
| **I1** | **Ablación e Importancia por Permutación** | `R4`: `HistGradientBoosting` base<br>`R5`: `HGB` sin likes de comentarios<br>`R6`: `Ridge` sin likes de comentarios | `C3`: `LogisticRegression` sin likes de comentarios |
| **I2** | **Selección Formal de Variables** | `R7`: `HGB` + `LassoCV` (`SelectFromModel`)<br>`R8`: `HGB` + `SelectKBest` (Info. Mutua, $k=15$) | — |
| **I3** | **Ingeniería de Variables y Reducción LSA** | `R9`: `HGB` + 10 variables nuevas<br>`R10`: `Ridge` + 10 variables nuevas<br>`R11`: `HGB` + nuevas + LSA (`TruncatedSVD`, 50 comp.) | `C4`: `HGB` fijo + 10 variables nuevas |
| **I4** | **Optimización Aleatoria y Boosting Avanzado** | `R12`: `HGB` afinado (`RandomizedSearchCV`)<br>`R13`: `LightGBM` + variables nuevas<br>`R14`: `XGBoost` + variables nuevas (**Ganador**) | `C5`: `HGB` afinado (`RandomizedSearchCV`) + nuevas |
| **I5** | **Ensambles y Apilamiento (*Stacking*)** | `R15`: Promedio (`HGB + Ridge + Híbrido`)<br>`R16`: Apilado (`HGB + Ridge + Híbrido` $\rightarrow$ reg. no negativa) | `C7`: Apilado (`LR + HGB + TF-IDF LR`) (**Ganador**) |
| **I6** | **Calibración e Intervalos Cuantílicos** | `QuantileRegressor` (intervalos de predicción al 80% nominal) | `C6`: `CalibratedClassifierCV` (isotónica) sobre `HGB` |

#### Nuevas Características Incorporadas en la Etapa III (27 Predictores)
Además de las 16 variables base (`n_com`, `sent_mean`, `sent_std`, `pos_share`, `neg_share`, `neu_share`, `sent_w`, `clike_mean`, `clike_max`, `len_mean`, `words_mean`, `vader_mean`, `vader_std`, `q_share`, `excl_share`, `upper_share`, `age_days`) y la categoría `Keyword`, la iteración `I3` incorpora 10 predictores léxicos y de sentimiento extremo:
* **Atributos del Título:** `title_len` (longitud del título), `title_vader` (polaridad VADER del título), `title_q` (presencia de signos de interrogación en el título).
* **Extremos e Intensidad Emocional en Comentarios:** `vader_max`, `vader_min`, `strong_pos` (proporción de comentarios altamente positivos) y `strong_neg` (proporción de comentarios altamente negativos).
* **Riqueza Léxica y Expresividad:** `lex_div` (diversidad léxica tipo *type-token ratio*) y `emoji_share` (proporción de comentarios con emojis).

#### Criterios Estadísticos de Selección de Modelos
* **Prueba $t$ corregida de Nadeau y Bengio (2003):** Evalúa la significancia estadística de las mejoras entre iteraciones corrigiendo la correlación entre pliegues de validación cruzada repetida ($2 \times 5$).
* **Regla de 1 Error Estándar (1-SE Rule, Hastie et al., 2009):** Entre los modelos cuyo desempeño se ubica dentro de un error estándar del mínimo RMSE (o máximo AUC-ROC), selecciona la alternativa de menor complejidad estructural.

## Resultados Principales

### 1. Resultados de la Etapa III (`salida/results_iii.json`)

#### Regresión — Variable Objetivo: $\ln(\text{Views} + 1)$ (Escenario B sin *Likes* del video)

| ID | Modelo | Iteración | RMSE (CV) | MAE (CV) | $R^2$ (CV) | $\rho$ Spearman |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: |
| `R0` | Referencia (`DummyRegressor`) | `I0` | 2.487 | 2.002 | -0.003 | — |
| `R1` | `RidgeCV` - B (Etapa II) | `I0` | 1.245 | 0.885 | 0.748 | — |
| `R2` | `GradientBoosting` (`GridSearchCV`, Etapa II) | `I0` | 1.242 | 0.900 | 0.749 | — |
| `R3` | Híbrido `B + TF-IDF + RidgeCV` (Etapa II) | `I0` | 1.217 | 0.872 | 0.759 | — |
| **`R14`** | **`XGBoost + nuevas` (Modelo Elegido)** | **`I4`** | **1.191 ($\pm$ 0.062)** | **0.877** | **0.769 ($\pm$ 0.025)** | **0.876** |

* **Decisión por Regla 1-SE:** El modelo **`XGBoost + nuevas` (`R14`)** alcanza el menor error (**RMSE = `1.191`**, **$R^2$ = `0.769`**) con menor complejidad estructural que los ensambles `Promedio (HGB + Ridge + Híbrido)` (`R15`) y `Apilado` (`R16`), ubicándose bajo el umbral 1-SE (`1.211`, $\text{SE} = 0.020$).
* **Intervalos de Predicción (`I6`):** Mediante regresión cuantílica (`QuantileRegressor`) para un nivel nominal del **80%**, se obtuvo una **cobertura empírica del `71.1%` ($\pm 2.0\%$)** con un ancho medio de `2.621` unidades logarítmicas y baja heterocedasticidad residual ($\rho = -0.061$).

#### Clasificación — Variable Objetivo: `es_popular` ($\text{Views} > \text{mediana}$)

| ID | Modelo | Iteración | Accuracy | F1-Score | AUC-ROC | Brier Score | ECE |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: | :---: |
| `C0` | Referencia (`DummyClassifier`) | `I0` | 50.0% | 0.000 | 0.500 | 0.500 | 0.500 |
| `C1` | `LogisticRegression` - B (Etapa II) | `I0` | 86.1% | 0.862 | 0.934 ($\pm$ 0.010) | 0.101 | 0.042 |
| `C2` | `GradientBoosting` (`GridSearchCV`, Etapa II) | `I0` | 85.9% | 0.860 | 0.935 ($\pm$ 0.006) | 0.103 | 0.045 |
| `C3` | `LogisticRegression` sin likes de comentarios | `I1` | 81.7% | 0.820 | 0.900 ($\pm$ 0.011) | 0.127 | 0.048 |
| `C4` | `HGB` fijo + variables nuevas | `I3` | 85.1% | 0.852 | 0.935 ($\pm$ 0.007) | 0.102 | 0.039 |
| `C5` | `HGB` afinado (`RandomizedSearchCV`) + nuevas | `I4` | 84.8% | 0.849 | 0.931 ($\pm$ 0.009) | 0.105 | 0.039 |
| `C6` | `HGB` calibrado (isotónica) + nuevas | `I6` | 85.0% | 0.853 | 0.930 ($\pm$ 0.006) | 0.108 | 0.054 |
| **`C7`** | **`Apilado (LR + HGB + TF-IDF LR)` (Elegido)** | **`I5`** | **86.7%** | **0.865** | **0.939 ($\pm$ 0.008)** | **0.102** | **0.076** |

* **Impacto de la Ablación (`I1`):** Eliminar las métricas de likes en comentarios (`C3`) reduce el AUC-ROC de `0.934` a `0.900`, confirmando el alto valor predictivo de la aprobación comunitaria en los comentarios.
* **Modelo Ganador de Clasificación (`C7`):** El ensamble apilado que integra regresión logística tabular, `HistGradientBoosting` y regresión logística sobre n-gramas TF-IDF del título alcanza el mayor poder discriminativo (**AUC-ROC = `0.939`**, **Precisión = `87.5%`**, **Especificidad = `87.8%`**), superando el umbral 1-SE (`0.937`).

### 2. Referencia Histórica de la Etapa II (Escenario con Métricas Completas en $\log_{10}\text{Views}$)
* **Regresión (Gradient Boosting Regressor):** $R^2$ CV = `0.852`, RMSE CV = `0.341` unidades $\log_{10}$ ($R^2$ Test = `0.858`, RMSE Test = `0.334`, MAE Test = `0.239`).
* **Clasificación (Random Forest Classifier):** Accuracy = `91.2%`, AUC-ROC = `0.965`, F1-Score = `0.912`.

## Requisitos e Instalación

El proyecto requiere **Python 3.10+** (probado y ejecutado en **Python 3.14.7** con `scikit-learn 1.9.1`, `lightgbm 4.7.0` y `xgboost 3.4.1`).

### 1. Clonar el repositorio
```bash
git clone https://github.com/Truder3105/Tarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-.git
cd Tarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-
```

### 2. Crear y activar un entorno virtual
```bash
python3 -m venv venv
source venv/bin/activate        # En Linux/macOS
# venv\Scripts\activate          # En Windows
```

### 3. Instalar dependencias
Instale las librerías base de [requirements.txt](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Frequirements.txt) y los paquetes de gradient boosting utilizados en la Etapa III:
```bash
pip install -r requirements.txt
pip install lightgbm xgboost
```

## Ejecución del Proyecto

El flujo completo consta de tres pasos modulares que admiten parámetros por línea de comandos:

### Paso 1: Ejecutar el Análisis Base (Etapas I y II)
```bash
python3 analisis.py --data data --out salida
```
* **Modo de prueba rápida:** Añada `--rapido` para reducir el número de iteraciones en mallas de búsqueda.
* **Salida:** Procesa [data/videos-stats.csv](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fdata%2Fvideos-stats.csv) y [data/comments.csv](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fdata%2Fcomments.csv), guarda métricas en `salida/results.json`, serializa `salida/modelo_final.joblib` y genera las figuras `fig_*.png`.

### Paso 2: Ejecutar la Experimentación Iterativa (Etapa III)
```bash
python3 iteracion.py --data data --out salida
```
* **Modo de prueba rápida:** Añada `--rapido` para validación rápida del pipeline.
* **Salida:** Ejecuta las iteraciones `I0` a `I6`, realiza pruebas estadísticas de Nadeau-Bengio y selección 1-SE, serializa el mejor pipeline en [salida/modelo\_final\_iii.joblib](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Fmodelo_final_iii.joblib), exporta metadatos en [salida/modelo\_iii\_metadata.json](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Fmodelo_iii_metadata.json), registra todos los resultados en [salida/results\_iii.json](https://www.google.com/search?q=https%3A%2F%2Fgithub.com%2FTruder3105%2FTarea-informe-de-trabajo-Proyecto-estadisticas-de-Youtube-Avances-II-%2Fblob%2Fmain%2Fsalida%2Fresults_iii.json) y genera las 6 figuras `salida/fig3_*.png`.

## Estructura de los Informes Generados (Norma APA 7)

Los documentos generados en `salida/` cumplen con las directrices de presentación académica APA 7:
1. **Página de Título Institucional:** Título del informe, autores, afiliación, asignatura, docente y fecha.
2. **Resumen y Palabras Clave:** Síntesis ejecutiva del problema, diseño experimental y hallazgos cuantitativos.
3. **Introducción y Justificación:** Contexto del problema, formulación de preguntas e hipótesis por iteración.
4. **Método y Diseño Experimental:** Descripción del corpus (Patil, s. f.), ingeniería de variables con VADER y TF-IDF/LSA, arquitectura sin fuga de datos, protocolo de validación cruzada y criterios estadísticos (Nadeau & Bengio, 2003; regla 1-SE).
5. **Resultados y Diagnóstico:** Tablas comparativas de las iteraciones `I0`–`I6`, pruebas de ablación, análisis por subgrupos temáticos (`Keyword`), intervalos de predicción cuantílicos y curvas de calibración.
6. **Discusión y Conclusiones:** Interpretación de hallazgos, limitaciones metodológicas y recomendaciones prácticas.
7. **Referencias Bibliográficas:** Citaciones normalizadas en APA 7 (Breiman, 2001; Bergstra & Bengio, 2012; Chen & Guestrin, 2016; Guyon & Elisseeff, 2003; Hastie et al., 2009; Ke et al., 2017; Koenker & Bassett, 1978; Nadeau & Bengio, 2003; Tibshirani, 1996; Wolpert, 1992).

## Licencia y Créditos

* **Dataset Fuente:** YouTube Statistics Dataset (Patil, s. f., Kaggle: `advaypatil/youtube-statistics`).
* **Desarrollo:** Proyecto académico de analítica de datos y aprendizaje automático.
