# Proyecto Estadísticas de YouTube – pipeline de informes (Avances I y II)

## Qué hace
`analisis.py` calcula TODAS las cifras reales a partir de los CSV de Kaggle (calidad, estadística no trivial,
suficiencia, 15 modelos, importancia de variables) y guarda `results.json`, figuras y `modelo_final.joblib`.
`informes.py` toma `Plantilla_APA7.docx` + `results.json` y genera los dos informes en Word (APA 7).

## Pasos
1. `pip install -r requirements.txt`
2. Descargue de Kaggle (advaypatil/youtube-statistics) `videos-stats.csv` y `comments.csv` a una carpeta `data/`.
3. `python analisis.py --data data --out salida`      (prueba rápida: añada `--rapido`)
4. Edite el bloque CONFIG de `informes.py` (integrantes, universidad, curso, profesor, fechas).
5. `python informes.py --plantilla Plantilla_APA7.docx --salida salida`
6. Abra los .docx, revíselos y expórtelos a PDF (Archivo > Guardar como > PDF), que es el formato exigido.

## Importante
- Los números del texto salen de `results.json`; no se escriben a mano. Si cambian los datos, regenere.
- Revise y personalice la redacción: usted debe poder defender cada decisión en la reunión con el profesor.
- La ejecución completa puede tardar varios minutos (búsquedas en malla anidadas).
