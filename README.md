# Revisión de hoja de vida para ATS

Aplicación en Streamlit para revisar un CV, extraer el texto de PDF/DOCX, generar una versión de texto plano y comparar términos con una oferta laboral. El análisis es heurístico: **no garantiza superar filtros ATS ni sustituye una revisión humana**.

## Probar localmente

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Admite CV en PDF con texto seleccionable, DOCX, TXT o Markdown. La oferta puede subirse como TXT/Markdown o pegarse en el campo. Los archivos se procesan en memoria durante la sesión; esta app no los guarda en una base de datos.

## Publicar

Despliega este repositorio en Streamlit Community Cloud apuntando a `app.py`. No se requieren secretos ni variables privadas.
