import tempfile
from pathlib import Path
import streamlit as st
from ats_core import analyze, load_docx_bytes, load_pdf_bytes, normalize_cv

st.set_page_config(page_title="CV ATS | Revisión de hoja de vida", page_icon="📄", layout="centered")
st.title("Revisa tu hoja de vida para ATS")
st.write("Carga tu CV y, si quieres, la oferta de trabajo. Te mostraré una revisión orientativa, una versión en texto limpio y términos de la oferta que podrías incorporar **solo si reflejan tu experiencia real**.")
st.info("Tus archivos se procesan durante esta sesión y no se guardan en una base de datos. Evita subir información que no quieras procesar en esta aplicación.")

cv_file = st.file_uploader("Sube tu CV", type=["pdf", "docx", "txt", "md"], help="PDF con texto seleccionable, DOCX o texto plano.")
job_file = st.file_uploader("Oferta de trabajo (opcional)", type=["txt", "md"], help="También puedes pegar aquí el texto de la oferta.")
job_text = st.text_area("O pega la oferta de trabajo", height=150, placeholder="Pega aquí la descripción del puesto para comparar palabras clave.")

if cv_file:
    data = cv_file.getvalue()
    warnings=[]
    try:
        suffix=Path(cv_file.name).suffix.lower()
        if suffix == ".pdf":
            cv_text,warnings=load_pdf_bytes(data)
        elif suffix == ".docx":
            cv_text,warnings=load_docx_bytes(data)
        else:
            cv_text=data.decode("utf-8-sig", errors="replace")
        if job_file:
            job_text=job_file.getvalue().decode("utf-8-sig", errors="replace")
        if st.button("Analizar mi CV", type="primary"):
            cleaned,_=normalize_cv(cv_text)
            score, positives, notes=analyze(cleaned,job_text)
            st.subheader("Resultado orientativo")
            st.metric("Indicador de estructura y contenido", f"{score}/100")
            st.caption("No es el puntaje de un ATS real ni garantiza pasar filtros automáticos.")
            col1,col2=st.columns(2)
            with col1:
                st.markdown("#### Lo que ya está")
                if positives:
                    for item in positives: st.write("✅ " + item)
                else: st.write("No se detectaron elementos de esta lista.")
            with col2:
                st.markdown("#### Qué revisar")
                if notes:
                    for item in notes: st.write("• " + item)
                else: st.write("No se detectaron recomendaciones básicas adicionales.")
            if warnings:
                for item in warnings: st.warning(item)
            st.subheader("CV en texto limpio")
            st.caption("Se conserva el contenido extraído y se organizan encabezados reconocibles. Revisa que la extracción haya mantenido el orden correcto.")
            st.text_area("Vista previa", cleaned, height=350)
            st.download_button("Descargar CV limpio (.txt)", cleaned, file_name="cv_ats.txt", mime="text/plain; charset=utf-8")
            report=["REVISIÓN BÁSICA DE CV PARA ATS",f"Puntaje orientativo de estructura: {score}/100","","FORTALEZAS",*(f"- {x}" for x in positives),"","A REVISAR",*(f"- {x}" for x in notes)]
            if warnings: report += ["","ADVERTENCIAS",*(f"- {x}" for x in warnings)]
            report += ["","Nota: ningún programa puede garantizar que un CV pase todos los ATS. El puntaje es heurístico; adapta el CV a la vacante y conserva solo información verdadera."]
            st.download_button("Descargar informe (.txt)", "\n".join(report)+"\n", file_name="informe_cv_ats.txt", mime="text/plain; charset=utf-8")
    except Exception as exc:
        st.error(f"No pude leer este archivo: {exc}")
else:
    st.caption("Carga un archivo para empezar. Si tu PDF es un escaneo, primero necesitará OCR.")

st.divider()
st.caption("Consejo: usa un diseño sencillo, encabezados claros y logros concretos. No añadas palabras clave que no representen tu experiencia.")
