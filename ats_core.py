#!/usr/bin/env python3
"""Auditor y limpiador básico de CV para ATS. Solo reorganiza/limpia texto; no inventa datos.

Uso:
  python3 cv_ats.py mi_cv.pdf --oferta oferta.txt
  python3 cv_ats.py mi_cv.docx --salida cv_ats.txt --reporte informe.txt
  python3 cv_ats.py --plantilla cv_base.txt

PDF: requiere PyMuPDF (pip install pymupdf). DOCX: lectura con biblioteca estándar.
"""
from __future__ import annotations
import argparse
import re
import sys
import unicodedata
import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

SECTION_ALIASES = {
    "perfil profesional": "PERFIL PROFESIONAL", "resumen profesional": "PERFIL PROFESIONAL",
    "sobre mí": "PERFIL PROFESIONAL", "sobre mi": "PERFIL PROFESIONAL", "objetivo profesional": "PERFIL PROFESIONAL",
    "experiencia": "EXPERIENCIA LABORAL", "experiencia laboral": "EXPERIENCIA LABORAL",
    "experiencia profesional": "EXPERIENCIA LABORAL", "trayectoria profesional": "EXPERIENCIA LABORAL",
    "educación": "EDUCACIÓN", "educacion": "EDUCACIÓN", "formación": "EDUCACIÓN", "formacion": "EDUCACIÓN",
    "formación académica": "EDUCACIÓN", "formacion academica": "EDUCACIÓN", "estudios": "EDUCACIÓN",
    "habilidades": "HABILIDADES", "competencias": "HABILIDADES", "aptitudes": "HABILIDADES",
    "habilidades técnicas": "HABILIDADES", "habilidades tecnicas": "HABILIDADES",
    "idiomas": "IDIOMAS", "certificaciones": "CERTIFICACIONES", "certificados": "CERTIFICACIONES",
    "cursos": "CURSOS Y CERTIFICACIONES", "cursos y certificaciones": "CURSOS Y CERTIFICACIONES",
    "proyectos": "PROYECTOS", "proyectos destacados": "PROYECTOS",
    "voluntariado": "VOLUNTARIADO", "referencias": "REFERENCIAS",
}
ORDER = ["PERFIL PROFESIONAL", "EXPERIENCIA LABORAL", "EDUCACIÓN", "HABILIDADES", "IDIOMAS", "CURSOS Y CERTIFICACIONES", "PROYECTOS", "VOLUNTARIADO", "CERTIFICACIONES", "REFERENCIAS"]
ACTION_VERBS = ("lideré", "dirigí", "desarrollé", "implementé", "mejoré", "reduje", "aumenté", "gestioné", "coordiné", "creé", "automaticé", "diseñé", "analicé", "atendí", "supervisé", "optimicé", "logré", "incrementé", "organicé", "capacité", "lidero", "desarrollo", "implemento", "gestiono", "coordino", "creo", "diseño", "analizo")
STOP = set("para como desde hasta entre sobre bajo ante tras durante mediante donde cuando quien quienes cuyo cuya los las unos unas del al con por sin que y o en de la el un una es son fue han sus su se mi mis tu tus más mas muy también experiencia profesional trabajo oferta empresa puesto años año perfil equipo resultados responsabilidad responsabilidades habilidad habilidades conocimiento conocimientos importante requerido requerida deseable".split())


def load_docx_bytes(data: bytes) -> tuple[str, list[str]]:
    warnings=[]
    with zipfile.ZipFile(BytesIO(data)) as z:
        xml=z.read("word/document.xml")
    root=ET.fromstring(xml)
    ns={"w":"http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs=[]
    for p in root.findall(".//w:p", ns):
        text="".join((node.text or "") for node in p.findall(".//w:t",ns))
        if text.strip(): paragraphs.append(text.strip())
    if root.findall(".//w:tbl",ns): warnings.append("El DOCX contiene tablas; el texto de sus celdas puede no conservar el orden visual. Revisa el resultado.")
    if root.findall(".//w:cols",ns): warnings.append("El DOCX parece usar columnas; verifica manualmente el orden del texto extraído.")
    return "\n".join(paragraphs),warnings


def load_pdf_bytes(data: bytes) -> tuple[str,list[str]]:
    try:
        import fitz
    except ImportError:
        raise RuntimeError("Para leer PDF instala PyMuPDF: python3 -m pip install pymupdf")
    doc=fitz.open(stream=data, filetype="pdf")
    text="\n".join(page.get_text("text") for page in doc)
    warnings=[]
    if not text.strip(): warnings.append("No se pudo extraer texto del PDF; podría ser un escaneo. Convierte el documento con OCR y vuelve a intentarlo.")
    if len(doc)>2: warnings.append(f"El PDF tiene {len(doc)} páginas. Para muchos perfiles, conviene revisar si puede condensarse a 1–2.")
    return text,warnings


def read_docx(path: Path) -> tuple[str,list[str]]:
    return load_docx_bytes(Path(path).read_bytes())


def read_pdf(path: Path) -> tuple[str,list[str]]:
    return load_pdf_bytes(Path(path).read_bytes())


def load_text(path: Path) -> tuple[str,list[str]]:
    path=Path(path)
    suffix=path.suffix.lower()
    if suffix==".docx": return load_docx_bytes(path.read_bytes())
    if suffix==".pdf": return load_pdf_bytes(path.read_bytes())
    if suffix in (".txt",".md"): return path.read_text(encoding="utf-8",errors="replace"),[]
    raise RuntimeError("Formato no compatible. Usa .txt, .md, .docx o .pdf.")


def canonical_heading(line: str) -> str | None:
    key=re.sub(r"\s+"," ",line.strip().strip(":|—–- ").lower())
    return SECTION_ALIASES.get(key)


def clean_line(line: str) -> str:
    line=line.replace("\u00a0"," ").replace("\t"," ")
    line=re.sub(r"\s+"," ",line).strip()
    line=re.sub(r"^[•●▪◦‣*-]\s*", "- ", line)
    return line


def normalize_cv(text: str) -> tuple[str,list[str]]:
    lines=[clean_line(x) for x in text.replace("\r","\n").split("\n")]
    lines=[x for x in lines if x]
    if not lines: return "",[]
    name=lines[0]
    # Divide en secciones reconocidas; conserva literalmente el resto del contenido.
    sections={}; current="INFORMACIÓN DE CONTACTO"; sections[current]=[]
    body=[]
    for line in lines[1:]:
        heading=canonical_heading(line)
        if heading:
            current=heading
            sections.setdefault(current,[])
        else:
            sections.setdefault(current,[]).append(line)
    # Lines at top before first heading are contact. If no contact details, nothing is fabricated.
    output=[name]
    for section in ["INFORMACIÓN DE CONTACTO"]+[s for s in ORDER if s in sections]+[s for s in sections if s not in ORDER and s!="INFORMACIÓN DE CONTACTO"]:
        content=sections.get(section,[])
        if content:
            output.extend(["",section,*content])
    return "\n".join(output).strip()+"\n",[]


def words(text: str) -> list[str]:
    return [w.lower() for w in re.findall(r"[A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9+#.]+",text)]


def analyze(text: str, job_text: str="") -> tuple[int,list[str],list[str]]:
    notes=[]; positives=[]; score=0
    word_count=len(words(text))
    email=bool(re.search(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}",text))
    phone=bool(re.search(r"(?:\+?\d[\d ()-]{7,}\d)",text))
    if email: score+=10; positives.append("Incluye un correo electrónico detectable.")
    else: notes.append("Añade un correo profesional en texto normal.")
    if phone: score+=10; positives.append("Incluye un teléfono detectable.")
    else: notes.append("Añade un teléfono con código de país si buscas fuera de tu ciudad.")
    headings={canonical_heading(x) for x in text.splitlines()}
    if "EXPERIENCIA LABORAL" in headings: score+=15; positives.append("Sección de experiencia identificable.")
    else: notes.append("Incluye una sección titulada ‘Experiencia laboral’ si corresponde a tu trayectoria.")
    if "EDUCACIÓN" in headings: score+=10; positives.append("Sección de educación identificable.")
    else: notes.append("Añade ‘Educación’ o ‘Formación’ si aplica.")
    if "HABILIDADES" in headings: score+=10; positives.append("Sección de habilidades identificable.")
    else: notes.append("Incluye habilidades relevantes y específicas para el puesto.")
    if 250<=word_count<=900: score+=10; positives.append(f"Extensión de {word_count} palabras, dentro de un rango práctico para revisar.")
    elif word_count<250: notes.append(f"El CV es corto ({word_count} palabras); verifica si faltan logros, tareas o formación relevantes.")
    else: notes.append(f"El CV tiene {word_count} palabras; revisa redundancias y prioriza lo relevante.")
    bullets=sum(1 for l in text.splitlines() if l.startswith("- "))
    if bullets: score+=5; positives.append(f"Usa {bullets} viñetas legibles en texto plano.")
    else: notes.append("Usa viñetas simples (‘- ’) para responsabilidades y logros, no cuadros de texto o gráficos.")
    if re.search(r"\b(?:19|20)\d{2}\b",text): score+=5; positives.append("Hay fechas en formato de año detectable.")
    else: notes.append("Incluye fechas claras para cada empleo y estudio, sin dejar periodos ambiguos.")
    if re.search(r"\d+\s*%|\$\s?\d+|\b\d+\s+(?:personas|clientes|proyectos|ventas|horas)\b",text,re.I):
        score+=10; positives.append("Incluye al menos un resultado cuantificado.")
    else: notes.append("Cuando puedas respaldarlo, cuantifica logros (porcentajes, volumen, tiempo, presupuesto o clientes).")
    verb_hits=[v for v in ACTION_VERBS if re.search(r"\b"+re.escape(v)+r"\b",text.lower())]
    if verb_hits: score+=5; positives.append("Hay verbos de acción: "+", ".join(verb_hits[:6])+".")
    else: notes.append("Empieza algunas viñetas con verbos de acción concretos (por ejemplo: desarrollé, reduje, coordiné).")
    if job_text.strip():
        cvset={w for w in words(text) if w not in STOP and len(w)>2}
        jobset={w for w in words(job_text) if w not in STOP and len(w)>2}
        matched=sorted(cvset & jobset)
        missing=sorted(jobset-cvset)
        if jobset:
            pct=round(100*len(matched)/len(jobset))
            score+=min(15,round(pct*0.15))
            positives.append(f"Coincidencia léxica simple con la oferta: {pct}% de palabras relevantes detectadas; esto no equivale al puntaje real de un ATS.")
            if missing: notes.append("Palabras de la oferta que no aparecen en el CV: "+", ".join(missing[:25])+". Añádelas solo si describen de verdad tu experiencia.")
        else: notes.append("No se detectaron palabras comparables en la oferta.")
    else:
        notes.append("Si tienes una oferta concreta, usa --oferta para identificar términos relevantes que faltan.")
    return min(100,score),positives,notes


def template() -> str:
    return """NOMBRE Y APELLIDO
Ciudad, país | Teléfono | Correo | LinkedIn (URL)

PERFIL PROFESIONAL
Resumen de 3–4 líneas: especialidad, años de experiencia si corresponde, fortalezas y el tipo de aporte que puedes demostrar.

EXPERIENCIA LABORAL
Cargo — Empresa | Ciudad | MM/AAAA–MM/AAAA
- Acción + responsabilidad + contexto.
- Logro verificable con resultado, volumen, tiempo o porcentaje.

EDUCACIÓN
Título — Institución | Ciudad | Año

HABILIDADES
- Herramientas, conocimientos técnicos y competencias relevantes para la vacante.

IDIOMAS
Idioma — nivel real

CURSOS Y CERTIFICACIONES
Nombre — institución | año
"""


def main() -> int:
    ap=argparse.ArgumentParser(description="Revisa un CV, genera una versión de texto simple y compara palabras con una oferta.")
    ap.add_argument("cv",nargs="?",help="Archivo .txt, .md, .docx o .pdf")
    ap.add_argument("--oferta",help="Archivo .txt/.md con la descripción del puesto")
    ap.add_argument("--salida",help="Archivo de salida ATS-friendly (predeterminado: <nombre>_ats.txt)")
    ap.add_argument("--reporte",help="Ruta del informe (predeterminado: <nombre>_informe.txt)")
    ap.add_argument("--plantilla",metavar="ARCHIVO",help="Crea una plantilla de CV en texto plano y termina")
    args=ap.parse_args()
    if args.plantilla:
        Path(args.plantilla).write_text(template(),encoding="utf-8")
        print(f"Plantilla creada: {args.plantilla}")
        return 0
    if not args.cv:
        ap.error("indica un CV o usa --plantilla archivo.txt")
    cv_path=Path(args.cv)
    if not cv_path.exists(): raise RuntimeError(f"No encuentro el archivo: {cv_path}")
    text,warnings=load_text(cv_path)
    if not text.strip(): raise RuntimeError("El archivo no contiene texto legible.")
    job=""
    if args.oferta:
        job_path=Path(args.oferta)
        if not job_path.exists(): raise RuntimeError(f"No encuentro la oferta: {job_path}")
        job=job_path.read_text(encoding="utf-8",errors="replace")
    cleaned,_=normalize_cv(text)
    score,positives,notes=analyze(cleaned,job)
    out=Path(args.salida) if args.salida else cv_path.with_name(cv_path.stem+"_ats.txt")
    report=Path(args.reporte) if args.reporte else cv_path.with_name(cv_path.stem+"_informe.txt")
    out.write_text(cleaned,encoding="utf-8")
    report_lines=["REVISIÓN BÁSICA DE CV PARA ATS",f"Puntaje orientativo de estructura: {score}/100","","FORTALEZAS",*(f"- {x}" for x in positives),"","A REVISAR",*(f"- {x}" for x in notes)]
    if warnings: report_lines += ["","ADVERTENCIAS DE EXTRACCIÓN",*(f"- {x}" for x in warnings)]
    report_lines += ["","Nota: ningún programa puede garantizar que un CV pase todos los ATS. El puntaje es una comprobación heurística; adapta el CV a la vacante y conserva solo información verdadera."]
    report.write_text("\n".join(report_lines)+"\n",encoding="utf-8")
    print(f"Versión ATS: {out}\nInforme: {report}\nPuntaje orientativo: {score}/100")
    if warnings:
        for w in warnings: print("Advertencia:",w)
    return 0

if __name__=="__main__":
    try: raise SystemExit(main())
    except (RuntimeError,OSError,zipfile.BadZipFile) as e:
        print(f"Error: {e}",file=sys.stderr); raise SystemExit(2)
