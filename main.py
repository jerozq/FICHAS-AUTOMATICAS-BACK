# backend/main.py
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(BASE_DIR, ".env"), override=True)
except ImportError:
    pass
import shutil
import uuid
from typing import Dict, Any, List
import tempfile
import zipfile
import requests
import base64

# Imports for document processing (adapted from the original main.py)
from docxtpl import DocxTemplate
from openpyxl import load_workbook
try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False
import subprocess
from PyPDF2 import PdfMerger
from datetime import datetime
import time

try:
    import win32com.client as win32
    WIN32_AVAILABLE = True
except Exception:
    WIN32_AVAILABLE = False

import docx
import re
import json
import unicodedata

app = FastAPI(title="Fichas Automáticas API")

# Setup CORS to allow the React frontend to communicate with this API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # For local development
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

RUTA_FORMATOS = os.path.join(BASE_DIR, "formatos")
WORKSPACE_ROOT = os.path.dirname(os.path.dirname(BASE_DIR))
# En Vercel, solo podemos escribir en /tmp
RUTA_SALIDA = os.path.join(tempfile.gettempdir(), "fichas_salida")
os.makedirs(RUTA_FORMATOS, exist_ok=True)
os.makedirs(RUTA_SALIDA, exist_ok=True)

# === AUXILIARY FUNCTIONS ===
def calcular_edad(fecha_nacimiento_str: str) -> str:
    try:
        if not fecha_nacimiento_str: return ""
        fecha_nac = datetime.strptime(fecha_nacimiento_str, "%d-%m-%Y")
        hoy = datetime.today()
        edad = hoy.year - fecha_nac.year - ((hoy.month, hoy.day) < (fecha_nac.month, fecha_nac.day))
        return str(edad)
    except Exception:
        return ""

def escribir_celda(ws, celda_ref: str, valor: str):
    celda = ws[celda_ref]
    for rango in ws.merged_cells.ranges:
        if celda.coordinate in rango:
            celda_superior = ws.cell(rango.min_row, rango.min_col)
            celda_superior.value = valor
            return
    celda.value = valor

EXTRACTABLE_FIELDS = {
    "primer_nombre", "segundo_nombre", "primer_apellido", "segundo_apellido",
    "tipo_documento", "cedula", "expedida_en", "fecha_expedicion", "fecha_nacimiento",
    "nacionalidad", "sexo", "direccion", "barrio", "vereda", "departamento",
    "municipio", "correo", "celular", "telefono", "nivel_educativo", "estado_civil",
    "n_hijos", "estrato", "situacion_laboral", "vivienda", "ingreso_mensual", "cargo",
    "rus", "ruc", "lugar_recepcion", "fecha_recepcion", "conducta_punible", "numero_proceso",
    "fecha_hora_captura", "fiscal", "juez", "privado_libertad", "centro_reclusion", "resumen_hechos"
}

DEFAULT_GEMINI_MODELS = [
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-2.0-flash-001",
    "gemini-2.0-flash-lite",
    "gemini-2.0-flash-lite-001",
]

def _strip_markdown_fences(text: str) -> str:
    if not text:
        return ""
    cleaned = text.strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```[a-zA-Z0-9_-]*", "", cleaned).strip()
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3].strip()
    return cleaned

def _safe_str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()

def _normalize_choice(value: str, mapping: Dict[str, str]) -> str:
    normalized = _safe_str(value).lower()
    if not normalized:
        return ""
    return mapping.get(normalized, normalized)

def _digits_only(value: str) -> str:
    return re.sub(r"\D", "", _safe_str(value))

def _normalize_date_for_html(value: str) -> str:
    raw = _safe_str(value)
    if not raw:
        return ""

    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", raw):
        return raw

    m = re.fullmatch(r"(\d{1,2})[-/](\d{1,2})[-/](\d{4})", raw)
    if m:
        day = int(m.group(1))
        month = int(m.group(2))
        year = int(m.group(3))
        if 1 <= day <= 31 and 1 <= month <= 12:
            return f"{year:04d}-{month:02d}-{day:02d}"

    return ""

def _is_reasonable_text(value: str, max_len: int = 120) -> bool:
    text = _safe_str(value)
    if not text:
        return False
    if len(text) > max_len:
        return False
    return "\n" not in text and "\r" not in text

def _is_valid_email(value: str) -> bool:
    return bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", _safe_str(value)))

def _normalize_text_token(text: str) -> str:
    plain = unicodedata.normalize("NFD", _safe_str(text))
    plain = "".join(ch for ch in plain if unicodedata.category(ch) != "Mn")
    return plain.lower().strip()

def _contains_forbidden_context(value: str) -> bool:
    txt = _normalize_text_token(value)
    forbidden = (
        "delito", "conducta punible", "hechos", "preliminar", "preliminares",
        "direccion", "celular", "telefono", "correo", "fiscal", "juez",
        "radicado", "proceso", "captura"
    )
    return any(token in txt for token in forbidden)

def _is_valid_person_name(value: str) -> bool:
    text = _safe_str(value)
    if not text or len(text) > 35:
        return False
    if _contains_forbidden_context(text):
        return False
    if re.search(r"\d", text):
        return False

    words = [w for w in re.split(r"\s+", text) if w]
    if not (1 <= len(words) <= 3):
        return False

    for w in words:
        if not re.fullmatch(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]+", w):
            return False
    return True

def _is_valid_short_code(value: str) -> bool:
    text = _safe_str(value)
    if not text or len(text) > 25:
        return False
    if _contains_forbidden_context(text):
        return False
    # RUS/RUC suelen ser códigos cortos; bloquear frases con muchos espacios.
    if len([w for w in text.split(" ") if w]) > 3:
        return False
    return bool(re.fullmatch(r"[A-Za-z0-9\-_/\.]+", text))

def _is_valid_cargo(value: str) -> bool:
    text = _safe_str(value)
    if not text or len(text) > 90:
        return False
    if _contains_forbidden_context(text):
        return False
    if any(sep in text for sep in ("\n", "\r", ".", ";", ":")):
        return False
    if len([w for w in text.split(" ") if w]) > 8:
        return False
    return True

def _normalize_extracted_data(payload: Dict[str, Any], source: str = "local") -> Dict[str, Any]:
    if not isinstance(payload, dict):
        return {}

    is_gemini = source == "gemini"
    normalized: Dict[str, Any] = {}
    for key, raw_value in payload.items():
        if key not in EXTRACTABLE_FIELDS:
            continue

        if key == "privado_libertad":
            if isinstance(raw_value, bool):
                normalized[key] = raw_value
                continue
            val = _safe_str(raw_value).lower()
            normalized[key] = val in ("si", "sí", "true", "1", "x", "yes")
            continue

        val = _safe_str(raw_value)
        if not val:
            continue

        if is_gemini and key in ("primer_nombre", "segundo_nombre", "primer_apellido", "segundo_apellido"):
            if not _is_valid_person_name(val):
                continue

        if is_gemini and key in ("rus", "ruc"):
            if not _is_valid_short_code(val):
                continue

        if is_gemini and key == "cargo":
            if not _is_valid_cargo(val):
                continue

        if key == "correo" and not _is_valid_email(val):
            continue

        if key in ("celular", "telefono"):
            digits = _digits_only(val)
            if not (7 <= len(digits) <= 12):
                continue
            normalized[key] = digits
            continue

        if key == "cedula":
            digits = _digits_only(val)
            if not (5 <= len(digits) <= 15):
                continue
            normalized[key] = digits
            continue

        if key == "n_hijos":
            digits = _digits_only(val)
            if not digits:
                continue
            n_hijos_int = int(digits)
            if n_hijos_int > 30:
                continue
            normalized[key] = str(n_hijos_int)
            continue

        if key == "estrato":
            digits = _digits_only(val)
            if digits not in ("0", "1", "2", "3", "4", "5", "6"):
                continue
            normalized[key] = digits
            continue

        if key == "numero_proceso":
            proc = re.sub(r"\s+", "", val)
            if not re.fullmatch(r"\d{11,25}(-\d{1,4})?", proc):
                continue
            normalized[key] = proc
            continue

        if key == "fecha_recepcion":
            normalized_date = _normalize_date_for_html(val)
            if not normalized_date:
                continue
            normalized[key] = normalized_date
            continue

        if key in ("primer_nombre", "segundo_nombre", "primer_apellido", "segundo_apellido") and not _is_reasonable_text(val, 40):
            continue

        if key in ("departamento", "municipio", "expedida_en", "nacionalidad") and not _is_reasonable_text(val, 60):
            continue

        if is_gemini and key in ("ruc", "rus", "fiscal", "juez", "conducta_punible", "cargo") and not _is_reasonable_text(val, 140):
            continue

        if key == "tipo_documento":
            type_doc = _normalize_choice(val, {
                "c.c": "cc", "cc": "cc", "cedula": "cc", "cédula": "cc",
                "ti": "ti", "t.i": "ti", "ce": "ce", "c.e": "ce"
            })
            if type_doc in ("cc", "ti", "ce"):
                normalized[key] = type_doc
        elif key == "sexo":
            sexo = _normalize_choice(val, {
                "masculino": "masculino", "m": "masculino",
                "femenino": "femenino", "f": "femenino",
                "otro": "otro"
            })
            if sexo in ("masculino", "femenino", "otro"):
                normalized[key] = sexo
        elif key == "estado_civil":
            estado = _normalize_choice(val, {
                "soltero": "soltero",
                "casado": "casado",
                "separado": "separado",
                "viudo": "viudo",
                "union libre": "union libre",
                "unión libre": "union libre"
            })
            if estado in ("soltero", "casado", "separado", "viudo", "union libre"):
                normalized[key] = estado
        elif key == "situacion_laboral":
            laboral = _normalize_choice(val, {
                "dependiente": "dependiente",
                "independiente": "independiente",
                "desempleado": "desempleado",
                "estudiante": "estudiante",
                "otros": "otros",
                "otro": "otros"
            })
            if laboral in ("dependiente", "independiente", "desempleado", "estudiante", "otros"):
                normalized[key] = laboral
        elif key == "vivienda":
            vivienda = _normalize_choice(val, {
                "propia": "propia",
                "arrendada": "arrendada",
                "familiar": "familiar",
                "otros": "otros",
                "otro": "otros"
            })
            if vivienda in ("propia", "arrendada", "familiar", "otros"):
                normalized[key] = vivienda
        else:
            normalized[key] = val

    return normalized

def extraer_datos_con_gemini(texto_plano: str) -> Dict[str, Any]:
    api_key = os.getenv("GEMINI_API_KEY")
    if not api_key:
        return {}

    preferred_model = _safe_str(os.getenv("GEMINI_MODEL"))
    modelos = [preferred_model] if preferred_model else []
    modelos.extend(model for model in DEFAULT_GEMINI_MODELS if model not in modelos)

    prompt = f"""
Extrae del siguiente texto los datos para prellenar un formulario judicial.

Reglas estrictas:
- Devuelve SOLO un JSON valido (sin markdown).
- No inventes informacion.
- Si no estas seguro de un campo, OMITELO.
- No cruces campos: jamas pongas delito/hechos en nombres o apellidos; ni direccion/celular en rus/ruc; ni resumen_hechos en cargo.
- Usa EXACTAMENTE estos nombres de campo:
    primer_nombre, segundo_nombre, primer_apellido, segundo_apellido,
    tipo_documento, cedula, expedida_en, fecha_expedicion, fecha_nacimiento,
    nacionalidad, sexo, direccion, barrio, vereda, departamento, municipio,
    correo, celular, telefono, nivel_educativo, estado_civil, n_hijos, estrato,
    situacion_laboral, vivienda, ingreso_mensual, cargo, rus, ruc,
    lugar_recepcion, fecha_recepcion, conducta_punible, numero_proceso,
    fecha_hora_captura, fiscal, juez, privado_libertad, centro_reclusion, resumen_hechos.
- tipo_documento: cc, ti o ce.
- sexo: masculino, femenino u otro.
- estado_civil: soltero, casado, separado, viudo, union libre.
- situacion_laboral: dependiente, independiente, desempleado, estudiante u otros.
- vivienda: propia, arrendada, familiar u otros.
- privado_libertad: true o false.

Texto a analizar:
{texto_plano}
"""

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json"
        }
    }

    last_error = None
    for modelo in modelos:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{modelo}:generateContent?key={api_key}"
        try:
            response = requests.post(url, json=payload, timeout=45)
            response.raise_for_status()
            data = response.json()
            text_out = (
                data.get("candidates", [{}])[0]
                    .get("content", {})
                    .get("parts", [{}])[0]
                    .get("text", "")
            )
            if not text_out:
                continue
            parsed = json.loads(_strip_markdown_fences(text_out))
            return _normalize_extracted_data(parsed, source="gemini")
        except Exception as exc:
            last_error = exc
            print(f"Fallo Gemini con modelo {modelo}: {exc}")
            continue

    print(f"No se pudo extraer con Gemini: {last_error}")
    return {}

# === EXTRACTION LOGIC ===
@app.post("/api/extract")
async def extract_data_from_docx(file: UploadFile = File(...)):
    # Create a temporary file to save the uploaded word document
    temp_dir = tempfile.mkdtemp()
    temp_file_path = os.path.join(temp_dir, file.filename)
    
    with open(temp_file_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
        
    if not os.getenv("GEMINI_API_KEY"):
        raise HTTPException(status_code=503, detail="GEMINI_API_KEY no esta configurada en el backend")

    try:
        doc = docx.Document(temp_file_path)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"No se pudo leer el archivo Word: {str(e)}")
        
    texto_completo = []
    for p in doc.paragraphs:
        t = p.text.strip()
        if t: texto_completo.append(t)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                text = c.text.strip(' \n\r\t')
                if text and text not in texto_completo:
                    texto_completo.append(text)
    texto_plano = "\n".join(texto_completo)
    gemini_data = extraer_datos_con_gemini(texto_plano)

    # Clean up temp file
    try:
        shutil.rmtree(temp_dir)
    except Exception:
        pass

    return {
        "extracted_data": gemini_data,
        "meta": {
            "gemini_enabled": True,
            "gemini_used": len(gemini_data) > 0,
            "fields_detected": len(gemini_data)
        }
    }

# === DOCUMENT GENERATION LOGIC ===

def llenar_excel1(plantilla, salida, datos):
    wb = load_workbook(plantilla)
    ws = wb.active

    escribir_celda(ws, "K5", str(datos.get("rus", "")))
    escribir_celda(ws, "W5", str(datos.get("ruc", "")))
    escribir_celda(ws, "D8", str(datos.get("lugar_recepcion", "")))
    escribir_celda(ws, "R8", str(datos.get("fecha_recepcion", "")))
    escribir_celda(ws, "D38", str(datos.get("primer_apellido", "")))
    escribir_celda(ws, "L38", str(datos.get("segundo_apellido", "")))
    escribir_celda(ws, "S38", str(datos.get("primer_nombre", "")))
    escribir_celda(ws, "Z38", str(datos.get("segundo_nombre", "")))
    escribir_celda(ws, "C43", str(datos.get("cedula", "")))
    escribir_celda(ws, "K43", str(datos.get("expedida_en", "")))
    escribir_celda(ws, "S43", str(datos.get("fecha_expedicion", "")))
    escribir_celda(ws, "AA43", str(datos.get("nacionalidad", "")))
    escribir_celda(ws, "I45", str(datos.get("direccion", "")))
    escribir_celda(ws, "U45", str(datos.get("barrio", "")))
    escribir_celda(ws, "H47", str(datos.get("departamento", "")))
    escribir_celda(ws, "O47", str(datos.get("municipio", "")))
    escribir_celda(ws, "Y47", str(datos.get("vereda", "")))
    escribir_celda(ws, "AA47", str(datos.get("correo", "")))
    escribir_celda(ws, "T49", str(datos.get("telefono", "")))
    escribir_celda(ws, "AA49", str(datos.get("celular", "")))
    escribir_celda(ws, "F52", str(datos.get("fecha_nacimiento", "")))
    escribir_celda(ws, "L51", calcular_edad(datos.get("fecha_nacimiento", "")))
    escribir_celda(ws, "AA55", str(datos.get("nivel_educativo", "")))
    escribir_celda(ws, "AC61", str(datos.get("n_hijos", "")))
    escribir_celda(ws, "H68", str(datos.get("cargo", "")))
    escribir_celda(ws, "S68", str(datos.get("empresa", "")))
    escribir_celda(ws, "H72", str(datos.get("ingreso_mensual", "")))
    escribir_celda(ws, "C74", str(datos.get("estrato", "")))
    escribir_celda(ws, "H88", str(datos.get("conducta_punible", "")))
    escribir_celda(ws, "I90", str(datos.get("numero_proceso", "")))
    escribir_celda(ws, "AA88", str(datos.get("fecha_hora_captura", "")))
    escribir_celda(ws, "H92", str(datos.get("fiscal", "")))
    escribir_celda(ws, "L92", str(datos.get("juez", "")))
    escribir_celda(ws, "A107", str(datos.get("resumen_hechos", "")))

    nombre_completo = f"{datos.get('primer_nombre','')} {datos.get('segundo_nombre','')} {datos.get('primer_apellido','')} {datos.get('segundo_apellido','')}".strip()
    escribir_celda(ws, "R116", "NO FIRMA PORQUE SE HIZO VIRTUAL")
    escribir_celda(ws, "Q119", str(datos.get("cedula", "")))
    escribir_celda(ws, "P119", str(datos.get("tipo_documento", "CC")).upper())

    tipo_doc = datos.get("tipo_documento", "").strip().lower()
    escribir_celda(ws, "F41", "X" if tipo_doc == "cc" else "")
    escribir_celda(ws, "I41", "X" if tipo_doc == "ti" else "")
    escribir_celda(ws, "K41", "X" if tipo_doc == "ce" else "")

    estado = datos.get("estado_civil", "").strip().lower()
    escribir_celda(ws, "R51", "X" if estado == "casado" else "")
    escribir_celda(ws, "T51", "X" if estado == "soltero" else "")
    escribir_celda(ws, "W51", "X" if estado == "viudo" else "")
    escribir_celda(ws, "Z51", "X" if estado == "separado" else "")
    escribir_celda(ws, "AC51", "X" if estado in ("union libre", "unión libre", "unionlibre") else "")
    escribir_celda(ws, "N61", str(datos.get("nombre_conyuge", "")) if estado in ("casado", "union libre", "unión libre", "unionlibre") else "")

    sexo = datos.get("sexo", "").strip().lower()
    escribir_celda(ws, "C53", "X" if sexo == "femenino" else "")
    escribir_celda(ws, "F53", "X" if sexo in ("masculino", "m") else "")

    laboral = datos.get("situacion_laboral", "").strip().lower()
    escribir_celda(ws, "L65", "X" if laboral in ("dependiente", "trabajador dependiente", "empleado") else "")
    escribir_celda(ws, "S65", "X" if laboral in ("independiente", "trabajador independiente", "indep") else "")
    escribir_celda(ws, "Y64", "X" if laboral in ("desempleado", "sin empleo") else "")
    escribir_celda(ws, "AC65", "X" if laboral in ("estudiante",) else "")

    vivienda = datos.get("vivienda", "").strip().lower()
    escribir_celda(ws, "I76", "X" if vivienda in ("propia", "propiedad", "propio") else "")
    escribir_celda(ws, "R76", "X" if vivienda in ("arrendada", "arrendar", "arriendo", "arrendado") else "")
    escribir_celda(ws, "M76", "X" if vivienda in ("familiar", "familiar/otros", "familiar ") else "")

    wb.save(salida)

def llenar_excel2(plantilla, salida, datos):
    wb = load_workbook(plantilla)
    ws = wb.active

    ws["D7"] = str(datos.get("fecha_recepcion", ""))
    ws["E9"] = str(datos.get("conducta_punible", ""))
    nombre_completo = f"{datos.get('primer_nombre','')} {datos.get('segundo_nombre','')} {datos.get('primer_apellido','')} {datos.get('segundo_apellido','')}".strip()
    ws["D8"] = nombre_completo
    ws["D15"] = str(datos.get("centro_reclusion", ""))

    privado = str(datos.get("privado_libertad", "false")).lower() in ("true", "1", "si", "sí")
    ws["F14"] = "X" if privado else ""
    ws["H14"] = "X" if not privado else ""

    ws["G40"] = "NO FIRMA PORQUE SE HIZO VIRTUAL"

    wb.save(salida)

def llenar_word(plantilla, salida, datos):
    doc = DocxTemplate(plantilla)
    context = {
        "primer_nombre": datos.get("primer_nombre", ""),
        "segundo_nombre": datos.get("segundo_nombre", ""),
        "primer_apellido": datos.get("primer_apellido", ""),
        "segundo_apellido": datos.get("segundo_apellido", ""),
        "nombre_completo": f"{datos.get('primer_nombre','')} {datos.get('segundo_nombre','')} {datos.get('primer_apellido','')} {datos.get('segundo_apellido','')}".strip(),
        "cedula": datos.get("cedula", ""),
        "tipo_documento": datos.get("tipo_documento", ""),
        "fecha_recepcion": datos.get("fecha_recepcion", ""),
        "lugar_recepcion": datos.get("lugar_recepcion", ""),
        "conducta_punible": datos.get("conducta_punible", ""),
        "centro_reclusion": datos.get("centro_reclusion", ""),
        "resumen_hechos": datos.get("resumen_hechos", ""),
        "firma": "NO FIRMA PORQUE SE HIZO VIRTUAL"
    }
    doc.render(context)
    doc.save(salida)

def obtener_soffice_path():
    candidates = [
        os.path.join(WORKSPACE_ROOT, "LibreOfficePortable", "App", "libreoffice", "program", "soffice.exe"),
        os.path.join(WORKSPACE_ROOT, "LibreOfficePortable", "App", "libreoffice", "program", "soffice.com"),
        shutil.which("soffice"),
        shutil.which("libreoffice"),
    ]
    for candidate in candidates:
        if candidate and os.path.exists(candidate):
            return candidate
    return None

def convertir_con_libreoffice(file_path: str):
    soffice_path = obtener_soffice_path()
    if not soffice_path:
        return None

    out_dir = os.path.dirname(os.path.abspath(file_path))
    pdf_path = os.path.splitext(file_path)[0] + ".pdf"
    try:
        subprocess.run(
            [soffice_path, "--headless", "--convert-to", "pdf", os.path.abspath(file_path), "--outdir", out_dir],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        if os.path.exists(pdf_path):
            return pdf_path
    except Exception as e:
        print(f"No se pudo convertir con LibreOffice: {e}")
    return None

def convertir_docx_a_pdf_windows(docx_path):
    if not WIN32_AVAILABLE:
        return None
    word = None
    document = None
    pdf_path = docx_path.replace(".docx", ".pdf")
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True)
        document.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)
        document.Close(False)
        word.Quit()
        if os.path.exists(pdf_path):
            return pdf_path
    except Exception as e:
        print(f"No se pudo convertir DOCX con Word: {e}")
        try:
            if document:
                document.Close(False)
        except Exception:
            pass
        try:
            if word:
                word.Quit()
        except Exception:
            pass
    return None

def convertir_docx_a_pdf(docx_path):
    pdf_path = docx_path.replace(".docx", ".pdf")
    try:
        if DOCX2PDF_AVAILABLE:
            docx2pdf_convert(docx_path)
            if os.path.exists(pdf_path):
                return pdf_path
    except Exception as e:
        print("No se pudo convertir DOCX a PDF:", e)
    pdf_local = convertir_docx_a_pdf_windows(docx_path)
    if pdf_local:
        return pdf_local
    return convertir_con_libreoffice(docx_path)

def convertir_documento_local(file_path: str):
    ext = file_path.split('.')[-1].lower()
    if ext == 'docx':
        return convertir_docx_a_pdf(file_path)
    if ext == 'xlsx':
        pdf_path = file_path.replace('.xlsx', '.pdf')
        ok, error = convertir_xlsx_a_pdf_windows(file_path, pdf_path)
        if ok and os.path.exists(pdf_path):
            return pdf_path
        if error:
            print(error)
        return convertir_con_libreoffice(file_path)
    return None

def convertir_documento_cloudconvert(file_path: str):
    api_token = os.getenv("CLOUD_CONVERT_API_SECRET")
    if not api_token:
        return None

    ext = file_path.split('.')[-1].lower()
    if ext not in ['docx', 'xlsx']:
        return None

    api_base = os.getenv("CLOUD_CONVERT_API_BASE", "https://api.cloudconvert.com/v2")
    filename = os.path.basename(file_path)
    output_filename = os.path.splitext(filename)[0] + ".pdf"

    try:
        with open(file_path, "rb") as f:
            file_b64 = base64.b64encode(f.read()).decode("ascii")

        headers = {
            "Authorization": f"Bearer {api_token}",
            "Content-Type": "application/json",
        }
        payload = {
            "tasks": {
                "import-my-file": {
                    "operation": "import/base64",
                    "file": file_b64,
                    "filename": filename,
                },
                "convert-my-file": {
                    "operation": "convert",
                    "input": "import-my-file",
                    "input_format": ext,
                    "output_format": "pdf",
                    "filename": output_filename,
                },
                "export-my-file": {
                    "operation": "export/url",
                    "input": "convert-my-file",
                    "inline": False,
                    "archive_multiple_files": False,
                },
            }
        }

        create_response = requests.post(f"{api_base}/jobs", headers=headers, json=payload, timeout=60)
        create_response.raise_for_status()
        job_data = create_response.json().get("data", {})
        job_id = job_data.get("id")
        if not job_id:
            print(f"Error de CloudConvert: respuesta sin job id: {create_response.text}")
            return None

        last_job_data = job_data
        for _ in range(60):
            job_response = requests.get(f"{api_base}/jobs/{job_id}", headers=headers, timeout=30)
            job_response.raise_for_status()
            last_job_data = job_response.json().get("data", {})
            status = last_job_data.get("status")
            if status == "finished":
                break
            if status == "error":
                break
            time.sleep(2)

        if last_job_data.get("status") != "finished":
            print(f"Error de CloudConvert: job no finalizado correctamente: {last_job_data}")
            return None

        export_task = None
        for task in last_job_data.get("tasks", []):
            if task.get("name") == "export-my-file" or task.get("operation") == "export/url":
                export_task = task
                break

        if not export_task:
            print(f"Error de CloudConvert: no se encontro task export/url: {last_job_data}")
            return None

        files = export_task.get("result", {}).get("files", [])
        if not files:
            print(f"Error de CloudConvert: export/url sin archivos: {export_task}")
            return None

        download_url = files[0].get("url")
        if not download_url:
            print(f"Error de CloudConvert: archivo exportado sin URL: {files[0]}")
            return None

        download_response = requests.get(download_url, timeout=120)
        download_response.raise_for_status()
        out_path = file_path.replace(f".{ext}", ".pdf")
        with open(out_path, "wb") as out:
            out.write(download_response.content)
        return out_path if os.path.exists(out_path) else None
    except Exception as e:
        print(f"Error llamando a CloudConvert: {e}")
        return None

def convertir_documento_a_pdf(file_path: str):
    pdf_path = convertir_documento_cloudconvert(file_path)
    if pdf_path and os.path.exists(pdf_path):
        return pdf_path
    pdf_path = convertir_documento_api(file_path)
    if pdf_path and os.path.exists(pdf_path):
        return pdf_path
    return convertir_documento_local(file_path)

def convertir_xlsx_a_pdf_windows(xlsx_path, output_pdf_path):
    if not WIN32_AVAILABLE:
        return False, "win32com no disponible"
    excel = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(os.path.abspath(xlsx_path), UpdateLinks=False, ReadOnly=True)
        for sheet in wb.Worksheets:
            try:
                sheet.Visible = True
                sheet.Activate()
                excel.ActiveWindow.View = 1
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = False
                time.sleep(0.2)
            except Exception:
                pass
        wb.ExportAsFixedFormat(
            Type=0,
            Filename=os.path.abspath(output_pdf_path),
            Quality=0,
            IncludeDocProperties=True,
            IgnorePrintAreas=False,
            OpenAfterPublish=False
        )
        wb.Close(SaveChanges=False)
        excel.Quit()
        return True, None
    except Exception as e:
        try:
            if excel: excel.Quit()
        except: pass
        return False, f"Error al exportar Excel a PDF: {e}"

def unir_pdfs(lista_pdfs, salida_final):
    merger = PdfMerger()
    for pdf in lista_pdfs:
        if pdf and os.path.exists(pdf):
            merger.append(pdf)
    merger.write(salida_final)
    merger.close()

def convertir_documento_api(file_path: str):
    api_secret = os.getenv("CONVERT_API_SECRET")
    if not api_secret:
        print("ADVERTENCIA: CONVERT_API_SECRET no está configurado.")
        return None
        
    ext = file_path.split('.')[-1].lower()
    if ext not in ['docx', 'xlsx']:
        return None
        
    url = f"https://v2.convertapi.com/convert/{ext}/to/pdf?Secret={api_secret}"
    try:
        with open(file_path, "rb") as f:
            files = {"File": f}
            response = requests.post(url, files=files, timeout=60)
            
        if response.status_code == 200:
            data = response.json()
            if "Files" in data and len(data["Files"]) > 0:
                file_data = base64.b64decode(data["Files"][0]["FileData"])
                out_path = file_path.replace(f".{ext}", ".pdf")
                with open(out_path, "wb") as out:
                    out.write(file_data)
                return out_path
        print(f"Error de ConvertAPI: {response.text}")
    except Exception as e:
        print(f"Error llamando a ConvertAPI: {e}")
    return None

@app.post("/api/generate")
async def generate_documents(datos: Dict[str, Any]):
    uid = str(uuid.uuid4())[:8]
    nombre_base = f"{datos.get('primer_nombre','sin_nombre')}_{datos.get('cedula','')}_{uid}".replace(" ", "_")
    
    excel1_out = os.path.join(RUTA_SALIDA, f"formato1_{nombre_base}.xlsx")
    excel2_out = os.path.join(RUTA_SALIDA, f"formato2_{nombre_base}.xlsx")
    word_out = os.path.join(RUTA_SALIDA, f"formato3_{nombre_base}.docx")
    
    formato1_in = os.path.join(RUTA_FORMATOS, "formato1.xlsx")
    formato2_in = os.path.join(RUTA_FORMATOS, "formato2.xlsx")
    formato3_in = os.path.join(RUTA_FORMATOS, "formato3.docx")
    
    if not os.path.exists(formato1_in) or not os.path.exists(formato2_in) or not os.path.exists(formato3_in):
        raise HTTPException(status_code=500, detail="Faltan las plantillas en la carpeta backend/formatos")

    try:
        # Fill templates
        llenar_excel1(formato1_in, excel1_out, datos)
        llenar_excel2(formato2_in, excel2_out, datos)
        llenar_word(formato3_in, word_out, datos)

        # Intentar ConvertAPI y, si falla, convertir localmente.
        pdf_excel1 = convertir_documento_a_pdf(excel1_out)
        pdf_word = convertir_documento_a_pdf(word_out)
        pdf_excel2 = convertir_documento_a_pdf(excel2_out)
        
        # Unir PDFs en el orden solicitado: Formato 1, Formato 3, Formato 2
        orden = [pdf_excel1, pdf_word, pdf_excel2]
        orden_existentes = [p for p in orden if p and os.path.exists(p)]
        
        nombre_completo = " ".join([
            str(datos.get('primer_nombre', '')).strip(),
            str(datos.get('segundo_nombre', '')).strip(),
            str(datos.get('primer_apellido', '')).strip(),
            str(datos.get('segundo_apellido', '')).strip(),
        ]).strip()
        nombre_completo = re.sub(r'\s+', ' ', nombre_completo)
        nombre_completo = re.sub(r'[<>:"/\\|?*]', '', nombre_completo)
        if not nombre_completo:
            nombre_completo = "SIN NOMBRE"

        salida_final_name = f"FICHA {nombre_completo}.pdf"
        salida_final = os.path.join(RUTA_SALIDA, salida_final_name)
        
        if len(orden_existentes) > 0:
            unir_pdfs(orden_existentes, salida_final)
        else:
            raise HTTPException(status_code=500, detail="No se pudo convertir ningún documento a PDF. ConvertAPI esta agotada y el fallback local tambien fallo.")
            
        if not os.path.exists(salida_final):
            raise HTTPException(status_code=500, detail="No se pudo crear el archivo PDF consolidado")

        # Devolver el archivo directamente
        return FileResponse(salida_final, media_type='application/pdf', filename=salida_final_name)

    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"Error durante la generación: {str(e)}\n{traceback.format_exc()}")

@app.get("/api/download/{filename}")
async def download_file(filename: str):
    file_path = os.path.join(RUTA_SALIDA, filename)
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Archivo no encontrado")
    return FileResponse(file_path, media_type='application/pdf', filename=filename)

@app.get("/api/template")
async def download_template():
    template_path = os.path.join(RUTA_FORMATOS, "PLANTILLA_INGRESO_DATOS.docx")
    if not os.path.exists(template_path):
        raise HTTPException(status_code=404, detail="Plantilla no encontrada")
    return FileResponse(
        template_path,
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        filename="Plantilla_Defensoria.docx"
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="127.0.0.1", port=8000)
