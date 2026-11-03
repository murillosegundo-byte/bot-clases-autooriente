"""
Bot de Seguimiento de Clases — CEA Auto Oriente (Versión Unificada lista para Render y PC).
Incluye:
- Formatos Oficiales A.O.C.E.A - F - 59 (A1), F - 60 (A2), F - 61 (B1) y F - 62 (C1 / Recat C1).
- Extracción Híbrida (Regex avanzado + Gemini API opcional).
- Integración con Google Sheets (Alumnos, Instructores, Registro_Clases).
- Integración con Google Drive / enlaces directos de evidencia fotográfica.
- Generador de PDFs en memoria con ReportLab.
- Servidor HTTP Health Check en segundo plano para Render Web Service.
"""

import difflib
import io
import json
import logging
import os
import re
import sys
import threading
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from telegram import ReactionTypeEmoji, Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# =============================================================================
# 1. CONFIGURACIÓN Y CREDENCIALES
# =============================================================================
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "TU_TELEGRAM_BOT_TOKEN").strip()
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "TU_ID_DE_GOOGLE_SHEETS").strip()
GOOGLE_CREDENTIALS_FILE = os.getenv(
    "GOOGLE_CREDENTIALS_FILE", str(BASE_DIR / "credentials.json")
).strip()
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON", "").strip()
GOOGLE_DRIVE_ROOT_FOLDER_ID = os.getenv("GOOGLE_DRIVE_ROOT_FOLDER_ID", "").strip()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
USE_GEMINI_VISION = os.getenv("USE_GEMINI_VISION", "false").lower() in (
    "true",
    "1",
    "yes",
    "si",
)
PORT = int(os.getenv("PORT", "10000"))

SHEET_ALUMNOS = "Alumnos"
SHEET_INSTRUCTORES = "Instructores"
SHEET_REGISTRO = "Registro_Clases"
TZ_COLOMBIA = timezone(timedelta(hours=-5))

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


def obtener_ruta_logo() -> Optional[Path]:
    """Busca el logo tanto en la raíz del repositorio como en la carpeta assets/."""
    candidatos = [
        BASE_DIR / "logo_autooriente.png",
        BASE_DIR / "assets" / "logo_autooriente.png",
    ]
    for ruta in candidatos:
        if ruta.exists():
            return ruta
    return None


def obtener_credenciales_google(scopes: List[str]) -> Credentials:
    """Carga credenciales desde GOOGLE_CREDENTIALS_JSON (Render) o archivo local."""
    if GOOGLE_CREDENTIALS_JSON:
        info = json.loads(GOOGLE_CREDENTIALS_JSON)
        return Credentials.from_service_account_info(info, scopes=scopes)

    render_secret_path = Path("/etc/secrets/credentials.json")
    if render_secret_path.exists():
        return Credentials.from_service_account_file(
            str(render_secret_path), scopes=scopes
        )

    return Credentials.from_service_account_file(
        GOOGLE_CREDENTIALS_FILE, scopes=scopes
    )


# =============================================================================
# 2. CATÁLOGO OFICIAL DE FORMATOS Y TEMARIOS (F-59, F-60, F-61 y F-62)
# =============================================================================
FORMATOS_OFICIALES: Dict[str, Dict[str, Any]] = {
    "A1": {
        "codigo_formato": "A.O.C.E.A - F - 59",
        "version": "01",
        "fecha_formato": "Mayo 05 de 2021",
        "elaborado_por": "Director General",
        "aprobado_por": "Dirección General",
        "titulo": "FORMATO\nCONTENIDO PROGRAMÁTICO SEGUIMIENTO Y CONTROL DE CLASES PRÁCTICAS A1",
        "duracion_minutos": 45,
        "filas": [
            {
                "etiqueta": "1.0",
                "clases": [1],
                "etapa": "pista",
                "tema": "Inspección pre operacional, documentos y equipos de prevención, adaptación viso espacial, ajuste del asiento, timón y espejos.\nExplicación elementos de seguridad activa y pasiva. Identificación del tablero de instrumentos, mandos y pedales. Encendido del motor y testigos.",
            },
            {
                "etiqueta": "2.0",
                "clases": [2],
                "etapa": "pista",
                "tema": "Maniobra de puesta en marcha con técnica punto de contacto del embrague. Conducción del vehículo hacia adelante.\nOperación del selector del vehículo. Coordinación: Acelerador - freno - embrague - desaceleración.",
            },
            {
                "etiqueta": "3.0",
                "clases": [3],
                "etapa": "via",
                "tema": "Conducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Eslalon.\nConducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Primeros indicios de parqueo.",
            },
            {
                "etiqueta": "4.0",
                "clases": [4],
                "etapa": "via",
                "tema": "Conducción del vehículo en vía urbana en terreno plano. Pares y arranques.",
            },
            {
                "etiqueta": "5.0",
                "clases": [5],
                "etapa": "via",
                "tema": "Arranque del vehículo en terreno pendiente, de frente y reversa.\nConducción en terreno plano con maniobras de adelantamiento.",
            },
            {
                "etiqueta": "6.0",
                "clases": [6],
                "etapa": "via",
                "tema": "Parqueo y estacionamiento con y sin uso de espejo. Refuerzo.\nConducción en vías urbanas usando las señales lumínicas, corporales y acústicas. Cambio de calzada. Cambio de carril.",
            },
            {
                "etiqueta": "7.0",
                "clases": [7],
                "etapa": "via",
                "tema": "Conducción en vías urbanas, glorietas, intersecciones, deprimidos, puentes elevados, zonas escolares.\nConducción en vías urbanas (Buses, taxis, trasporte Masivo)",
            },
            {
                "etiqueta": "8.0",
                "clases": [8],
                "etapa": "via",
                "tema": "Conducción en vías principales urbanas de la ciudad.\nConducción en vías nacionales. Piedecuesta - Bucaramanga - Girón - peaje vía aeropuerto.",
            },
        ],
    },
    "A2": {
        "codigo_formato": "A.O.C.E.A - F - 60",
        "version": "04",
        "fecha_formato": "Octubre 29 de 2026",
        "elaborado_por": "Director General",
        "aprobado_por": "Dirección General",
        "titulo": "FORMATO\nCONTENIDO PROGRAMÁTICO SEGUIMIENTO Y CONTROL DE CLASES PRÁCTICAS A2",
        "duracion_minutos": 45,
        "filas": [
            {
                "etiqueta": "1.0",
                "clases": [1],
                "etapa": "pista",
                "tema": "Inspección pre operacional, documentos y equipos de prevención, adaptación viso espacial, correcta postura en el manillar y espejos.",
            },
            {
                "etiqueta": "2.0",
                "clases": [2],
                "etapa": "pista",
                "tema": "Explicación elementos de seguridad activa y pasiva. Identificación del tablero de instrumentos, mandos y pedales. Encendido del motor y testigos.",
            },
            {
                "etiqueta": "3.0",
                "clases": [3],
                "etapa": "pista",
                "tema": "Maniobra de carreteo o maniobra cero cale, maniobra de puesta en marcha con técnica punto de contacto del embrague.",
            },
            {
                "etiqueta": "4.0",
                "clases": [4],
                "etapa": "pista",
                "tema": "Operación del selector del vehículo. Coordinación: Acelerador - freno - embrague - desaceleración.",
            },
            {
                "etiqueta": "5.0",
                "clases": [5],
                "etapa": "pista",
                "tema": "Circular sobre una franja de anchura limitada, ejercicio de equilibrio y control de la motocicleta, Giro en U.",
            },
            {
                "etiqueta": "6.0",
                "clases": [6],
                "etapa": "via",
                "tema": "Conducción del vehículo hacia adelante con giros haciendo uso conos Eslalon a baja velocidad.",
            },
            {
                "etiqueta": "7.0",
                "clases": [7],
                "etapa": "via",
                "tema": "Conducción del vehículo hacia adelante con giros haciendo uso conos Aceleración, embrague y frenado Controlado.",
            },
            {
                "etiqueta": "8.0",
                "clases": [8],
                "etapa": "via",
                "tema": "Conducción: sorteo de obstáculos a alta velocidad (máx. 40 km/h) y frenado de emergencia. Slalom, cajón de frenado y evasión de barrera, Eslalon a alta velocidad",
            },
            {
                "etiqueta": "9.0",
                "clases": [9],
                "etapa": "via",
                "tema": "Arranque del vehículo en terreno plano y Desplazamientos en ocho.",
            },
            {
                "etiqueta": "10.0",
                "clases": [10],
                "etapa": "via",
                "tema": "Arranque del vehículo en terreno inclinado, con y sin acompañante.",
            },
            {
                "etiqueta": "11.0",
                "clases": [11],
                "etapa": "via",
                "tema": "Maniobra de avance y detención, Interacción con señales de tránsito horizontales y verticales.",
            },
            {
                "etiqueta": "12.0",
                "clases": [12],
                "etapa": "via",
                "tema": "Conducción en vías usando las señales lumínicas, corporales y acústicas. Cambio de calzada. Cambio de carril.",
            },
            {
                "etiqueta": "13.0",
                "clases": [13],
                "etapa": "via",
                "tema": "Conducción del vehículo en vía en terreno plano. Pares y arranques. Barrio de poco flujo vehicular, interacción con señales de tránsito",
            },
            {
                "etiqueta": "14.0",
                "clases": [14],
                "etapa": "via",
                "tema": "Refuerzo de pruebas CALE según resolución del Ministerio de Transporte.",
            },
            {
                "etiqueta": "15.0",
                "clases": [15],
                "etapa": "via",
                "tema": "Evaluación general del curso de conducción de motocicleta, interacción con señales de tránsito.",
            },
        ],
    },
    "B1": {
        "codigo_formato": "A.O.C.E.A - F - 61",
        "version": "02",
        "fecha_formato": "abril 05 de 2025",
        "elaborado_por": "Director General",
        "aprobado_por": "Dirección General",
        "titulo": "FORMATO\nCONTENIDO PROGRAMÁTICO SEGUIMIENTO Y CONTROL DE CLASES PRÁCTICAS B1",
        "duracion_minutos": 45,
        "filas": [
            {
                "etiqueta": "1.0",
                "clases": [1],
                "etapa": "pista",
                "tema": "Inspección pre operacional, documentos y equipos de prevención, adaptación viso espacial, ajuste del asiento, timón y espejos.",
            },
            {
                "etiqueta": "2.0",
                "clases": [2],
                "etapa": "pista",
                "tema": "Explicación elementos de seguridad activa y pasiva. Identificación del tablero de instrumentos, mandos y pedales. Encendido del motor y testigos.",
            },
            {
                "etiqueta": "3.0",
                "clases": [3],
                "etapa": "pista",
                "tema": "Maniobra de puesta en marcha con técnica punto de contacto del embrague. Conducción del vehículo hacia adelante.",
            },
            {
                "etiqueta": "4.0",
                "clases": [4],
                "etapa": "pista",
                "tema": "Operación del selector del vehículo. Coordinación: Acelerador - freno - embrague - desaceleración.",
            },
            {
                "etiqueta": "5.0",
                "clases": [5],
                "etapa": "pista",
                "tema": "Conducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Eslalon.",
            },
            {
                "etiqueta": "6.0",
                "clases": [6],
                "etapa": "pista",
                "tema": "Conducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Primeros indicios de parqueo.",
            },
            {
                "etiqueta": "7.0",
                "clases": [7],
                "etapa": "via",
                "tema": "Conducción del vehículo en vía urbana en terreno plano. Pares y arranques.",
            },
            {
                "etiqueta": "8.0",
                "clases": [8],
                "etapa": "via",
                "tema": "Conducción del vehículo en vía urbana en terreno plano. Pares y arranques.",
            },
            {
                "etiqueta": "9.0",
                "clases": [9],
                "etapa": "via",
                "tema": "Arranque del vehículo en terreno pendiente, de frente y reversa.",
            },
            {
                "etiqueta": "10.0",
                "clases": [10],
                "etapa": "via",
                "tema": "Arranque del vehículo en terreno pendiente, de frente y reversa.",
            },
            {
                "etiqueta": "11.0",
                "clases": [11],
                "etapa": "via",
                "tema": "Conducción en terreno plano con maniobras de adelantamiento.",
            },
            {
                "etiqueta": "12.0",
                "clases": [12],
                "etapa": "via",
                "tema": "Parqueo y estacionamiento con y sin uso de espejo. Refuerzo.",
            },
            {
                "etiqueta": "13.0",
                "clases": [13],
                "etapa": "via",
                "tema": "Parqueo y estacionamiento con y sin uso de espejo. Refuerzo.",
            },
            {
                "etiqueta": "14.0",
                "clases": [14],
                "etapa": "via",
                "tema": "Conducción en vías urbanas usando las señales lumínicas, corporales y acústicas. Cambio de calzada. Cambio de carril.",
            },
            {
                "etiqueta": "15.0",
                "clases": [15],
                "etapa": "via",
                "tema": "Conducción en vías urbanas, glorietas, intersecciones, deprimidos, puentes elevados, zonas escolares.",
            },
            {
                "etiqueta": "16.0",
                "clases": [16],
                "etapa": "via",
                "tema": "Conducción en vías urbanas (Buses, taxis, trasporte Masivo)",
            },
            {
                "etiqueta": "17.0",
                "clases": [17],
                "etapa": "via",
                "tema": "Conducción en vías urbanas (Buses, taxis, trasporte Masivo)",
            },
            {
                "etiqueta": "18.0",
                "clases": [18],
                "etapa": "via",
                "tema": "Conducción en vías principales urbanas de la ciudad.",
            },
            {
                "etiqueta": "19.0",
                "clases": [19],
                "etapa": "via",
                "tema": "Conducción en vías principales urbanas de la ciudad.",
            },
            {
                "etiqueta": "20.0",
                "clases": [20],
                "etapa": "via",
                "tema": "Conducción en vías nacionales. Piedecuesta - Bucaramanga - Girón - peaje vía aeropuerto.",
            },
        ],
    },
    "C1": {
        "codigo_formato": "A.O.C.E.A - F - 62",
        "version": "02",
        "fecha_formato": "abril 05 de 2025",
        "elaborado_por": "Director General",
        "aprobado_por": "Dirección General",
        "titulo": "FORMATO\nCONTENIDO PROGRAMÁTICO SEGUIMIENTO Y CONTROL DE CLASES PRÁCTICAS C1",
        "duracion_minutos": 45,
        "filas": [
            {
                "etiqueta": "1.0",
                "clases": [1],
                "etapa": "pista",
                "tema": "Inspección pre operacional, documentos y equipos de prevención, adaptación viso espacial, ajuste del asiento, timón y espejos.",
            },
            {
                "etiqueta": "2.0",
                "clases": [2],
                "etapa": "pista",
                "tema": "Explicación elementos de seguridad activa y pasiva. Identificación del tablero de instrumentos, mandos y pedales. Encendido del motor y testigos.",
            },
            {
                "etiqueta": "3,4",
                "clases": [3, 4],
                "etapa": "pista",
                "tema": "Maniobra de puesta en marcha con técnica punto de contacto del embrague. Conducción del vehículo hacia adelante.",
            },
            {
                "etiqueta": "5,6",
                "clases": [5, 6],
                "etapa": "pista",
                "tema": "Operación del selector del vehículo. Coordinación: Acelerador - freno - embrague - desaceleración.",
            },
            {
                "etiqueta": "7,8",
                "clases": [7, 8],
                "etapa": "pista",
                "tema": "Conducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Eslalon.",
            },
            {
                "etiqueta": "9,10",
                "clases": [9, 10],
                "etapa": "pista",
                "tema": "Conducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Primeros indicios de parqueo.",
            },
            {
                "etiqueta": "11,12",
                "clases": [11, 12],
                "etapa": "via",
                "tema": "Conducción del vehículo en vía urbana en terreno plano. Pares y arranques.",
            },
            {
                "etiqueta": "13,14",
                "clases": [13, 14],
                "etapa": "via",
                "tema": "Conducción del vehículo en vía urbana en terreno plano. Pares y arranques.",
            },
            {
                "etiqueta": "15,16",
                "clases": [15, 16],
                "etapa": "via",
                "tema": "Arranque del vehículo en terreno pendiente, de frente y reversa.",
            },
            {
                "etiqueta": "17,18",
                "clases": [17, 18],
                "etapa": "via",
                "tema": "Conducción en terreno plano con maniobras de adelantamiento.",
            },
            {
                "etiqueta": "19,20",
                "clases": [19, 20],
                "etapa": "via",
                "tema": "Parqueo y estacionamiento con y sin uso de espejo. Refuerzo.",
            },
            {
                "etiqueta": "21,22",
                "clases": [21, 22],
                "etapa": "via",
                "tema": "Conducción en vías urbanas usando las señales lumínicas, corporales y acústicas. Cambio de calzada. Cambio de carril.",
            },
            {
                "etiqueta": "23,24",
                "clases": [23, 24],
                "etapa": "via",
                "tema": "Conducción en vías urbanas, glorietas, intersecciones, deprimidos, puentes elevados, zonas escolares.",
            },
            {
                "etiqueta": "25,26",
                "clases": [25, 26],
                "etapa": "via",
                "tema": "Conducción en vías urbanas (Buses, taxis, trasporte Masivo)",
            },
            {
                "etiqueta": "27,28",
                "clases": [27, 28],
                "etapa": "via",
                "tema": "Conducción en vías principales urbanas de la ciudad.",
            },
            {
                "etiqueta": "29,30",
                "clases": [29, 30],
                "etapa": "via",
                "tema": "Conducción en vías nacionales. Piedecuesta - Bucaramanga - Girón - peaje vía aeropuerto.",
            },
        ],
    },
    "RECAT_C1": {
        "codigo_formato": "A.O.C.E.A - F - 62",
        "version": "02",
        "fecha_formato": "abril 05 de 2025",
        "elaborado_por": "Director General",
        "aprobado_por": "Dirección General",
        "titulo": "FORMATO\nCONTENIDO PROGRAMÁTICO SEGUIMIENTO Y CONTROL DE CLASES PRÁCTICAS C1 (RECATEGORIZACIÓN)",
        "duracion_minutos": 45,
        "filas": [
            {
                "etiqueta": "21,22",
                "clases": [21, 22],
                "etapa": "pista",
                "tema": "Conducción en vías urbanas usando las señales lumínicas, corporales y acústicas. Cambio de calzada. Cambio de carril.",
            },
            {
                "etiqueta": "23,24",
                "clases": [23, 24],
                "etapa": "via",
                "tema": "Conducción en vías urbanas, glorietas, intersecciones, deprimidos, puentes elevados, zonas escolares.",
            },
            {
                "etiqueta": "25,26",
                "clases": [25, 26],
                "etapa": "via",
                "tema": "Conducción en vías urbanas (Buses, taxis, trasporte Masivo)",
            },
            {
                "etiqueta": "27,28",
                "clases": [27, 28],
                "etapa": "via",
                "tema": "Conducción en vías principales urbanas de la ciudad.",
            },
            {
                "etiqueta": "29,30",
                "clases": [29, 30],
                "etapa": "via",
                "tema": "Conducción en vías nacionales. Piedecuesta - Bucaramanga - Girón - peaje vía aeropuerto.",
            },
        ],
    },
}


def normalizar_categoria(cat: str) -> str:
    if not cat:
        return "B1"
    limpio = cat.strip().upper().replace("-", "_").replace(" ", "_")
    if "RECAT" in limpio:
        return "RECAT_C1"
    if "A1" in limpio:
        return "A1"
    if "A2" in limpio or "MOTO" in limpio:
        return "A2"
    if "C1" in limpio:
        return "C1"
    return "B1"


def obtener_tema_oficial(categoria: str, clases: List[int]) -> str:
    cat_key = normalizar_categoria(categoria)
    formato = FORMATOS_OFICIALES.get(cat_key, FORMATOS_OFICIALES["B1"])
    if not clases:
        return "Sin especificar"
    primera = clases[0]
    for fila in formato["filas"]:
        if primera in fila["clases"]:
            return fila["tema"]
    return f"Práctica de conducción clase {', '.join(map(str, clases))}"


# =============================================================================
# 3. EXTRACTOR HÍBRIDO DE REPORTES (REGEX + GEMINI OPCIONAL)
# =============================================================================
@dataclass
class ReporteExtraido:
    alumno: str = "Desconocido"
    clases_lista: List[int] = field(default_factory=list)
    clases_texto: str = "No identificada"
    tema: str = "Sin especificar"
    fecha_timemark: Optional[str] = None
    hora_timemark: Optional[str] = None
    instructor_detectado: Optional[str] = None


def extraer_numeros_clase(texto_clase: str) -> List[int]:
    numeros = [int(n) for n in re.findall(r"\d+", texto_clase) if 1 <= int(n) <= 35]
    if not numeros:
        return []
    if len(numeros) == 2 and numeros[1] > numeros[0] and (numeros[1] - numeros[0]) <= 4:
        return list(range(numeros[0], numeros[1] + 1))
    vistos = []
    for n in numeros:
        if n not in vistos:
            vistos.append(n)
    return vistos


def formatear_clases(clases: List[int]) -> str:
    if not clases:
        return "No identificada"
    if len(clases) == 1:
        return str(clases[0])
    return "-".join(str(c) for c in clases)


def parsear_reporte_regex(texto: str) -> ReporteExtraido:
    resultado = ReporteExtraido()
    if not texto or not texto.strip():
        return resultado

    lineas = [l.strip() for l in texto.split("\n") if l.strip()]
    posibles_alumnos = []
    posibles_temas = []

    patron_clase = re.compile(
        r"(?:clases?|cl\.?)\s*[:#.]?\s*(\d{1,2}(?:\s*[-/,yY&aA]\s*\d{1,2})*)",
        re.IGNORECASE,
    )
    patron_prefijo_alumno = re.compile(
        r"^(?:alumn[oa]s?|aprendiz|estudiante|nombre)\s*[:.\-]?\s*(.+)$",
        re.IGNORECASE,
    )
    patron_prefijo_instructor = re.compile(
        r"^(?:instructor[a]?|profe(?:sor)?)\s*[:.\-]?\s*(.+)$",
        re.IGNORECASE,
    )
    patron_prefijo_tema = re.compile(
        r"^(?:temas?|actividad|pr[aá]ctica)\s*[:.\-]?\s*(.+)$",
        re.IGNORECASE,
    )

    for linea in lineas:
        m_inst = patron_prefijo_instructor.match(linea)
        if m_inst:
            resultado.instructor_detectado = m_inst.group(1).strip(" .-")
            continue

        m_alum = patron_prefijo_alumno.match(linea)
        if m_alum and not patron_clase.search(linea):
            resultado.alumno = m_alum.group(1).strip(" .-")
            continue

        m_tema = patron_prefijo_tema.match(linea)
        if m_tema:
            posibles_temas.append(m_tema.group(1).strip(" .-"))
            continue

        m_clase = patron_clase.search(linea)
        if m_clase:
            raw_clases = m_clase.group(1)
            clases_nums = extraer_numeros_clase(raw_clases)
            if clases_nums:
                resultado.clases_lista = clases_nums
                resultado.clases_texto = formatear_clases(clases_nums)

            resto = linea[: m_clase.start()] + " " + linea[m_clase.end() :]
            resto = re.sub(r"^[\s.\-:;]+|[\s.\-:;]+$", "", resto)
            if resto:
                m_alum_inline = patron_prefijo_alumno.match(resto)
                if m_alum_inline:
                    resultado.alumno = m_alum_inline.group(1).strip(" .-")
                else:
                    posibles_temas.append(resto)
            continue

        palabras = linea.split()
        tiene_digitos = any(ch.isdigit() for ch in linea)
        palabras_tema_clave = {
            "conduccion", "conducción", "arranque", "punto", "contacto", "vias", "vías",
            "glorietas", "trafico", "tráfico", "parqueo", "estacionamiento", "pista",
            "pendiente", "loma", "cambios", "freno", "espejos", "maniobras", "reverso",
            "intersecciones", "señales", "senales", "urbana", "urbanas", "nacionales",
        }
        contiene_palabra_tema = any(
            p.lower().strip(".,;:()") in palabras_tema_clave for p in palabras
        )

        if not tiene_digitos and 1 <= len(palabras) <= 5 and not contiene_palabra_tema:
            posibles_alumnos.append(linea.strip(" .-"))
        else:
            posibles_temas.append(linea.strip(" .-"))

    if resultado.alumno == "Desconocido" and posibles_alumnos:
        resultado.alumno = posibles_alumnos[0]

    if posibles_temas:
        resultado.tema = " - ".join(posibles_temas)

    return resultado


def parsear_reporte(texto: str, imagen_bytes: Optional[bytes] = None) -> ReporteExtraido:
    return parsear_reporte_regex(texto)


# =============================================================================
# 4. SERVICIO DE GOOGLE SHEETS Y DRIVE
# =============================================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

HEADERS_ALUMNOS = ["ID_Matricula", "Documento", "Nombre_Completo", "Categoria", "Estado"]
HEADERS_INSTRUCTORES = ["Telegram_User_ID", "Nombre_Instructor", "Categoria_Habilitada"]
HEADERS_REGISTRO = [
    "Timestamp",
    "Fecha_Reporte",
    "Hora_Reporte",
    "Instructor",
    "ID_Alumno",
    "Nombre_Alumno",
    "Categoria",
    "Clases_Reportadas",
    "Tema_Registrado",
    "URL_Foto_Drive",
]


def _normalizar_texto(texto: str) -> str:
    if not texto:
        return ""
    texto_nfd = unicodedata.normalize("NFD", str(texto))
    sin_tildes = "".join(c for c in texto_nfd if unicodedata.category(c) != "Mn")
    return " ".join(sin_tildes.lower().strip().split())


class SheetsService:
    def __init__(self):
        self._client: Optional[gspread.Client] = None
        self._spreadsheet: Optional[gspread.Spreadsheet] = None

    def conectar(self) -> gspread.Spreadsheet:
        if self._spreadsheet is None:
            creds = obtener_credenciales_google(SCOPES)
            self._client = gspread.authorize(creds)
            self._spreadsheet = self._client.open_by_key(GOOGLE_SHEET_ID)
        return self._spreadsheet

    def obtener_o_crear_hoja(self, nombre_hoja: str, encabezados: List[str]) -> gspread.Worksheet:
        book = self.conectar()
        try:
            ws = book.worksheet(nombre_hoja)
            if not ws.row_values(1):
                ws.append_row(encabezados)
            return ws
        except gspread.WorksheetNotFound:
            ws = book.add_worksheet(title=nombre_hoja, rows=1000, cols=len(encabezados) + 2)
            ws.append_row(encabezados)
            return ws

    def inicializar_estructura(self) -> None:
        self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
        self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
        self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)

    def buscar_instructor(self, telegram_user_id: int, nombre_fallback: str) -> str:
        try:
            ws = self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
            for fila in ws.get_all_records():
                if str(fila.get("Telegram_User_ID", "")).strip() == str(telegram_user_id):
                    nombre = str(fila.get("Nombre_Instructor", "")).strip()
                    if nombre:
                        return nombre
        except Exception as exc:
            logger.warning("No se pudo consultar hoja Instructores: %s", exc)
        return nombre_fallback

    def buscar_alumno(self, criterio: str) -> Optional[Dict[str, Any]]:
        if not criterio or criterio == "Desconocido":
            return None
        try:
            ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            registros = ws.get_all_records()
        except Exception as exc:
            logger.error("Error leyendo hoja Alumnos: %s", exc)
            return None

        criterio_limpio = _normalizar_texto(criterio)
        mapa_nombres: Dict[str, Dict[str, Any]] = {}

        for fila in registros:
            id_mat = str(fila.get("ID_Matricula", "")).strip()
            doc = str(fila.get("Documento", "")).strip()
            nombre = str(fila.get("Nombre_Completo", "")).strip()

            if criterio_limpio in (_normalizar_texto(id_mat), _normalizar_texto(doc)) and criterio_limpio:
                return fila

            nombre_norm = _normalizar_texto(nombre)
            if nombre_norm:
                if criterio_limpio == nombre_norm or criterio_limpio in nombre_norm:
                    return fila
                mapa_nombres[nombre_norm] = fila

        if mapa_nombres:
            coincidencias = difflib.get_close_matches(
                criterio_limpio, list(mapa_nombres.keys()), n=1, cutoff=0.65
            )
            if coincidencias:
                return mapa_nombres[coincidencias[0]]
        return None

    def registrar_clase(
        self,
        timestamp: str,
        fecha_reporte: str,
        hora_reporte: str,
        instructor: str,
        id_alumno: str,
        nombre_alumno: str,
        categoria: str,
        clases_reportadas: str,
        tema_registrado: str,
        url_foto_drive: str,
    ) -> None:
        ws = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
        ws.append_row(
            [
                timestamp,
                fecha_reporte,
                hora_reporte,
                instructor,
                id_alumno,
                nombre_alumno,
                normalizar_categoria(categoria),
                clases_reportadas,
                tema_registrado,
                url_foto_drive,
            ],
            value_input_option="USER_ENTERED",
        )

    def obtener_clases_alumno(self, criterio_alumno: str) -> Dict[str, Any]:
        info_alumno = self.buscar_alumno(criterio_alumno)
        ws_reg = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
        todas_filas = ws_reg.get_all_records()

        criterio_norm = _normalizar_texto(criterio_alumno)
        nombre_oficial_norm = (
            _normalizar_texto(info_alumno.get("Nombre_Completo", ""))
            if info_alumno
            else criterio_norm
        )
        id_oficial_norm = (
            _normalizar_texto(str(info_alumno.get("ID_Matricula", "")))
            if info_alumno
            else ""
        )
        doc_oficial_norm = (
            _normalizar_texto(str(info_alumno.get("Documento", "")))
            if info_alumno
            else ""
        )

        clases_filtradas = []
        for fila in todas_filas:
            id_fila = _normalizar_texto(str(fila.get("ID_Alumno", "")))
            nom_fila = _normalizar_texto(str(fila.get("Nombre_Alumno", "")))

            coincide_id = id_fila and id_fila in (id_oficial_norm, doc_oficial_norm, criterio_norm)
            coincide_nombre = nom_fila and (
                criterio_norm in nom_fila
                or (nombre_oficial_norm and nombre_oficial_norm in nom_fila)
                or (nombre_oficial_norm and nom_fila in nombre_oficial_norm)
            )
            if coincide_id or coincide_nombre:
                clases_filtradas.append(fila)

        if not info_alumno and clases_filtradas:
            ultima = clases_filtradas[-1]
            info_alumno = {
                "ID_Matricula": ultima.get("ID_Alumno") or "S/N",
                "Documento": ultima.get("ID_Alumno") or "Pendiente",
                "Nombre_Completo": ultima.get("Nombre_Alumno") or criterio_alumno,
                "Categoria": normalizar_categoria(str(ultima.get("Categoria", "B1"))),
                "Estado": "En curso",
            }

        return {"alumno": info_alumno, "registros": clases_filtradas}


sheets_service = SheetsService()


def subir_foto_evidencia(
    imagen_bytes: bytes,
    documento_alumno: str,
    nombre_alumno: str,
    clases_texto: str,
    fecha_reporte: str,
    fallback_file_id: str = "",
    fallback_url: str = "",
) -> str:
    if not GOOGLE_DRIVE_ROOT_FOLDER_ID:
        return fallback_url or f"telegram_file_id:{fallback_file_id}"
    try:
        creds = obtener_credenciales_google(["https://www.googleapis.com/auth/drive"])
        service = build("drive", "v3", credentials=creds, cache_discovery=False)
        nombre_archivo = (
            f"Clase_{clases_texto}_{fecha_reporte}.jpg".replace("/", "-").replace(" ", "_")
        )
        metadata = {"name": nombre_archivo, "parents": [GOOGLE_DRIVE_ROOT_FOLDER_ID]}
        media = MediaIoBaseUpload(
            io.BytesIO(imagen_bytes), mimetype="image/jpeg", resumable=False
        )
        archivo = (
            service.files()
            .create(body=metadata, media_body=media, fields="id, webViewLink")
            .execute()
        )
        return archivo.get(
            "webViewLink", f"https://drive.google.com/file/d/{archivo.get('id')}/view"
        )
    except Exception as exc:
        logger.warning("Usando enlace directo de Telegram como respaldo de foto: %s", exc)
        return fallback_url or f"telegram_file_id:{fallback_file_id}"


# =============================================================================
# 5. GENERADOR DE PDF OFICIAL (A.O.C.E.A - F - 59, 60, 61 y 62)
# =============================================================================
def mapear_clases_registradas(registros: List[Dict[str, Any]]) -> Dict[int, Dict[str, str]]:
    mapa: Dict[int, Dict[str, str]] = {}
    for fila in registros:
        raw_clases = str(fila.get("Clases_Reportadas", ""))
        nums = extraer_numeros_clase(raw_clases)
        fecha = str(fila.get("Fecha_Reporte", "")).strip()
        hora = str(fila.get("Hora_Reporte", "")).strip()
        instructor = str(fila.get("Instructor", "")).strip()
        tema_rep = str(fila.get("Tema_Registrado", "")).strip()
        for num in nums:
            mapa[num] = {
                "fecha": fecha,
                "hora": hora,
                "instructor": instructor,
                "tema_reportado": tema_rep,
            }
    return mapa


def generar_pdf_seguimiento(
    datos_alumno: Dict[str, Any], registros: List[Dict[str, Any]]
) -> io.BytesIO:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=1.0 * cm,
        rightMargin=1.0 * cm,
        topMargin=0.8 * cm,
        bottomMargin=0.8 * cm,
        title="Formato Oficial Seguimiento Clases CEA Auto Oriente",
    )

    categoria_key = normalizar_categoria(str(datos_alumno.get("Categoria", "B1")))
    cfg = FORMATOS_OFICIALES.get(categoria_key, FORMATOS_OFICIALES["B1"])
    mapa_clases = mapear_clases_registradas(registros)

    nombre_alumno = str(datos_alumno.get("Nombre_Completo", "")).strip()
    documento_alumno = str(datos_alumno.get("Documento", "")).strip()

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "TituloOficial",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.5,
        leading=10.5,
        alignment=TA_CENTER,
    )
    estilo_meta = ParagraphStyle(
        "MetaOficial",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.2,
        leading=9.0,
        alignment=TA_CENTER,
    )
    estilo_celda = ParagraphStyle(
        "CeldaTema",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.8,
        leading=8.2,
        alignment=TA_LEFT,
    )
    estilo_celda_centro = ParagraphStyle(
        "CeldaCentro",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.8,
        leading=8.2,
        alignment=TA_CENTER,
    )
    estilo_cabecera_tabla = ParagraphStyle(
        "CabeceraTabla",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.0,
        leading=8.5,
        alignment=TA_CENTER,
    )

    elementos = []
    logo_path = obtener_ruta_logo()
    if logo_path:
        celda_logo = Image(
            str(logo_path), width=4.6 * cm, height=1.25 * cm, kind="proportional"
        )
    else:
        celda_logo = Paragraph("<b>CEA AUTO ORIENTE</b>", estilo_titulo)

    titulo_html = cfg["titulo"].replace("\n", "<br/>")
    filas_encabezado = [
        [celda_logo, Paragraph(f"<b>{titulo_html}</b>", estilo_titulo), ""],
        [
            Paragraph(f"<b>Código:</b> {cfg['codigo_formato']}", estilo_meta),
            Paragraph(f"<b>Revisión:</b> {cfg['version']}", estilo_meta),
            Paragraph(f"<b>Fecha:</b> {cfg['fecha_formato']}", estilo_meta),
        ],
        [
            Paragraph(f"<b>Elaborado por:</b> {cfg['elaborado_por']}", estilo_meta),
            Paragraph(
                f"<b>Revisado y Aprobado por:</b> {cfg['aprobado_por']}", estilo_meta
            ),
            "",
        ],
    ]

    tabla_encabezado = Table(filas_encabezado, colWidths=[5.5 * cm, 8.5 * cm, 5.5 * cm])
    tabla_encabezado.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
                ("SPAN", (1, 0), (2, 0)),
                ("SPAN", (1, 2), (2, 2)),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    elementos.append(tabla_encabezado)
    elementos.append(Spacer(1, 0.15 * cm))

    filas_tabla = [
        [Paragraph("<b>CURSO DE TÉCNICAS DE CONDUCCIÓN</b>", estilo_titulo), "", "", "", ""],
        [
            Paragraph("<b>1.0</b>", estilo_cabecera_tabla),
            Paragraph("<b>clases</b>", estilo_cabecera_tabla),
            Paragraph("<b>TEMAS POR CLASES DE 45 MINUTOS</b>", estilo_cabecera_tabla),
            Paragraph("<b>Nombre del<br/>Firma Instructor</b>", estilo_cabecera_tabla),
            Paragraph("<b>Nombre del<br/>Firma Alumno</b>", estilo_cabecera_tabla),
        ],
    ]

    estilos_tabla = [
        ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 1), colors.HexColor("#EAEAEA")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]

    indices_pista = []
    indices_via = []

    for item in cfg["filas"]:
        idx_actual = len(filas_tabla)
        if item["etapa"] == "pista":
            indices_pista.append(idx_actual)
            etiqueta_etapa = Paragraph("<b>25 % en pista</b>", estilo_celda_centro)
        else:
            indices_via.append(idx_actual)
            etiqueta_etapa = Paragraph(
                "<b>75 % en vías urbanas y<br/>vías Nacionales</b>", estilo_celda_centro
            )

        reg_encontrado = None
        for num_c in item["clases"]:
            if num_c in mapa_clases:
                reg_encontrado = mapa_clases[num_c]
                break

        if reg_encontrado:
            texto_instructor = (
                f"<b>{reg_encontrado['instructor']}</b><br/>"
                f"<font size=6>{reg_encontrado['fecha']} {reg_encontrado['hora']}</font>"
            )
            texto_alumno = f"{nombre_alumno}"
        else:
            texto_instructor = ""
            texto_alumno = ""

        tema_html = item["tema"].replace("\n", "<br/>")
        filas_tabla.append(
            [
                etiqueta_etapa,
                Paragraph(f"<b>{item['etiqueta']}</b>", estilo_celda_centro),
                Paragraph(tema_html, estilo_celda),
                Paragraph(texto_instructor, estilo_celda_centro),
                Paragraph(texto_alumno, estilo_celda_centro),
            ]
        )

    if len(indices_pista) > 1:
        estilos_tabla.append(("SPAN", (0, indices_pista[0]), (0, indices_pista[-1])))
    if len(indices_via) > 1:
        estilos_tabla.append(("SPAN", (0, indices_via[0]), (0, indices_via[-1])))

    tabla_clases = Table(
        filas_tabla,
        colWidths=[2.3 * cm, 1.3 * cm, 9.5 * cm, 3.2 * cm, 3.2 * cm],
        repeatRows=2,
    )
    tabla_clases.setStyle(TableStyle(estilos_tabla))
    elementos.append(tabla_clases)
    elementos.append(Spacer(1, 0.15 * cm))

    filas_pie = [
        [Paragraph("<b>OBSERVACIONES AL PROCESO DE APRENDIZAJE:</b>", estilo_celda), ""],
        [Paragraph("&nbsp;", estilo_celda), ""],
        [
            Paragraph(
                "Concluida las clases se realizara evaluación práctica, calificada y aprobada por SICOV.",
                estilo_celda,
            ),
            "",
        ],
        [
            Paragraph(
                "En constancia de haber realizado las clases de Técnicas de Conducción por el alumno abajo identificado,",
                estilo_celda,
            ),
            "",
        ],
        [
            Paragraph(
                f"<b>Nombre Alumno:</b> {nombre_alumno}<br/><br/>"
                f"<b>Firma del alumno:</b> ____________________________________<br/>"
                f"<b>CC, TI, CE.:</b> {documento_alumno}",
                estilo_celda,
            ),
            Paragraph(
                "<br/><br/>____________________________________<br/><b>Recibido de Archivo</b>",
                estilo_celda_centro,
            ),
        ],
    ]

    tabla_pie = Table(filas_pie, colWidths=[11.5 * cm, 8.0 * cm])
    tabla_pie.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, 1), 0.6, colors.black),
                ("LINEBELOW", (0, 0), (-1, 0), 0.4, colors.grey),
                ("SPAN", (0, 0), (1, 0)),
                ("SPAN", (0, 1), (1, 1)),
                ("SPAN", (0, 2), (1, 2)),
                ("SPAN", (0, 3), (1, 3)),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    elementos.append(tabla_pie)

    doc.build(elementos)
    buffer.seek(0)
    return buffer


# =============================================================================
# 6. MANEJADORES DE TELEGRAM
# =============================================================================
async def procesar_reporte_foto(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    if not msg or not msg.photo or not msg.caption:
        return

    caption = msg.caption.strip()
    user = update.effective_user
    nombre_telegram = user.full_name if user else "Instructor"
    telegram_user_id = user.id if user else 0

    dt_col = msg.date.astimezone(TZ_COLOMBIA)
    timestamp_str = dt_col.strftime("%Y-%m-%d %H:%M:%S")
    fecha_reporte = dt_col.strftime("%Y-%m-%d")
    hora_reporte = dt_col.strftime("%H:%M")

    try:
        foto_obj = await msg.photo[-1].get_file()
        foto_bytes = bytes(await foto_obj.download_as_bytearray())

        reporte = parsear_reporte(caption, imagen_bytes=foto_bytes)
        if reporte.fecha_timemark:
            fecha_reporte = reporte.fecha_timemark
        if reporte.hora_timemark:
            hora_reporte = reporte.hora_timemark

        instructor = sheets_service.buscar_instructor(telegram_user_id, nombre_telegram)
        if reporte.instructor_detectado and instructor == nombre_telegram:
            instructor = reporte.instructor_detectado

        ficha_alumno = sheets_service.buscar_alumno(reporte.alumno)
        if ficha_alumno:
            id_alumno = str(
                ficha_alumno.get("Documento") or ficha_alumno.get("ID_Matricula") or "S/N"
            ).strip()
            nombre_alumno_oficial = str(
                ficha_alumno.get("Nombre_Completo") or reporte.alumno
            ).strip()
            categoria = normalizar_categoria(str(ficha_alumno.get("Categoria", "B1")))
        else:
            id_alumno = "Pendiente"
            nombre_alumno_oficial = reporte.alumno
            categoria = "C1" if any(c > 20 for c in reporte.clases_lista) else "B1"

        tema_final = reporte.tema
        if tema_final == "Sin especificar" and reporte.clases_lista:
            tema_final = obtener_tema_oficial(categoria, reporte.clases_lista)

        url_foto_drive = subir_foto_evidencia(
            imagen_bytes=foto_bytes,
            documento_alumno=id_alumno,
            nombre_alumno=nombre_alumno_oficial,
            clases_texto=reporte.clases_texto,
            fecha_reporte=fecha_reporte,
            fallback_file_id=foto_obj.file_id,
            fallback_url=foto_obj.file_path or "",
        )

        sheets_service.registrar_clase(
            timestamp=timestamp_str,
            fecha_reporte=fecha_reporte,
            hora_reporte=hora_reporte,
            instructor=instructor,
            id_alumno=id_alumno,
            nombre_alumno=nombre_alumno_oficial,
            categoria=categoria,
            clases_reportadas=reporte.clases_texto,
            tema_registrado=tema_final,
            url_foto_drive=url_foto_drive,
        )

        try:
            await msg.set_reaction(reaction=[ReactionTypeEmoji("✅")])
        except Exception:
            pass

        await msg.reply_text(
            f"✅ Clase {reporte.clases_texto} registrada para {nombre_alumno_oficial} "
            f"(Categoría {categoria} | Instructor: {instructor})"
        )
    except Exception as exc:
        logger.exception("Error procesando el reporte fotográfico: %s", exc)
        await msg.reply_text(f"⚠️ Error al registrar en Google Sheets: {exc}")


async def cmd_ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    texto = (
        "🚗 *Bot de Seguimiento de Clases — CEA Auto Oriente*\n\n"
        "📸 *1. Registro de Clases (Instructores):*\n"
        "Envía una foto con el nombre del alumno, número de clase y tema.\n\n"
        "📄 *2. Generar Formato PDF para Carpeta/Matrícula:*\n"
        "`/reporte [Documento o Nombre del Alumno]`\n\n"
        "⚙️ *3. Inicializar Hojas en Google Sheets:*\n"
        "`/inicializar`"
    )
    await update.effective_message.reply_text(texto, parse_mode="Markdown")


async def cmd_inicializar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    try:
        sheets_service.inicializar_estructura()
        await msg.reply_text(
            "✅ Hojas `Alumnos`, `Instructores` y `Registro_Clases` verificadas en Google Sheets.",
            parse_mode="Markdown",
        )
    except Exception as exc:
        await msg.reply_text(f"❌ Error conectando con Google Sheets: {exc}")


async def cmd_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    if not context.args:
        await msg.reply_text(
            "⚠️ Indica el nombre o documento del alumno.\nEjemplo: `/reporte Camilo Torres`",
            parse_mode="Markdown",
        )
        return

    criterio = " ".join(context.args).strip()
    aviso = await msg.reply_text(
        f"⏳ Generando formato PDF para *{criterio}*...", parse_mode="Markdown"
    )

    try:
        resultado = sheets_service.obtener_clases_alumno(criterio)
        datos_alumno = resultado["alumno"]
        registros = resultado["registros"]

        if not datos_alumno and not registros:
            await aviso.edit_text(
                f"❌ No se encontraron registros para: *{criterio}*.",
                parse_mode="Markdown",
            )
            return

        pdf_buffer = generar_pdf_seguimiento(datos_alumno, registros)
        categoria = normalizar_categoria(str(datos_alumno.get("Categoria", "B1")))
        codigo_fmt = FORMATOS_OFICIALES.get(categoria, FORMATOS_OFICIALES["B1"])[
            "codigo_formato"
        ].replace(" ", "")
        nombre_limpio = (
            str(datos_alumno.get("Nombre_Completo", criterio)).strip().replace(" ", "_")
        )
        filename = f"{codigo_fmt}_{categoria}_{nombre_limpio}.pdf"

        await msg.reply_document(
            document=pdf_buffer,
            filename=filename,
            caption=(
                f"📄 *Formato Oficial {codigo_fmt}*\n"
                f"👤 *Alumno:* {datos_alumno.get('Nombre_Completo')}\n"
                f"🪪 *Documento:* {datos_alumno.get('Documento')}\n"
                f"🏷 *Categoría:* {categoria}\n"
                f"📊 *Reportes registrados:* {len(registros)}"
            ),
            parse_mode="Markdown",
        )
        await aviso.delete()
    except Exception as exc:
        logger.exception("Error generando reporte PDF: %s", exc)
        await aviso.edit_text(f"⚠️ Error al generar el PDF: {exc}")


# =============================================================================
# 7. SERVIDOR HEALTH CHECK (RENDER) Y ARRANQUE DEL BOT
# =============================================================================
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write("Bot CEA Auto Oriente activo (200 OK)".encode("utf-8"))

    def log_message(self, format, *args):
        return


def iniciar_servidor_healthcheck() -> None:
    try:
        server = HTTPServer(("0.0.0.0", PORT), HealthCheckHandler)
        hilo = threading.Thread(target=server.serve_forever, daemon=True)
        hilo.start()
        logger.info("🌐 Servidor Health Check activo en puerto %s", PORT)
    except Exception as exc:
        logger.warning("No se pudo iniciar servidor Health Check: %s", exc)


def main() -> None:
    if not TELEGRAM_BOT_TOKEN or TELEGRAM_BOT_TOKEN == "TU_TELEGRAM_BOT_TOKEN":
        logger.error("⚠️ Configura TELEGRAM_BOT_TOKEN en .env o en Render.")
        sys.exit(1)

    iniciar_servidor_healthcheck()

    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler(["start", "ayuda", "help"], cmd_ayuda))
    app.add_handler(CommandHandler("inicializar", cmd_inicializar))
    app.add_handler(CommandHandler("reporte", cmd_reporte))

    filtro_reporte = filters.PHOTO & filters.CaptionRegex(r"(?i)\b(clases?|cl\.?)\b")
    app.add_handler(MessageHandler(filtro_reporte, procesar_reporte_foto))

    logger.info("🚗 Bot CEA Auto Oriente iniciado y escuchando mensajes...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
