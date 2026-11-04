"""
Bot de Seguimiento de Clases — CEA Auto Oriente (Versión con Selección de 1 o 2 Clases + Validación de Nombre y Cédula).
Incluye:
- Flujo interactivo con Botones para el Instructor:
  1) El instructor envía la foto.
  2) El bot despliega el MENÚ DE CATEGORÍAS (A1, A2, B1, C1, Recat C1).
  3) En el SUBMENÚ DE CLASES, el instructor puede marcar 1 sola clase (ej. Clase 15) o 2 clases al tiempo (ej. Clase 1 y Clase 2) con casillas ✅ y pulsar "Confirmar Clases".
  4) El bot solicita con indicaciones claras el NOMBRE COMPLETO (Nombre y Apellido) y NÚMERO DE CÉDULA (CC/TI/CE) para que Google Sheets quede 100% preciso (y registra automáticamente al alumno en la hoja 'Alumnos' si es nuevo).
- Comando /reporte [Cédula o Nombre] para emitir el PDF oficial con el logo de Auto Oriente.
"""

import difflib
import io
import json
import logging
import os
import re
import sys
import threading
import time
import unicodedata
import urllib.request
from dataclasses import dataclass, field
from datetime import datetime, timedelta, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

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
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReactionTypeEmoji,
    Update,
)
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
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
_RAW_SHEET_ID = os.getenv(
    "GOOGLE_SHEET_ID", "1xouy9DMkp4oZffZVx9E1y2IwY5qxRi5u2yr3p5wpj6U"
).strip()
GOOGLE_SHEET_ID = (
    "1xouy9DMkp4oZffZVx9E1y2IwY5qxRi5u2yr3p5wpj6U"
    if _RAW_SHEET_ID.startswith("{")
    else _RAW_SHEET_ID
)
GOOGLE_CREDENTIALS_FILE = os.getenv(
    "GOOGLE_CREDENTIALS_FILE", str(BASE_DIR / "credentials.json")
).strip()
GOOGLE_CREDENTIALS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON", "").strip()
GOOGLE_DRIVE_ROOT_FOLDER_ID = os.getenv("GOOGLE_DRIVE_ROOT_FOLDER_ID", "").strip()
RENDER_EXTERNAL_URL = os.getenv(
    "RENDER_EXTERNAL_URL", "https://bot-clases-autooriente.onrender.com"
).strip()
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

# Sesiones interactivas por (chat_id, user_id)
SESIONES_REPORTE: Dict[Tuple[int, int], Dict[str, Any]] = {}


def obtener_ruta_logo() -> Optional[Path]:
    candidatos = [
        BASE_DIR / "logo_autooriente.png",
        BASE_DIR / "assets" / "logo_autooriente.png",
    ]
    for ruta in candidatos:
        if ruta.exists():
            return ruta
    return None


def obtener_credenciales_google(scopes: List[str]) -> Credentials:
    for raw_val in (GOOGLE_CREDENTIALS_JSON, GOOGLE_CREDENTIALS_FILE, _RAW_SHEET_ID):
        if raw_val and raw_val.lstrip().startswith("{"):
            info = json.loads(raw_val)
            return Credentials.from_service_account_info(info, scopes=scopes)

    for candidate in (
        Path("/etc/secrets/credentials.json"),
        BASE_DIR / "credentials.json",
        Path(GOOGLE_CREDENTIALS_FILE),
    ):
        if candidate.is_file():
            return Credentials.from_service_account_file(str(candidate), scopes=scopes)

    raise FileNotFoundError(
        "Falta configurar GOOGLE_CREDENTIALS_JSON o Secret File credentials.json en Render -> Environment."
    )


# =============================================================================
# 2. CATÁLOGO OFICIAL DE FORMATOS Y TEMARIOS (F-59, F-60, F-61 y F-62)
# =============================================================================
FORMATOS_OFICIALES: Dict[str, Dict[str, Any]] = {
    "A1": {
        "nombre_corto": "🏍️ A1 (Motocicleta)",
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
                "resumen": "Inspección preoperacional y mandos",
                "tema": "Inspección pre operacional, documentos y equipos de prevención, adaptación viso espacial, ajuste del asiento, timón y espejos.\nExplicación elementos de seguridad activa y pasiva. Identificación del tablero de instrumentos, mandos y pedales. Encendido del motor y testigos.",
            },
            {
                "etiqueta": "2.0",
                "clases": [2],
                "etapa": "pista",
                "resumen": "Punto de contacto y selector",
                "tema": "Maniobra de puesta en marcha con técnica punto de contacto del embrague. Conducción del vehículo hacia adelante.\nOperación del selector del vehículo. Coordinación: Acelerador - freno - embrague - desaceleración.",
            },
            {
                "etiqueta": "3.0",
                "clases": [3],
                "etapa": "via",
                "resumen": "Giros con conos, eslalon y parqueo",
                "tema": "Conducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Eslalon.\nConducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Primeros indicios de parqueo.",
            },
            {
                "etiqueta": "4.0",
                "clases": [4],
                "etapa": "via",
                "resumen": "Vía urbana terreno plano (pares/arranques)",
                "tema": "Conducción del vehículo en vía urbana en terreno plano. Pares y arranques.",
            },
            {
                "etiqueta": "5.0",
                "clases": [5],
                "etapa": "via",
                "resumen": "Terreno pendiente y adelantamiento",
                "tema": "Arranque del vehículo en terreno pendiente, de frente y reversa.\nConducción en terreno plano con maniobras de adelantamiento.",
            },
            {
                "etiqueta": "6.0",
                "clases": [6],
                "etapa": "via",
                "resumen": "Parqueo y señales lumínicas/cambio carril",
                "tema": "Parqueo y estacionamiento con y sin uso de espejo. Refuerzo.\nConducción en vías urbanas usando las señales lumínicas, corporales y acústicas. Cambio de calzada. Cambio de carril.",
            },
            {
                "etiqueta": "7.0",
                "clases": [7],
                "etapa": "via",
                "resumen": "Glorietas, puentes y tráfico masivo",
                "tema": "Conducción en vías urbanas, glorietas, intersecciones, deprimidos, puentes elevados, zonas escolares.\nConducción en vías urbanas (Buses, taxis, trasporte Masivo)",
            },
            {
                "etiqueta": "8.0",
                "clases": [8],
                "etapa": "via",
                "resumen": "Vías principales y vías nacionales",
                "tema": "Conducción en vías principales urbanas de la ciudad.\nConducción en vías nacionales. Piedecuesta - Bucaramanga - Girón - peaje vía aeropuerto.",
            },
        ],
    },
    "A2": {
        "nombre_corto": "🏍️ A2 (Moto 15 clases)",
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
                "resumen": "Inspección preoperacional y postura",
                "tema": "Inspección pre operacional, documentos y equipos de prevención, adaptación viso espacial, correcta postura en el manillar y espejos.",
            },
            {
                "etiqueta": "2.0",
                "clases": [2],
                "etapa": "pista",
                "resumen": "Seguridad activa/pasiva y encendido",
                "tema": "Explicación elementos de seguridad activa y pasiva. Identificación del tablero de instrumentos, mandos y pedales. Encendido del motor y testigos.",
            },
            {
                "etiqueta": "3.0",
                "clases": [3],
                "etapa": "pista",
                "resumen": "Carreteo y punto de contacto embrague",
                "tema": "Maniobra de carreteo o maniobra cero cale, maniobra de puesta en marcha con técnica punto de contacto del embrague.",
            },
            {
                "etiqueta": "4.0",
                "clases": [4],
                "etapa": "pista",
                "resumen": "Selector: acelerador, freno y embrague",
                "tema": "Operación del selector del vehículo. Coordinación: Acelerador - freno - embrague - desaceleración.",
            },
            {
                "etiqueta": "5.0",
                "clases": [5],
                "etapa": "pista",
                "resumen": "Franja limitada, equilibrio y Giro en U",
                "tema": "Circular sobre una franja de anchura limitada, ejercicio de equilibrio y control de la motocicleta, Giro en U.",
            },
            {
                "etiqueta": "6.0",
                "clases": [6],
                "etapa": "via",
                "resumen": "Eslalon con conos a baja velocidad",
                "tema": "Conducción del vehículo hacia adelante con giros haciendo uso conos Eslalon a baja velocidad.",
            },
            {
                "etiqueta": "7.0",
                "clases": [7],
                "etapa": "via",
                "resumen": "Giros con conos y frenado controlado",
                "tema": "Conducción del vehículo hacia adelante con giros haciendo uso conos Aceleración, embrague y frenado Controlado.",
            },
            {
                "etiqueta": "8.0",
                "clases": [8],
                "etapa": "via",
                "resumen": "Obstáculos y frenado de emergencia",
                "tema": "Conducción: sorteo de obstáculos a alta velocidad (máx. 40 km/h) y frenado de emergencia. Slalom, cajón de frenado y evasión de barrera, Eslalon a alta velocidad",
            },
            {
                "etiqueta": "9.0",
                "clases": [9],
                "etapa": "via",
                "resumen": "Arranque en plano y desplazamientos en 8",
                "tema": "Arranque del vehículo en terreno plano y Desplazamientos en ocho.",
            },
            {
                "etiqueta": "10.0",
                "clases": [10],
                "etapa": "via",
                "resumen": "Terreno inclinado con/sin acompañante",
                "tema": "Arranque del vehículo en terreno inclinado, con y sin acompañante.",
            },
            {
                "etiqueta": "11.0",
                "clases": [11],
                "etapa": "via",
                "resumen": "Avance, detención y señales de tránsito",
                "tema": "Maniobra de avance y detención, Interacción con señales de tránsito horizontales y verticales.",
            },
            {
                "etiqueta": "12.0",
                "clases": [12],
                "etapa": "via",
                "resumen": "Señales lumínicas y cambio de carril",
                "tema": "Conducción en vías usando las señales lumínicas, corporales y acústicas. Cambio de calzada. Cambio de carril.",
            },
            {
                "etiqueta": "13.0",
                "clases": [13],
                "etapa": "via",
                "resumen": "Pares, arranques y señales en barrio",
                "tema": "Conducción del vehículo en vía en terreno plano. Pares y arranques. Barrio de poco flujo vehicular, interacción con señales de tránsito",
            },
            {
                "etiqueta": "14.0",
                "clases": [14],
                "etapa": "via",
                "resumen": "Refuerzo de pruebas CALE",
                "tema": "Refuerzo de pruebas CALE según resolución del Ministerio de Transporte.",
            },
            {
                "etiqueta": "15.0",
                "clases": [15],
                "etapa": "via",
                "resumen": "Evaluación general curso motocicleta",
                "tema": "Evaluación general del curso de conducción de motocicleta, interacción con señales de tránsito.",
            },
        ],
    },
    "B1": {
        "nombre_corto": "🚗 B1 (Carro Particular 20 clases)",
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
                "resumen": "Inspección preoperacional y adaptación",
                "tema": "Inspección pre operacional, documentos y equipos de prevención, adaptación viso espacial, ajuste del asiento, timón y espejos.",
            },
            {
                "etiqueta": "2.0",
                "clases": [2],
                "etapa": "pista",
                "resumen": "Seguridad activa/pasiva, mandos y motor",
                "tema": "Explicación elementos de seguridad activa y pasiva. Identificación del tablero de instrumentos, mandos y pedales. Encendido del motor y testigos.",
            },
            {
                "etiqueta": "3.0",
                "clases": [3],
                "etapa": "pista",
                "resumen": "Arranque con punto de contacto",
                "tema": "Maniobra de puesta en marcha con técnica punto de contacto del embrague. Conducción del vehículo hacia adelante.",
            },
            {
                "etiqueta": "4.0",
                "clases": [4],
                "etapa": "pista",
                "resumen": "Selector: acelerador, freno y embrague",
                "tema": "Operación del selector del vehículo. Coordinación: Acelerador - freno - embrague - desaceleración.",
            },
            {
                "etiqueta": "5.0",
                "clases": [5],
                "etapa": "pista",
                "resumen": "Giros con conos y eslalon",
                "tema": "Conducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Eslalon.",
            },
            {
                "etiqueta": "6.0",
                "clases": [6],
                "etapa": "pista",
                "resumen": "Giros con conos e indicios de parqueo",
                "tema": "Conducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Primeros indicios de parqueo.",
            },
            {
                "etiqueta": "7.0",
                "clases": [7],
                "etapa": "via",
                "resumen": "Vía urbana terreno plano (pares/arranques)",
                "tema": "Conducción del vehículo en vía urbana en terreno plano. Pares y arranques.",
            },
            {
                "etiqueta": "8.0",
                "clases": [8],
                "etapa": "via",
                "resumen": "Vía urbana terreno plano (pares/arranques)",
                "tema": "Conducción del vehículo en vía urbana en terreno plano. Pares y arranques.",
            },
            {
                "etiqueta": "9.0",
                "clases": [9],
                "etapa": "via",
                "resumen": "Arranque en pendiente frente y reversa",
                "tema": "Arranque del vehículo en terreno pendiente, de frente y reversa.",
            },
            {
                "etiqueta": "10.0",
                "clases": [10],
                "etapa": "via",
                "resumen": "Arranque en pendiente frente y reversa",
                "tema": "Arranque del vehículo en terreno pendiente, de frente y reversa.",
            },
            {
                "etiqueta": "11.0",
                "clases": [11],
                "etapa": "via",
                "resumen": "Terreno plano y adelantamiento",
                "tema": "Conducción en terreno plano con maniobras de adelantamiento.",
            },
            {
                "etiqueta": "12.0",
                "clases": [12],
                "etapa": "via",
                "resumen": "Parqueo y estacionamiento con/sin espejo",
                "tema": "Parqueo y estacionamiento con y sin uso de espejo. Refuerzo.",
            },
            {
                "etiqueta": "13.0",
                "clases": [13],
                "etapa": "via",
                "resumen": "Parqueo y estacionamiento (Refuerzo)",
                "tema": "Parqueo y estacionamiento con y sin uso de espejo. Refuerzo.",
            },
            {
                "etiqueta": "14.0",
                "clases": [14],
                "etapa": "via",
                "resumen": "Señales lumínicas y cambio de carril",
                "tema": "Conducción en vías urbanas usando las señales lumínicas, corporales y acústicas. Cambio de calzada. Cambio de carril.",
            },
            {
                "etiqueta": "15.0",
                "clases": [15],
                "etapa": "via",
                "resumen": "Glorietas, puentes y zonas escolares",
                "tema": "Conducción en vías urbanas, glorietas, intersecciones, deprimidos, puentes elevados, zonas escolares.",
            },
            {
                "etiqueta": "16.0",
                "clases": [16],
                "etapa": "via",
                "resumen": "Tráfico masivo (buses y taxis)",
                "tema": "Conducción en vías urbanas (Buses, taxis, trasporte Masivo)",
            },
            {
                "etiqueta": "17.0",
                "clases": [17],
                "etapa": "via",
                "resumen": "Tráfico masivo (buses y taxis)",
                "tema": "Conducción en vías urbanas (Buses, taxis, trasporte Masivo)",
            },
            {
                "etiqueta": "18.0",
                "clases": [18],
                "etapa": "via",
                "resumen": "Vías principales urbanas de la ciudad",
                "tema": "Conducción en vías principales urbanas de la ciudad.",
            },
            {
                "etiqueta": "19.0",
                "clases": [19],
                "etapa": "via",
                "resumen": "Vías principales urbanas de la ciudad",
                "tema": "Conducción en vías principales urbanas de la ciudad.",
            },
            {
                "etiqueta": "20.0",
                "clases": [20],
                "etapa": "via",
                "resumen": "Vías nacionales (Piedecuesta-Girón-Aeropuerto)",
                "tema": "Conducción en vías nacionales. Piedecuesta - Bucaramanga - Girón - peaje vía aeropuerto.",
            },
        ],
    },
    "C1": {
        "nombre_corto": "🚕 C1 (Servicio Público 30 clases)",
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
                "resumen": "Inspección preoperacional y adaptación",
                "tema": "Inspección pre operacional, documentos y equipos de prevención, adaptación viso espacial, ajuste del asiento, timón y espejos.",
            },
            {
                "etiqueta": "2.0",
                "clases": [2],
                "etapa": "pista",
                "resumen": "Seguridad activa/pasiva y encendido",
                "tema": "Explicación elementos de seguridad activa y pasiva. Identificación del tablero de instrumentos, mandos y pedales. Encendido del motor y testigos.",
            },
            {
                "etiqueta": "3,4",
                "clases": [3, 4],
                "etapa": "pista",
                "resumen": "Arranque con punto de contacto",
                "tema": "Maniobra de puesta en marcha con técnica punto de contacto del embrague. Conducción del vehículo hacia adelante.",
            },
            {
                "etiqueta": "5,6",
                "clases": [5, 6],
                "etapa": "pista",
                "resumen": "Selector: acelerador, freno y embrague",
                "tema": "Operación del selector del vehículo. Coordinación: Acelerador - freno - embrague - desaceleración.",
            },
            {
                "etiqueta": "7,8",
                "clases": [7, 8],
                "etapa": "pista",
                "resumen": "Giros con conos y eslalon",
                "tema": "Conducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Eslalon.",
            },
            {
                "etiqueta": "9,10",
                "clases": [9, 10],
                "etapa": "pista",
                "resumen": "Conos y primeros indicios de parqueo",
                "tema": "Conducción del vehículo hacia atrás y hacia adelante con giros haciendo uso de los conos. Primeros indicios de parqueo.",
            },
            {
                "etiqueta": "11,12",
                "clases": [11, 12],
                "etapa": "via",
                "resumen": "Vía urbana en plano (pares y arranques)",
                "tema": "Conducción del vehículo en vía urbana en terreno plano. Pares y arranques.",
            },
            {
                "etiqueta": "13,14",
                "clases": [13, 14],
                "etapa": "via",
                "resumen": "Vía urbana en plano (pares y arranques)",
                "tema": "Conducción del vehículo en vía urbana en terreno plano. Pares y arranques.",
            },
            {
                "etiqueta": "15,16",
                "clases": [15, 16],
                "etapa": "via",
                "resumen": "Arranque en pendiente frente y reversa",
                "tema": "Arranque del vehículo en terreno pendiente, de frente y reversa.",
            },
            {
                "etiqueta": "17,18",
                "clases": [17, 18],
                "etapa": "via",
                "resumen": "Terreno plano y adelantamiento",
                "tema": "Conducción en terreno plano con maniobras de adelantamiento.",
            },
            {
                "etiqueta": "19,20",
                "clases": [19, 20],
                "etapa": "via",
                "resumen": "Parqueo y estacionamiento con/sin espejo",
                "tema": "Parqueo y estacionamiento con y sin uso de espejo. Refuerzo.",
            },
            {
                "etiqueta": "21,22",
                "clases": [21, 22],
                "etapa": "via",
                "resumen": "Señales lumínicas y cambio de carril",
                "tema": "Conducción en vías urbanas usando las señales lumínicas, corporales y acústicas. Cambio de calzada. Cambio de carril.",
            },
            {
                "etiqueta": "23,24",
                "clases": [23, 24],
                "etapa": "via",
                "resumen": "Glorietas, puentes y zonas escolares",
                "tema": "Conducción en vías urbanas, glorietas, intersecciones, deprimidos, puentes elevados, zonas escolares.",
            },
            {
                "etiqueta": "25,26",
                "clases": [25, 26],
                "etapa": "via",
                "resumen": "Tráfico masivo (buses y taxis)",
                "tema": "Conducción en vías urbanas (Buses, taxis, trasporte Masivo)",
            },
            {
                "etiqueta": "27,28",
                "clases": [27, 28],
                "etapa": "via",
                "resumen": "Vías principales urbanas de la ciudad",
                "tema": "Conducción en vías principales urbanas de la ciudad.",
            },
            {
                "etiqueta": "29,30",
                "clases": [29, 30],
                "etapa": "via",
                "resumen": "Vías nacionales (Piedecuesta-Girón-Aeropuerto)",
                "tema": "Conducción en vías nacionales. Piedecuesta - Bucaramanga - Girón - peaje vía aeropuerto.",
            },
        ],
    },
    "RECAT_C1": {
        "nombre_corto": "🔄 Recategorización C1 (Clases 21-30)",
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
                "resumen": "Señales lumínicas y cambio de carril",
                "tema": "Conducción en vías urbanas usando las señales lumínicas, corporales y acústicas. Cambio de calzada. Cambio de carril.",
            },
            {
                "etiqueta": "23,24",
                "clases": [23, 24],
                "etapa": "via",
                "resumen": "Glorietas, puentes y zonas escolares",
                "tema": "Conducción en vías urbanas, glorietas, intersecciones, deprimidos, puentes elevados, zonas escolares.",
            },
            {
                "etiqueta": "25,26",
                "clases": [25, 26],
                "etapa": "via",
                "resumen": "Tráfico masivo (buses y taxis)",
                "tema": "Conducción en vías urbanas (Buses, taxis, trasporte Masivo)",
            },
            {
                "etiqueta": "27,28",
                "clases": [27, 28],
                "etapa": "via",
                "resumen": "Vías principales urbanas de la ciudad",
                "tema": "Conducción en vías principales urbanas de la ciudad.",
            },
            {
                "etiqueta": "29,30",
                "clases": [29, 30],
                "etapa": "via",
                "resumen": "Vías nacionales (Piedecuesta-Girón-Aeropuerto)",
                "tema": "Conducción en vías nacionales. Piedecuesta - Bucaramanga - Girón - peaje vía aeropuerto.",
            },
        ],
    },
}


def normalizar_categoria(cat: str) -> str:
    if not cat:
        return "B1"
    limpio = cat.strip().upper().replace("-", "_").replace(" ", "_")
    if "RECAT" in limpio or limpio == "RC1":
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
    temas_vistos = []
    for c in clases:
        for fila in formato["filas"]:
            if c in fila["clases"]:
                t = fila["tema"].replace("\n", " ")
                if t not in temas_vistos:
                    temas_vistos.append(t)
    if temas_vistos:
        return " / ".join(temas_vistos)
    return f"Práctica de conducción clase {', '.join(map(str, clases))}"


# =============================================================================
# 3. EXTRACTOR DE REPORTES Y VALIDACIÓN DE NOMBRE + CÉDULA
# =============================================================================
@dataclass
class ReporteExtraido:
    alumno: str = "Desconocido"
    documento_detectado: Optional[str] = None
    categoria_detectada: Optional[str] = None
    clases_lista: List[int] = field(default_factory=list)
    clases_texto: str = "No identificada"
    tema: str = "Sin especificar"


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
    return sorted(vistos)


def formatear_clases(clases: List[int]) -> str:
    if not clases:
        return "No identificada"
    ordenadas = sorted(set(clases))
    if len(ordenadas) == 1:
        return str(ordenadas[0])
    return "-".join(str(c) for c in ordenadas)


def extraer_nombre_y_cedula(texto: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Extrae el Nombre Completo y el Número de Documento (CC/TI/CE) de un texto como:
      - 'Sergio Andrés Pérez 1098765432'
      - 'CC 1098765432 - Sergio Andrés Pérez'
      - 'Alumno: Camilo Torres CC: 13959739'
    """
    if not texto or not texto.strip():
        return None, None

    limpio = texto.strip()
    # Buscar número de documento de 6 a 12 dígitos
    m_doc = re.search(r"\b(\d{6,12})\b", limpio)
    documento = m_doc.group(1) if m_doc else None

    # Limpiar palabras como CC, TI, CE, Alumno, Cédula y el número para dejar solo el nombre
    sin_doc = re.sub(r"\b\d{5,12}\b", "", limpio)
    sin_prefijos = re.sub(
        r"\b(?:alumn[oa]s?|nombre|cc|ti|ce|c[eé]dula|documento|doc|id)\b\s*[:#.\-]?",
        "",
        sin_doc,
        flags=re.IGNORECASE,
    )
    # Quitar caracteres especiales sobrantes
    nombre_limpio = re.sub(r"[^A-Za-zÁÉÍÓÚáéíóúÑñÜü\s]", " ", sin_prefijos)
    palabras = [p for p in nombre_limpio.split() if len(p) >= 2]

    nombre_formateado = " ".join(palabras).title() if palabras else None
    return nombre_formateado, documento


def parsear_reporte_regex(texto: str) -> ReporteExtraido:
    resultado = ReporteExtraido()
    if not texto or not texto.strip():
        return resultado

    texto_limpio = re.sub(
        r"\d{1,2}:\d{2}\s*(?:AM|PM|a\.?\s*m\.?|p\.?\s*m\.?)\s*$",
        "",
        texto.strip(),
        flags=re.IGNORECASE,
    )

    m_cat = re.search(
        r"\b(recat[_\s\-]*c1|rc1|a1|a2|b1|c1)\b", texto_limpio, re.IGNORECASE
    )
    if m_cat:
        resultado.categoria_detectada = normalizar_categoria(m_cat.group(1))

    patron_clase = re.compile(
        r"(?:clases?|cl\.?)\s*[:#.]?\s*(\d{1,2}(?:\s*(?:[-/,yY&aA]|\s+y\s+)\s*\d{1,2})*)",
        re.IGNORECASE,
    )
    m_clase = patron_clase.search(texto_limpio)
    if m_clase:
        nums = extraer_numeros_clase(m_clase.group(1))
        if nums:
            resultado.clases_lista = nums
            resultado.clases_texto = formatear_clases(nums)

    # Extraer posible nombre y cédula quitando la parte de clase
    texto_sin_clase = patron_clase.sub("", texto_limpio)
    nombre, doc = extraer_nombre_y_cedula(texto_sin_clase)
    if nombre:
        resultado.alumno = nombre
    if doc:
        resultado.documento_detectado = doc

    return resultado


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

    def listar_alumnos_activos(self, categoria: Optional[str] = None) -> List[Dict[str, Any]]:
        try:
            ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            registros = ws.get_all_records()
            activos = []
            for f in registros:
                nombre = str(f.get("Nombre_Completo", "")).strip()
                if not nombre:
                    continue
                estado = str(f.get("Estado", "En curso")).strip().lower()
                if "finalizado" in estado:
                    continue
                if categoria:
                    cat_alumno = normalizar_categoria(str(f.get("Categoria", "")))
                    if cat_alumno != normalizar_categoria(categoria):
                        continue
                activos.append(f)
            return activos if activos else [f for f in registros if str(f.get("Nombre_Completo", "")).strip()]
        except Exception as exc:
            logger.warning("No se pudo listar alumnos activos: %s", exc)
            return []

    def registrar_o_actualizar_alumno(
        self, documento: str, nombre_completo: str, categoria: str
    ) -> Dict[str, Any]:
        """
        Verifica si el alumno ya existe en la pestaña 'Alumnos'.
        Si no existe, lo agrega automáticamente con su Documento, Nombre_Completo y Categoría
        para mantener la base maestra siempre actualizada y precisa.
        """
        existente = self.buscar_alumno(documento) or self.buscar_alumno(nombre_completo)
        if existente:
            return existente

        try:
            ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            filas = ws.get_all_records()
            nuevo_id = f"MAT-{len(filas) + 1:03d}"
            cat_norm = normalizar_categoria(categoria)
            nueva_fila = [nuevo_id, documento, nombre_completo.title(), cat_norm, "En curso"]
            ws.append_row(nueva_fila, value_input_option="USER_ENTERED")
            return {
                "ID_Matricula": nuevo_id,
                "Documento": documento,
                "Nombre_Completo": nombre_completo.title(),
                "Categoria": cat_norm,
                "Estado": "En curso",
            }
        except Exception as exc:
            logger.warning("No se pudo agregar automáticamente en hoja Alumnos: %s", exc)
            return {
                "ID_Matricula": documento,
                "Documento": documento,
                "Nombre_Completo": nombre_completo.title(),
                "Categoria": normalizar_categoria(categoria),
                "Estado": "En curso",
            }

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
                if criterio_limpio == nombre_norm or (
                    len(criterio_limpio.split()) >= 2 and criterio_limpio in nombre_norm
                ):
                    return fila
                mapa_nombres[nombre_norm] = fila

        if mapa_nombres and len(criterio_limpio.split()) >= 2:
            coincidencias = difflib.get_close_matches(
                criterio_limpio, list(mapa_nombres.keys()), n=1, cutoff=0.72
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
    imagen_bytes: Optional[bytes],
    documento_alumno: str,
    nombre_alumno: str,
    clases_texto: str,
    fecha_reporte: str,
    fallback_file_id: str = "",
    fallback_url: str = "",
) -> str:
    if not GOOGLE_DRIVE_ROOT_FOLDER_ID or not imagen_bytes:
        return fallback_url or (
            f"telegram_file_id:{fallback_file_id}" if fallback_file_id else "Sin foto adjunta"
        )
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
# 6. CONSTRUCCIÓN DE MENÚS INTERACTIVOS (CATEGORÍAS -> MULTI-CLASE -> ALUMNO)
# =============================================================================
def teclado_categorias() -> InlineKeyboardMarkup:
    botones = [
        [
            InlineKeyboardButton("🏍️ A1 (F-59)", callback_data="cat|A1"),
            InlineKeyboardButton("🏍️ A2 Moto (F-60)", callback_data="cat|A2"),
        ],
        [
            InlineKeyboardButton("🚗 B1 Particular (F-61)", callback_data="cat|B1"),
            InlineKeyboardButton("🚕 C1 Público (F-62)", callback_data="cat|C1"),
        ],
        [
            InlineKeyboardButton("🔄 Recategorización C1 (21-30)", callback_data="cat|RECAT_C1"),
        ],
        [
            InlineKeyboardButton("❌ Cancelar", callback_data="cancelar"),
        ],
    ]
    return InlineKeyboardMarkup(botones)


def obtener_clases_y_temas_de_indices(categoria: str, indices_seleccionados: List[int]) -> Tuple[str, str]:
    """Calcula el texto de clases (ej. '1-2' o '15') y el tema combinado de las filas marcadas."""
    cat_key = normalizar_categoria(categoria)
    cfg = FORMATOS_OFICIALES.get(cat_key, FORMATOS_OFICIALES["B1"])
    clases_nums: List[int] = []
    temas: List[str] = []

    for idx in sorted(indices_seleccionados):
        if 0 <= idx < len(cfg["filas"]):
            fila = cfg["filas"][idx]
            for c in fila["clases"]:
                if c not in clases_nums:
                    clases_nums.append(c)
            tema_limpio = fila["tema"].replace("\n", " ")
            if tema_limpio not in temas:
                temas.append(tema_limpio)

    return formatear_clases(clases_nums), " / ".join(temas)


def teclado_clases_por_categoria(
    categoria: str, indices_seleccionados: Optional[List[int]] = None
) -> InlineKeyboardMarkup:
    """
    Submenú de clases con casillas de verificación (⬜ / ✅):
    Permite al instructor marcar 1 sola clase (ej. Clase 15) o 2 clases al tiempo (ej. Clase 1 y Clase 2)
    y luego tocar el botón 'Confirmar Clases'.
    """
    cat_key = normalizar_categoria(categoria)
    cfg = FORMATOS_OFICIALES.get(cat_key, FORMATOS_OFICIALES["B1"])
    seleccionados = set(indices_seleccionados or [])
    botones = []

    # Si ya marcó al menos 1 clase, mostrar botón superior de Confirmar
    if seleccionados:
        clases_txt, _ = obtener_clases_y_temas_de_indices(cat_key, list(seleccionados))
        botones.append(
            [
                InlineKeyboardButton(
                    f"💾 CONFIRMAR CLASE(S) {clases_txt} ➡️",
                    callback_data="confirmar_clases",
                )
            ]
        )

    for idx, fila in enumerate(cfg["filas"]):
        etiqueta = fila["etiqueta"].replace(".0", "")
        resumen = fila.get("resumen") or fila["tema"][:30]
        icono = "✅" if idx in seleccionados else "⬜"
        texto_boton = f"{icono} Clase {etiqueta}: {resumen}"
        botones.append(
            [InlineKeyboardButton(texto_boton, callback_data=f"tog|{cat_key}|{idx}")]
        )

    # Botón inferior de Confirmar también al final de la lista para mayor comodidad
    if seleccionados:
        clases_txt, _ = obtener_clases_y_temas_de_indices(cat_key, list(seleccionados))
        botones.append(
            [
                InlineKeyboardButton(
                    f"💾 CONFIRMAR CLASE(S) {clases_txt} ➡️",
                    callback_data="confirmar_clases",
                )
            ]
        )

    botones.append(
        [
            InlineKeyboardButton("🔄 Limpiar selección", callback_data=f"limpiar|{cat_key}"),
            InlineKeyboardButton("⬅️ Categorías", callback_data="volver_cat"),
        ]
    )
    botones.append([InlineKeyboardButton("❌ Cancelar", callback_data="cancelar")])
    return InlineKeyboardMarkup(botones)


def teclado_seleccion_alumno(categoria: str) -> InlineKeyboardMarkup:
    """Muestra botones con los alumnos activos en Google Sheets + botón de volver."""
    alumnos = sheets_service.listar_alumnos_activos(categoria=categoria)
    botones = []
    for f in alumnos[:8]:
        nombre = str(f.get("Nombre_Completo", "")).strip()
        doc = str(f.get("Documento", "")).strip()
        if nombre and doc:
            botones.append(
                [
                    InlineKeyboardButton(
                        f"👤 {nombre} — CC: {doc}",
                        callback_data=f"alu|{doc}",
                    )
                ]
            )
    botones.append(
        [
            InlineKeyboardButton("⬅️ Volver a Clases", callback_data=f"cat|{categoria}"),
            InlineKeyboardButton("❌ Cancelar", callback_data="cancelar"),
        ]
    )
    return InlineKeyboardMarkup(botones)


def mensaje_indicaciones_alumno(categoria: str, clases_texto: str, tema_oficial: str) -> str:
    """Mensaje claro pidiendo Nombre Completo y Cédula al instructor."""
    return (
        f"✅ *Categoría:* `{categoria}` | *Clase(s) marcada(s):* `{clases_texto}`\n"
        f"📝 *Tema:* {tema_oficial}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👨‍🏫 *INDICACIÓN PARA EL INSTRUCTOR (Paso Final):*\n"
        f"Para que el registro en Google Sheets sea 100% preciso, escribe ahora en el chat el "
        f"*NOMBRE COMPLETO (Nombre y Apellido)* y el *NÚMERO DE CÉDULA (CC / TI / CE)* del alumno:\n\n"
        f"📌 *Ejemplo de cómo debes escribirlo:*\n"
        f"`Sergio Andrés Pérez 1098765432`\n\n"
        f"*(O si el alumno ya aparece en los botones de abajo, solo tócalo)* 👇"
    )


async def guardar_en_sheets_y_confirmar(
    chat_id: int,
    user_id: int,
    mensaje_respuesta,
) -> None:
    """Guarda el alumno en 'Alumnos' (si es nuevo) y registra la clase en 'Registro_Clases'."""
    sesion = SESIONES_REPORTE.pop((chat_id, user_id), None)
    if not sesion:
        return

    try:
        instructor = sheets_service.buscar_instructor(user_id, sesion["instructor_nombre"])
        nombre_alumno = sesion.get("nombre_alumno_limpio") or "Desconocido"
        documento_alumno = sesion.get("documento_detectado") or "Pendiente"
        categoria = sesion["categoria"]
        clases_texto = sesion["clases_texto"]
        tema_oficial = sesion["tema_oficial"]

        # Asegurar que el alumno quede creado/actualizado en la pestaña 'Alumnos'
        ficha = sheets_service.registrar_o_actualizar_alumno(
            documento=documento_alumno,
            nombre_completo=nombre_alumno,
            categoria=categoria,
        )
        id_alumno_final = str(ficha.get("Documento") or documento_alumno).strip()
        nombre_final = str(ficha.get("Nombre_Completo") or nombre_alumno).strip()

        url_foto_drive = subir_foto_evidencia(
            imagen_bytes=sesion.get("foto_bytes"),
            documento_alumno=id_alumno_final,
            nombre_alumno=nombre_final,
            clases_texto=clases_texto,
            fecha_reporte=sesion["fecha_reporte"],
            fallback_file_id=sesion.get("file_id", ""),
            fallback_url=sesion.get("file_url", ""),
        )

        sheets_service.registrar_clase(
            timestamp=sesion["timestamp"],
            fecha_reporte=sesion["fecha_reporte"],
            hora_reporte=sesion["hora_reporte"],
            instructor=instructor,
            id_alumno=id_alumno_final,
            nombre_alumno=nombre_final,
            categoria=categoria,
            clases_reportadas=clases_texto,
            tema_registrado=tema_oficial,
            url_foto_drive=url_foto_drive,
        )

        texto_exito = (
            f"✅ *CLASE REGISTRADA EXITOSAMENTE*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"👤 *Alumno:* {nombre_final}\n"
            f"🪪 *Documento (CC/TI):* `{id_alumno_final}`\n"
            f"🏷 *Categoría:* `{categoria}`\n"
            f"📚 *Clase(s) registrada(s):* `{clases_texto}`\n"
            f"📝 *Tema Oficial:* {tema_oficial}\n"
            f"👨‍🏫 *Instructor:* {instructor}\n"
            f"🕒 *Fecha y Hora:* {sesion['fecha_reporte']} {sesion['hora_reporte']}"
        )
        await mensaje_respuesta.edit_text(texto_exito, parse_mode="Markdown")
    except Exception as exc:
        logger.exception("Error guardando sesión en Google Sheets: %s", exc)
        await mensaje_respuesta.edit_text(f"⚠️ Error al guardar en Google Sheets: {exc}")


# =============================================================================
# 7. MANEJADORES DE MENSAJES Y CALLBACKS DE BOTONES
# =============================================================================
async def procesar_mensaje_imagen(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Cuando el instructor envía la foto de la clase:
    1) Revisa si en el comentario ya puso Nombre Completo + Cédula.
    2) Despliega el Menú de Categorías para iniciar la selección de clase(s).
    """
    msg = update.effective_message
    user = update.effective_user
    if not msg or not user:
        return

    foto_obj = None
    if msg.photo:
        foto_obj = await msg.photo[-1].get_file()
    elif msg.document and msg.document.mime_type and msg.document.mime_type.startswith("image/"):
        foto_obj = await msg.document.get_file()

    if not foto_obj:
        return

    foto_bytes = bytes(await foto_obj.download_as_bytearray())
    caption = (msg.caption or "").strip()
    dt_col = msg.date.astimezone(TZ_COLOMBIA)

    nombre_cap, doc_cap = extraer_nombre_y_cedula(caption)
    # Solo considerar completo el nombre si tiene al menos 2 palabras (Nombre y Apellido) y Cédula
    tiene_datos_completos = bool(
        nombre_cap and len(nombre_cap.split()) >= 2 and doc_cap
    )

    SESIONES_REPORTE[(msg.chat_id, user.id)] = {
        "foto_bytes": foto_bytes,
        "file_id": foto_obj.file_id,
        "file_url": foto_obj.file_path or "",
        "timestamp": dt_col.strftime("%Y-%m-%d %H:%M:%S"),
        "fecha_reporte": dt_col.strftime("%Y-%m-%d"),
        "hora_reporte": dt_col.strftime("%H:%M"),
        "instructor_nombre": user.full_name or "Instructor",
        "nombre_alumno_limpio": nombre_cap if tiene_datos_completos else None,
        "documento_detectado": doc_cap if tiene_datos_completos else None,
        "texto_parcial_alumno": caption,
        "categoria": None,
        "indices_clases": [],
        "clases_texto": None,
        "tema_oficial": None,
        "esperando_nombre_texto": False,
    }

    if tiene_datos_completos:
        resumen_alu = f"👤 *Alumno detectado:* {nombre_cap} — *CC:* `{doc_cap}`\n\n"
    else:
        resumen_alu = (
            "ℹ️ *Recuerda:* En el último paso el bot te pedirá el **Nombre Completo y Cédula** del alumno.\n\n"
        )

    await msg.reply_text(
        f"📸 *Foto de evidencia recibida*\n"
        f"{resumen_alu}"
        f"👇 *Paso 1:* Selecciona la **Categoría** del curso:",
        reply_markup=teclado_categorias(),
        parse_mode="Markdown",
    )


async def cmd_menu_clase(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    user = update.effective_user
    if not msg or not user:
        return

    ahora = datetime.now(TZ_COLOMBIA)
    SESIONES_REPORTE[(msg.chat_id, user.id)] = {
        "foto_bytes": None,
        "file_id": "",
        "file_url": "Registro por menú sin foto",
        "timestamp": ahora.strftime("%Y-%m-%d %H:%M:%S"),
        "fecha_reporte": ahora.strftime("%Y-%m-%d"),
        "hora_reporte": ahora.strftime("%H:%M"),
        "instructor_nombre": user.full_name or "Instructor",
        "nombre_alumno_limpio": None,
        "documento_detectado": None,
        "texto_parcial_alumno": "",
        "categoria": None,
        "indices_clases": [],
        "clases_texto": None,
        "tema_oficial": None,
        "esperando_nombre_texto": False,
    }

    await msg.reply_text(
        "📋 *Menú de Registro de Clases — CEA Auto Oriente*\n\n"
        "👇 *Paso 1:* Selecciona la **Categoría** del curso:",
        reply_markup=teclado_categorias(),
        parse_mode="Markdown",
    )


def _instrucciones_submenu_clases(cat_key: str, indices_seleccionados: List[int]) -> str:
    cfg = FORMATOS_OFICIALES.get(cat_key, FORMATOS_OFICIALES["B1"])
    if indices_seleccionados:
        clases_txt, tema_txt = obtener_clases_y_temas_de_indices(cat_key, indices_seleccionados)
        estado_sel = (
            f"✅ *Clase(s) marcada(s):* `{clases_txt}`\n"
            f"📝 *Tema:* {tema_txt}\n\n"
            f"👉 *Puedes marcar una segunda clase si fue bloque doble, o tocar el botón* "
            f"**`💾 CONFIRMAR CLASE(S) {clases_txt} ➡️`** *para continuar.*"
        )
    else:
        estado_sel = (
            "👇 *Paso 2:* Toca **1 clase** (si fue clase individual) o toca **2 clases** "
            "(si fue bloque doble, ej. Clase 1 y Clase 2) y luego pulsa **`💾 CONFIRMAR CLASE(S)`**:"
        )

    return (
        f"🏷 *Categoría:* {cfg['nombre_corto']} (`{cfg['codigo_formato']}`)\n\n"
        f"{estado_sel}"
    )


async def manejar_botones_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Maneja los toques del instructor en Categorías -> Checkboxes de 1 o 2 Clases -> Confirmar -> Alumno."""
    query = update.callback_query
    if not query or not query.message or not query.from_user:
        return

    await query.answer()
    chat_id = query.message.chat_id
    user_id = query.from_user.id
    data = query.data or ""

    sesion = SESIONES_REPORTE.get((chat_id, user_id))
    if not sesion:
        ahora = datetime.now(TZ_COLOMBIA)
        sesion = {
            "foto_bytes": None,
            "file_id": "",
            "file_url": "Registro por botón",
            "timestamp": ahora.strftime("%Y-%m-%d %H:%M:%S"),
            "fecha_reporte": ahora.strftime("%Y-%m-%d"),
            "hora_reporte": ahora.strftime("%H:%M"),
            "instructor_nombre": query.from_user.full_name or "Instructor",
            "nombre_alumno_limpio": None,
            "documento_detectado": None,
            "texto_parcial_alumno": "",
            "categoria": None,
            "indices_clases": [],
            "clases_texto": None,
            "tema_oficial": None,
            "esperando_nombre_texto": False,
        }
        SESIONES_REPORTE[(chat_id, user_id)] = sesion

    if data == "cancelar":
        SESIONES_REPORTE.pop((chat_id, user_id), None)
        await query.message.edit_text("❌ Registro cancelado.")
        return

    if data == "volver_cat":
        sesion["indices_clases"] = []
        sesion["esperando_nombre_texto"] = False
        await query.message.edit_text(
            "👇 *Paso 1:* Selecciona la **Categoría** del curso:",
            reply_markup=teclado_categorias(),
            parse_mode="Markdown",
        )
        return

    # 1. Seleccionar CATEGORÍA -> Mostrar Submenú de Clases con casillas ⬜
    if data.startswith("cat|"):
        cat_key = data.split("|", 1)[1]
        sesion["categoria"] = cat_key
        sesion["indices_clases"] = []
        sesion["esperando_nombre_texto"] = False
        await query.message.edit_text(
            _instrucciones_submenu_clases(cat_key, []),
            reply_markup=teclado_clases_por_categoria(cat_key, []),
            parse_mode="Markdown",
        )
        return

    # 2. Limpiar selección de clases
    if data.startswith("limpiar|"):
        cat_key = data.split("|", 1)[1]
        sesion["indices_clases"] = []
        await query.message.edit_text(
            _instrucciones_submenu_clases(cat_key, []),
            reply_markup=teclado_clases_por_categoria(cat_key, []),
            parse_mode="Markdown",
        )
        return

    # 3. Marcar o desmarcar 1 o 2 clases en el submenú (Checkboxes ✅ / ⬜)
    if data.startswith("tog|"):
        _, cat_key, idx_str = data.split("|")
        idx = int(idx_str)
        sesion["categoria"] = cat_key
        actuales: List[int] = sesion.get("indices_clases", [])

        if idx in actuales:
            actuales.remove(idx)
        else:
            # Permitir marcar hasta 2 bloques/clases a la vez; si marca una 3ra, reemplaza la más antigua
            if len(actuales) >= 2:
                actuales.pop(0)
            actuales.append(idx)

        sesion["indices_clases"] = sorted(actuales)
        await query.message.edit_text(
            _instrucciones_submenu_clases(cat_key, sesion["indices_clases"]),
            reply_markup=teclado_clases_por_categoria(cat_key, sesion["indices_clases"]),
            parse_mode="Markdown",
        )
        return

    # 4. El instructor pulsó "💾 CONFIRMAR CLASE(S)"
    if data == "confirmar_clases":
        cat_key = sesion.get("categoria") or "B1"
        indices = sesion.get("indices_clases", [])
        if not indices:
            return

        clases_txt, tema_txt = obtener_clases_y_temas_de_indices(cat_key, indices)
        sesion["clases_texto"] = clases_txt
        sesion["tema_oficial"] = tema_txt

        # Si ya teníamos Nombre Completo + Cédula desde el comentario de la foto, guardar directo
        if sesion.get("nombre_alumno_limpio") and sesion.get("documento_detectado"):
            await guardar_en_sheets_y_confirmar(chat_id, user_id, query.message)
        else:
            sesion["esperando_nombre_texto"] = True
            await query.message.edit_text(
                mensaje_indicaciones_alumno(cat_key, clases_txt, tema_txt),
                reply_markup=teclado_seleccion_alumno(cat_key),
                parse_mode="Markdown",
            )
        return

    # 5. El instructor seleccionó un alumno existente de la lista de botones
    if data.startswith("alu|"):
        doc_alu = data.split("|", 1)[1]
        ficha = sheets_service.buscar_alumno(doc_alu)
        if ficha:
            sesion["nombre_alumno_limpio"] = str(ficha.get("Nombre_Completo", "")).strip()
            sesion["documento_detectado"] = str(ficha.get("Documento", doc_alu)).strip()
        else:
            sesion["nombre_alumno_limpio"] = doc_alu
            sesion["documento_detectado"] = doc_alu

        sesion["esperando_nombre_texto"] = False
        await guardar_en_sheets_y_confirmar(chat_id, user_id, query.message)
        return


async def procesar_texto_chat(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Valida que el instructor envíe el Nombre Completo (Nombre y Apellido) y la Cédula
    cuando está en el Paso Final del registro.
    """
    msg = update.effective_message
    user = update.effective_user
    if not msg or not user or not msg.text:
        return

    texto = msg.text.strip()
    if texto.startswith("/"):
        return

    sesion = SESIONES_REPORTE.get((msg.chat_id, user.id))
    if sesion and sesion.get("esperando_nombre_texto"):
        nombre, documento = extraer_nombre_y_cedula(texto)

        # Verificar si escribió la cédula o nombre de un alumno que ya existe en la hoja 'Alumnos'
        ficha_existente = sheets_service.buscar_alumno(documento or nombre or texto)
        if ficha_existente:
            nombre_oficial = str(ficha_existente.get("Nombre_Completo", "")).strip()
            doc_oficial = str(ficha_existente.get("Documento", "")).strip()
            if nombre_oficial and doc_oficial and doc_oficial != "Pendiente":
                sesion["nombre_alumno_limpio"] = nombre_oficial
                sesion["documento_detectado"] = doc_oficial
                sesion["esperando_nombre_texto"] = False
                aviso = await msg.reply_text("⏳ Guardando registro en Google Sheets...")
                await guardar_en_sheets_y_confirmar(msg.chat_id, user.id, aviso)
                return

        # Validar que haya escrito Nombre + Apellido (mínimo 2 palabras) Y Número de Documento
        if not nombre or len(nombre.split()) < 2 or not documento:
            faltantes = []
            if not nombre or len(nombre.split()) < 2:
                faltantes.append("• **Nombre y Apellido completos** (ej. *Sergio Andrés Pérez*)")
            if not documento:
                faltantes.append("• **Número de Cédula / Documento** (sin puntos, ej. *1098765432*)")

            lista_faltantes = "\n".join(faltantes)
            await msg.reply_text(
                f"⚠️ *Faltan datos para que la hoja de cálculo quede exacta:*\n"
                f"{lista_faltantes}\n\n"
                f"✍️ *Por favor envíalo completo en un solo mensaje así:*\n"
                f"`Sergio Andrés Pérez 1098765432`",
                parse_mode="Markdown",
            )
            return

        sesion["nombre_alumno_limpio"] = nombre
        sesion["documento_detectado"] = documento
        sesion["esperando_nombre_texto"] = False
        aviso = await msg.reply_text("⏳ Guardando alumno y clase en Google Sheets...")
        await guardar_en_sheets_y_confirmar(msg.chat_id, user.id, aviso)
        return


async def cmd_ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    texto = (
        "🚗 *Bot de Seguimiento de Clases — CEA Auto Oriente*\n\n"
        "📸 *Pasos para el Instructor:*\n"
        "1️⃣ Envía la **foto de la clase** al chat.\n"
        "2️⃣ Toca la **Categoría** (`A1`, `A2`, `B1`, `C1`, `Recat C1`).\n"
        "3️⃣ En el listado de clases puedes marcar **1 sola clase** (ej. `Clase 15`) o **2 clases al tiempo** (ej. `Clase 1` y `Clase 2`) y pulsar **`💾 CONFIRMAR CLASE(S)`**.\n"
        "4️⃣ Escribe el **Nombre Completo y Cédula** del alumno (ej. `Sergio Andrés Pérez 1098765432`).\n\n"
        "📄 *Generar PDF Oficial para Carpeta/Matrícula:*\n"
        "`/reporte [Cédula o Nombre del Alumno]`\n"
        "Ejemplo: `/reporte 1098765432` o `/reporte Camilo Torres`\n\n"
        "⚙️ *Inicializar Hojas en Google Sheets:*\n"
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
            "⚠️ Indica el nombre o cédula del alumno.\nEjemplo: `/reporte Camilo Torres`",
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
# 8. SERVIDOR HEALTH CHECK + AUTO-PING (RENDER KEEP-ALIVE) Y ARRANQUE DEL BOT
# =============================================================================
class HealthCheckHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        self.send_response(200)
        self.send_header("Content-type", "text/plain; charset=utf-8")
        self.end_headers()
        self.wfile.write("Bot CEA Auto Oriente activo (200 OK)".encode("utf-8"))

    def log_message(self, format, *args):
        return


def _loop_keep_alive() -> None:
    while True:
        time.sleep(600)
        if RENDER_EXTERNAL_URL:
            try:
                urllib.request.urlopen(RENDER_EXTERNAL_URL, timeout=15).read()
            except Exception:
                pass


def iniciar_servidor_healthcheck() -> None:
    try:
        server = HTTPServer(("0.0.0.0", PORT), HealthCheckHandler)
        hilo = threading.Thread(target=server.serve_forever, daemon=True)
        hilo.start()
        hilo_ping = threading.Thread(target=_loop_keep_alive, daemon=True)
        hilo_ping.start()
        logger.info("🌐 Servidor Health Check + Keep-Alive activo en puerto %s", PORT)
    except Exception as exc:
        logger.warning("No se pudo iniciar servidor Health Check: %s", exc)


def main() -> None:
    if not TELEGRAM_BOT_TOKEN or TELEGRAM_BOT_TOKEN == "TU_TELEGRAM_BOT_TOKEN":
        logger.error("⚠️ Configura TELEGRAM_BOT_TOKEN en .env o en Render.")
        sys.exit(1)

    iniciar_servidor_healthcheck()

    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler(["start", "ayuda", "help"], cmd_ayuda))
    app.add_handler(CommandHandler(["clase", "menu", "registrar"], cmd_menu_clase))
    app.add_handler(CommandHandler("inicializar", cmd_inicializar))
    app.add_handler(CommandHandler("reporte", cmd_reporte))

    app.add_handler(CallbackQueryHandler(manejar_botones_callback))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.IMAGE, procesar_mensaje_imagen))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), procesar_texto_chat))

    logger.info("🚗 Bot CEA Auto Oriente iniciado...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
