"""
Bot de Seguimiento de Clases — CEA Auto Oriente (Versión Completa en Línea).
Incluye:
1. Menú de Categorías limpio (🏍️ A1 Moto, 🏍️ A2 Moto, 🚗 B1 Carro Particular, 🚕 C1 Servicio Público, 🔄 Recategorización C1).
2. Selección de 1 clase o 2 clases al tiempo con casillas (⬜ / ✅).
3. Validación de Horario de Envío:
   - Para 2 clases: entre el minuto 40 y el minuto 50.
   - Para 1 clase: entre el minuto 18 y el minuto 28.
   - Si se envía antes o después, advierte al instructor y notifica al administrador / registra la novedad en Google Sheets.
4. Auto-llenado de las 3 hojas de Google Sheets: 'Alumnos', 'Instructores' y 'Registro_Clases'.
5. Generación del PDF Oficial (F-59, F-60, F-61 y F-62) con el comando /reporte.
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
ADMIN_CHAT_ID_ENV = os.getenv("ADMIN_CHAT_ID", "").strip()
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
# Mapa de seguridad: message_id del menú del bot -> user_id propietario del turno
PROPIETARIO_MENSAJE_BOT: Dict[int, int] = {}
# Lista de chats de administradores suscritos a alertas con /alertas
ADMIN_CHATS_ALERTA: set = set()
if ADMIN_CHAT_ID_ENV.lstrip("-").isdigit():
    ADMIN_CHATS_ALERTA.add(int(ADMIN_CHAT_ID_ENV))


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
# 2. CATÁLOGO OFICIAL DE FORMATOS Y TEMARIOS
# =============================================================================
FORMATOS_OFICIALES: Dict[str, Dict[str, Any]] = {
    "A1": {
        "nombre_corto": "🏍️ A1 Moto",
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
        "nombre_corto": "🏍️ A2 Moto",
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
        "nombre_corto": "🚗 B1 Carro Particular",
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
        "nombre_corto": "🚕 C1 Servicio Público",
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
        "nombre_corto": "🔄 Recategorización C1",
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


def validar_minuto_envio(minuto: int, cantidad_clases: int) -> Tuple[bool, str, str]:
    """
    Valida si la hora de envío de la foto cumple los tiempos estipulados:
    - Si son 2 clases (bloque doble): debe enviarse entre el minuto 40 y el minuto 50.
    - Si es 1 sola clase: debe enviarse entre el minuto 18 y el minuto 28.
    Retorna: (es_valido, estado_corto_sheets, mensaje_alerta_instructor)
    """
    if cantidad_clases >= 2:
        min_ini, min_fin = 40, 50
        tipo_txt = "bloque de 2 clases (mitad de clase: minuto 40 al 50)"
    else:
        min_ini, min_fin = 18, 28
        tipo_txt = "1 sola clase (mitad de clase: minuto 18 al 28)"

    if min_ini <= minuto <= min_fin:
        return (
            True,
            f"EN TIEMPO (minuto :{minuto:02d})",
            "✅ *Validación de horario:* Enviada dentro del tiempo estipulado.",
        )

    momento = "ANTES" if minuto < min_ini else "DESPUÉS"
    estado_sheets = f"FUERA DE TIEMPO - {momento} (minuto :{minuto:02d})"
    alerta = (
        f"\n\n⚠️ *NOVEDAD EN HORA DE ENVÍO ({momento} DE TIEMPO):*\n"
        f"La evidencia fue enviada en el **minuto :{minuto:02d}**, y para {tipo_txt} "
        f"debe enviarse entre el **minuto {min_ini} y el minuto {min_fin}**.\n"
        f"🚨 *La hora de envío no está dentro de los tiempos estipulados y se informa al administrador.*"
    )
    return False, estado_sheets, alerta


# =============================================================================
# 3. EXTRACTOR DE REPORTES Y VALIDACIÓN DE NOMBRE + CÉDULA
# =============================================================================
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
    if not texto or not texto.strip():
        return None, None

    limpio = texto.strip()
    m_doc = re.search(r"\b(\d{6,12})\b", limpio)
    documento = m_doc.group(1) if m_doc else None

    sin_doc = re.sub(r"\b\d{5,12}\b", "", limpio)
    sin_prefijos = re.sub(
        r"\b(?:alumn[oa]s?|nombre|cc|ti|ce|c[eé]dula|documento|doc|id)\b\s*[:#.\-]?",
        "",
        sin_doc,
        flags=re.IGNORECASE,
    )
    nombre_limpio = re.sub(r"[^A-Za-zÁÉÍÓÚáéíóúÑñÜü\s]", " ", sin_prefijos)
    palabras = [p for p in nombre_limpio.split() if len(p) >= 2]

    nombre_formateado = " ".join(palabras).title() if palabras else None
    return nombre_formateado, documento


# =============================================================================
# 4. SERVICIO DE GOOGLE SHEETS Y DRIVE
# =============================================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

HEADERS_ALUMNOS = ["ID_Matricula", "Documento", "Nombre_Completo", "Categoria", "Estado"]
HEADERS_INSTRUCTORES = [
    "Telegram_User_ID",
    "Nombre_Instructor",
    "Categoria_Habilitada",
    "Documento_Instructor",
]
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
    "Validacion_Horario",
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
        self._cache_instructores: Dict[int, Dict[str, str]] = {}

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
            primera_fila = ws.row_values(1)
            if not primera_fila:
                ws.append_row(encabezados)
            elif len(primera_fila) < len(encabezados):
                # Si se agregó la columna 'Validacion_Horario', asegurar que el encabezado exista
                for col_idx, titulo in enumerate(encabezados, start=1):
                    if col_idx > len(primera_fila):
                        ws.update_cell(1, col_idx, titulo)
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

    def obtener_instructor_registrado(self, telegram_user_id: int) -> Optional[Dict[str, str]]:
        """
        Verifica si el instructor ya tiene Nombre Oficial (Nombre + Apellido) y Cédula
        registrados en la pestaña 'Instructores'.
        """
        if telegram_user_id in self._cache_instructores:
            cached = self._cache_instructores[telegram_user_id]
            if cached.get("documento") and len(cached.get("nombre", "").split()) >= 2:
                return cached

        try:
            ws = self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
            registros = ws.get_all_records()
            for fila in registros:
                if str(fila.get("Telegram_User_ID", "")).strip() == str(telegram_user_id):
                    nombre = str(fila.get("Nombre_Instructor", "")).strip()
                    doc = str(fila.get("Documento_Instructor", "")).strip()
                    info = {
                        "nombre": nombre,
                        "documento": doc,
                        "categorias": str(fila.get("Categoria_Habilitada", "")).strip(),
                    }
                    if nombre and len(nombre.split()) >= 2 and doc:
                        self._cache_instructores[telegram_user_id] = info
                        return info
                    return None
        except Exception as exc:
            logger.warning("No se pudo consultar hoja Instructores: %s", exc)
        return None

    def guardar_instructor_oficial(
        self,
        telegram_user_id: int,
        nombre_oficial: str,
        documento_oficial: str,
        categoria_inicial: str = "",
    ) -> Dict[str, str]:
        """
        Guarda o actualiza en la pestaña 'Instructores' el Nombre Oficial y la Cédula
        del instructor vinculados a su Telegram_User_ID (o vincula una fila pre-cargada por cédula).
        """
        nombre_limpio = nombre_oficial.strip().title()
        doc_limpio = documento_oficial.strip()
        cat_norm = normalizar_categoria(categoria_inicial) if categoria_inicial else ""

        info = {
            "nombre": nombre_limpio,
            "documento": doc_limpio,
            "categorias": cat_norm,
        }
        self._cache_instructores[telegram_user_id] = info

        try:
            ws = self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
            registros = ws.get_all_records()
            for idx_fila, fila in enumerate(registros, start=2):
                tg_id_fila = str(fila.get("Telegram_User_ID", "")).strip()
                doc_fila = str(fila.get("Documento_Instructor", "")).strip()

                # Coincide por Telegram_User_ID o por Documento pre-cargado por administración
                if tg_id_fila == str(telegram_user_id) or (doc_fila and doc_fila == doc_limpio):
                    cats_actuales = [
                        c.strip()
                        for c in str(fila.get("Categoria_Habilitada", "")).split(",")
                        if c.strip()
                    ]
                    if cat_norm and cat_norm not in cats_actuales:
                        cats_actuales.append(cat_norm)
                    cats_str = ", ".join(cats_actuales)
                    ws.update_cell(idx_fila, 1, str(telegram_user_id))
                    ws.update_cell(idx_fila, 2, nombre_limpio)
                    ws.update_cell(idx_fila, 3, cats_str)
                    ws.update_cell(idx_fila, 4, doc_limpio)
                    info["categorias"] = cats_str
                    return info

            ws.append_row(
                [str(telegram_user_id), nombre_limpio, cat_norm, doc_limpio],
                value_input_option="USER_ENTERED",
            )
        except Exception as exc:
            logger.warning("No se pudo guardar instructor oficial en Sheets: %s", exc)
        return info

    def registrar_o_actualizar_instructor(
        self,
        telegram_user_id: int,
        nombre_telegram: str,
        categoria_dictada: str,
        documento_instructor: str = "",
    ) -> str:
        """
        Busca al instructor en la pestaña 'Instructores'.
        - Si ya existe, devuelve su Nombre_Instructor oficial y actualiza sus categorías habilitadas.
        - Si no existe aún, lo agrega automáticamente en la pestaña 'Instructores' con su nombre y cédula.
        """
        cat_norm = normalizar_categoria(categoria_dictada)
        try:
            ws = self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
            registros = ws.get_all_records()
            for idx_fila, fila in enumerate(registros, start=2):
                if str(fila.get("Telegram_User_ID", "")).strip() == str(telegram_user_id):
                    nombre_oficial = str(fila.get("Nombre_Instructor", "")).strip() or nombre_telegram
                    doc_actual = str(fila.get("Documento_Instructor", "")).strip() or documento_instructor
                    cats_actuales = [
                        c.strip()
                        for c in str(fila.get("Categoria_Habilitada", "")).split(",")
                        if c.strip()
                    ]
                    if cat_norm not in cats_actuales:
                        cats_actuales.append(cat_norm)
                        ws.update_cell(idx_fila, 3, ", ".join(cats_actuales))
                    if documento_instructor and not str(fila.get("Documento_Instructor", "")).strip():
                        ws.update_cell(idx_fila, 4, documento_instructor)
                    self._cache_instructores[telegram_user_id] = {
                        "nombre": nombre_oficial,
                        "documento": doc_actual,
                        "categorias": ", ".join(cats_actuales),
                    }
                    return nombre_oficial

            # Si no estaba registrado en la hoja 'Instructores', agregarlo automáticamente
            ws.append_row(
                [str(telegram_user_id), nombre_telegram, cat_norm, documento_instructor],
                value_input_option="USER_ENTERED",
            )
            self._cache_instructores[telegram_user_id] = {
                "nombre": nombre_telegram,
                "documento": documento_instructor,
                "categorias": cat_norm,
            }
            return nombre_telegram
        except Exception as exc:
            logger.warning("No se pudo actualizar hoja Instructores: %s", exc)
            return nombre_telegram

    def registrar_o_actualizar_alumno(
        self, documento: str, nombre_completo: str, categoria: str
    ) -> Dict[str, Any]:
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
        validacion_horario: str = "EN TIEMPO",
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
                validacion_horario,
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
# 6. CONSTRUCCIÓN DE MENÚS INTERACTIVOS CON BLOQUEO POR INSTRUCTOR (OWNER_ID)
# =============================================================================
def teclado_categorias(owner_id: int) -> InlineKeyboardMarkup:
    """Menú limpio sin códigos F-59/F-60/F-61/F-62 y protegido por owner_id."""
    botones = [
        [
            InlineKeyboardButton("🏍️ A1 Moto", callback_data=f"cat|A1|{owner_id}"),
            InlineKeyboardButton("🏍️ A2 Moto", callback_data=f"cat|A2|{owner_id}"),
        ],
        [
            InlineKeyboardButton("🚗 B1 Carro Particular", callback_data=f"cat|B1|{owner_id}"),
            InlineKeyboardButton("🚕 C1 Servicio Público", callback_data=f"cat|C1|{owner_id}"),
        ],
        [
            InlineKeyboardButton("🔄 Recategorización C1", callback_data=f"cat|RECAT_C1|{owner_id}"),
        ],
        [
            InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar|{owner_id}"),
        ],
    ]
    return InlineKeyboardMarkup(botones)


def obtener_clases_y_temas_de_indices(
    categoria: str, indices_seleccionados: List[int]
) -> Tuple[str, str, int]:
    """Devuelve (clases_texto, tema_combinado, cantidad_total_clases)."""
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

    return formatear_clases(clases_nums), " / ".join(temas), len(clases_nums)


def teclado_clases_por_categoria(
    categoria: str,
    indices_seleccionados: Optional[List[int]],
    owner_id: int,
) -> InlineKeyboardMarkup:
    cat_key = normalizar_categoria(categoria)
    cfg = FORMATOS_OFICIALES.get(cat_key, FORMATOS_OFICIALES["B1"])
    seleccionados = set(indices_seleccionados or [])
    botones = []

    if seleccionados:
        clases_txt, _, _ = obtener_clases_y_temas_de_indices(cat_key, list(seleccionados))
        botones.append(
            [
                InlineKeyboardButton(
                    f"💾 CONFIRMAR CLASE(S) {clases_txt} ➡️",
                    callback_data=f"confirmar_clases|{owner_id}",
                )
            ]
        )

    for idx, fila in enumerate(cfg["filas"]):
        etiqueta = fila["etiqueta"].replace(".0", "")
        resumen = fila.get("resumen") or fila["tema"][:30]
        icono = "✅" if idx in seleccionados else "⬜"
        texto_boton = f"{icono} Clase {etiqueta}: {resumen}"
        botones.append(
            [InlineKeyboardButton(texto_boton, callback_data=f"tog|{cat_key}|{idx}|{owner_id}")]
        )

    if seleccionados:
        clases_txt, _, _ = obtener_clases_y_temas_de_indices(cat_key, list(seleccionados))
        botones.append(
            [
                InlineKeyboardButton(
                    f"💾 CONFIRMAR CLASE(S) {clases_txt} ➡️",
                    callback_data=f"confirmar_clases|{owner_id}",
                )
            ]
        )

    botones.append(
        [
            InlineKeyboardButton(
                "🔄 Limpiar selección", callback_data=f"limpiar|{cat_key}|{owner_id}"
            ),
            InlineKeyboardButton("⬅️ Categorías", callback_data=f"volver_cat|{owner_id}"),
        ]
    )
    botones.append([InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar|{owner_id}")])
    return InlineKeyboardMarkup(botones)


def teclado_seleccion_alumno(categoria: str, owner_id: int) -> InlineKeyboardMarkup:
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
                        callback_data=f"alu|{doc}|{owner_id}",
                    )
                ]
            )
    botones.append(
        [
            InlineKeyboardButton("⬅️ Volver a Clases", callback_data=f"cat|{categoria}|{owner_id}"),
            InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar|{owner_id}"),
        ]
    )
    return InlineKeyboardMarkup(botones)


def _encabezado_instructor(sesion: Dict[str, Any]) -> str:
    nom = sesion.get("instructor_nombre") or "Instructor"
    doc = sesion.get("instructor_documento") or ""
    if doc:
        return f"👨‍🏫 *Instructor:* {nom} *(CC: {doc})*\n"
    return f"👨‍🏫 *Instructor:* {nom}\n"


def mensaje_indicaciones_alumno(
    sesion: Dict[str, Any], categoria: str, clases_texto: str, tema_oficial: str
) -> str:
    cfg = FORMATOS_OFICIALES.get(categoria, FORMATOS_OFICIALES["B1"])
    return (
        f"{_encabezado_instructor(sesion)}"
        f"✅ *Categoría:* {cfg['nombre_corto']} | *Clase(s) marcada(s):* `{clases_texto}`\n"
        f"📝 *Tema:* {tema_oficial}\n\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"✍️ *PASO FINAL PARA {sesion.get('instructor_nombre', 'EL INSTRUCTOR').upper()}:*\n"
        f"Para que el registro en Google Sheets sea 100% preciso, escribe ahora en el chat el "
        f"*NOMBRE COMPLETO (Nombre y Apellido)* y el *NÚMERO DE CÉDULA (CC / TI / CE)* del alumno:\n\n"
        f"📌 *Ejemplo de cómo debes escribirlo:*\n"
        f"`Sergio Andrés Pérez 1098765432`\n\n"
        f"*(O si el alumno ya aparece en los botones de abajo, solo tócalo)* 👇"
    )


async def guardar_en_sheets_y_confirmar(
    chat_id: int,
    user_id: int,
    context: ContextTypes.DEFAULT_TYPE,
    mensaje_respuesta,
) -> None:
    """
    1. Registra/actualiza al Instructor en la pestaña 'Instructores' con su nombre y cédula oficial.
    2. Registra/actualiza al Alumno en la pestaña 'Alumnos'.
    3. Valida el minuto de envío (40-50 para 2 clases / 18-28 para 1 clase).
    4. Registra la fila en 'Registro_Clases' y notifica al administrador si está fuera de tiempo.
    """
    sesion = SESIONES_REPORTE.pop((chat_id, user_id), None)
    if not sesion:
        return

    menu_msg_id = sesion.get("menu_msg_id")
    if menu_msg_id:
        PROPIETARIO_MENSAJE_BOT.pop(menu_msg_id, None)

    try:
        categoria = sesion["categoria"]
        nombre_alumno = sesion.get("nombre_alumno_limpio") or "Desconocido"
        documento_alumno = sesion.get("documento_detectado") or "Pendiente"
        clases_texto = sesion["clases_texto"]
        cantidad_clases = sesion.get("cantidad_clases", 2)
        tema_oficial = sesion["tema_oficial"]
        minuto_envio = sesion.get("minuto_envio", 45)
        doc_instructor = sesion.get("instructor_documento", "")

        # 1. Auto-registrar o actualizar al Instructor en la hoja 'Instructores'
        instructor = sheets_service.registrar_o_actualizar_instructor(
            telegram_user_id=user_id,
            nombre_telegram=sesion["instructor_nombre"],
            categoria_dictada=categoria,
            documento_instructor=doc_instructor,
        )

        # 2. Auto-registrar o actualizar al Alumno en la hoja 'Alumnos'
        ficha = sheets_service.registrar_o_actualizar_alumno(
            documento=documento_alumno,
            nombre_completo=nombre_alumno,
            categoria=categoria,
        )
        id_alumno_final = str(ficha.get("Documento") or documento_alumno).strip()
        nombre_final = str(ficha.get("Nombre_Completo") or nombre_alumno).strip()

        # 3. Validar el minuto de envío según cantidad de clases (2 clases: 40-50 | 1 clase: 18-28)
        en_tiempo, estado_horario_sheets, alerta_horario_txt = validar_minuto_envio(
            minuto=minuto_envio, cantidad_clases=cantidad_clases
        )

        # 4. Subir evidencia o enlace directo
        url_foto_drive = subir_foto_evidencia(
            imagen_bytes=sesion.get("foto_bytes"),
            documento_alumno=id_alumno_final,
            nombre_alumno=nombre_final,
            clases_texto=clases_texto,
            fecha_reporte=sesion["fecha_reporte"],
            fallback_file_id=sesion.get("file_id", ""),
            fallback_url=sesion.get("file_url", ""),
        )

        # 5. Guardar en 'Registro_Clases'
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
            validacion_horario=estado_horario_sheets,
        )

        cfg = FORMATOS_OFICIALES.get(categoria, FORMATOS_OFICIALES["B1"])
        doc_inst_txt = f" (CC: `{doc_instructor}`)" if doc_instructor else ""
        texto_exito = (
            f"✅ *CLASE REGISTRADA EXITOSAMENTE*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"👤 *Alumno:* {nombre_final}\n"
            f"🪪 *Documento (CC/TI):* `{id_alumno_final}`\n"
            f"🏷 *Categoría:* {cfg['nombre_corto']}\n"
            f"📚 *Clase(s) registrada(s):* `{clases_texto}`\n"
            f"📝 *Tema Oficial:* {tema_oficial}\n"
            f"👨‍🏫 *Instructor:* {instructor}{doc_inst_txt}\n"
            f"🕒 *Fecha y Hora de envío:* {sesion['fecha_reporte']} {sesion['hora_reporte']}\n"
            f"{alerta_horario_txt}"
        )
        await mensaje_respuesta.edit_text(texto_exito, parse_mode="Markdown")

        # Si el menú original citando la foto del instructor era otro mensaje distinto, actualizarlo también
        if menu_msg_id and getattr(mensaje_respuesta, "message_id", None) != menu_msg_id:
            try:
                await context.bot.edit_message_text(
                    chat_id=chat_id,
                    message_id=menu_msg_id,
                    text=texto_exito,
                    parse_mode="Markdown",
                )
                await mensaje_respuesta.delete()
            except Exception:
                pass

        # 6. Si el envío fue fuera de los minutos estipulados, notificar a los administradores suscritos
        if not en_tiempo and ADMIN_CHATS_ALERTA:
            msg_admin = (
                f"🚨 *REPORTE ENVIADO FUERA DEL TIEMPO ESTIPULADO*\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"👨‍🏫 *Instructor:* {instructor}{doc_inst_txt}\n"
                f"👤 *Alumno:* {nombre_final} (CC: `{id_alumno_final}`)\n"
                f"🏷 *Categoría:* {cfg['nombre_corto']} | *Clase(s):* `{clases_texto}`\n"
                f"🕒 *Hora de envío:* `{sesion['hora_reporte']}` ({estado_horario_sheets})\n"
                f"📌 *Regla:* {'Minuto 40 al 50 (2 clases)' if cantidad_clases >= 2 else 'Minuto 18 al 28 (1 clase)'}"
            )
            for admin_chat in list(ADMIN_CHATS_ALERTA):
                if admin_chat != chat_id:
                    try:
                        await context.bot.send_message(
                            chat_id=admin_chat, text=msg_admin, parse_mode="Markdown"
                        )
                    except Exception:
                        pass

    except Exception as exc:
        logger.exception("Error guardando sesión en Google Sheets: %s", exc)
        await mensaje_respuesta.edit_text(f"⚠️ Error al guardar en Google Sheets: {exc}")


# =============================================================================
# 7. MANEJADORES DE MENSAJES Y CALLBACKS DE BOTONES (MULTI-INSTRUCTOR GRUPAL)
# =============================================================================
def _texto_menu_paso_1(sesion: Dict[str, Any]) -> str:
    tiene_alu = bool(
        sesion.get("nombre_alumno_limpio") and sesion.get("documento_detectado")
    )
    if tiene_alu:
        resumen_alu = (
            f"👤 *Alumno detectado:* {sesion['nombre_alumno_limpio']} — "
            f"*CC:* `{sesion['documento_detectado']}`\n\n"
        )
    else:
        resumen_alu = (
            "ℹ️ *Recuerda:* En el último paso el bot te pedirá el **Nombre Completo y Cédula** del alumno.\n\n"
        )

    return (
        f"📸 *Foto de evidencia recibida ({sesion['hora_reporte']})*\n"
        f"{_encabezado_instructor(sesion)}"
        f"{resumen_alu}"
        f"👇 *Paso 1:* Selecciona la **Categoría** del curso:"
    )


async def procesar_mensaje_imagen(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
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
    tiene_datos_completos = bool(
        nombre_cap and len(nombre_cap.split()) >= 2 and doc_cap
    )

    # Verificar si este instructor ya está registrado con su Nombre Oficial y Cédula en 'Instructores'
    perfil_inst = sheets_service.obtener_instructor_registrado(user.id)
    nombre_inst_oficial = (
        perfil_inst["nombre"] if perfil_inst else (user.full_name or "Instructor")
    )
    doc_inst_oficial = perfil_inst["documento"] if perfil_inst else ""

    sesion = {
        "owner_id": user.id,
        "foto_msg_id": msg.message_id,
        "menu_msg_id": None,
        "foto_bytes": foto_bytes,
        "file_id": foto_obj.file_id,
        "file_url": foto_obj.file_path or "",
        "timestamp": dt_col.strftime("%Y-%m-%d %H:%M:%S"),
        "fecha_reporte": dt_col.strftime("%Y-%m-%d"),
        "hora_reporte": dt_col.strftime("%H:%M"),
        "minuto_envio": dt_col.minute,
        "instructor_nombre": nombre_inst_oficial,
        "instructor_documento": doc_inst_oficial,
        "nombre_alumno_limpio": nombre_cap if tiene_datos_completos else None,
        "documento_detectado": doc_cap if tiene_datos_completos else None,
        "categoria": None,
        "indices_clases": [],
        "clases_texto": None,
        "cantidad_clases": 2,
        "tema_oficial": None,
        "esperando_registro_instructor": not bool(perfil_inst),
        "esperando_nombre_texto": False,
    }
    SESIONES_REPORTE[(msg.chat_id, user.id)] = sesion

    # Si es la primera vez del instructor (no tiene Nombre Oficial + Cédula en 'Instructores'), pedírselo 1 sola vez
    if not perfil_inst:
        msg_bot = await msg.reply_text(
            f"👋 *¡Hola! Identificación Oficial de Instructor (Única vez)*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"📸 Tu foto de las `{dt_col.strftime('%H:%M')}` ya quedó guardada.\n\n"
            f"Para que en **Google Sheets** y en los **PDF oficiales** aparezca tu nombre real y cédula "
            f"(en lugar del alias de Telegram), escribe ahora en el chat **TU NOMBRE COMPLETO y TU CÉDULA de Instructor**:\n\n"
            f"📌 *Ejemplo:*\n"
            f"`Segundo Murillo 13959739`\n\n"
            f"*(Solo tendrás que hacerlo esta primera vez; en tus próximas fotos el bot te reconocerá automáticamente).* 👇",
            reply_to_message_id=msg.message_id,
            parse_mode="Markdown",
        )
        sesion["menu_msg_id"] = msg_bot.message_id
        PROPIETARIO_MENSAJE_BOT[msg_bot.message_id] = user.id
        return

    # Si ya está registrado como instructor oficial, responder citando su foto con su menú bloqueado para él
    msg_bot = await msg.reply_text(
        _texto_menu_paso_1(sesion),
        reply_markup=teclado_categorias(user.id),
        reply_to_message_id=msg.message_id,
        parse_mode="Markdown",
    )
    sesion["menu_msg_id"] = msg_bot.message_id
    PROPIETARIO_MENSAJE_BOT[msg_bot.message_id] = user.id


async def cmd_registrar_instructor(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Comando /soy o /instructor para que un instructor registre o actualice
    su Nombre Completo Oficial y Cédula en la hoja 'Instructores'.
    """
    msg = update.effective_message
    user = update.effective_user
    if not msg or not user:
        return

    texto_args = " ".join(context.args).strip() if context.args else ""
    nombre_inst, doc_inst = extraer_nombre_y_cedula(texto_args)

    if not nombre_inst or len(nombre_inst.split()) < 2 or not doc_inst:
        perfil_actual = sheets_service.obtener_instructor_registrado(user.id)
        estado_actual = (
            f"✅ *Actualmente estás registrado como:* {perfil_actual['nombre']} (CC: `{perfil_actual['documento']}`)\n\n"
            if perfil_actual
            else "ℹ️ *Aún no has vinculado tu Nombre Completo y Cédula de Instructor.*\n\n"
        )
        await msg.reply_text(
            f"👨‍🏫 *Registro Oficial de Instructor — CEA Auto Oriente*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"{estado_actual}"
            f"Para registrarte o actualizar tus datos oficiales, escribe:\n"
            f"`/soy Nombre Apellido Cédula`\n\n"
            f"📌 *Ejemplo:*\n"
            f"`/soy Segundo Murillo 13959739`",
            reply_to_message_id=msg.message_id,
            parse_mode="Markdown",
        )
        return

    info = sheets_service.guardar_instructor_oficial(
        telegram_user_id=user.id,
        nombre_oficial=nombre_inst,
        documento_oficial=doc_inst,
    )
    await msg.reply_text(
        f"✅ *INSTRUCTOR VINCULADO EXITOSAMENTE*\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👨‍🏫 *Nombre Oficial:* {info['nombre']}\n"
        f"🪪 *Cédula:* `{info['documento']}`\n"
        f"🆔 *Telegram ID:* `{user.id}`\n\n"
        f"A partir de ahora todas tus fotos en el grupo quedarán registradas con tu nombre oficial.",
        reply_to_message_id=msg.message_id,
        parse_mode="Markdown",
    )


async def cmd_menu_clase(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    user = update.effective_user
    if not msg or not user:
        return

    ahora = datetime.now(TZ_COLOMBIA)
    perfil_inst = sheets_service.obtener_instructor_registrado(user.id)
    nombre_inst_oficial = (
        perfil_inst["nombre"] if perfil_inst else (user.full_name or "Instructor")
    )
    doc_inst_oficial = perfil_inst["documento"] if perfil_inst else ""

    sesion = {
        "owner_id": user.id,
        "foto_msg_id": msg.message_id,
        "menu_msg_id": None,
        "foto_bytes": None,
        "file_id": "",
        "file_url": "Registro por menú sin foto",
        "timestamp": ahora.strftime("%Y-%m-%d %H:%M:%S"),
        "fecha_reporte": ahora.strftime("%Y-%m-%d"),
        "hora_reporte": ahora.strftime("%H:%M"),
        "minuto_envio": ahora.minute,
        "instructor_nombre": nombre_inst_oficial,
        "instructor_documento": doc_inst_oficial,
        "nombre_alumno_limpio": None,
        "documento_detectado": None,
        "categoria": None,
        "indices_clases": [],
        "clases_texto": None,
        "cantidad_clases": 2,
        "tema_oficial": None,
        "esperando_registro_instructor": not bool(perfil_inst),
        "esperando_nombre_texto": False,
    }
    SESIONES_REPORTE[(msg.chat_id, user.id)] = sesion

    if not perfil_inst:
        msg_bot = await msg.reply_text(
            "👋 *¡Hola! Identificación Oficial de Instructor (Única vez)*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Escribe ahora en el chat **TU NOMBRE COMPLETO y TU CÉDULA de Instructor** "
            "(ej. `Segundo Murillo 13959739`) para vincular tu cuenta de Telegram:",
            reply_to_message_id=msg.message_id,
            parse_mode="Markdown",
        )
        sesion["menu_msg_id"] = msg_bot.message_id
        PROPIETARIO_MENSAJE_BOT[msg_bot.message_id] = user.id
        return

    msg_bot = await msg.reply_text(
        f"📋 *Menú de Registro de Clases — CEA Auto Oriente*\n"
        f"{_encabezado_instructor(sesion)}\n"
        f"👇 *Paso 1:* Selecciona la **Categoría** del curso:",
        reply_markup=teclado_categorias(user.id),
        reply_to_message_id=msg.message_id,
        parse_mode="Markdown",
    )
    sesion["menu_msg_id"] = msg_bot.message_id
    PROPIETARIO_MENSAJE_BOT[msg_bot.message_id] = user.id


def _instrucciones_submenu_clases(
    sesion: Dict[str, Any], cat_key: str, indices_seleccionados: List[int]
) -> str:
    cfg = FORMATOS_OFICIALES.get(cat_key, FORMATOS_OFICIALES["B1"])
    if indices_seleccionados:
        clases_txt, tema_txt, _ = obtener_clases_y_temas_de_indices(cat_key, indices_seleccionados)
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
        f"{_encabezado_instructor(sesion)}"
        f"🏷 *Categoría seleccionada:* {cfg['nombre_corto']}\n\n"
        f"{estado_sel}"
    )


def _extraer_owner_id_de_callback(data: str, msg_id: int) -> Optional[int]:
    partes = data.split("|")
    if len(partes) >= 2 and partes[-1].isdigit():
        return int(partes[-1])
    return PROPIETARIO_MENSAJE_BOT.get(msg_id)


async def manejar_botones_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    if not query or not query.message or not query.from_user:
        return

    chat_id = query.message.chat_id
    user_id = query.from_user.id
    msg_id = query.message.message_id
    data = query.data or ""

    # 1. CONTROL DE TURNOS EN GRUPO: Validar que quien presiona el botón sea el instructor dueño de la foto
    owner_id = _extraer_owner_id_de_callback(data, msg_id)
    if owner_id is not None and user_id != owner_id:
        sesion_dueno = SESIONES_REPORTE.get((chat_id, owner_id))
        nombre_dueno = (
            sesion_dueno.get("instructor_nombre", "otro compañero")
            if sesion_dueno
            else "otro compañero"
        )
        await query.answer(
            f"⛔ Este menú pertenece al reporte del instructor {nombre_dueno}. "
            f"Por favor usa únicamente los botones de tu propia foto.",
            show_alert=True,
        )
        return

    await query.answer()

    # Quitar el sufijo |owner_id para procesar la acción limpia
    partes_raw = data.split("|")
    if len(partes_raw) >= 2 and partes_raw[-1].isdigit():
        accion_limpia = "|".join(partes_raw[:-1])
    else:
        accion_limpia = data

    sesion = SESIONES_REPORTE.get((chat_id, user_id))
    if not sesion:
        ahora = datetime.now(TZ_COLOMBIA)
        perfil_inst = sheets_service.obtener_instructor_registrado(user_id)
        sesion = {
            "owner_id": user_id,
            "foto_msg_id": msg_id,
            "menu_msg_id": msg_id,
            "foto_bytes": None,
            "file_id": "",
            "file_url": "Registro por botón",
            "timestamp": ahora.strftime("%Y-%m-%d %H:%M:%S"),
            "fecha_reporte": ahora.strftime("%Y-%m-%d"),
            "hora_reporte": ahora.strftime("%H:%M"),
            "minuto_envio": ahora.minute,
            "instructor_nombre": (
                perfil_inst["nombre"]
                if perfil_inst
                else (query.from_user.full_name or "Instructor")
            ),
            "instructor_documento": perfil_inst["documento"] if perfil_inst else "",
            "nombre_alumno_limpio": None,
            "documento_detectado": None,
            "categoria": None,
            "indices_clases": [],
            "clases_texto": None,
            "cantidad_clases": 2,
            "tema_oficial": None,
            "esperando_registro_instructor": False,
            "esperando_nombre_texto": False,
        }
        SESIONES_REPORTE[(chat_id, user_id)] = sesion
        PROPIETARIO_MENSAJE_BOT[msg_id] = user_id

    if accion_limpia == "cancelar":
        SESIONES_REPORTE.pop((chat_id, user_id), None)
        PROPIETARIO_MENSAJE_BOT.pop(msg_id, None)
        await query.message.edit_text(
            f"❌ Registro cancelado por *{sesion.get('instructor_nombre', 'el instructor')}*.",
            parse_mode="Markdown",
        )
        return

    if accion_limpia == "volver_cat":
        sesion["indices_clases"] = []
        sesion["esperando_nombre_texto"] = False
        await query.message.edit_text(
            f"{_encabezado_instructor(sesion)}"
            f"👇 *Paso 1:* Selecciona la **Categoría** del curso:",
            reply_markup=teclado_categorias(user_id),
            parse_mode="Markdown",
        )
        return

    if accion_limpia.startswith("cat|"):
        cat_key = accion_limpia.split("|", 1)[1]
        sesion["categoria"] = cat_key
        sesion["indices_clases"] = []
        sesion["esperando_nombre_texto"] = False
        await query.message.edit_text(
            _instrucciones_submenu_clases(sesion, cat_key, []),
            reply_markup=teclado_clases_por_categoria(cat_key, [], user_id),
            parse_mode="Markdown",
        )
        return

    if accion_limpia.startswith("limpiar|"):
        cat_key = accion_limpia.split("|", 1)[1]
        sesion["indices_clases"] = []
        await query.message.edit_text(
            _instrucciones_submenu_clases(sesion, cat_key, []),
            reply_markup=teclado_clases_por_categoria(cat_key, [], user_id),
            parse_mode="Markdown",
        )
        return

    if accion_limpia.startswith("tog|"):
        _, cat_key, idx_str = accion_limpia.split("|")
        idx = int(idx_str)
        sesion["categoria"] = cat_key
        actuales: List[int] = sesion.get("indices_clases", [])

        if idx in actuales:
            actuales.remove(idx)
        else:
            if len(actuales) >= 2:
                actuales.pop(0)
            actuales.append(idx)

        sesion["indices_clases"] = sorted(actuales)
        await query.message.edit_text(
            _instrucciones_submenu_clases(sesion, cat_key, sesion["indices_clases"]),
            reply_markup=teclado_clases_por_categoria(
                cat_key, sesion["indices_clases"], user_id
            ),
            parse_mode="Markdown",
        )
        return

    if accion_limpia == "confirmar_clases":
        cat_key = sesion.get("categoria") or "B1"
        indices = sesion.get("indices_clases", [])
        if not indices:
            return

        clases_txt, tema_txt, cant_clases = obtener_clases_y_temas_de_indices(cat_key, indices)
        sesion["clases_texto"] = clases_txt
        sesion["tema_oficial"] = tema_txt
        sesion["cantidad_clases"] = cant_clases

        if sesion.get("nombre_alumno_limpio") and sesion.get("documento_detectado"):
            await guardar_en_sheets_y_confirmar(chat_id, user_id, context, query.message)
        else:
            sesion["esperando_nombre_texto"] = True
            await query.message.edit_text(
                mensaje_indicaciones_alumno(sesion, cat_key, clases_txt, tema_txt),
                reply_markup=teclado_seleccion_alumno(cat_key, user_id),
                parse_mode="Markdown",
            )
        return

    if accion_limpia.startswith("alu|"):
        doc_alu = accion_limpia.split("|", 1)[1]
        ficha = sheets_service.buscar_alumno(doc_alu)
        if ficha:
            sesion["nombre_alumno_limpio"] = str(ficha.get("Nombre_Completo", "")).strip()
            sesion["documento_detectado"] = str(ficha.get("Documento", doc_alu)).strip()
        else:
            sesion["nombre_alumno_limpio"] = doc_alu
            sesion["documento_detectado"] = doc_alu

        sesion["esperando_nombre_texto"] = False
        await guardar_en_sheets_y_confirmar(chat_id, user_id, context, query.message)
        return


async def procesar_texto_chat(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    user = update.effective_user
    if not msg or not user or not msg.text:
        return

    texto = msg.text.strip()
    if texto.startswith("/"):
        return

    sesion = SESIONES_REPORTE.get((msg.chat_id, user.id))
    if not sesion:
        return

    # CASO A: El instructor está registrando por primera vez su propio Nombre Oficial + Cédula
    if sesion.get("esperando_registro_instructor"):
        nombre_inst, doc_inst = extraer_nombre_y_cedula(texto)
        if not nombre_inst or len(nombre_inst.split()) < 2 or not doc_inst:
            await msg.reply_text(
                "⚠️ *Para registrarte como Instructor Oficial, envía tu Nombre Completo (Nombre y Apellido) y tu Cédula:*\n\n"
                "📌 *Ejemplo:*\n"
                "`Segundo Murillo 13959739`",
                reply_to_message_id=msg.message_id,
                parse_mode="Markdown",
            )
            return

        info = sheets_service.guardar_instructor_oficial(
            telegram_user_id=user.id,
            nombre_oficial=nombre_inst,
            documento_oficial=doc_inst,
        )
        sesion["instructor_nombre"] = info["nombre"]
        sesion["instructor_documento"] = info["documento"]
        sesion["esperando_registro_instructor"] = False

        # Actualizar el mensaje del menú que cita la foto original del instructor para pasar al Paso 1
        menu_msg_id = sesion.get("menu_msg_id")
        texto_paso1 = (
            f"✅ *Instructor vinculado:* {info['nombre']} *(CC: {info['documento']})*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"{_texto_menu_paso_1(sesion)}"
        )
        if menu_msg_id:
            try:
                await context.bot.edit_message_text(
                    chat_id=msg.chat_id,
                    message_id=menu_msg_id,
                    text=texto_paso1,
                    reply_markup=teclado_categorias(user.id),
                    parse_mode="Markdown",
                )
                return
            except Exception:
                pass

        nuevo_menu = await msg.reply_text(
            texto_paso1,
            reply_markup=teclado_categorias(user.id),
            reply_to_message_id=sesion.get("foto_msg_id") or msg.message_id,
            parse_mode="Markdown",
        )
        sesion["menu_msg_id"] = nuevo_menu.message_id
        PROPIETARIO_MENSAJE_BOT[nuevo_menu.message_id] = user.id
        return

    # CASO B: El instructor está enviando el Nombre Completo y Cédula del Alumno para su foto
    if sesion.get("esperando_nombre_texto"):
        nombre, documento = extraer_nombre_y_cedula(texto)

        ficha_existente = sheets_service.buscar_alumno(documento or nombre or texto)
        if ficha_existente:
            nombre_oficial = str(ficha_existente.get("Nombre_Completo", "")).strip()
            doc_oficial = str(ficha_existente.get("Documento", "")).strip()
            if nombre_oficial and doc_oficial and doc_oficial != "Pendiente":
                sesion["nombre_alumno_limpio"] = nombre_oficial
                sesion["documento_detectado"] = doc_oficial
                sesion["esperando_nombre_texto"] = False
                aviso = await msg.reply_text(
                    "⏳ Guardando registro en Google Sheets...",
                    reply_to_message_id=sesion.get("foto_msg_id") or msg.message_id,
                )
                await guardar_en_sheets_y_confirmar(msg.chat_id, user.id, context, aviso)
                return

        if not nombre or len(nombre.split()) < 2 or not documento:
            faltantes = []
            if not nombre or len(nombre.split()) < 2:
                faltantes.append("• **Nombre y Apellido completos** (ej. *Sergio Andrés Pérez*)")
            if not documento:
                faltantes.append("• **Número de Cédula / Documento** (sin puntos, ej. *1098765432*)")

            lista_faltantes = "\n".join(faltantes)
            await msg.reply_text(
                f"👨‍🏫 *{sesion.get('instructor_nombre', 'Instructor')}, faltan datos del alumno:*\n"
                f"{lista_faltantes}\n\n"
                f"✍️ *Por favor envíalo completo en un solo mensaje así:*\n"
                f"`Sergio Andrés Pérez 1098765432`",
                reply_to_message_id=msg.message_id,
                parse_mode="Markdown",
            )
            return

        sesion["nombre_alumno_limpio"] = nombre
        sesion["documento_detectado"] = documento
        sesion["esperando_nombre_texto"] = False
        aviso = await msg.reply_text(
            "⏳ Guardando alumno, instructor y clase en Google Sheets...",
            reply_to_message_id=sesion.get("foto_msg_id") or msg.message_id,
        )
        await guardar_en_sheets_y_confirmar(msg.chat_id, user.id, context, aviso)
        return


async def cmd_alertas_admin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Activa las notificaciones al administrador en el chat actual cuando una foto llegue fuera de horario."""
    msg = update.effective_message
    if not msg:
        return
    ADMIN_CHATS_ALERTA.add(msg.chat_id)
    await msg.reply_text(
        "🔔 *Alertas de Administrador activadas en este chat.*\n"
        "Recibirás un aviso automático cada vez que un instructor envíe una foto fuera de los minutos estipulados "
        "(minuto 40 al 50 para 2 clases / minuto 18 al 28 para 1 clase).",
        reply_to_message_id=msg.message_id,
        parse_mode="Markdown",
    )


async def cmd_ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    texto = (
        "🚗 *Bot de Seguimiento de Clases — CEA Auto Oriente*\n\n"
        "👨‍🏫 *Vinculación Oficial del Instructor (1 sola vez):*\n"
        "`/soy [Tu Nombre Completo] [Tu Cédula]`\n"
        "*(Si aún no lo has hecho, el bot te lo pedirá automáticamente en tu primera foto).* \n\n"
        "📸 *Pasos para el Instructor en el Grupo:*\n"
        "1️⃣ Envía la **foto de la clase** al grupo.\n"
        "   • *Si son 2 clases:* enviar entre el **minuto 40 y el minuto 50**.\n"
        "   • *Si es 1 sola clase:* enviar entre el **minuto 18 y el minuto 28**.\n"
        "2️⃣ El bot responderá **citando tu foto** y tus botones estarán bloqueados para que ningún otro compañero los toque por error.\n"
        "3️⃣ Toca la **Categoría** (`🏍️ A1 Moto`, `🏍️ A2 Moto`, `🚗 B1 Carro Particular`, `🚕 C1 Servicio Público`, `🔄 Recategorización C1`).\n"
        "4️⃣ Marca **1 clase** o **2 clases** (`✅`) y pulsa **`💾 CONFIRMAR CLASE(S)`**.\n"
        "5️⃣ Escribe el **Nombre Completo y Cédula** del alumno (ej. `Sergio Andrés Pérez 1098765432`).\n\n"
        "📄 *Generar PDF Oficial para Carpeta/Matrícula:*\n"
        "`/reporte [Cédula o Nombre del Alumno]`\n\n"
        "🔔 *Recibir alertas de horario (Administrador):*\n"
        "`/alertas`\n\n"
        "⚙️ *Inicializar Hojas en Google Sheets:*\n"
        "`/inicializar`"
    )
    await msg.reply_text(
        texto,
        reply_to_message_id=msg.message_id,
        parse_mode="Markdown",
    )


async def cmd_inicializar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    try:
        sheets_service.inicializar_estructura()
        await msg.reply_text(
            "✅ Hojas `Alumnos`, `Instructores` y `Registro_Clases` verificadas en Google Sheets.",
            reply_to_message_id=msg.message_id,
            parse_mode="Markdown",
        )
    except Exception as exc:
        await msg.reply_text(
            f"❌ Error conectando con Google Sheets: {exc}",
            reply_to_message_id=msg.message_id,
        )


async def cmd_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    if not context.args:
        await msg.reply_text(
            "⚠️ Indica el nombre o cédula del alumno.\nEjemplo: `/reporte Segundo Murillo`",
            reply_to_message_id=msg.message_id,
            parse_mode="Markdown",
        )
        return

    criterio = " ".join(context.args).strip()
    aviso = await msg.reply_text(
        f"⏳ Generando formato PDF para *{criterio}*...",
        reply_to_message_id=msg.message_id,
        parse_mode="Markdown",
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
            reply_to_message_id=msg.message_id,
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
    app.add_handler(CommandHandler(["soy", "instructor", "mi_nombre"], cmd_registrar_instructor))
    app.add_handler(CommandHandler(["clase", "menu", "registrar"], cmd_menu_clase))
    app.add_handler(CommandHandler("alertas", cmd_alertas_admin))
    app.add_handler(CommandHandler("inicializar", cmd_inicializar))
    app.add_handler(CommandHandler("reporte", cmd_reporte))

    app.add_handler(CallbackQueryHandler(manejar_botones_callback))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.IMAGE, procesar_mensaje_imagen))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), procesar_texto_chat))

    logger.info("🚗 Bot CEA Auto Oriente iniciado...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
