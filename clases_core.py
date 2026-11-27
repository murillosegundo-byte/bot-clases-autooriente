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

import base64
import difflib
import io
import json
import uuid
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
from reportlab.lib.pagesizes import landscape, letter
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
SHEET_CIERRES = "Cierres_Jornada"
SHEET_INASISTENCIAS = "Inasistencias"
SHEET_MANTENIMIENTOS = "Mantenimientos"
SHEET_VEHICULOS = "Vehiculos_PESV"
TZ_COLOMBIA = timezone(timedelta(hours=-5))

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


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
    runtime_json = os.getenv("GOOGLE_CREDENTIALS_JSON_RUNTIME", "").strip()
    for raw_val in (runtime_json, GOOGLE_CREDENTIALS_JSON, GOOGLE_CREDENTIALS_FILE, _RAW_SHEET_ID):
        if raw_val and raw_val.lstrip().startswith("{"):
            info = json.loads(raw_val)
            return Credentials.from_service_account_info(info, scopes=scopes)

    for candidate in (
        Path("/tmp/cea_google_credentials.json"),
        Path("/etc/secrets/credentials.json"),
        BASE_DIR / "credentials.json",
        Path(GOOGLE_CREDENTIALS_FILE),
    ):
        if candidate.is_file():
            return Credentials.from_service_account_file(str(candidate), scopes=scopes)

    raise FileNotFoundError(
        "Falta configurar GOOGLE_CREDENTIALS_JSON o subir credentials.json desde el botón 'Conectar Google Sheets'."
    )


# =============================================================================
# 2. CATÁLOGO OFICIAL DE FORMATOS Y TEMARIOS
# =============================================================================
FORMATOS_OFICIALES: Dict[str, Dict[str, Any]] = {
    "A1": {
        "nombre_corto": "🏍️ A1 Moto",
        "codigo_formato": "A.O.C.E.A - F - 59",
        "version": "05",
        "fecha_formato": "Noviembre 06 de 2026",
        "total_clases": 8,
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
        "version": "05",
        "fecha_formato": "Noviembre 06 de 2026",
        "total_clases": 15,
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
        "version": "05",
        "fecha_formato": "Noviembre 06 de 2026",
        "total_clases": 20,
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
        "version": "05",
        "fecha_formato": "Noviembre 06 de 2026",
        "total_clases": 30,
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
        "version": "05",
        "fecha_formato": "Noviembre 06 de 2026",
        "total_clases": 10,
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


def normalizar_categoria(cat: Optional[str]) -> str:
    if not cat:
        return "B1"
    limpio = str(cat).strip().upper().replace("-", "_").replace(" ", "_")
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
def extraer_numeros_clase(texto_clase: Optional[str]) -> List[int]:
    if not texto_clase:
        return []
    numeros = [int(n) for n in re.findall(r"\d+", str(texto_clase)) if 1 <= int(n) <= 35]
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


def obtener_temas_para_clases(categoria: str, clases: List[int]) -> str:
    """
    Devuelve el texto concatenado del temario oficial (A.O.C.E.A - F - 59 a F - 62)
    para la lista de clases dictadas en la categoría indicada.
    """
    cat_key = normalizar_categoria(categoria)
    cfg = FORMATOS_OFICIALES.get(cat_key, FORMATOS_OFICIALES["B1"])
    nums_set = set(clases or [])
    temas: List[str] = []
    for fila in cfg.get("filas", []):
        clases_fila = set(fila.get("clases", []))
        if clases_fila & nums_set:
            t_str = str(fila.get("tema", "")).strip()
            if t_str and t_str not in temas:
                temas.append(t_str)
    if not temas:
        return f"Clase práctica {cat_key} ({formatear_clases(clases)})"
    return " / ".join(temas)


# Palabras del temario/clases que NUNCA hacen parte del nombre de un ciudadano
PALABRAS_CORTE_CLASE = (
    r"\b(?:"
    r"clases?|temas?|categor[ií]as?|bloque|bloques|horas?|fecha|"
    r"a1|a2|b1|b2|b3|c1|c2|c3|rc1|recategorizaci[oó]n|"
    r"ex[aá]men(?:es)?|pr[aá]ctic[oa]s?|evaluaci[oó]n(?:es)?|"
    r"parqueos?|estacionamientos?|espejos?|retrovisor(?:es)?|"
    r"conducci[oó]n|inspecci[oó]n|preoperacional|pre|operacional|"
    r"adaptaci[oó]n|seguridad|tablero|instrumentos|mandos|pedales|encendido|motor|"
    r"maniobras?|carreteo|embrague|selector|acelerador|frenos?|frenado|desaceleraci[oó]n|"
    r"franja|equilibrio|giros?|conos|eslalon|slalom|obst[aá]culos|emergencia|"
    r"arranques?|terrenos?|planos?|inclinados?|pendientes?|reversa|adelantamientos?|"
    r"avances?|detenci[oó]n|se[ñn]ales|tr[aá]nsito|lum[ií]nicas|corporales|ac[uú]sticas|"
    r"calzadas?|carril(?:es)?|pares|barrios?|flujo|vehicular|refuerzos?|cale|"
    r"glorietas?|intersecciones|deprimidos|puentes|zonas|escolares|"
    r"urban[oa]s?|nacional(?:es)?|principales|pista|v[ií]as?|transporte|masivo|buses|taxis|"
    r"motos?|motocicletas?|carros?|veh[ií]culos?|instructor(?:es)?|profes?(?:or)?|notas?|teor[ií]a|te[oó]ric[oa]s?"
    r")\b"
)


def limpiar_solo_nombre_persona(texto_bruto: Optional[str]) -> Optional[str]:
    """
    Recibe un texto que puede traer pegada la clase o el tema (ej.
    'Maria Porras Clase Parqueo Estacionamiento Con Uso De Espejos Retrovisores',
    'Juan Figueroa Clase Exámen Práctico', o
    'Clase 13 y 14 conducción en vías urbanas refuercen pruebas Caleb alumno María angarita')
    y devuelve ÚNICAMENTE el nombre y apellido limpios del ciudadano.
    """
    if not texto_bruto or not str(texto_bruto).strip():
        return None
    texto_bruto = str(texto_bruto)

    # 0. Si el instructor puso explícitamente 'alumno María Angarita' al final de una descripción de clase
    m_explicito = re.search(
        r"\b(?:alumn[oa]s?|aprendiz|ciudadan[oa]s?)\b\s*[:#.\-]?\s*([A-Za-zÁÉÍÓÚáéíóúÑñÜü\s]{3,})",
        texto_bruto,
        flags=re.IGNORECASE,
    )
    if m_explicito:
        fragmento_alu = re.split(
            PALABRAS_CORTE_CLASE, m_explicito.group(1), maxsplit=1, flags=re.IGNORECASE
        )[0].strip()
        palabras_exp = [p for p in fragmento_alu.split() if len(p) >= 2]
        if len(palabras_exp) >= 2:
            return " ".join(palabras_exp[:4]).title()

    # 1. Revisar línea por línea para descartar líneas que describan la clase o el tema
    lineas_validas: List[str] = []
    for linea in texto_bruto.splitlines():
        linea_str = linea.strip()
        if not linea_str:
            continue
        # Cortar en esa línea todo lo que aparezca desde la primera palabra de clase/tema
        corte = re.split(PALABRAS_CORTE_CLASE, linea_str, maxsplit=1, flags=re.IGNORECASE)[0]
        corte = corte.strip()
        if corte:
            lineas_validas.append(corte)

    unido = " ".join(lineas_validas) if lineas_validas else texto_bruto
    # 2. Asegurar corte global ante cualquier palabra de clase/tema
    unido = re.split(PALABRAS_CORTE_CLASE, unido, maxsplit=1, flags=re.IGNORECASE)[0]

    # 3. Quitar números de documento y prefijos como 'Alumno:', 'CC:', 'Nombre:'
    sin_doc = re.sub(r"\b\d+\b", " ", unido)
    sin_prefijos = re.sub(
        r"\b(?:alumn[oa]s?|ciudadan[oa]s?|aprendiz|nombre|cc|ti|ce|c[eé]dula|documento|doc|id)\b\s*[:#.\-]?",
        " ",
        sin_doc,
        flags=re.IGNORECASE,
    )

    # 4. Dejar únicamente letras válidas de nombres en español
    solo_letras = re.sub(r"[^A-Za-zÁÉÍÓÚáéíóúÑñÜü\s]", " ", sin_prefijos)
    conectores = {"de", "del", "la", "las", "los"}
    palabras: List[str] = []
    for p in solo_letras.split():
        if len(p) >= 2 or p.lower() in conectores:
            palabras.append(p)

    # Quitar conectores sueltos al inicio o al final (ej. si quedó un 'Con' o 'De' al final)
    while palabras and palabras[-1].lower() in (conectores | {"con", "sin", "por", "para", "en", "el", "al", "y"}):
        palabras.pop()
    while palabras and palabras[0].lower() in (conectores | {"con", "sin", "por", "para", "en", "el", "al", "y"}):
        palabras.pop(0)

    # Un nombre completo tiene máximo 4 palabras principales (2 nombres + 2 apellidos)
    if len(palabras) > 4:
        palabras = palabras[:4]

    if not palabras:
        return None
    return " ".join(palabras).title()


def extraer_nombre_y_cedula(texto: str) -> Tuple[Optional[str], Optional[str]]:
    if not texto or not texto.strip():
        return None, None

    limpio = texto.strip()
    m_doc = re.search(r"\b(\d{6,12})\b", limpio)
    documento = m_doc.group(1) if m_doc else None

    # Si el documento está en el texto, priorizar el texto que acompaña al documento o está antes de él
    nombre_formateado = limpiar_solo_nombre_persona(limpio)
    return nombre_formateado, documento


# =============================================================================
# 4. SERVICIO DE GOOGLE SHEETS Y DRIVE
# =============================================================================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

HEADERS_ALUMNOS = ["ID_Matricula", "Documento", "Nombre_Completo", "Categoria", "Estado", "Estado_RUNT", "Fecha_RUNT"]
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
    "Observaciones_Instructor",
]
HEADERS_CIERRES = [
    "Timestamp",
    "Fecha_Jornada",
    "Hora_Cierre",
    "Instructor",
    "Documento_Instructor",
    "Clases_Registradas_Bot",
    "Clases_Segun_Agenda",
    "Estado_Cruce",
    "Detalle_Alumnos_Dia",
    "URL_Captura_Agenda",
]
HEADERS_INASISTENCIAS = [
    "ID_Novedad",
    "Consecutivo",
    "Fecha",
    "Hora",
    "Documento_Alumno",
    "Nombre_Alumno",
    "Categoria",
    "Clases_Faltadas",
    "Nombre_Instructor",
    "No_Horas",
    "Valor_Hora",
    "Estado",
]
HEADERS_MANTENIMIENTOS = [
    "Timestamp",
    "Fecha",
    "Correo",
    "Nombre_Instructor",
    "Placa",
    "Tipo_Mantenimiento",
    "Establecimiento",
    "NIT_CC",
    "Actividad_Realizada",
    "Nombre_Tecnico",
    "Rol",
    "Observaciones",
    "Kilometraje_Actual",
    "Costo",
    "Proximo_Cambio",
]
HEADERS_VEHICULOS = [
    "Placa",
    "Tipo_Vehiculo",
    "Estado",
    "Fecha_Registro",
]

PLACAS_VEHICULOS_OFICIALES: List[Dict[str, str]] = [
    {"placa": "UDT984", "tipo": "🚗 Carro", "etiqueta": "🚗 UDT984"},
    {"placa": "FSP335", "tipo": "🚗 Carro", "etiqueta": "🚗 FSP335"},
    {"placa": "DUX342", "tipo": "🚗 Carro", "etiqueta": "🚗 DUX342"},
    {"placa": "YNE31F", "tipo": "🏍️ Moto", "etiqueta": "🏍️ YNE31F"},
    {"placa": "YNE30F", "tipo": "🏍️ Moto", "etiqueta": "🏍️ YNE30F"},
    {"placa": "AZA25G", "tipo": "🏍️ Moto", "etiqueta": "🏍️ AZA25G"},
]


def _texto_indica_inasistencia(texto: str) -> bool:
    """Detecta si un texto, observación o pie de foto reporta que el alumno NO asistió a la clase."""
    if not texto:
        return False
    t = _normalizar_texto(texto)
    patrones = (
        r"\bno\s+asisti[oó]?\b",
        r"\binasistencia\b",
        r"\bfalt[oó]\s+a\b",
        r"\bno\s+se\s+present[oó]?\b",
        r"\bperdi[oó]\s+(?:la\s+)?clase\b",
        r"\bno\s+lleg[oó]\b",
    )
    return any(re.search(p, t) for p in patrones)


def _es_registro_inasistencia(fila: Dict[str, Any]) -> bool:
    """
    Indica si una fila de Registro_Clases corresponde a una INASISTENCIA (el alumno no asistió).
    Las inasistencias NUNCA deben llenar las casillas de la tabla de control de clases dictadas,
    sino registrarse únicamente en el cuadro de OBSERVACIONES AL PROCESO DE APRENDIZAJE.
    """
    cl_rep = str(fila.get("Clases_Reportadas", ""))
    tema_reg = str(fila.get("Tema_Registrado", ""))
    obs_inst = str(fila.get("Observaciones_Instructor", ""))
    val_hor = str(fila.get("Validacion_Horario", ""))
    return (
        _texto_indica_inasistencia(cl_rep)
        or _texto_indica_inasistencia(tema_reg)
        or _texto_indica_inasistencia(obs_inst)
        or "inasistencia" in val_hor.lower()
        or "no asisti" in val_hor.lower()
    )

TALLERES_FRECUENTES_PESV: List[Dict[str, str]] = [
    {"nombre": "Motollantas Richard", "nit": "13721566"},
    {"nombre": "Richard", "nit": "901517095"},
    {"nombre": "Euromotor", "nit": "901492074"},
    {"nombre": "Lubricentro Ricel", "nit": "91251739"},
    {"nombre": "Taller Angarita", "nit": "91278337"},
    {"nombre": "Taller Oscar Rios", "nit": "91281433"},
    {"nombre": "Fayco", "nit": "0"},
    {"nombre": "La Estrella", "nit": "0"},
    {"nombre": "Italo Colombiano", "nit": "0"},
]

FIRMA_REP_LEGAL_B64 = "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"
FIRMA_REP_PESV_B64 = "iVBORw0KGgoAAAANSUhEUgAAALwAAAA3CAYAAAC2PiCwAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAAFxEAABcRAcom8z8AADmUSURBVHhexZ13nJfFtf/J6/6SvO69qVeTqFejsRdEsCVGESNijyYqqDEWigV7LygQWyyJYuwlFhQUKdJZOktnKbtLR3rvZdlegPM775nvZ3d4snrvP7/X75CTmTlz2pw5zzzzzPPs1yb79u0zMAt79+4NqHoWJEO5Z8+eb+X53wL80kNd8vJF/SqzdbUlp/L/B2T9SH1J28Surq5uP/+rq6utoqIi9IlHMuKlTqnYCwH6RBefStUB1UHplT/iASQrulB9kq2pqQml9NTW1gYE5E9aR5Y6mOoUwCObqYxKATyynYWUj/4mufq/QMqY1rNAH4PCoSw05gDwTfqgN2Y3SwfSPgaswCkwQtGAb/Ln/wXINpD1MaWl48r6TJ1+8dBmYpED1JfypbwC6t/GQyl7KS8gfmipv42NhTxQXbwpZvXCo/GoLl4A/aKl9gDq2JO/yFNmAb7Ul/oVHkRATosRoJ0KgaJnaQLa6EgHqXpKT1E8YArqo0z7U7kUFDwg5cGu6qKD8iVLR086bkC8WUz7sIN9SkDjpUQfdTCtSw91QHpUqk6/9GT1SY90SAagVB1QG1TiqJ6OGaQvtU8JT5qg6qfU2MULUhdm+WVTNAF6KisrgwyQ9kFDhrth1l/xUaJT/ZQh4SVMKaOipYKU34ZZQCZ1KB0Qg2xMpwYHqA2qLj+/CbIyoqme9qXYWB+2tABQT+ORRcmCyICpLsE3tUHioVghX1VVFXQIpE++wJ+NIW3xKSFTWYH4AekA4Ev1Atgi8aRbPNJJHR71gfJP/WkfqLwA08UhBWSF3wSpbtXFT6k2SP9+K7yY0rpASgWNKc6CZFI+jKbByGIWoEl32i96VialyYb8AFSHniKggKX86mNSQKAxHeIH1EafkDGrD1RdsSDBqad6U50C6RKPEEjbsisd0tsYQKdfPFle2ugCsK0LKUX6y8vL62NEQmuxkO7GUP3SnwXppmwMRJeOlJdSKFsh4QUQcZQSUB8ldFA0KQAo1RYNSOvSOX/+fHv00Udt9uzZoQ2kcpQEVcHSIATiRd//FAh4ZFdtyacoPSA82G5sUv8nlK/UG5vwbB1e8QhTX+lLfUv1p/rwVRcLdsUDDQRSfvrTdoqSkw7JNmYboNRY5Qe0LJ/0S09aRwYd0LIgPaB0pZDaQo/4UjkB9fqEV6eEGEBjE5ZtC6hrACCDF58QGDRokB144IH2ySefhDYgXal+dACiCVTPyoCyAWT75Bd1AF7pEF0o++pPS9WzAC21keWlrn7RVQeVZAB1thDEML1wUt3wUAfhgwe64g5QpjLiT+ugALpks/5QB9VPHd/ow7bsqIRP/GldPPJf/aDsyCfajCfVR10IpKX0grKlfkETdaBYToDUGYgCLoONGRXAA8CDHCU02QBI+F/+8pf2xRdfhDZ24MvWG9OfgnSLD0ztZOmMQeMDUr9EE69AferHpvzLgvRJJ6XiRluyagNpSfxBeCQPpjKUjEFxSunflOgg9fTCkZ6UJwvoQybtkyyAvB4oQc23EFuyRxt+AW2eT8QrHfJFNHhkQ7ZB+Z4FyaY6UjpyYYUXEwwKTMooQ6kyhKmLT/Kg2o3ByJF5nvCH2aeffhraklFdcrKbAm38E1/Kn5aiZ/tSgCfVw3goBZJP+1M9ogsF4qHUGLK8tEHZVFv9yJFwWuVpS5dQ8qIr4aETI2Q1l9DhEz+l5lgyIH2qp/S0LRANXdKHzdQeJSh++tL5g0f8KS+Q0kDpS0E6Uj58kY0U4APDObyYIMh5DQikDzo0kHbqOCWy6le7MZg4caIdfPAh9vrrb+Qo+6/WgrQugJZObFriX5YuUF19AKX4QPnfWB916QfELz7xpnXxAfAqXml/2oYHBLBFPbUpPkpiQD2L4qOfOUx1AvDUOL20tMwqqyq9HW3Bn8ZP9lOb0ieAJv3iS9uUqU+0lfDUZU+yjSE8qV8AZSqTrWdlUj8oQ8LLAXXArAFCoy4aVzF0EJr4VRdIVxbyRoyxH/3oAOvxakz4VJfsUadUXSg+6OIFAZVZyPKkckLRxCNbjfGAxCptp/zyD2RyuS1n6ULJgd/Un2JjMiD+pPMFiB+Arvmt83qNY3Wdbzly4wjjqfVEqfGFbm+DXclQl25oOi4Vyhf1qy0QTf5IZxZ0odIPfyqjumyKposfUJ/6oaf+gE0gYCQFOmRQjKIp+dP+xkBOZOGfH3xiBxxwkPX5ol9oow8fcBwEJJsOUJj2Cb8N0v5UJkXGQSmeFLK2NSm08RsapVZwgXihU1e8pJ9S+E1xVD/2NLFCjgC1v5VvaV32AErJ42tFTZXV7fVEx+6e3Fx6wldVuL66KCM9yNFPHWQ8fPYgG/SjkzoyKqUD3yULAmk9hVQHdtArP1SmsrKvtvigiwZAg4+yCZ2gAEbaMiDDKV18WUwh2xb8o8dbdtyxTW36tJmhjT4w62RjOgF4CSIgHjArnwJ0jUftFLNjkk+iA6ITD4KHD9RTXnxobBzZPvXT1kQA6kuRPk08iF3aKuFJATp6BbItfTW1NVZaWW5llRVW46s84qHP+fZ58hsrvPPvzfEzRuSxha/U5S8AD23FBJCf8CpGsi9+yiyk/fKZuiDVk9IBaALVxScZoP5Na9qBMRzWAOU8g1bwacspYQrZtqBb16ft+ONOtnlzF4Y2E4BOQLalP+ssAK9uqaJTl19Zu9IpXwXUhWmSUGZ1UerCB+iTTmjIQxMCyKifOGproz5Q/PBJDqCufskArKylpaX1/inJqIsXXams6kDwB9y31yp9PGxp0L7HV/vaap9rX+VJ+Gr3VT5Jj/KBEgQ0NvqpgwD9JSUloZQP0ic90LIADR544ZNt8UqH9Kkv5aGEB4BHMpT4F44lqagDkBBGRYdGnYmjFA8lKDlh2geqv1PH2+2oXx1n8+fHhA8Rz0HKL5k0qOgEoYmHulYhQHLql6+g5EUTH23FQLzClAdZQD5IHwgPpfhA6VRdOlN+leKl1GRTyg6gfslofqRXdJCYpQmq/j307fGY+kJDorOSBxu+lSHh93nJNqee3/XLD9Gop89yIDbpA8QjX1IfRAOzAI90pXrVTvtTW9CAlC6ApjhQ1q/waSkGCUgpyEDoT40w+DQggOoprcL3iB073G4tzz7Pli1dHmiNQSonm9RTP8WnftEoUz5K0eFVf9qWHuoAbcaPXrWlTzySh0+TD9DWBZjVTR2d8jeNIyUom/RThyb76hcvNACa7nrqp642SB3+Wt/G8MBaWeN3rJw+/OBhlS0NCU8pG0DWXtCTi4/s0cYGINvwqZ9S8upLgbbsgAJiqW0ROqQ7BemibEyv+CnDCp+FrFEJyXm1ZUB0UPrUFzGQbPnylXZhm0vtyiuusVUr1kZiI4CMSvSB0p/6oHoK8Cr4Ka90ApJNJykF+pFTXyorelqiizp8lKoDlNiBD1Cf6I3xw6s2QL/sMflcYNIBHUSX2rJHW75RIkvCs8J7b1jdw4rvCc++XXt4Fwwy8k96ZENIG7p8E8i+/KUuFCDfGMATLsBcv/SLjk75BMqu2lmQ3wC84ZSGihQDUiRle3K3OMHu3WW+ojRcdTJGO+Xb63vFQMu158wpsl//+hxrf8tttmH9lkDjYUkgeekSqK2BC+U7/bQlqzoDBcVDiUy6sgLSl4L4pRMUwKvgyxZIPbWVbYtPOkSjDtDGP61oGzZssJ07dwYe+awxiZ96GgeQOgh/miTUw1EkbZ+bal/l6+XQnRxL7vGLAt0N8sQ7tvEPm+hHFpBdSvkDwit+ATyprIB2qkc60rra4qcNoE99KULDNnV49kt4JUdEBIIuZyQAe23GjALr1au3DRs2wlasWBVoBEAKs7DHg8rxlyAvb5Sd0vwMu/GmjrZpy9ZAS+XkJJB1XjQNDJsKmugpvzCVo53VCz3rv/iyMqksMowdHtFToM02I01OeEFA/kMX0IYOTJ061U455RS77LLLbNmyZYGGvWy8kVFCoQtUXb5S19wC0Pbk7IP4WS/j8x778BsdyLLFxQbJy/atIU+QB1SnlC7a3I142BYt65uAuuIpn+FN29QVA2hZXY0h/cRHMmFLkyolMNSjExGB8vIKu+mmW+yqq64JKzUrD6sAvMiDWXBXAgo+6/2FnekrfOc777PNW7YFGvKyrVJ16aUt3tRX6gxEPktOvCkt1SFZSqH4AQKpAAHixw59oilRpFM6hOoD6Zec2ikdUFJSPvDAA3bSSSfZCSecYE888US9P6BAbSGyYGoXvwGNlzL1m1JY7TYiP7Lwk1AshLxA4+w90igB6ZMt/Ob9gGxlARuMg36h+OSLdFKmdfrVBtUWyocswqu40g5vWmlgDEJsc7VhHKFAsh07dtqll15uTz7ZNRIckANQinwW3ERI9/j/Zm+8+Y61ufBSe+zxp2zrtu2BxgqSHYT8oBQCaZ/qKTIw6cmifFWdUnZBZBUYtcXTmC1QQH8WUz5io/jKHjySFSip+XT62GOPtddee8169Ohhp59+ui1dujTwaIwCyae6tQiA2QuFfs0VJatf6gsILyt5ZSXJSxyIKbqJD7INfKkcunbv3r1fUkMXj/wQXXFRv2jcEaRD40jlQbUlJ0h5QMmrLyQ8lVSQRK+q4tbZEFj27O3bd/TV5qkcJTqpgYGpobB/95KTAKdYeUWlte9wq53bqrXdd/8jtm79xv10IMNg5VxjAC98qU35neoRZtvSwW1W/GkwVab88kltlfBKViiAJ9WlftXle2N8AJ9Os7JPmjTJVq5c6c89v7a333479IkXJClYrdElgI5f8jm1Bw0ZjYk+If1KPhah6mq2Y9Xewj98JdZxq1NXF/f1kbdh3iihZ+2qTgnCB4/4KKVLfqR0yQEaGwBNeuWDeFP9AupNCIAgNRxLmBhQFHrzzbft2muv94mYYhs3bgzGwv4xZ0hGhO6K1fq2B1i+YqWv7pdYu2tvsI6dOtuy5asCXbwA+sBURxbSATFI8QMpDaBM/QJok/AKDvQsX8pPqT5KEMCOki3lF4iWomQp5SNAH/rE9+qrr1rHjh39rroj9Hfq1Mnjfm3wG5AudCAnlHzqJyA6/IxbSQ8qeRoH9MQkr6lhrHG1p0SXxoBu6c3GFRtqC6jL/5Se+i399KeY6qM/racXs/ilT/UmOuKiAbMMKcm9y9uxnp8/yU477Qw7+OD/9j3mgyHZkU0HKYcAHXsBU6dO8wewK+yBBx+1u+95wDZuiqc08Eo2RdFTgM6goEcftSLFAadyIH7BI53CxkB06ZYO0SnR1ZjNFAHpyPZJnxBdxF86AWJ63333hX07PMBbb71lv//978OpDZDKSC9tbMp3ABoXZcoju9gRbNmyxYYNG+YL2pvWs2dPGz58qBUWzg5J7hZcJwkekYT3EXpJfPZPaJWKu0qAPiGQ8mdpqqdjAdAFDaBfPPJB44cukA7x1/8BCCjBOEgCFI3J5rx58+yoo46xJk2aWLt211pp6e5Ar7+inBH5gN5mdY/6zAYPGeoT+aA999xLdtfdD9j2HbsCXfwC9HwT0JcdMIC8fFAAQC4OBSDlh16yK776BlL71NOgpWXqKzyp/mwf9qGnSD+lAD7dJSS7evVqu+666+zLL78MbWDw4MHhtIY/jwQ0TyRt+vs1GjewZs0a6927d3geoB+MF0rDvn7Xrl32/vvv29lnn22HHHKItWzZ0m6++Wbr1u0pp79rI0eO8DkuCfriSo8c441jZsurBYg28aQu1JhB7OObfKWPOn2SA1J+8QLQFG/qyGfpoMYmWepCILxpTQkAzLV8LuoDotRqsGrVqnByQMLf3KG9Vfo+D6iqrcl9iNSgg4SvZgK4eLw9ePgI6z9osP2txz/sjrvvta0ebLglQ5kdCHXRAbVTHvGpJHhpsgU+L7nTOMWfKepslSfVvAXz/bmiIvTznMGYxa+yMaAPW2nwUxRk9YHip8THdKKhASTgOeecY3Pnzg1tYNy4cWEfn5eXF9rIIM/F0ljCUXbv3t2uuOKKcKQpn2N/tLNi5TK7/PeX2g9/8AO7+OKLwl+gcZFw+rZ7d4ktX77Munbtak8+2cW2hQOGhhhTxjE0jF9jApTY8Gqc8k3yIH204RV/2ic9tAFK8VAK4ZMc/AB0Qcob9vBiFpOMgQSVc1hg06ZN4cSAhO/U+XZP2h3hrL3GZXl7R2IBGK12mWpv1/AM4LR5S5YEfOIvf3HZO21LSYnfGP81GQDaBCj1Q3QFBloaROhq14PTA69PMn4CfD8yYWK+TZ0+LSa607kYkJUPlN8G8DS2qoCSl1+UaV2lYpvKU//DH/5gl1xySTjeExQXF9sZZ5xR/3fAklFc8B1+tZd4nM8999xwygOwYNGHnAfFqqorrH2HG+1nvzjAPv30I5ctC3xZ4I7+0ksvhQdnADsxXxrmAdAYhdiihEdtZAFotJGVvMZOKTnRaYsXuYZx7G8XgEcATbySh9aEYFAREaDEQb0wYJAAD6rt2rWz733ve3b8ySdZy9a/szfffTusnCRO+PrOZZEh4flDA1zb5tuH3l/2tXmLl/gK/5rddd8DVlIWg5y+aZV9gHrWL9ox4A1BBajjr9oAMiG4+OC+6QVYhd/W33jrTRs7flxoQ5eNepncOFLMgviwqbroWnmpIwuPUPrgpx9gDqBxBz3++OPtxhtvrE8QYPny5WG7oQSWDhCd2GKuJNO/f38766yzbM6cOaGNP7q4gCnT8u3Qww6yv77wTGgD9Ml3/BEv26Dscx4g+0A6LuQUl3RO1A+qjzq8mlMBdWEqByIHP3L4Kh4gLaHDIzv4RT3s4bPKqcMgR2AEuOV98MEHduihh9pZ57a0J3yv97o/UPEhEknFUaSMxi1NDFr+5Cl2/oUX2fSZs+ytd9+zjrffYTt3N+z/ZUe263VkfBKKJr5sP0CJXvSzqlflfCwsKrIHH37I5vrqBcjvFLL6pFNAv4JJn9rELCsLLU0C8QS/En5gzJgx9sMf/tDuueeeehqwefNmu+CCC8LPm8AvQKcWLPkDvPjii9amTRtbv359aItP/Z/0/Ked1PQ4j0G8IPbu5Uef4t/AyldiRz32738HpF8XhgA6MoB0yK7GKd/hha6xp7pTHtkXHyU08ab9aZkF6MKwhwekRHWBaLxQ+Pjjj238+PF200032a+OOdr6DxkU+uBmRSdxGKScCt9suPjDjzxmBx9yqBUVz7M33nzbrm7bzjZvjZ8WsAIzYOyAkgfUTh2mT6sV/ZTQVVc79Ie2B9S3W2El93+jxoy2W9q3tzXr4sdrddLhJZFAFpBOTRxtvvGePn26bd8eX5oByKqfMvWJUvVsW3ySAd555x377ne/a88++2xoCzhFufjii+2uu+6q5wWQ1SonRPfDDz8c+LlQsME39OID3njzNWt57m+9P14QnL4o6YS6QJCRHbDM78zqA0lyLVjIaZzIUKZzKx5KxSSlqy+VhweQXvWLLnm1BSlNYwPCXzwBGABhokQghZkzZ4aA843H008/bUcff6yd1LyZ9fKHnTrf0wEklAKBDmDL1m129tkt7brr/xT+gPiFF16y31/xB98ebQr92E+DkjpOYLnQpAte+ak6peQkSxs634SELU1uOwN8NWigXfmHK2197oiPPvRIXvoA2kym7LNV4BSDUxN8BuCRvGRTfdIFyGdQetXPtuHee++1H/hDpH7CRH1caFdffbXdfvvtQQZQnMM4c3ZBkpJjzTvvvLP+BCc9xiSm7N9/d/65Htt4UkbCc3qDvHQqkZHT/EgXNlI/4JMf0Bk3QB1EVrp0sQDyOUXZgF++KMb0A9IPQIeHPsVb/ZRZrP88mDIVSttA3759rUuXLtanTx/71ZFH2pPdu1qPN163ps1Psev//GebMnVKGEwWhg4dbkcccaTL9w/tV155zdq37+SrS26Fd4eZGGzhuIIBiAbgB0GABigYLMsaTIoCEp7tlR6o33rnbbvhxj9bye543Ba3NN6fe9cApOMWcOG98cYb4aXQ8OHDw6Sntijlm8qsjm8DVuNbb73VjjnmmP1+lQ1gnLfccku4s0o3kMZCtnhZxTPA3/72t9CGTr/iunPXDju31VnW+oJWVlYeYxCPGfdPmJhgMeZ6IKYPpK625kf1SG9Y6UUHU1pqK0X6ySO+09KFJLrs0AYkk0LaL9sC2v/yaYEMiAGgzW2WYyr2hz/+yU9sxKiRoa9g1ky73FfMo3yL06FjBxs1alTY6/MANmjwEH/YahXezm71lR54+eW/W4cOnbytb2kaBq4klt3GIPLHCQHhFwqQ50SJc2bobFtIamQ739XZbrz5prD6AyR8+Osf5yPQ+K4TEvgFnHzwYiY/Pz8kP33Y10UoXspUTsDY2A6wWq9bt84KCgrCceOECRNCie7LL788fCH51cCBIX7wrVixIjx8cg7/xz/+0bZt27pffLK22P5cdNFF4eIE6FdcgW3bt3qyn2sXXXKB7S7dGWis8GmSUjI2SsbHH+4QGxammpr4I68ad0zwuIIrDtIRde8JcsgjJ15iDI1jYnjFTwkPcyw6vkEDZCPGvuEuC3qtvi4UvzDs4UFBaHupBAFQ/tRT3azL409Zpw632cknnmKLFy0JfcBu3yMO81Wvve+NTz755HBWDzZteqK1bXu1LVq0IMdpvkL28Iey+3xf2ZBUOMWkEAAQ4GTitdd6WMdO7e3+B++3np9+Gp4BkoU4HK/xBw3IE8QpU6aEhCQ5OYceMWJE4FOwCHKbC9vY9X/6U30C8Nyher9+/exP3jdt2rTQTmHU2DH291dfsUWLF4V2iFGdy/qDOZ/UEiNQk40/CxYsCPtyVufWrVuH83Ve8px66qnhwZ+XPc2bNw9JfvDBB9uPf/xjxx9Y05NPtNNPb+GxPMlOPPGE8CHZj370I/v5zw+0885r6QtIO7vjjtt84Wgf/O3atZv7PsDGjplgb7/9rjVr1iy8dAIYO35phd/qz06tWrW0tte29fjFu5S2PPgO8oa1pobvjeKY4nk7Y0KP8/i4edCNd4W4UDLeKO+xqCWmDXnE6V14VvOS9zUsMFU+z9RZcBqDPSGm8UUXfpD08cJlPOSL07yOfp4Vw/Oi17EZ7Tqfy3CnwD9ho3/xhAApQAkQlJdffsWu/uO1dtavz7F2V11vu3bqDVyDPFcyyfLhhx9az56f+APeZF/RWEWiJoL2+ONdfH95t1VWeGI7mUHowqIEZ8yYEc6d/+M//92OPf5oa36qJ0XzFnbrbXda/sSp9nmfvjZy9EgrLfcV3IMB8IaS7024E/Xq1Ssc4bHKMwnykRXzN7/5jXXzO5WAfoBEeOSRR4KcvhNSH88er7/9pnX7S3fbtGVzoIUtkI8noP+vpjp+yAUQh48++ii8szj88MNDAnK+zucCzzzzTPgCkjP1gb6SM1ZWeJK22SlN7YEH77aRI4fZF30+84fLf9i7777t28C/ezzOtKuuvsq6/6WbPegLANihwy12/fXX+6LSzjp3vsseuP8RO+uss+2//uu/bMCAAcEX4sn4GQtval9++WW/2A6zhx96OPTTR1JTapXmu/eamvhSDmCOFEOAHK2t81XfL4r4qUG0wYVBQnqeheTzyyG8awGYYZ+JkOQkIsnPYQL0latX2TvvvWc9P/vUZs6eZZOmTLaNmzd5H/rYyzecHjEeLjqSHpfSLWu4gJyIzmi7YdEWhBU+C9G5WApe9b13s6YtrNlJLeyRBx63ncmnAeiQM/tDbHOVAQT0ttvusPvvf9ATBP5490iDyWrDHpSE/2rgAFu3fo1t2LTRFi5cbAv9rvL+Bx/bd7/3fTv3vFb+4LkmJ2Xh1s9R3JVXXhneFbDSA+hX4vL2kq8Q+V5EIJ9JELYUhYWFoY1P+AusWrPaLrrsUnuqW9eQ/AGQA3NJH1abXCzYSnGaRUIXFRX5NmRbva7GgK3OCy+84Hef39mcwoIctQHYBrVpc6E9//zzOUqEuOrxPU5V+Hx765bt9vZb7/idtWn4NgZg/JobFqMOHTqE9yiMlbeqAj0oAiQ8czd3bpE99NCD4YGZB+GRI0fbiuWrwwq/Z29NfcKzyscFguNG3t94gnpMWL1JciKsVTisuk4nQYFlfie/+NJL7IhfHWEv+4Xd+4vPw9w+9sTjtmDhPNuxc2tIeF10XFR1wY7Pa13DXSRcRNwJnAJNSB9jF4Y9fBZgzCb8Z5/2thOPP9lOaXqadbixk61aGZNNk0wZHYoJttcDUbeHM2JW8Jhw0J988im/Hd+53x0CBNAzZMiQsC3iRAQgZCk8+miX8Kb3T3/+k+3aHb8mBEhUbv3f+c53wh5WXxrKJ4DPbQ866KCwdQEUBICE4w0ne2Ag7RszdqwddsTh9uFHH4Y24wh9dJPwPgHhRMgRe7rAsqA+ZFM+7i6s1DffcpMnPwsJ+2C2CNHvOXMKfZtziif8X0MbIMkag16f9Q7byXQ7l+KsWbNCfEl6Pj3gTsNzC75E34KY3w2K7cwzz7Dvf//7YbtFzI855jh7zOOfnz/ZqmsrwhynK/w3AUmoRGRV5whbW5l/fvhPO+zwX/qWtWfYzpaVl9lUvzCHDh/md/PxtnLV0jBWXbjVfif9JiC5w0XmvsgWzwgphBWegaYOh9uDXzHh7DoXgMLCuda8GZ8VfMcOO+SXNnRIXEEAZJUgEZkQAoEx9IJR0TvvvucrRltbuyZ3Dp5LEgA/OnfuHL4b4ctAQlRRXeZ+5Ca+sNiDfoL97Oe/sNffzD2U5S6IRYsW2VFHHWX/9m//FvbNKUg/v2vJT3Xr2C/uCaP8c889Z6eddlr9Nyxx8qPPz//1eTv4sENt3ITxoV0fL8bpyU7CUyoO9GtFkh7qjBU6oD6AC/HII4+0v//9b64SHfSxcES/3333fTv++KaWN2JUaO/bx8THdxHoq6xsOB178cWXwnNBdhyyxcJw5pln2g033BDuGFzkrPr6M0KAl1AdO97i27HDLC9vhK1duyYcWPCboHfffb+NGjXWtu9gYWClbfjFBv4qbvDgITZh/CTbVbI7vORb7XfHsJ92HpKRJNQbeeCZ55613/z2rHAXB/hERVBdy2fc8S5CPCTD1ri4eK4/s/mF4XcykLsqwB2Yb7xIefKXOwJywv3+xA+g5Aoh4eMLGyaA5KizYUPyrEWL+C3Nu/5wJH6hJjwkgg/OTXqJbjBOXr9+/e2C1hfagtzv0tTL5IBTIG63GkBQ5sAAb7v9Tmvd5hJrd9314XhRfQAPuUcffXS4WNauXRt0ojtNLPaw3AV0QaQJONZXcbZR/Jw3gBzAg+41ba+xs84529b4xANYDT9rge8kZahHeyQp9mRTsQHTOKc87KuP8DsIqy+AX+iBlb3qo490sfY33+p3rRKnoZ8LNV6sIJMKkAi3+5aRhCYegOKg7QpbHe5yn332WWhzHMrYORWSb3PmzPRnpqZ+Ab4c2sD27TvCJ+HPPvtXn5syt1vjScz+mgfLvWFb9thjT/jzSnN/PnjMbrq5vR197DH2nC8Wte4juRT277m6Zo73Ir9rfb7lT5oY2lXuZ9ieON9ev7B9VO4/8YgXAneBHj1es8MOPdwuu/RKe9YXqt/63NxxZ2db7fPDYh3zVgt2Q/zB+mPJFLhKKviLdq+H/ZYHUjBm1Hi7rf0dtmBePHmhLwY+TrYShf0lWxolu+h5eSPDUeW0qTNCG7r6ACak2SnNfEI+9Vv9Ftuwcb2998F7dro/tHXr9qwtXrLc2nfsZPfcd48Hp+EDK16M/eIXv/AV6O7Qxi+d+uiMnYuIFY3tAz4zbAWSOwrPAC/nzq8F7Hubt2hhj3Z5IsgAHroQs/DTFihxW/zVv+yRsBGIm2MS35igDXe11f7AxqkNzx7s1QH6OOUAigrnWatzL7CPPowJWuNbCfa0xFQXBqcnwLatOzzZf2OXXnppvS78w6YSnrsJCwOfiKQAT0VlfFAdNnyI30mPtAH949YPKPEVm4R+/vkXXZeP2/hgjeNZnz9HDguaNz/Vevf+wm7t1Nl+deTR9vRzz9i7779rxfOK6xNRpyfUARaUP990o51/Qev67/3DSU54uPXtsCMJH2Np4TDggAN+Znfdda8Vzim2rdu2hu3POee2DHcK3qQzP0r4OM9xvkDf0sTkUOLBwM83VNbyg5sY9v0WDvpFQIARqiz325g71JDMKPbg89DgNK5M9u57QsK7w95mRQI4fz7zjF97MAeGtvJBFxVn3A8//Ij98vAj7ORmzaxps6bW4tTTrPNdd/teM26Derz2urVs1crmzY/fwwCDBg/yh9n/Y927cwLDeNxn950ViDoIvP/+e+H36SdPmRTaAhL11k632pW+r9X+H+ja5Sn79//4D+ube6bA1+CvJ1qtP6DxALePEwNfZRUjSu6IxDIsCOH0grh4YnmMOT8mjuh48YUX7eCDDrEPP/wo6A/PPh7beFxaZ0926WonntjMZs8uCv379rEC+kPcHt/S5J6P2N4ACxYs9DvFEeGNbLqIAPgMEEP+iOemm26ujwl+4J/OtQtmTrNTT29mffrkjjbdBr48+OBDdvFFl/hCxDsVHwd2Xe/IvDw7/JdH2FNPdrVReaPt1o53+DPYV+GPwt9862177PEnrLQsLk6s3PxUNzsHAat7U1/knn/hr6gLyU4O1tRVOVa6pXhxlJaVWtt2ba1Vq/PCQYCAmM6aPdsuvORi6/b002GsNXzW7uPhjkKOCpvU1njFbacYLgA3UuvBrXaDtf5EXoVxb/u0ugmfYC9JZjcXyr2ObGEoq6vL3aBfFPt8dQWDTHR6re/pTvUV89YOHa3OnQI4tyX2e3OrWunu8vDrwo899qR16/6cTZo8PdAFxXMX2HG+p33yqb+EwQLjxo/1xPx3u/SyS8LLFSaIC7aOW7/7X+v+A1u2brKW557tt96TfFK+tJmzCmxh7j3Baz162EE/+7nlj48nPOtXrbcWJ7WwU08+zVYs91u+02qqPXHJE2/U1Xky1/rKEWLGnpP9qSeH12tCyVrjfT4DtT42Xxi97qIxz2z75h3W+rwL7YILLvE7WXxYxlcWC2K8c+dWu7DN+XbNVX+wirLdIUHLa8t8HqqstGqX34Wd5vGtqY3JNHv2DDv4kF/Y/fffH9rArl0ltmTJUi9LQ5s7wl133ufPO8fZmtXrAi0FkqWsYrfd8+Advk3oaBUV+nR4r91x+612bbtrfMFzmvPVVvpcl1fbA/c8YGf5Ivbq3161ezvfbRPHTcjJmF80A/yB+zSbMSO+PSY2xGPzlq32yhuvW2Fu0erZu5f9tuU5Ph9xWxeeTXzRLa8r97Ii7DR4oL3O7858KwSQp1UsNl4v8UX4pVdes8/79AtjYJ7Iq7iyE3DuSsl/xQ9hGAXUmRwgTBKla4aug3+OFsMLBngTeYLK0RHHRHvcmPdEBfS57L13PWAnHNPUvs69vAorYFgFfaXx8puAZGLlw5833nzLb2OtfHWYEvp27CyxRx7tYvfc+6Bt2rzN7fuWxjMsnNd6PRyl5S6ohQsX2dVXXRP+euuUU1pY125Phgti4cJ5dvppLayTP8Txev/uO+4MzytdHuHP7YhD3B4IWBH58aKwrfA7CRdY8C8E14NNzWPAQzd3ykpf8ehXLAb2G+Cr+8H2l788E0hcMFVVvnAwgw5lviq2a3udHX/ciTZjGseV+zwBfBFyfVXc9v2ijvGIAtwd27e/2c44s4UNGNDXb///DC+mHnvscd86NRzhDhs23H7605/afffeHx5Q2e6wl+d4U74tX7XC9+E320MPPWqLFi6xsaMn2B+vaGvvvOlbIYbAVe92q13+vnvvtuannGwXX9jGPuv5sZMZe4TComI74aRmflf+R2gTE2DJ0iV2+m9+bW+EZzFPWPe9bdu2dvlll9kaf54AuAuwEavymNf6pDOn3bs/bzf+uaOVlfI5ilklV5DDV0OG2ZU+p3NzW23lZHDT6yHsXvoeHtcQkosqvwVyCsJKl8L/QhSYV7zQfnvGOXZX586elPEPk/GBu0P0hcSOd4/Yln8+cH9Qos7tbtKUfA9cw+lCte9ly8r9Cd1ZuW5AgtKYX2VllTZ96kyb6s8Smzavd5Z4Bxg3fqTv89tYi+bN7ICf/Nh+deihNin37fz+/lAX+v+z0u4ptzpffeM0qY+7mF/4Tqv2bQh3QaDGE+Xaq/7oW4H/DkeAAH57Du8Hw4aNso4db7eJE6eGdtSW0+z8pFaY0NwY586dbTfe1NYuvPg8u+rqy/2Z5K+2bPkSny/iGZlq66rtmWe7289/cYC1u/bqgBdd3MYKC+NPmOdU2bx5i6xduxv9GaKNb0Nb2l13PGhrlsWP/qIHxKHWujzxiP3swJ/aiy8+635A94XBL0qAbeUdd95uj3V5NLQlhw8vvfyS3eAX5Ny5cfwL5823pscdb2/20MXRMF7GCkyeXGAXXXilfdV/aCQ41Pok3+N3tcsuvyJc9IKQ5LjooDxosnLV136bX++3mDW2bsMK2+Tl6tVLbMPaFVaya4vt2LbBNq9fZZvWrbTSkq1WsmOzbXf+Eq/v2LnZZdf51bnVH3h22c6SLbZtxwY3us0qqnfZrvIttrt8m+0u2+63VOfdtt52l273O0O5/f3lF+zAA37iq/L91rff575tGWfb0b1jkz+IbAjltu2bbMv2DbbDbe10Gzvcn23Os9ntb3d9tXsqXfdO93m9bXXe8ooSq6gqtRK3sc192+Y6GNuOnZsClpVvt5344bSq6t3Ov9NteX/JJtu8bY3LbPDb+Q5bsLDQV8g+dkHrVnZt26ts7aplzrPFeTa4nY22ccvagPiEf+s2rLKNm1fbrtItVlLmPu7aYJu2rvExbIylt3eVOa/7s8nlKivLPK6b7fGH77Pu3Z6w7e7r9pJtvkfd6Rfido/tliBbUrrNL6JK21ay2bbif02JVe0rt227N9muim1WVrXTx7/Vx+XxrqSOji1+AS+zRUvm2Jp1X3tilYa+Uo9JZeVuKy3f6ePe7T5tsae6PmYnn3KCXX7Fxfb8i0/b0uULrdx5Sip2WVllaUjKFStX2HB/KJwwId+3QGutfHeFlZWUuk3Xia8+18VzZ9nQoQNs7boVHtcy17/L9+mlVunzXOF67r73Dru63R9sw6ZVfjcpCz5Xed/s2QV2/u/OtdNPbWFPd+1qz3Xvbof+7Gf26P33WY3f3Srczx3lPr5qt+dY6borqzx3/v6yXXPNNeGP24fnjbDPeve262+4wT76+CO/49ZYRXmFbwErfFHhFzVq3M8d0Z+qMmuSN2qojR03woYM7Wtjxg2xiVPybMiwz23I4E8tf8IgyxvxhQ0b2svGju1vUyYPtRnTR9q4cQP8ybiX5Y3+0kaN7Wf5k4fYuPyB3u5jI0f3tYnOV1A4xiYVDLOJ04bYmPwBljemT+CfkD/Yxoz5yvV+aTff3M4OP+JQa3Xeb/2W95KNGTvCRo8ZbsOGu66RQwIOHfGVDRn+lQ0c0s8GDe1vg4cNsH5ffRHKEaOG2MgxA23IiM9t3MTBNnHqiICTpuW5vX7W+8v3/Vb3mY0Y2ccGDfnUsaf16feBfdH3PZft43p7efs96933I+vdz/sG9LK+X31uAwf1tX79v/Db9En25z9d536NtFEeo+Gjh9rQvEE2dOTgUI7M0Xp+/pF9/OkHNmb8cNc70PoN7G1fDe5j/Qd+7rZ62udffhzag4ZHn0ePHmYjRwy2vGEDbWL+aBvv+KWPadyEoTZt+gi31991fGSD83rbxJl5Nm76UBs91cc3O8+mFo20STOH2ripA23clIE2bFQvGz2+r495sI2d0M8meBymF4wNOG6CJ2q+y4/Ps/yJY2ykz/WwEYNswOC+NjZ/lE2Znu/yQ2xm0Qxb8PVcGztxdBjbkJH9bZzP4WSf6/zJI2zGzAm+CMy24uLpNnvmZJs2ZbyN8nka7Ener//nNsrHM33GJLc9xsZPGO31yVZUPMtGj80LePd9ne2HP/lPu/+hO32seTZ+4gjfio503jGetP+w1ue3sp8f+BM70O+ovzvnbOv18T9twZzZrs9zYfyXlj9jmI95lI13n6ZOH2UFs8Zbj388bx06XW/nuewFvpV6ossTvvpPsunTptrwIUMtf9x4mz2jwApmTPR8HBbkpheMsSY8xK1es9xWrlpiW3ewyn/tV3qhfb1sji1fURzKNWsX+oqzylavXeCrNHeBpbbo65m2fGWxrdu4xNasX2grVs0NuGHT17Zx09JAW76u2Fatn2dLls+yonmTbOHiAre1wJYuK7INbmf9+hX+dD3dZhRMdvoyt+e61uLLUr8NL7JVa5bakmULrHDuTJvtkzKneGbAufPn2PxFxVbkK0vx/ALnKbSV7uNiL4sXTHff5tjajUtt3qICm1U40eYUTXQdk2zGrLE2u3CCzV843fVM9Hp+8GnBkkKb5xNe4DaK5rmu1cvtb6+8bIcddqh93vsL2+L72xXu17LV7pffESnnLiryB67ZNmfeLJtVXGDzFs7xVbXY5i6Y4zjb/St0e5Pd11lue4b7Mc15Z3s8F/mzQrEtXTzf1rmdFcsX2wIfS8Gc6TZvgcfUY49PRfOn2PyvC6zw6xlWuLTA5q6cY3OWTLdZ8/Pd16lWtMiTyrFgzmgfC2Oa7D5MtsLiSVY8b4avujM8rhNtvvszf0Gh76VnWsHMKT754236rKnBZ7DI72ZLViy2lb46L1y6INCme8xmFo33WEzxMc2yr5cW2uLFs22B+7fIE3+l3wmmF0zxpM+zaTOm+AUxzWbPKfALbIzNmjPD+RfZ0hVf+/x4XBfOdZ9m2WeffxIWqWnuU9E8j/niQps+c6rHpzB8TjFx4libMjXfvl6ywFYuW2Krln9tC/0uVbh4qhX7uOcvm2Xzl8x02Wmedwts1Vo+NfGY+DZsxoxp7p/n2zKfl6JCK5o9y2ZMmWwLiorc3zm2zPN44eJZnnfF+rSADQ47pWrf5vg+dG+Z70n9YYvjRP/HQyf18I8jx8APaLdHO+7dROPhpM717eEEx/d54VV57vQhx6D/c+CkB/s8VVPGdtDh8pwS8YSdPgxGPd7ve0XOdNnb8xqZUxm+oOQMFuChmVMPnSLFPWaURScQaxEFT3R50s47r7VtzJ2eMGq3ELygXuU+VdRV+kNTVaBHasNzB/XavRztsm/3OseJ/i9CLk5hk8lY/eEaecYe/I7+4ZdrsDKXL3X5ci9rHKvQ6/9q/bkBG/F5J8avoqrEb/28tCMOFcE2fcST82werqvd913lu31b5Lx7qv2hEK2MLMYA3TV7Kz2m0XfNLe8damv46I9TJ48BcfYx1Hl7j2N5VVX8j6XB5/2MYn9w3Xs5bqx0mWr3o9Kq91W69viOIAK63bb77l4ErzTrodcLHkb1oRrA8bdOtozjb6/v4T0Svno8w4sr9wtoEt/SucIQEIi540WvMZCaOpLKnfBdf3iocqxxQ5zC1NW6q9XI+YBdDc4AzBknIhwZ6hQHCA+6/vRRx1Eoc4tOniacIf58G841IC9Cwtl+GG78F8/54+lHRKiMAH3RBxC7DI2YQAcBSvqiz0wiAfaJ83+uOehZu269XdP2WuvSpXs43QF0QdQ4b43Hqc79gJ+TqEr+I2F+sdXUVvne1CfQ+3lYw1c8I+C8lWQcWAjvMwiSj4My/Dqv6w0nXcTLS86ROZXgdKLaXahyx6q8Hh9U3R8vmQNOjgC+MeHkjN69nlR1uQfksHAQ7EB3nT7x+M2CEMeBT4yPxGI8jIu27LgP7m84mfO5Dkd9xM6R/hTh5ey7whM/nKgFnjg/Yc78HxS37nZq/Jmk3Gr8Yq32i7YGn90ffI1fSMaY4gdTgL4qH6PmMnw45iWLBSds8ZfR/AIJecR4mVtfKIkzLeLm/A0/xEQwXWno8EDq+K2+n44cQIsvVWIZ3/ghG/mE9ON0Y7IpLeXn/JUSh9Eb6xGoBzl89DpvhOUH2qKeyCNe9astRLeQgIF14bcU49vV448/wd55573QRnn4dICrJwfok6+U2AEVC5Wyr08P6tF9yAJ6kKOvhhUqsRcgkdEccSzKvHG02OBP/Ish5/I+T4JwMfkFWRN/iInxAyEWzo8s70SwXVXJyr7HV36/kMMK7ncZ94Xv18M43RYvIuvfYjOW3HipB5/ww+U4O9cbz/i6P4470B15V8GLofpPfL2fODFOXubpv0QSYx9txnhGDDF2PejmJVc4xnSfXTJcWuFOk/OLseJX+H14lATFGMIhGL1TgRFIUABfuAoZTE6xBq06egSprOriF9CmL6WrLQwTlASYugBe8ad6QHzRePUXN5HP/ffbYrhdO/BtCd+Mf/RR/B0YAJn6CXbk+3f8kH5soV9/TEGbEh7Z1ye49EkmBfgVd/ppgwLosi9Z8Up3TIT0DyfwhTjiA+OPF4UAeb7fJ9GDfdfBD2uV+p2Kz0tCwnu7vIKXiTHuyOAnvBqb/Al00PlIvvBpiidfSHS3G+4sTg9vUx3RHXg8qYkvOp05YPhPafrdAt0aNzwgNMYcLyT3wfvqkx+dtB3Dr1Xk/KMMK7wCDahTDJSA+NJ2cDA3QbTVR6k6QB1d4vk2TPlFS2Wxx8RSyh/RaQNZX2gL1YYfvaxWcX9LO8aAP8g46KCDw4/HAqhCDn5sYJ+xy0f5R1/qk+wB+9l0pK4+gWQA8YPSo37x0MaeeFnpSXa+YIzbl3h7V9JTMkbZlT6NA+AzZ352pczvBiQ8vy2ENRIptSvELrHAD40L9N6Y5C7t3vs/H5O3QephhVeiOoavKZHnRZ4nOv+tKUrFC5SvskWbVZ0EJ+nRS9KX+1aYRA92c37CC4Y3rekk4Xy6KgGpkOpqi5b2pXTJy0HRANqaMECysh3/hrLhgYa+tC29oPwGpAOQTkrZpk9BC+j7SW738cdD4y99HXLIf4fv9gH2r+iWPV3o1KUTG6AmnHoaV9lqzBfpgV8y8ImfUjpFA6EpRtEGNFZh5o8+f9wN25t4MQulJ8rEWECTD2Hl9UWg0nWFpPTECX9KmdtiiVd+UeKD/ADZs5N4lezn2QV4m8Sscn/C4YLrrWVFTpI+bFsqWeV94cGW0wDFRHXsyGcSPlyM+O66tNrzbBXauQtV46z/DxOD9c46g5A2faoLBWkdSOWykMoLFTDq8kM6tHLAI37o8GjA8FBKv8pUj2QlL34AGg+U/JUPiQKw3bntttvshBNOCv+1EyD8LacHL/DnbMsGpXykLv/El5aSSf2BlrZVpnaESixu54w9/YPqKAdfXNljcnsC+4ofkT+m9odq16O4pT6gCzpJXuVy8fMFt0mC5hIHfnipg9m5qdfnyCpLMvJ9elhxvY92RZVfpK47PBDnEjTMtesKP71OsjOV3p9CagMfgl3oOUQPhwnoVvKHj8d4BvA+sIn2SAQX56VQCGRpomfrKqW8MaBfAUtBcoAmRAmCX/INmiaddmovnQgQkD0lDwgtRU6oPAJeJ4hRjm9L+GPv44470T75+DO317CSAcjJB/mpftmq15+pg/IVEI3+FETnh5RkF0BW8QFZ5ZV4LuV98dRCSc8qz8Wsi1q+AOiXv4ozK29IeJdllSdJteWAl5yRPRDfoKGDOsgeniRPV92wV3d+0biY6n+1zvvD59a+lXEmD7A7t3846vVThnZOl0eunh1fU7v4Bz++m5n9X+H5VkzpsZyzAAAAAElFTkSuQmCC"


def obtener_imagen_firma_rep_legal(width_cm: float = 3.6, height_cm: float = 1.4):
    for p in (BASE_DIR / "firma_segundo_murillo.png", BASE_DIR / "assets" / "firma_segundo_murillo.png"):
        if p.exists():
            return Image(str(p), width=width_cm * cm, height=height_cm * cm, kind="proportional")
    try:
        raw = base64.b64decode(FIRMA_REP_LEGAL_B64)
        return Image(io.BytesIO(raw), width=width_cm * cm, height=height_cm * cm, kind="proportional")
    except Exception:
        return Spacer(1, height_cm * cm)


def obtener_imagen_firma_pesv(width_cm: float = 3.8, height_cm: float = 1.4):
    for p in (BASE_DIR / "firma_marly_alfonso.png", BASE_DIR / "assets" / "firma_marly_alfonso.png"):
        if p.exists():
            return Image(str(p), width=width_cm * cm, height=height_cm * cm, kind="proportional")
    try:
        raw = base64.b64decode(FIRMA_REP_PESV_B64)
        return Image(io.BytesIO(raw), width=width_cm * cm, height=height_cm * cm, kind="proportional")
    except Exception:
        return Spacer(1, height_cm * cm)


SEED_INASISTENCIAS_HISTORICAS: List[List[str]] = [['a33227f4', '1', '05/05/2025', '11:01:06', '', 'Nicolás Vera', '', 'No especificada', 'Jesús Toro', '2', '11000', 'Cancelo Efectivo'], ['1dbfa69f', '2', '07/05/2025', '8:17:36', '', 'Sebastian Garzón', '', 'No especificada', 'Leidy Salamanca', '1', '10000', 'Cancelo Efectivo'], ['e5ca62c3', '3', '08/05/2025', '8:01:13', '', 'Jose Gómez', '', 'No especificada', 'Víctor León', '2', '20000', 'Cancelo Trasferencia'], ['c3e5ebf3', '4', '08/05/2025', '8:06:03', '', 'Yohana Gutiérrez', '', 'No especificada', 'Leidy Salamanca', '1', '10000', 'Cancelo Trasferencia'], ['ebfd7a61', '5', '09/05/2025', '6:40:12', '', 'Laura Ruiz', '', 'No especificada', 'Dairo Serrano', '1', '10000', 'Cancelo Trasferencia'], ['9578d3da', '6', '09/05/2025', '7:40:25', '', 'Yonnier Ariza', '', 'No especificada', 'Rolando Ramírez', '1', '10000', 'Cancelo Efectivo'], ['7297da84', '7', '09/05/2025', '8:23:31', '', 'Jorge Ayala', '', 'No especificada', 'Martin Sánchez', '2', '20000', 'Cancelo Efectivo'], ['b5eec4e7', '8', '10/05/2025', '6:45:20', '', 'Jose Gómez', '', 'No especificada', 'Víctor León', '2', '20000', 'Cancelo Trasferencia'], ['a2a15edb', '9', '13/05/2025', '10:17:43', '', 'Ana María Mora Camargo', '', 'No especificada', 'Rolando Ramírez', '2', '20000', 'Cancelo Trasferencia'], ['702ea743', '10', '13/05/2025', '16:28:38', '', 'José Murillo', '', 'No especificada', 'Jesús Toro', '1', '11500', 'Cancelo Efectivo'], ['f00a89f4', '11', '15/05/2025', '15:26:20', '', 'Wilmar Fontecha', '', 'No especificada', 'Víctor León', '2', '20000', 'Cancelo Efectivo'], ['9bbb1c59', '12', '19/05/2025', '6:17:57', '', 'Jose Gómez', '', 'No especificada', 'Víctor León', '2', '20000', 'Pendiente Pago'], ['ac858201', '13', '23/05/2025', '10:18:02', '', 'Melissa Castrillon', '', 'No especificada', 'Rolando Ramírez', '1', '10000', 'Cancelo Efectivo'], ['9f043e80', '14', '24/05/2025', '6:16:36', '', 'Jesús Guerrero', '', 'No especificada', 'Hermes Becerra', '2', '10000', 'Pendiente Pago'], ['9c9ac072', '15', '24/05/2025', '8:33:25', '', 'Laura Valderrama', '', 'No especificada', 'Jesús Toro', '1', '10000', 'Cancelo Trasferencia'], ['2ccaf9fd', '16', '24/05/2025', '9:39:42', '', 'Jose Hernández', '', 'No especificada', 'Víctor León', '2', '20000', 'Pendiente Pago'], ['4bfe78de', '17', '24/05/2025', '11:18:11', '', 'Diana Roa', '', 'No especificada', 'Víctor León', '2', '20000', 'Pendiente Pago'], ['1bff2ae0', '18', '28/05/2025', '6:18:32', '', 'Karen Moreno', '', 'No especificada', 'Leidy Salamanca', '2', '20000', 'Pendiente Pago'], ['d0741333', '19', '30/05/2025', '8:21:57', '', 'Maria Cuervo', '', 'No especificada', 'Víctor León', '2', '20000', 'Pendiente Pago'], ['dfdf3a8c', '20', '30/05/2025', '8:25:56', '', 'Adela Sánchez', '', 'No especificada', 'Rolando Ramírez', '2', '20000', 'Pendiente Pago'], ['eac20114', '21', '30/05/2025', '14:27:53', '', 'Sonia Gómez', '', 'No especificada', 'Víctor León', '2', '20000', 'Pendiente Pago'], ['2389c510', '22', '03/06/2025', '6:50:41', '', 'Edilson Ardilla', '', 'No especificada', 'Rolando Ramírez', '2', '20000', 'Pendiente Pago'], ['27de2f6e', '23', '03/06/2025', '15:33:04', '', 'Roberto Orduz', '', 'No especificada', 'Martin Sánchez', '2', '20000', 'Pendiente Pago'], ['8ab431eb', '24', '06/06/2025', '7:35:42', '', 'Jorge Cáceres', '', 'No especificada', 'Hermes Becerra', '1', '10000', 'Cancelo Efectivo'], ['c875dd58', '25', '09/06/2025', '13:47:29', '', 'Angie Dueñas', '', 'No especificada', 'Víctor León', '2', '20000', 'Pendiente Pago'], ['db5fbae6', '26', '10/06/2025', '11:39:01', '', 'Juan Larrota', '', 'No especificada', 'Alcibíades Bueno', '2', '20000', 'Pendiente Pago'], ['3a0f7ccc', '27', '10/06/2025', '17:16:19', '', 'José Gómez Gutiérrez', '', 'No especificada', 'Dairo Serrano', '2', '100000', 'Pendiente Pago'], ['c96135c6', '28', '11/06/2025', '6:26:14', '', 'Estefannny Patiño', '', 'No especificada', 'Leidy Salamanca', '1', '10000', 'Pendiente Pago'], ['b54f958f', '29', '11/06/2025', '14:28:27', '', 'Mayra Araque', '', 'No especificada', 'Dairo Serrano', '2', '10000', 'Pendiente Pago'], ['189568d3', '30', '13/06/2025', '11:31:49', '', 'Lina Arias', '', 'No especificada', 'Rolando Ramírez', '2', '20000', 'Pendiente Pago'], ['7a8d2dc7', '31', '14/06/2025', '8:27:06', '', 'Laura Valderrama', '', 'No especificada', 'Jesús Toro', '1', '10000', 'Cancelo Trasferencia'], ['e5054bab', '32', '17/06/2025', '7:51:30', '', 'Ashlee Acuña', '', 'No especificada', 'Jesús Toro', '2', '20000', 'Pendiente Pago'], ['941722eb', '33', '18/06/2025', '6:28:14', '', 'Keiner Castillo', '', 'No especificada', 'Jesús Toro', '2', '20000', 'Pendiente Pago'], ['7b16d6df', '34', '19/06/2025', '8:26:15', '', 'Dana Valentina Ayala', '', 'No especificada', 'Leidy Salamanca', '1', '10000', 'Cancelo Efectivo'], ['f15056e2', '35', '20/06/2025', '16:17:39', '', 'Lina Rojas', '', 'No especificada', 'Dairo Serrano', '2', '10000', 'Pendiente Pago'], ['d4379a8f', '36', '21/06/2025', '9:47:43', '', 'José Sampayo', '', 'No especificada', 'Martin Sánchez', '2', '20000', 'Pendiente Pago'], ['e2db4503', '37', '25/06/2025', '8:18:07', '', 'Jilber Tolado', '', 'No especificada', 'Leidy Salamanca', '2', '20000', 'Pendiente Pago'], ['d35fbc9c', '38', '25/06/2025', '10:38:57', '', 'Jenifer Mosquera', '', 'No especificada', 'Dairo Serrano', '2', '10000', 'Pendiente Pago'], ['5952a73f', '39', '28/06/2025', '7:00:32', '', 'Brayan Grandas', '', 'No especificada', 'Dairo Serrano', '1', '10000', 'Pendiente Pago'], ['8d56b925', '40', '01/07/2025', '6:08:05', '', 'José Sampayo', '', 'No especificada', 'Martin Sánchez', '2', '20000', 'Pendiente Pago'], ['0a8045bd', '41', '01/07/2025', '6:15:17', '', 'Toser Blanco', '', 'No especificada', 'Jesús Toro', '2', '20000', 'Pendiente Pago'], ['86c214f8', '42', '02/07/2025', '8:14:52', '', 'Dana Valentina Ayala', '', 'No especificada', 'Leidy Salamanca', '2', '20000', 'Pendiente Pago'], ['9948ad61', '43', '05/07/2025', '6:24:25', '', 'Neftali Bermudez', '', 'No especificada', 'Martin Sánchez', '2', '20000', 'Pendiente Pago'], ['3387d4c1', '44', '08/07/2025', '16:46:57', '', 'Juan Jaimes', '', 'No especificada', 'Hermes Becerra', '2', '10000', 'Pendiente Pago'], ['9835ab55', '45', '11/07/2025', '6:41:46', '', 'José Sampayo', '', 'No especificada', 'Martin Sánchez', '2', '20000', 'Pendiente Pago'], ['0c7aa53f', '46', '12/07/2025', '6:17:21', '', 'Johanny Toscano', '', 'No especificada', 'Rolando Ramírez', '2', '20000', 'Pendiente Pago'], ['2e112de0', '47', '12/07/2025', '6:18:22', '', 'Santiago Salamanca', '', 'No especificada', 'Leidy Salamanca', '2', '20000', 'Pendiente Pago'], ['062e3baf', '48', '14/07/2025', '5:59:57', '', 'Lina Arias', '', 'No especificada', 'Hermes Becerra', '2', '10000', 'Pendiente Pago'], ['a2022362', '49', '14/07/2025', '7:55:41', '', 'José Sampayo', '', 'No especificada', 'Martin Sánchez', '2', '20000', 'Pendiente Pago'], ['25280abd', '50', '15/07/2025', '6:45:49', '', 'Héctor Suárez', '', 'No especificada', 'Víctor León', '2', '20000', 'Pendiente Pago'], ['9b27724f', '51', '15/07/2025', '7:58:23', '', 'Carlos Arauz', '', 'No especificada', 'Martin Sánchez', '2', '20000', 'Pendiente Pago'], ['6a5f452e', '52', '17/07/2025', '6:18:00', '', 'Lina Arias', '', 'No especificada', 'Jesús Toro', '1', '10000', 'Pendiente Pago'], ['b262809d', '53', '18/07/2025', '14:54:13', '', 'Sara Gómez', '', 'No especificada', 'Leidy Salamanca', '2', '20000', 'Pendiente Pago'], ['66a3ca35', '54', '22/07/2025', '9:11:37', '', 'Anderson Parada', '', 'No especificada', 'Hermes Becerra', '2', '10000', 'Pendiente Pago'], ['c71f118c', '55', '23/07/2025', '16:20:00', '', 'Cristian Mota', '', 'No especificada', 'Dairo Serrano', '1', '10000', 'Pendiente Pago'], ['a48271e6', '56', '29/07/2025', '10:10:19', '', 'Angie Sarmiento', '', 'No especificada', 'Alcibíades Bueno', '2', '23000', 'Pendiente Pago'], ['61e179f8', '57', '01/08/2025', '6:16:11', '', 'Duvian Pamplona', '', 'No especificada', 'Víctor León', '2', '23000', 'Pendiente Pago'], ['15ee3577', '58', '01/08/2025', '8:26:34', '', 'Claudia Ruiz', '', 'No especificada', 'Leidy Salamanca', '2', '50000', 'Pendiente Pago'], ['9b184f20', '59', '04/08/2025', '10:54:27', '', 'Sandra Roya', '', 'No especificada', 'Dairo Serrano', '2', '', 'Cancelo Efectivo'], ['aa7f5cc0', '60', '15/08/2025', '16:09:59', '', 'Jose Gómez', '', 'No especificada', 'Dairo Serrano', '2', '10000', 'Pendiente Pago'], ['b1e70315', '61', '19/08/2025', '8:20:59', '', 'Luz Maldonado', '', 'No especificada', 'Leidy Salamanca', '1', '25000', 'Cancelo Efectivo'], ['f4a35cae', '62', '20/08/2025', '7:41:41', '', 'Esmeralda Ordóñez', '', 'No especificada', 'Rolando Ramírez', '2', '20000', 'Cancelo Trasferencia'], ['15ca9863', '63', '20/08/2025', '16:23:41', '', 'Ginna Daza', '', 'No especificada', 'Dairo Serrano', '2', '10000', 'Pendiente Pago'], ['dbf7a060', '64', '21/08/2025', '16:15:32', '', 'José Gomes', '', 'No especificada', 'Hermes Becerra', '2', '50000', 'Pendiente Pago'], ['7d14c063', '65', '21/08/2025', '16:16:20', '', 'Valeria Mateus', '', 'No especificada', 'Rolando Ramírez', '2', '20000', 'Pendiente Pago'], ['0c360348', '66', '25/08/2025', '8:28:32', '', 'Juan Uretra', '', 'No especificada', 'Víctor León', '2', '23000', 'Pendiente Pago'], ['3294a695', '67', '25/08/2025', '13:25:18', '', 'Jonathan Guerrero', '', 'No especificada', 'Leidy Salamanca', '1 , 2', '25000', 'Pendiente Pago'], ['b08d6e80', '68', '25/08/2025', '16:16:38', '', 'Ginnet Díaz Daza', '', 'No especificada', 'Dairo Serrano', '2', '100000', 'Pendiente Pago'], ['292b00e8', '69', '27/08/2025', '16:29:38', '', 'Karen Santamaría', '', 'No especificada', 'Jesús Toro', '2', '50000', 'Pendiente Pago'], ['2f6e2596', '70', '04/09/2025', '10:28:09', '', 'David Carrillo', '', 'No especificada', 'Leidy Salamanca', '1', '25000', 'Cancelo Efectivo'], ['7d70d5e3', '71', '05/09/2025', '15:41:21', '', 'Wendy Vega', '', 'No especificada', 'Martin Sánchez', '2', '25000', 'Pendiente Pago'], ['fc21d2b0', '72', '06/09/2025', '6:19:05', '', 'Luis Alejandro Becerra', '', 'No especificada', 'Rolando Ramírez', '2', '20000', 'Pendiente Pago'], ['8fb61908', '73', '06/09/2025', '6:22:24', '', 'Daniel Fernández', '', 'No especificada', 'Hermes Becerra', '1', '25000', 'Pendiente Pago'], ['09930cf1', '74', '08/09/2025', '10:23:07', '', 'David Carrillo', '', 'No especificada', 'Jesús Toro', '1', '25000', 'Pendiente Pago'], ['3efd4c17', '75', '08/09/2025', '14:07:13', '', 'Maryi Sanchez', '', 'No especificada', 'Alcibíades Bueno', '2', '23000', 'Pendiente Pago'], ['8c0d222a', '76', '10/09/2025', '6:12:02', '', 'Daniel Fernández', '', 'No especificada', 'Jesús Toro', '2', '50000', 'Pendiente Pago'], ['8f75cff5', '77', '10/09/2025', '10:09:02', '', 'María Amorocho', '', 'No especificada', 'Gabriel Roa', '2', '20000', 'Pendiente Pago'], ['daebd4d2', '78', '15/09/2025', '6:15:24', '', 'Nini Ochoa', '', 'No especificada', 'Jesús Toro', '2', '50000', 'Pendiente Pago'], ['53bed2e4', '79', '15/09/2025', '6:15:19', '', 'Daniel Fernández', '', 'No especificada', 'Hermes Becerra', '2', '50000', 'Pendiente Pago'], ['9fe8d3aa', '80', '15/09/2025', '12:24:33', '', 'Marly Rodríguez', '', 'No especificada', 'Hermes Becerra', '1', '25000', 'Pendiente Pago'], ['f41386db', '81', '16/09/2025', '6:17:49', '', 'Daniel Fernández', '', 'No especificada', 'Hermes Becerra', '2', '50000', 'Pendiente Pago'], ['b89dbc31', '82', '17/09/2025', '13:51:14', '', 'Jesus Contreras', '', 'No especificada', 'Jesús Toro', '1', '25000', 'Pendiente Pago'], ['4f3cce6c', '83', '24/09/2025', '8:14:34', '', 'Eider Salas', '', 'No especificada', 'Martin Sánchez', '2', '50000', 'Pendiente Pago'], ['d983b388', '84', '27/09/2025', '8:24:31', '', 'Juan Navas', '', 'No especificada', 'Gabriel Roa', '2', '50000', 'Pendiente Pago'], ['9f58d0da', '85', '27/09/2025', '12:51:48', '', 'Luis Becerra', '', 'No especificada', 'Leidy Salamanca', '1', '25000', 'Cancelo Trasferencia'], ['e1450199', '86', '29/09/2025', '8:18:42', '', 'Claudia Ruiz', '', 'No especificada', 'Rolando Ramírez', '2', '20000', 'Pendiente Pago'], ['c9637624', '87', '03/10/2025', '7:10:43', '', 'Daniel Fernández', '', 'No especificada', 'Jesús Toro', '2', '50000', 'Pendiente Pago'], ['f7aba590', '88', '04/10/2025', '7:28:48', '', 'Sergio Díaz', '', 'No especificada', 'Víctor León', '2', '50000', 'Pendiente Pago'], ['7fc3c0a6', '89', '15/10/2025', '14:36:12', '', 'Kevin Gómez', '', 'No especificada', 'Víctor León', '2', '23000', 'Pendiente Pago'], ['d212a891', '90', '15/10/2025', '14:40:52', '', 'Laura Marín', '', 'No especificada', 'Víctor León', '2', '23000', 'Pendiente Pago'], ['9cdd77a2', '91', '16/10/2025', '6:13:11', '', 'María Cristina Vargas', '', 'No especificada', 'Leidy Salamanca', '2', '50000', 'Pendiente Pago'], ['7d1b5caf', '92', '24/10/2025', '10:22:25', '', 'Didier Amador', '', 'No especificada', 'Hermes Becerra', '1', '25000', 'Pendiente Pago'], ['e750ce57', '93', '27/10/2025', '6:20:07', '', 'Mariana Cano', '', 'No especificada', 'Leidy Salamanca', '2', '50000', 'Pendiente Pago'], ['f14c12af', '94', '27/10/2025', '12:52:42', '', 'Paula Hernandez', '', 'No especificada', 'Alcibíades Bueno', '2', '11500', 'Pendiente Pago'], ['f7a52aec', '95', '29/10/2025', '8:26:56', '', 'Williexy Gómez', '', 'No especificada', 'Leidy Salamanca', '1', '25000', 'Pendiente Pago'], ['7ba4b8de', '96', '31/10/2025', '20:25:57', '', 'Diego Navas', '', 'No especificada', 'Jesús Toro', '2', '50000', 'Pendiente Pago'], ['89486d63', '97', '10/11/2025', '14:37:53', '', 'Jhoan Carreño', '', 'No especificada', 'Hermes Becerra', '1', '25000', 'Pendiente Pago'], ['adaa85d4', '98', '12/11/2025', '8:15:35', '', 'Angelly Medina', '', 'No especificada', 'Leidy Salamanca', '1', '25000', 'Pendiente Pago'], ['c2f0982b', '99', '13/11/2025', '8:35:21', '', 'Camilo Pacheco', '', 'No especificada', 'Hermes Becerra', '2', '25000', 'Pendiente Pago'], ['013f1513', '100', '18/11/2025', '11:24:32', '', 'Cristian Nuncira', '', 'No especificada', 'Víctor León', '2', '23000', 'Pendiente Pago'], ['a12468a5', '101', '25/11/2025', '6:26:37', '', 'Edisson Villamizar', '', 'No especificada', 'Leidy Salamanca', '2', '50000', 'Pendiente Pago'], ['e0a90243', '102', '27/11/2025', '6:18:41', '', 'Edisson Villamizar', '', 'No especificada', 'Leidy Salamanca', '1', '25000', 'Pendiente Pago'], ['bf87add7', '103', '18/12/2025', '14:18:51', '', 'Angie Cogollo', '', 'No especificada', 'Leidy Salamanca', '1 , 2', '25000', 'Pendiente Pago'], ['0e9f409c', '104', '29/12/2025', '16:20:30', '', 'Héctor Cortez', '', 'No especificada', 'Leidy Salamanca', '1', '25000', 'Pendiente Pago'], ['01371ed5', '105', '19/01/2026', '14:15:44', '', 'María Sepúlveda', '', 'No especificada', 'Hermes Becerra', '2', '25000', 'Pendiente Pago'], ['68057525', '106', '20/01/2026', '10:28:46', '', 'Jordan Valderrama', '', 'No especificada', 'Leidy Salamanca', '1', '25000', 'Pendiente Pago'], ['a630ee5f', '107', '26/01/2026', '8:19:18', '', 'Ana Castillo', '', 'No especificada', 'Leidy Salamanca', '1', '25000', 'Pendiente Pago'], ['50df5f8c', '108', '04/02/2026', '6:10:13', '', 'José Ardilla', '', 'No especificada', 'Hermes Becerra', '2', '25000', 'Pendiente Pago'], ['ccf2b829', '109', '09/02/2026', '10:19:23', '', 'Alejandro Romero', '', 'No especificada', 'Rolando Ramírez', '2', '25000', 'Pendiente Pago'], ['3d6ce47f', '110', '13/03/2026', '12:16:15', '', 'Isaías Romero', '', 'No especificada', 'Hermes Becerra', '2', '25000', 'Pendiente Pago'], ['509a62bf', '111', '16/03/2026', '14:44:07', '', 'Valentina Vargas', '', 'No especificada', 'Martin Sánchez', '2', '25000', 'Pendiente Pago'], ['5bd86467', '112', '17/03/2026', '11:58:52', '', 'Isaías Romero', '', 'No especificada', 'Hermes Becerra', '2', '25000', 'Pendiente Pago'], ['b1931ffb', '113', '17/03/2026', '14:04:32', '', 'Valentina Vargas', '', 'No especificada', 'Martin Sánchez', '2', '25000', 'Pendiente Pago'], ['23c09fc4', '114', '18/03/2026', '7:58:24', '', 'Natalia Novoa', '', 'No especificada', 'Rolando Ramírez', '2', '50000', 'Cancelo Efectivo'], ['9dc4f62f', '115', '21/04/2026', '9:53:55', '', 'Angelica Valencia', '', 'No especificada', 'Martin Sánchez', '2', '25000', 'Pendiente Pago'], ['fbbb0ce9', '116', '28/04/2026', '7:43:48', '', 'Anderson García', '', 'No especificada', 'Rolando Ramírez', '2', '25000', 'Pendiente Pago'], ['74cd1472', '117', '20/05/2026', '16:06:42', '', 'José Cogollo', '', 'No especificada', 'Hermes Becerra', '2', '50000', 'Cancelo Efectivo'], ['024a792f', '118', '25/05/2026', '14:25:38', '', 'Danna Pabón', '', 'No especificada', 'Hermes Becerra', '2', '50000', 'Pendiente Pago'], ['2ee53c43', '119', '17/06/2026', '8:12:47', '', 'Carolina Calderón', '', 'No especificada', 'Rolando Ramírez', '2', '25000', 'Pendiente Pago'], ['d9e3d916', '120', '26/06/2026', '7:33:13', '', 'Karen Garcia', '', 'No especificada', 'Oscar Eduardo Mantilla Garcia', '1', '25000', 'Pendiente Pago'], ['3d3c3335', '121', '30/06/2026', '7:12:20', '', 'Kevin Cruz', '', 'No especificada', 'Rolando Antonio Ramirez Giraldo', '2', '25000', 'Pendiente Pago'], ['2234455c', '122', '30/06/2026', '11:05:41', '', 'Daniel Forero', '', 'No especificada', 'Oscar Eduardo Mantilla Garcia', '2', '50000', 'Pendiente Pago'], ['c2fef868', '123', '30/06/2026', '14:17:50', '', 'Dayana Ramos', '', 'No especificada', 'Rolando Antonio Ramirez Giraldo', '2', '25000', 'Pendiente Pago'], ['08b1aeba', '124', '01/07/2026', '10:45:28', '', 'Daniel Forero', '', 'No especificada', 'Oscar Eduardo Mantilla Garcia', '1', '25000', 'Pendiente Pago'], ['c9b7cb9f', '125', '02/07/2026', '7:06:18', '', 'Eduar Guerrero', '', 'No especificada', 'Victor Eduardo Leon Anaya', '2', '26000', 'Pendiente Pago'], ['d4a6d658', '126', '02/07/2026', '7:48:14', '', 'Karen Guío', '', 'No especificada', 'Hermes Giovanny Becerra Montañez', '2', '25000', 'Pendiente Pago'], ['4cc6f7bb', '127', '07/07/2026', '9:57:29', '', 'Ingrid Reyes', '', 'No especificada', 'Rolando Antonio Ramirez Giraldo', '2', '25000', 'Cancelo Trasferencia'], ['d95ad604', '128', '28/07/2026', '11:48:25', '', 'Jose Gabriel Hernández', '', 'No especificada', 'Rolando Antonio Ramirez Giraldo', '2', '25000', 'Cancelo Efectivo'], ['baf7083f', '129', '04/08/2026', '10:24:08', '', 'Omar Rueda', '', 'No especificada', 'Hermes Giovanny Becerra Montañez', '1', '25000', 'Pendiente Pago'], ['461dd9f1', '130', '10/08/2026', '10:32:46', '', 'Yois Cavanzo', '', 'No especificada', 'Victor Eduardo Leon Anaya', '2', '26000', 'Pendiente Pago'], ['82e1e698', '131', '10/08/2026', '14:21:57', '', 'Juan Osorio', '', 'No especificada', 'Oscar Eduardo Mantilla Garcia', '1', '25000', 'Pendiente Pago'], ['0f0db7fe', '132', '12/08/2026', '12:20:25', '', 'Juan Osorio', '', 'No especificada', 'Oscar Eduardo Mantilla Garcia', '1', '25000', 'Pendiente Pago'], ['c54e7d18', '133', '12/08/2026', '14:26:52', '', 'Viviana Fermin', '', 'No especificada', 'Oscar Eduardo Mantilla Garcia', '1', '25000', 'Pendiente Pago'], ['01940ec7', '134', '20/08/2026', '15:29:18', '', 'Nicolas Almeida', '', 'No especificada', 'Martin Fernando Sanchez Parra', '2', '25000', 'Pendiente Pago'], ['40122d2d', '134', '19/08/2026', '14:47:38', '', 'Johan Casique', '', 'No especificada', 'Oscar Eduardo Mantilla Garcia', '1', '25000', 'Pendiente Pago'], ['c1115bb6', '135', '22/08/2026', '8:20:19', '', 'Jhoann Jaimes', '', 'No especificada', 'Oscar Eduardo Mantilla Garcia', '1', '25000', 'Pendiente Pago'], ['3517d17a', '136', '27/08/2026', '9:23:26', '', 'Sebastian Arango', '', 'No especificada', 'Martin Fernando Sanchez Parra', '2', '25000', 'Pendiente Pago'], ['55c75a7b', '137', '29/08/2026', '10:15:43', '', 'Juan Krispin', '', 'No especificada', 'Hermes Giovanny Becerra Montañez', '2', '25000', 'Pendiente Pago'], ['beb17316', '138', '31/08/2026', '6:56:58', '', 'Michaell Zuñiga', '', 'No especificada', 'Oscar Eduardo Mantilla Garcia', '2', '50000', 'Pendiente Pago'], ['b245200a', '139', '05/09/2026', '8:17:08', '', 'Diego Rincón', '', 'No especificada', 'Martin Fernando Sanchez Parra', '2', '25000', 'Pendiente Pago'], ['33df5226', '140', '10/09/2026', '6:18:44', '', 'Laura Pabón', '', 'No especificada', 'Victor Eduardo Leon Anaya', '2', '26000', 'Pendiente Pago'], ['742fbf0c', '141', '12/09/2026', '16:12:25', '', 'Jeison Lopez', '', 'No especificada', 'Rolando Antonio Ramirez Giraldo', '2', '25000', 'Pendiente Pago'], ['93ce0f11', '142', '14/09/2026', '7:04:09', '', 'Luis Granados', '', 'No especificada', 'Martin Fernando Sanchez Parra', '2', '25000', 'Pendiente Pago'], ['cd0f01db', '143', '14/09/2026', '15:39:11', '', 'Miguel Arias', '', 'No especificada', 'Martin Fernando Sanchez Parra', '1', '25000', 'Pendiente Pago'], ['2be1d034', '144', '19/09/2026', '7:09:05', '', 'Luis Granados', '', 'No especificada', 'Martin Fernando Sanchez Parra', '1', '25000', 'Pendiente Pago'], ['2996abc5', '145', '19/09/2026', '10:44:20', '', 'Maidelith Tettay', '', 'No especificada', 'Rolando Antonio Ramirez Giraldo', '2', '25000', 'Pendiente Pago'], ['30f681a2', '146', '19/09/2026', '11:20:45', '', 'Miguel Arias', '', 'No especificada', 'Martin Fernando Sanchez Parra', '2', '25000', 'Pendiente Pago'], ['ad2af1c7', '147', '10/10/2026', '6:34:15', '', 'Jhon Fonseca', '', 'No especificada', 'Martin Fernando Sanchez Parra', '2', '25000', 'Pendiente Pago'], ['4361106c', '148', '15/10/2026', '10:56:50', '', 'Nelson Meló', '', 'No especificada', 'Oscar Eduardo Mantilla Garcia', '1', '25000', 'Pendiente Pago'], ['1fd57799', '149', '16/10/2026', '8:21:46', '', 'Juan Herrera', '', 'No especificada', 'Hermes Giovanny Becerra Montañez', '2', '25000', 'Pendiente Pago'], ['f6688e48', '150', '26/10/2026', '15:25:31', '', 'Kevin Remolina', '', 'No especificada', 'Martin Fernando Sanchez Parra', '2', '25000', 'Pendiente Pago'], ['22319d3c', '151', '27/10/2026', '9:46:12', '', 'Juan Francisco Navas', '', 'No especificada', 'Rolando Antonio Ramirez Giraldo', '1', '25000', 'Cancelo Efectivo'], ['e648f8b7', '152', '03/11/2026', '12:42:28', '', 'Keiner Ospina', '', 'No especificada', 'Hermes Giovanny Becerra Montañez', '2', '25000', 'Pendiente Pago']]

SEED_MANTENIMIENTOS_HISTORICOS: List[List[str]] = [['2025-07-15 08:00:00', 'Julio', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'UDT984', 'Preventivo', 'La estrella', '0', 'Mantenimiento', 'Cambio de rodillo izquierdo y tensión de frenos', 'MECANICO', 'Cambio de balinera', '0', '0', '0'], ['2025-07-15 08:00:00', 'Julio', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE31F', 'Preventivo', 'Motollantas Richard', '13721566', 'Mantenimiento', 'Motollantas Richard', 'MECANICO', 'Cambio de aceite', '0', '0', '0'], ['2025-07-15 08:00:00', 'Julio', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE30F', 'Preventivo', 'Montallantas Richard', '13721566', 'Mantenimiento', 'Montallantas Richard', 'MECANICO', 'Cambio de aceite y tensión de cadena', '0', '0', '0'], ['2025-07-15 08:00:00', 'Julio', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'UDT984', 'Preventivo', 'Italo colombiano', '0', 'Mantenimiento', 'Italo colombiano', 'MECANICO', 'Cambio Batería', '0', '0', '0'], ['2025-07-15 08:00:00', 'Julio', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE31F', 'Preventivo', 'Moto llantas Richard', '13721566', 'Mantenimiento', 'Moto llantas Richard', 'MECANICO', 'Se realizo cambio de llanta, arreglo de venti, cambio de direccionales, cambió de valineras traseras', '0', '0', '0'], ['2025-07-15 08:00:00', 'Julio', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE30F', 'Preventivo', 'Montallantas Richard', '13721566', 'Mantenimiento', 'Montallantas Richard', 'MECANICO', 'Tensión de cadena calibración de llantas', '0', '0', '0'], ['2025-08-15 08:00:00', 'Agosto', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE31F', 'Preventivo', 'Motollantas Richard', '', 'Mantenimiento', 'Motollantas Richard', 'MECANICO', 'Cambio de aceite, y cambio de retenedor del motor de arranque, más tensión de cadena', '0', '0', '0'], ['2025-08-15 08:00:00', 'Agosto', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'AZA25G', 'Preventivo', 'Richard', '901517095', 'Mantenimiento', 'Richard', 'MECANICO', 'Cambio de aceite', '0', '0', '0'], ['2025-08-15 08:00:00', 'Agosto', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE30F', 'Preventivo', 'Montallantas Richard', '37558489', 'Mantenimiento', 'Montallantas Richard', 'MECANICO', 'Cambio de aceite y cambio de relación', '0', '0', '0'], ['2025-08-15 08:00:00', 'Agosto', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'FSP335', 'Preventivo', 'Taller Angarita', '91278337', 'Mantenimiento', 'Taller Angarita', 'MECANICO', 'Latoneria y pintura', '0', '0', '0'], ['2025-08-15 08:00:00', 'Agosto', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'DUX342', 'Preventivo', 'Ricel', '', 'Mantenimiento', 'Ricel', 'MECANICO', 'Cambio de aceite y todos los filtros', '0', '0', '0'], ['2025-08-15 08:00:00', 'Agosto', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'FSP335', 'Preventivo', 'Euromotor', '901492074', 'Mantenimiento', 'Euromotor', 'MECANICO', 'Cambio de instalación de alta', '0', '0', '0'], ['2025-08-15 08:00:00', 'Agosto', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE30F', 'Preventivo', 'Moto llantas Richard', '', 'Mantenimiento', 'Moto llantas Richard', 'MECANICO', 'Tensión de cadena', '0', '0', '0'], ['2025-08-15 08:00:00', 'Agosto', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'AZA25G', 'Preventivo', 'Richard', '901517095', 'Mantenimiento', 'Richard', 'MECANICO', 'Cambio de guaya del embrague y tensión de cadena', '0', '0', '0'], ['2025-08-15 08:00:00', 'Agosto', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE30F', 'Preventivo', 'Cambio de guaya de embrague', '13721566', 'Mantenimiento', 'Cambio de guaya de embrague', 'MECANICO', 'Cambio de guaya de embrague', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE31F', 'Preventivo', 'CAMBIO DE ACEITE', '', 'Mantenimiento', 'Motollantas Richard', 'MECANICO', 'Cambio de aceite', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'FSP335', 'Preventivo', 'CAMBIO DE ACEITE', '91251739', 'Mantenimiento', 'Lubricentro Risel', 'MECANICO', 'Cambio Aceite y filtros', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE30F', 'Preventivo', 'CAMBIO DE ACEITE', '13721566', 'Mantenimiento', 'Moto llantas Richard', 'MECANICO', 'Cambio de aceite y tensión de cadena', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'AZA25G', 'Preventivo', 'CAMBIO DE ACEITE', '901517095', 'Mantenimiento', 'RICHARD', 'MECANICO', 'Cambio de aceite', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'FSP335', 'Preventivo', 'CAMBIO DE LLANTAS', '', 'Mantenimiento', 'Fayco', 'MECANICO', 'Cambio de llantas', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'UDT984', 'Preventivo', 'Cambio deposito del agua', '901517095', 'Mantenimiento', 'Euromotor', 'MECANICO', 'Cambio soporte de moto y y caja', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'UDT984', 'Preventivo', 'Cambio Bomba agua', '901517095', 'Mantenimiento', 'Euromotor', 'MECANICO', 'Cambio soporte de moto y y caja', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'YNE31F', 'Preventivo', 'CAMBIO DE ACEITE', '', 'Mantenimiento', 'Motollantas Richard', 'MECANICO', 'Cambio de aceite', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'FSP335', 'Preventivo', 'BALANCEO', '91281433', 'Mantenimiento', 'Taller Oscar Rios', 'MECANICO', 'Cambio de soportes de amortiguador. Copelas', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'DUX342', 'Preventivo', 'CAMBIO DE ACEITE', '', 'Mantenimiento', 'Ricel', 'MECANICO', 'Cambio de aceite sin filtro de aire acondicionado', '0', '0', '0'], ['2025-09-15 08:00:00', 'Septiembre', 'autoorientecrm@gmail.com', 'CEA Auto Oriente', 'AZA25G', 'Preventivo', 'MANTENIMINTO FRENOS', '901517095', 'Mantenimiento', 'RICHARD', 'MECANICO', 'Cambio de pastillas traseras', '0', '0', '0']]



def _normalizar_texto(texto: str) -> str:
    if not texto:
        return ""
    texto_nfd = unicodedata.normalize("NFD", str(texto))
    sin_tildes = "".join(c for c in texto_nfd if unicodedata.category(c) != "Mn")
    return " ".join(sin_tildes.lower().strip().split())


class _CeldaDummy:
    def __init__(self, value: str = ""):
        self.value = value


class _MemoriaWorksheet:
    """Hoja en memoria cuando Google Sheets no está conectado aún o no hay red disponible."""

    def __init__(self, title: str, encabezados: List[str]):
        self.title = title
        self._rows: List[List[str]] = [list(encabezados)]
        self.row_count: int = 1000

    def row_values(self, idx: int) -> List[str]:
        if 1 <= idx <= len(self._rows):
            return list(self._rows[idx - 1])
        return []

    def get_all_values(self) -> List[List[str]]:
        return [list(r) for r in self._rows]

    def append_row(self, values: List[Any], **kwargs: Any) -> None:
        self._rows.append([str(v) for v in values])

    def add_rows(self, count: int) -> None:
        self.row_count += max(1, count)

    def update_cell(self, row: int, col: int, value: Any) -> None:
        while len(self._rows) < row:
            self._rows.append([])
        fila = self._rows[row - 1]
        while len(fila) < col:
            fila.append("")
        fila[col - 1] = str(value)

    def cell(self, row: int, col: int) -> _CeldaDummy:
        if 1 <= row <= len(self._rows):
            fila = self._rows[row - 1]
            if 1 <= col <= len(fila):
                return _CeldaDummy(str(fila[col - 1]))
        return _CeldaDummy("")

    def delete_rows(self, start_index: int, end_index: Optional[int] = None) -> None:
        fin = end_index if end_index is not None else start_index
        for idx in range(fin, start_index - 1, -1):
            if 1 <= idx <= len(self._rows):
                self._rows.pop(idx - 1)

    def update(self, values: List[List[Any]] = None, range_name: str = "A2", **kwargs: Any) -> None:
        if not values:
            return
        m = re.match(r"([A-Za-z]+)(\d+)", str(range_name or "A2").strip())
        if m:
            col_letters = m.group(1).upper()
            start_col = 0
            for ch in col_letters:
                start_col = start_col * 26 + (ord(ch) - ord("A") + 1)
            start_row = int(m.group(2))
        else:
            start_col = 1
            m_num = re.search(r"(\d+)", str(range_name or "A2"))
            start_row = int(m_num.group(1)) if m_num else 2
        for offset, row_vals in enumerate(values):
            r_idx = start_row + offset
            while len(self._rows) < r_idx:
                self._rows.append([])
            fila_dest = self._rows[r_idx - 1]
            for c_off, val in enumerate(row_vals):
                c_idx = start_col + c_off
                while len(fila_dest) < c_idx:
                    fila_dest.append("")
                fila_dest[c_idx - 1] = str(val)


def _a1_col_to_index(col_letters: str) -> int:
    idx = 0
    for ch in col_letters.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx


def _parsear_inicio_rango_a1(range_name: str) -> Optional[Tuple[int, int]]:
    """Extrae (fila_1based, col_1based) iniciales desde un rango A1 como 'F12:G12' o 'Hoja!A5:K5'."""
    if not range_name or not isinstance(range_name, str):
        return None
    rango = range_name.split("!")[-1].strip()
    inicio = rango.split(":")[0].strip().replace("$", "")
    m = re.match(r"^([A-Za-z]+)(\d+)$", inicio)
    if not m:
        return None
    return int(m.group(2)), _a1_col_to_index(m.group(1))


class _CachedWorksheet:
    """Wrapper thread-safe con caché de 120 segundos, parcheo en RAM y respaldo anti-429 de Google Sheets."""

    def __init__(self, ws: Any, cache_dict: Dict[str, Tuple[float, List[List[str]]]], lock: Optional[threading.RLock] = None):
        self._ws = ws
        self._cache = cache_dict
        self._lock = lock or threading.RLock()
        self.title = getattr(ws, "title", "Sheet")

    @property
    def row_count(self) -> int:
        return getattr(self._ws, "row_count", 1000)

    def _invalidate(self) -> None:
        with self._lock:
            self._cache.pop(self.title, None)

    def get_all_values(self) -> List[List[str]]:
        now = time.time()
        with self._lock:
            cached = self._cache.get(self.title)
            if cached and (now - cached[0] < 120.0):
                return [list(r) for r in cached[1]]
        last_err = None
        for intento in range(3):
            try:
                vals = self._ws.get_all_values()
                with self._lock:
                    self._cache[self.title] = (time.time(), vals)
                return [list(r) for r in vals]
            except Exception as exc:
                last_err = exc
                if intento < 2:
                    time.sleep(0.9 * (intento + 1))
        with self._lock:
            cached = self._cache.get(self.title)
            if cached:
                logger.warning("Usando caché de respaldo para hoja '%s' tras límite/error temporal de Google API: %s", self.title, last_err)
                return [list(r) for r in cached[1]]
        if last_err:
            raise last_err
        return []

    def row_values(self, idx: int) -> List[str]:
        vals = self.get_all_values()
        if 1 <= idx <= len(vals):
            return list(vals[idx - 1])
        return []

    def append_row(self, values: List[Any], **kwargs: Any) -> Any:
        res = self._ws.append_row(values, **kwargs)
        with self._lock:
            cached = self._cache.get(self.title)
            if cached:
                vals = [list(r) for r in cached[1]]
                vals.append(["" if v is None else str(v) for v in values])
                self._cache[self.title] = (time.time(), vals)
            else:
                self._cache.pop(self.title, None)
        return res

    def add_rows(self, count: int) -> Any:
        res = self._ws.add_rows(count)
        return res

    def update_cell(self, row: int, col: int, value: Any) -> Any:
        res = self._ws.update_cell(row, col, value)
        with self._lock:
            cached = self._cache.get(self.title)
            if cached:
                vals = [list(r) for r in cached[1]]
                while len(vals) < row:
                    vals.append([])
                while len(vals[row - 1]) < col:
                    vals[row - 1].append("")
                vals[row - 1][col - 1] = "" if value is None else str(value)
                self._cache[self.title] = (time.time(), vals)
            else:
                self._cache.pop(self.title, None)
        return res

    def cell(self, row: int, col: int) -> Any:
        vals = self.get_all_values()
        if 1 <= row <= len(vals):
            fila = vals[row - 1]
            if 1 <= col <= len(fila):
                return _CeldaDummy(str(fila[col - 1]))
        return self._ws.cell(row, col)

    def delete_rows(self, start_index: int, end_index: Optional[int] = None) -> Any:
        self._invalidate()
        if end_index is not None:
            return self._ws.delete_rows(start_index, end_index)
        return self._ws.delete_rows(start_index)

    def update(self, *args: Any, **kwargs: Any) -> Any:
        res = self._ws.update(*args, **kwargs)
        # Parchear el rango en caché RAM para evitar re-descargar toda la hoja tras cada edición
        range_name = kwargs.get("range_name")
        values = kwargs.get("values")
        if range_name is None and values is None and len(args) >= 2:
            if isinstance(args[0], str) and isinstance(args[1], list):
                range_name, values = args[0], args[1]
            elif isinstance(args[0], list) and isinstance(args[1], str):
                values, range_name = args[0], args[1]
        elif values is None and len(args) >= 1 and isinstance(args[0], list):
            values = args[0]
        elif range_name is None and len(args) >= 1 and isinstance(args[0], str):
            range_name = args[0]

        inicio = _parsear_inicio_rango_a1(range_name) if isinstance(range_name, str) else None
        if inicio and isinstance(values, list) and all(isinstance(r, (list, tuple)) for r in values):
            row0, col0 = inicio
            with self._lock:
                cached = self._cache.get(self.title)
                if cached:
                    vals = [list(r) for r in cached[1]]
                    for r_off, row_data in enumerate(values):
                        r_idx = row0 + r_off
                        while len(vals) < r_idx:
                            vals.append([])
                        need_cols = col0 + len(row_data) - 1
                        while len(vals[r_idx - 1]) < need_cols:
                            vals[r_idx - 1].append("")
                        for c_off, cell_val in enumerate(row_data):
                            vals[r_idx - 1][col0 - 1 + c_off] = "" if cell_val is None else str(cell_val)
                    self._cache[self.title] = (time.time(), vals)
                    return res

        self._invalidate()
        return res


class SheetsService:
    def __init__(self):
        self._client: Optional[gspread.Client] = None
        self._spreadsheet: Optional[gspread.Spreadsheet] = None
        self._cache_instructores: Dict[int, Dict[str, str]] = {}
        self._cache_lecturas: Dict[str, Tuple[float, List[List[str]]]] = {}
        self._worksheets_cache: Dict[str, _CachedWorksheet] = {}
        self._memoria_hojas: Dict[str, _MemoriaWorksheet] = {}
        self._ultimo_fallo_ts: float = 0.0
        self._ultimo_error_msg: str = ""
        self._sembrado_ejecutado: bool = False
        self._cache_lock = threading.RLock()

    @property
    def conectado_en_vivo(self) -> bool:
        if self._spreadsheet is None:
            try:
                self.conectar()
            except Exception:
                pass
        return self._spreadsheet is not None

    def configurar_credenciales_en_vivo(self, raw_json: str) -> Dict[str, Any]:
        """Permite subir o pegar credentials.json directamente desde el Dashboard Web."""
        try:
            info = json.loads(raw_json)
            if "client_email" not in info or "private_key" not in info:
                return {
                    "ok": False,
                    "message": "El archivo JSON no contiene 'client_email' ni 'private_key' de Service Account.",
                }
            os.environ["GOOGLE_CREDENTIALS_JSON_RUNTIME"] = raw_json
            try:
                Path("/tmp/cea_google_credentials.json").write_text(raw_json, encoding="utf-8")
            except Exception:
                pass
            self._spreadsheet = None
            self._client = None
            self._ultimo_fallo_ts = 0.0
            self._cache_lecturas.clear()
            self._worksheets_cache.clear()
            self._sembrado_ejecutado = False
            self.conectar()
            return {
                "ok": True,
                "message": f"Conectado exitosamente en vivo a Google Sheets ({info.get('client_email')}).",
            }
        except Exception as exc:
            self._ultimo_error_msg = str(exc)
            return {
                "ok": False,
                "message": f"Error conectando a Google Sheets: {exc}",
            }

    def conectar(self) -> gspread.Spreadsheet:
        if self._spreadsheet is None:
            ahora = time.time()
            if self._ultimo_fallo_ts > 0 and (ahora - self._ultimo_fallo_ts < 30.0):
                raise ConnectionError(self._ultimo_error_msg or "Reintento en espera tras fallo de conexión.")
            try:
                creds = obtener_credenciales_google(SCOPES)
                self._client = gspread.authorize(creds)
                self._spreadsheet = self._client.open_by_key(GOOGLE_SHEET_ID)
                self._ultimo_fallo_ts = 0.0
                self._ultimo_error_msg = ""
            except Exception as exc:
                self._ultimo_fallo_ts = ahora
                self._ultimo_error_msg = str(exc)
                raise
        return self._spreadsheet

    def obtener_o_crear_hoja(self, nombre_hoja: str, encabezados: List[str]) -> Any:
        if nombre_hoja in self._worksheets_cache and self._spreadsheet is not None:
            return self._worksheets_cache[nombre_hoja]
        try:
            book = self.conectar()
            try:
                ws_raw = book.worksheet(nombre_hoja)
                ws = _CachedWorksheet(ws_raw, self._cache_lecturas, lock=self._cache_lock)
                primera_fila = ws.row_values(1)
                if not primera_fila:
                    ws.append_row(encabezados)
                elif len(primera_fila) < len(encabezados):
                    for col_idx, titulo in enumerate(encabezados, start=1):
                        if col_idx > len(primera_fila):
                            ws.update_cell(1, col_idx, titulo)
                self._worksheets_cache[nombre_hoja] = ws
                return ws
            except gspread.WorksheetNotFound:
                ws_raw = book.add_worksheet(title=nombre_hoja, rows=1000, cols=len(encabezados) + 2)
                ws_raw.append_row(encabezados)
                ws = _CachedWorksheet(ws_raw, self._cache_lecturas, lock=self._cache_lock)
                self._worksheets_cache[nombre_hoja] = ws
                return ws
        except Exception as exc:
            if nombre_hoja in self._worksheets_cache:
                logger.warning("Manteniendo hoja cacheada '%s' ante límite/error temporal de Google API: %s", nombre_hoja, exc)
                return self._worksheets_cache[nombre_hoja]
            if nombre_hoja not in self._memoria_hojas:
                logger.info("Usando hoja en memoria para '%s' (Google Sheets offline/pendiente): %s", nombre_hoja, exc)
                self._memoria_hojas[nombre_hoja] = _MemoriaWorksheet(nombre_hoja, encabezados)
            return self._memoria_hojas[nombre_hoja]

    def _agregar_fila_al_final(self, ws: gspread.Worksheet, valores_fila: List[Any]) -> int:
        """
        Inserta siempre en la primera fila completamente libre al final de la hoja,
        evitando que las Tablas de Google Sheets (Tabla_1, Tabla_2, Tabla_3) sobrescriban la fila 2.
        """
        todas_las_filas = ws.get_all_values()
        ultima_fila_ocupada = 1
        for idx, fila in enumerate(todas_las_filas, start=1):
            if any(str(celda).strip() for celda in fila):
                ultima_fila_ocupada = idx

        siguiente_fila = max(2, ultima_fila_ocupada + 1)
        if siguiente_fila > ws.row_count:
            ws.add_rows(100)

        ws.update(
            values=[valores_fila],
            range_name=f"A{siguiente_fila}",
            value_input_option="USER_ENTERED",
        )
        return siguiente_fila

    def _leer_filas_diccionario(
        self, ws: gspread.Worksheet, encabezados_esperados: List[str]
    ) -> List[Dict[str, str]]:
        """
        Lee todas las filas usando get_all_values() para que funcione incluso si hay columnas
        fuera de una Tabla de Google Sheets (ej. Documento_Instructor en columna D).
        """
        valores = ws.get_all_values()
        if not valores or len(valores) < 2:
            return []

        cabecera_real = [str(c).strip() for c in valores[0]]
        resultado: List[Dict[str, str]] = []
        num_cols = max(len(cabecera_real), len(encabezados_esperados))

        for fila in valores[1:]:
            if not any(str(c).strip() for c in fila):
                continue
            item: Dict[str, str] = {}
            for i in range(num_cols):
                val = str(fila[i]).strip() if i < len(fila) else ""
                if i < len(cabecera_real) and cabecera_real[i]:
                    item[cabecera_real[i]] = val
                if i < len(encabezados_esperados):
                    item[encabezados_esperados[i]] = val
            resultado.append(item)
        return resultado

    def limpiar_nombres_contaminados(self) -> int:
        """
        Revisa las hojas 'Alumnos' y 'Registro_Clases' y limpia automáticamente
        cualquier nombre que haya quedado con palabras de clase o temario pegadas
        (ej. 'Juan Figueroa Clase Exámen Práctico' -> 'Juan Figueroa').
        """
        corregidos = 0
        try:
            ws_alu = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            filas_alu = ws_alu.get_all_values()
            for idx_fila in range(2, len(filas_alu) + 1):
                fila = filas_alu[idx_fila - 1]
                if len(fila) >= 3 and str(fila[2]).strip():
                    original = str(fila[2]).strip()
                    limpio = limpiar_solo_nombre_persona(original)
                    if limpio and _normalizar_texto(limpio) == "martin sanchez juan herrera":
                        limpio = "Juan Herrera"
                    if limpio and limpio != original:
                        ws_alu.update_cell(idx_fila, 3, limpio)
                        corregidos += 1

            ws_reg = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
            filas_reg = ws_reg.get_all_values()
            for idx_fila in range(2, len(filas_reg) + 1):
                fila = filas_reg[idx_fila - 1]
                if len(fila) >= 6 and str(fila[5]).strip():
                    original = str(fila[5]).strip()
                    limpio = limpiar_solo_nombre_persona(original)
                    if limpio and limpio != original:
                        ws_reg.update_cell(idx_fila, 6, limpio)
                        corregidos += 1

            # Corregir categoría en Registro_Clases si el tema registrado pertenece inequívocamente a Moto (A2)
            # pero quedó guardado por error como carro (B1/C1), y eliminar duplicados exactos por (alumno, categoría, clase)
            filas_reg_act = ws_reg.get_all_values()
            clases_vistas_por_alumno: set = set()
            filas_a_eliminar: List[int] = []
            frases_exclusivas_moto_a2 = (
                "postura en el manillar",
                "maniobra cero cale",
                "franja de anchura limitada",
                "control de la motocicleta",
                "desplazamientos en ocho",
                "con y sin acompañante",
                "pruebas cale",
                "conducción de motocicleta",
            )
            for idx_fila in range(2, len(filas_reg_act) + 1):
                fila = filas_reg_act[idx_fila - 1]
                if not any(str(c).strip() for c in fila):
                    continue
                id_alu = _normalizar_texto(str(fila[4]).strip()) if len(fila) > 4 else ""
                nom_alu = _normalizar_texto(str(fila[5]).strip()) if len(fila) > 5 else ""
                cat_alu = normalizar_categoria(str(fila[6]).strip()) if len(fila) > 6 else "B1"
                clases_rep = str(fila[7]).strip() if len(fila) > 7 else ""
                tema_rep_low = (str(fila[8]).strip().lower() if len(fila) > 8 else "")
                if cat_alu in ("B1", "C1") and any(fr in tema_rep_low for fr in frases_exclusivas_moto_a2):
                    try:
                        ws_reg.update_cell(idx_fila, 7, "A2")
                        cat_alu = "A2"
                        corregidos += 1
                    except Exception:
                        pass
                es_ref_fila = "refuerzo" in clases_rep.lower() or "adicional" in clases_rep.lower()
                nums_cl = tuple(sorted(extraer_numeros_clase(clases_rep)))
                clave_alu = id_alu or nom_alu
                if clave_alu and nums_cl and not es_ref_fila:
                    clave_unica = (clave_alu, cat_alu, nums_cl)
                    if clave_unica in clases_vistas_por_alumno:
                        filas_a_eliminar.append(idx_fila)
                    else:
                        clases_vistas_por_alumno.add(clave_unica)

            for idx_del in reversed(filas_a_eliminar):
                try:
                    ws_reg.delete_rows(idx_del)
                    corregidos += 1
                except Exception as exc_del:
                    logger.warning("No se pudo eliminar fila duplicada %s: %s", idx_del, exc_del)

            # Asegurar que el alumno Juan Manuel Melo Hortua (B1 - 1095920107) esté en el listado de Alumnos
            self.registrar_o_actualizar_alumno(
                documento="1095920107",
                nombre_completo="Juan Manuel Melo Hortua",
                categoria="B1",
            )
            # Asegurar que Jeridson Gonzalez (1098826670) tenga sus dos matrículas activas (A2 Moto y C1 Carro)
            self.registrar_o_actualizar_alumno(
                documento="1098826670",
                nombre_completo="Jeridson Gonzalez",
                categoria="A2",
            )
            self.registrar_o_actualizar_alumno(
                documento="1098826670",
                nombre_completo="Jeridson Gonzalez",
                categoria="C1",
            )

            # Sincronizar Estado = 'Finalizado' / 'En curso' separando estrictamente por categoría
            self.sincronizar_estados_alumnos_finalizados()
            self.sembrar_datos_historicos_appsheet()
        except Exception as exc:
            logger.warning("No se pudo ejecutar limpieza automática de nombres: %s", exc)
        return corregidos

    def _calcular_mapa_clases_por_alumno(
        self,
    ) -> Tuple[Dict[Tuple[str, str], set], Dict[Tuple[str, str], set]]:
        """
        Recorre 'Registro_Clases' y devuelve (clases_por_doc_cat, clases_por_nombre_cat)
        con el conjunto de números de clase únicos ya dictados a cada alumno POR CATEGORÍA.
        Así, cuando un mismo alumno cursa simultáneamente Moto (A2) y Carro (B1/C1),
        el avance y bloqueo de clases de Moto es 100% independiente del de Carro.
        """
        por_doc: Dict[Tuple[str, str], set] = {}
        por_nom: Dict[Tuple[str, str], set] = {}
        try:
            ws_reg = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
            filas_reg = self._leer_filas_diccionario(ws_reg, HEADERS_REGISTRO)
            for r in filas_reg:
                if _es_registro_inasistencia(r):
                    continue
                doc_r = _normalizar_texto(str(r.get("ID_Alumno", "")))
                nom_r = _normalizar_texto(
                    limpiar_solo_nombre_persona(str(r.get("Nombre_Alumno", "")))
                    or str(r.get("Nombre_Alumno", ""))
                )
                cat_r = normalizar_categoria(str(r.get("Categoria", "B1")))
                nums = extraer_numeros_clase(str(r.get("Clases_Reportadas", "")))
                if not nums:
                    continue
                if doc_r and doc_r not in ("pendiente", "desconocido", "s/n"):
                    por_doc.setdefault((doc_r, cat_r), set()).update(nums)
                if nom_r:
                    por_nom.setdefault((nom_r, cat_r), set()).update(nums)
        except Exception as exc:
            logger.warning("No se pudo calcular mapa de clases por alumno: %s", exc)
        return por_doc, por_nom

    def sincronizar_estados_alumnos_finalizados(self) -> int:
        """
        Verifica en la hoja 'Alumnos' quiénes ya cumplieron el total mínimo de clases de SU CATEGORÍA
        (8 en A1, 15 en A2, 20 en B1, 30 en C1, 10 en RECAT_C1) y actualiza su Estado a 'Finalizado'.
        Si un alumno con múltiples cursos fue marcado 'Finalizado' por error antes de cumplir las horas
        de esa categoría específica, lo reactiva a 'En curso'.
        """
        actualizados = 0
        try:
            por_doc, por_nom = self._calcular_mapa_clases_por_alumno()
            ws_alu = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            filas_alu = ws_alu.get_all_values()
            for idx_fila in range(2, len(filas_alu) + 1):
                fila = filas_alu[idx_fila - 1]
                if len(fila) < 3:
                    continue
                doc_a = _normalizar_texto(str(fila[1]).strip()) if len(fila) > 1 else ""
                nom_a = _normalizar_texto(str(fila[2]).strip()) if len(fila) > 2 else ""
                cat_a = normalizar_categoria(str(fila[3]).strip()) if len(fila) > 3 else "B1"
                est_a = str(fila[4]).strip() if len(fila) > 4 else "En curso"

                cfg = FORMATOS_OFICIALES.get(cat_a, FORMATOS_OFICIALES["B1"])
                total_req = cfg.get("total_clases", 20)
                clases_alu = set()
                if doc_a and (doc_a, cat_a) in por_doc:
                    clases_alu.update(por_doc[(doc_a, cat_a)])
                if nom_a and (nom_a, cat_a) in por_nom:
                    clases_alu.update(por_nom[(nom_a, cat_a)])

                clases_validas = {c for c in clases_alu if 1 <= c <= total_req}
                if len(clases_validas) >= total_req and "finalizado" not in est_a.lower():
                    ws_alu.update_cell(idx_fila, 5, "Finalizado")
                    actualizados += 1
                elif len(clases_validas) < total_req and "finalizado" in est_a.lower():
                    ws_alu.update_cell(idx_fila, 5, "En curso")
                    actualizados += 1
        except Exception as exc:
            logger.warning("Error sincronizando estados de alumnos finalizados: %s", exc)
        return actualizados

    def actualizar_estado_alumno(
        self,
        documento_o_nombre: str,
        nuevo_estado: str,
        categoria: Optional[str] = None,
    ) -> None:
        """
        Actualiza la columna Estado ('En curso' / 'Finalizado') de un alumno en la hoja Alumnos
        únicamente para la fila que coincida con la categoría indicada (si se especifica).
        """
        try:
            ws_alu = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            filas_alu = ws_alu.get_all_values()
            crit_norm = _normalizar_texto(documento_o_nombre)
            cat_buscada = normalizar_categoria(categoria) if categoria else None
            for idx_fila in range(2, len(filas_alu) + 1):
                fila = filas_alu[idx_fila - 1]
                doc_a = _normalizar_texto(str(fila[1]).strip()) if len(fila) > 1 else ""
                nom_a = _normalizar_texto(str(fila[2]).strip()) if len(fila) > 2 else ""
                cat_a = normalizar_categoria(str(fila[3]).strip()) if len(fila) > 3 else "B1"
                if crit_norm and crit_norm in (doc_a, nom_a):
                    if cat_buscada and cat_a != cat_buscada:
                        continue
                    ws_alu.update_cell(idx_fila, 5, nuevo_estado)
                    break
        except Exception as exc:
            logger.warning("No se pudo actualizar estado del alumno %s: %s", documento_o_nombre, exc)

    def obtener_documento_por_nombre_instructor(self, nombre_instructor: str) -> str:
        """Busca la cédula oficial de un instructor por su nombre en la hoja Instructores."""
        if not nombre_instructor:
            return ""
        nom_buscado = _normalizar_texto(nombre_instructor)
        try:
            ws_inst = self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
            registros = self._leer_filas_diccionario(ws_inst, HEADERS_INSTRUCTORES)
            for fila in reversed(registros):
                nom_f = _normalizar_texto(str(fila.get("Nombre_Instructor", "")))
                doc_f = str(fila.get("Documento_Instructor", "")).strip()
                if doc_f and nom_f and (nom_f == nom_buscado or nom_buscado in nom_f or nom_f in nom_buscado):
                    return doc_f
        except Exception:
            pass
        return ""

    def inicializar_estructura(self) -> None:
        self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
        self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
        self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
        self.obtener_o_crear_hoja(SHEET_CIERRES, HEADERS_CIERRES)
        self.obtener_o_crear_hoja(SHEET_INASISTENCIAS, HEADERS_INASISTENCIAS)
        self.obtener_o_crear_hoja(SHEET_MANTENIMIENTOS, HEADERS_MANTENIMIENTOS)
        self.obtener_o_crear_hoja(SHEET_VEHICULOS, HEADERS_VEHICULOS)
        self.sembrar_datos_historicos_appsheet()
        self.limpiar_nombres_contaminados()

    def sembrar_datos_historicos_appsheet(self) -> None:
        """
        Si las hojas 'Inasistencias', 'Mantenimientos' o 'Vehiculos_PESV' están recién creadas (sin datos),
        carga automáticamente el historial proveniente de la base de datos de AppSheet.
        """
        if self._sembrado_ejecutado:
            return
        self._sembrado_ejecutado = True
        try:
            ws_veh = self.obtener_o_crear_hoja(SHEET_VEHICULOS, HEADERS_VEHICULOS)
            filas_veh = ws_veh.get_all_values()
            datos_veh = [f for f in filas_veh[1:] if any(str(c).strip() for c in f)]
            if not datos_veh and PLACAS_VEHICULOS_OFICIALES:
                fecha_semilla = datetime.now(TZ_COLOMBIA).strftime("%Y-%m-%d")
                filas_semilla_veh = [
                    [
                        v["placa"].upper(),
                        "Moto" if "moto" in v["tipo"].lower() else "Carro",
                        "Activo",
                        fecha_semilla,
                    ]
                    for v in PLACAS_VEHICULOS_OFICIALES
                ]
                ws_veh.update(
                    values=filas_semilla_veh,
                    range_name="A2",
                    value_input_option="USER_ENTERED",
                )
                logger.info("✅ Sembrados %d vehículos iniciales en Vehiculos_PESV.", len(filas_semilla_veh))
        except Exception as exc_v:
            logger.warning("No se pudo sembrar vehículos iniciales en Vehiculos_PESV: %s", exc_v)
        try:
            ws_ina = self.obtener_o_crear_hoja(SHEET_INASISTENCIAS, HEADERS_INASISTENCIAS)
            filas_ina = ws_ina.get_all_values()
            datos_ina = [f for f in filas_ina[1:] if any(str(c).strip() for c in f)]
            if not datos_ina and SEED_INASISTENCIAS_HISTORICAS:
                if len(SEED_INASISTENCIAS_HISTORICAS) + 5 > ws_ina.row_count:
                    ws_ina.add_rows(len(SEED_INASISTENCIAS_HISTORICAS) + 50)
                ws_ina.update(
                    values=SEED_INASISTENCIAS_HISTORICAS,
                    range_name="A2",
                    value_input_option="USER_ENTERED",
                )
                logger.info("✅ Sembrados %d registros históricos de Inasistencias.", len(SEED_INASISTENCIAS_HISTORICAS))
        except Exception as exc_i:
            logger.warning("No se pudo sembrar histórico de Inasistencias: %s", exc_i)

        try:
            ws_man = self.obtener_o_crear_hoja(SHEET_MANTENIMIENTOS, HEADERS_MANTENIMIENTOS)
            filas_man = ws_man.get_all_values()
            datos_man = [f for f in filas_man[1:] if any(str(c).strip() for c in f)]
            if not datos_man and SEED_MANTENIMIENTOS_HISTORICOS:
                if len(SEED_MANTENIMIENTOS_HISTORICOS) + 5 > ws_man.row_count:
                    ws_man.add_rows(len(SEED_MANTENIMIENTOS_HISTORICOS) + 50)
                ws_man.update(
                    values=SEED_MANTENIMIENTOS_HISTORICOS,
                    range_name="A2",
                    value_input_option="USER_ENTERED",
                )
                logger.info("✅ Sembrados %d registros históricos de Mantenimientos.", len(SEED_MANTENIMIENTOS_HISTORICOS))
        except Exception as exc_m:
            logger.warning("No se pudo sembrar histórico de Mantenimientos: %s", exc_m)

    def listar_alumnos_activos(
        self, categoria: Optional[str] = None, modo: str = "activos"
    ) -> List[Dict[str, Any]]:
        """
        Lista alumnos filtrando automáticamente según su progreso de clases:
        - modo='activos': Solo devuelve alumnos 'En curso' que AÚN NO HAN CUMPLIDO el mínimo de clases
          de su categoría (15 en A1/A2, 20 en B1, 30 en C1, 10 en RECAT_C1).
          Si detecta que un alumno ya cumplió todas sus clases, lo marca 'Finalizado' y lo oculta.
        - modo='refuerzo': Devuelve únicamente los alumnos que YA FINALIZARON su curso, para cuando
          el instructor necesite dictar una Clase de Refuerzo adicional.
        """
        try:
            por_doc, por_nom = self._calcular_mapa_clases_por_alumno()
            ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            registros = self._leer_filas_diccionario(ws, HEADERS_ALUMNOS)
            resultado = []
            filtrar_cat = categoria and categoria.strip().upper() not in ("TODAS", "REFUERZO")

            for idx_r, f in enumerate(registros, start=2):
                raw_nom = str(f.get("Nombre_Completo", "")).strip()
                nombre = limpiar_solo_nombre_persona(raw_nom) or raw_nom
                if not nombre:
                    continue
                doc = str(f.get("Documento", "")).strip()
                cat_alumno = normalizar_categoria(str(f.get("Categoria", "B1")))
                estado = str(f.get("Estado", "En curso")).strip().lower()

                cfg = FORMATOS_OFICIALES.get(cat_alumno, FORMATOS_OFICIALES["B1"])
                total_req = cfg.get("total_clases", 20)

                doc_norm = _normalizar_texto(doc)
                nom_norm = _normalizar_texto(nombre)
                clases_vistas: set = set()
                if doc_norm and (doc_norm, cat_alumno) in por_doc:
                    clases_vistas.update(por_doc[(doc_norm, cat_alumno)])
                if nom_norm and (nom_norm, cat_alumno) in por_nom:
                    clases_vistas.update(por_nom[(nom_norm, cat_alumno)])

                clases_validas = {c for c in clases_vistas if 1 <= c <= total_req}
                cant_completadas = len(clases_validas)
                ya_cumplio_minimo = cant_completadas >= total_req

                # Sincronizar el estado en Sheets solo en modo interactivo del bot (no durante modo='todos' del dashboard para ahorrar cuota)
                if modo != "todos":
                    if cant_completadas >= total_req and "finalizado" not in estado:
                        try:
                            ws.update_cell(idx_r, 5, "Finalizado")
                        except Exception:
                            pass
                    elif cant_completadas < total_req and "finalizado" in estado:
                        try:
                            ws.update_cell(idx_r, 5, "En curso")
                        except Exception:
                            pass

                if filtrar_cat and cat_alumno != normalizar_categoria(categoria):
                    continue

                raw_runt = str(f.get("Estado_RUNT", "")).strip()
                est_runt = "Cargue RUNT" if "cargue" in raw_runt.lower() or "cargado" in raw_runt.lower() else "No Registrado RUNT"
                fecha_runt = str(f.get("Fecha_RUNT", "")).strip()

                item_alu = dict(f)
                item_alu["Categoria"] = cat_alumno
                item_alu["Clases_Completadas"] = cant_completadas
                item_alu["Clases_Requeridas"] = total_req
                item_alu["Curso_Completado"] = ya_cumplio_minimo
                item_alu["Clases_Vistas_Lista"] = sorted(clases_validas)
                item_alu["Estado_Real"] = "Finalizado" if ya_cumplio_minimo else "En curso"
                item_alu["Estado_RUNT"] = est_runt
                item_alu["Fecha_RUNT"] = fecha_runt

                if modo == "todos":
                    resultado.append(item_alu)
                elif modo == "refuerzo":
                    if ya_cumplio_minimo:
                        resultado.append(item_alu)
                else:
                    # En modo 'activos', ocultar SIEMPRE a quienes ya cumplieron el mínimo de clases
                    if not ya_cumplio_minimo:
                        resultado.append(item_alu)

            # Mostrar primero los alumnos más recientes (de abajo hacia arriba)
            return list(reversed(resultado))
        except Exception as exc:
            logger.warning("No se pudo listar alumnos (%s): %s", modo, exc)
            return []

    def obtener_instructor_registrado(self, telegram_user_id: int) -> Optional[Dict[str, str]]:
        """
        Verifica si el instructor ya está registrado en la pestaña 'Instructores'.
        Recorre de abajo hacia arriba para devolver el último instructor vinculado a ese Telegram_User_ID.
        """
        if telegram_user_id in self._cache_instructores:
            cached = self._cache_instructores[telegram_user_id]
            if len(cached.get("nombre", "").split()) >= 2:
                return cached

        try:
            ws = self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
            registros = self._leer_filas_diccionario(ws, HEADERS_INSTRUCTORES)
            for fila in reversed(registros):
                if str(fila.get("Telegram_User_ID", "")).strip() == str(telegram_user_id):
                    nombre = str(fila.get("Nombre_Instructor", "")).strip()
                    doc = str(fila.get("Documento_Instructor", "")).strip()
                    info = {
                        "nombre": nombre,
                        "documento": doc,
                        "categorias": str(fila.get("Categoria_Habilitada", "")).strip(),
                    }
                    if nombre and len(nombre.split()) >= 2:
                        self._cache_instructores[telegram_user_id] = info
                        return info
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
        Guarda en la pestaña 'Instructores' el Nombre Oficial y la Cédula del instructor.
        - Si el instructor ya existe (misma Cédula o mismo Nombre), actualiza únicamente su fila.
        - Si es un instructor nuevo (otra Cédula u otro Nombre), SIEMPRE crea una fila nueva al final
          sin borrar jamás al instructor anterior (incluso si se envía desde el mismo Telegram).
        """
        nombre_limpio = nombre_oficial.strip().title()
        doc_limpio = documento_oficial.strip()
        cat_norm = normalizar_categoria(categoria_inicial) if categoria_inicial else ""
        nombre_norm = _normalizar_texto(nombre_limpio)

        info = {
            "nombre": nombre_limpio,
            "documento": doc_limpio,
            "categorias": cat_norm,
        }
        self._cache_instructores[telegram_user_id] = info

        try:
            ws = self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
            todas_filas = ws.get_all_values()

            for idx_fila in range(2, len(todas_filas) + 1):
                fila = todas_filas[idx_fila - 1]
                nom_fila = str(fila[1]).strip() if len(fila) > 1 else ""
                cats_fila = str(fila[2]).strip() if len(fila) > 2 else ""
                doc_fila = str(fila[3]).strip() if len(fila) > 3 else ""

                mismo_doc = bool(doc_fila and doc_limpio and doc_fila == doc_limpio)
                mismo_nombre = bool(nom_fila and _normalizar_texto(nom_fila) == nombre_norm)

                # Solo actualiza si es exactamente la misma persona (misma cédula o mismo nombre)
                if mismo_doc or mismo_nombre:
                    cats_actuales = [c.strip() for c in cats_fila.split(",") if c.strip()]
                    if cat_norm and cat_norm not in cats_actuales:
                        cats_actuales.append(cat_norm)
                    cats_str = ", ".join(cats_actuales)
                    ws.update(
                        values=[[str(telegram_user_id), nombre_limpio, cats_str, doc_limpio or doc_fila]],
                        range_name=f"A{idx_fila}:D{idx_fila}",
                        value_input_option="USER_ENTERED",
                    )
                    info["categorias"] = cats_str
                    return info

            # Es un instructor nuevo: agregar siempre en una fila nueva debajo sin tocar las anteriores
            self._agregar_fila_al_final(
                ws, [str(telegram_user_id), nombre_limpio, cat_norm, doc_limpio]
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
        Busca al instructor en la pestaña 'Instructores' por su Cédula, Nombre o Telegram_User_ID.
        - Si ya existe ese instructor, actualiza sus categorías habilitadas sin borrar a nadie.
        - Si es un instructor nuevo, lo agrega en una nueva fila al final de 'Instructores'.
        """
        cat_norm = normalizar_categoria(categoria_dictada)
        nombre_limpio = nombre_telegram.strip()
        nombre_norm = _normalizar_texto(nombre_limpio)
        doc_limpio = documento_instructor.strip()

        try:
            ws = self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
            todas_filas = ws.get_all_values()

            # 1. Buscar coincidencia exacta por Documento o por Nombre del instructor
            fila_encontrada_idx = None
            for idx_fila in range(len(todas_filas), 1, -1):
                fila = todas_filas[idx_fila - 1]
                if not any(str(c).strip() for c in fila):
                    continue
                tg_fila = str(fila[0]).strip() if len(fila) > 0 else ""
                nom_fila = str(fila[1]).strip() if len(fila) > 1 else ""
                doc_fila = str(fila[3]).strip() if len(fila) > 3 else ""

                if (doc_limpio and doc_fila == doc_limpio) or (
                    nombre_norm and _normalizar_texto(nom_fila) == nombre_norm
                ):
                    fila_encontrada_idx = idx_fila
                    break

            # 2. Si no coincidió por nombre/documento, buscar la última fila con ese Telegram_User_ID
            if fila_encontrada_idx is None:
                for idx_fila in range(len(todas_filas), 1, -1):
                    fila = todas_filas[idx_fila - 1]
                    tg_fila = str(fila[0]).strip() if len(fila) > 0 else ""
                    if tg_fila == str(telegram_user_id):
                        fila_encontrada_idx = idx_fila
                        break

            if fila_encontrada_idx is not None:
                fila = todas_filas[fila_encontrada_idx - 1]
                nombre_oficial = (str(fila[1]).strip() if len(fila) > 1 else "") or nombre_limpio
                cats_fila = str(fila[2]).strip() if len(fila) > 2 else ""
                doc_actual = (str(fila[3]).strip() if len(fila) > 3 else "") or doc_limpio

                cats_actuales = [c.strip() for c in cats_fila.split(",") if c.strip()]
                if cat_norm and cat_norm not in cats_actuales:
                    cats_actuales.append(cat_norm)
                cats_str = ", ".join(cats_actuales)

                ws.update(
                    values=[[str(telegram_user_id), nombre_oficial, cats_str, doc_actual]],
                    range_name=f"A{fila_encontrada_idx}:D{fila_encontrada_idx}",
                    value_input_option="USER_ENTERED",
                )
                self._cache_instructores[telegram_user_id] = {
                    "nombre": nombre_oficial,
                    "documento": doc_actual,
                    "categorias": cats_str,
                }
                return nombre_oficial

            # Si no estaba registrado en la hoja 'Instructores', agregarlo en una fila nueva al final
            self._agregar_fila_al_final(
                ws, [str(telegram_user_id), nombre_limpio, cat_norm, doc_limpio]
            )
            self._cache_instructores[telegram_user_id] = {
                "nombre": nombre_limpio,
                "documento": doc_limpio,
                "categorias": cat_norm,
            }
            return nombre_limpio
        except Exception as exc:
            logger.warning("No se pudo actualizar hoja Instructores: %s", exc)
            return nombre_limpio

    def registrar_o_actualizar_alumno(
        self,
        documento: str,
        nombre_completo: str,
        categoria: str,
        estado_runt: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Registra o recupera la matrícula de un alumno para la CATEGORÍA indicada.
        - Si el alumno ya está matriculado en esa misma categoría, devuelve esa matrícula (y actualiza Estado_RUNT si se especifica).
        - Si el alumno ya existe en otra categoría (ej. A2 Moto) y ahora inicia un segundo curso
          (ej. C1 Carro), crea automáticamente una nueva fila de matrícula (MAT-XXX) para la
          nueva categoría sin sobrescribir ni bloquear el otro curso.
        """
        cat_norm = normalizar_categoria(categoria)
        nombre_depurado = limpiar_solo_nombre_persona(nombre_completo) or nombre_completo.strip().title()
        runt_norm = ""
        if estado_runt:
            runt_norm = "Cargue RUNT" if ("cargue" in estado_runt.lower() or "cargado" in estado_runt.lower()) else "No Registrado RUNT"

        # 1. Buscar si ya tiene matrícula en ESTA MISMA categoría
        existente_misma_cat = (
            self.buscar_alumno(documento, categoria=cat_norm, solo_misma_categoria=True)
            or self.buscar_alumno(nombre_depurado, categoria=cat_norm, solo_misma_categoria=True)
        )
        if existente_misma_cat:
            nom_actual = str(existente_misma_cat.get("Nombre_Completo", "")).strip()
            doc_actual = str(existente_misma_cat.get("Documento", "")).strip()
            runt_actual = str(existente_misma_cat.get("Estado_RUNT", "")).strip() or "No Registrado RUNT"
            nom_limpio = nombre_depurado if (nombre_depurado and nombre_depurado != "Desconocido") else (limpiar_solo_nombre_persona(nom_actual) or nom_actual)
            doc_limpio = documento.strip() if (documento and documento.strip() not in ("", "Pendiente", "Desconocido")) else doc_actual
            runt_final = runt_norm if runt_norm else ("Cargue RUNT" if "cargue" in runt_actual.lower() else "No Registrado RUNT")
            fecha_runt_final = (
                datetime.now(TZ_COLOMBIA).strftime("%Y-%m-%d %H:%M")
                if runt_final == "Cargue RUNT"
                else str(existente_misma_cat.get("Fecha_RUNT", "")).strip()
            )
            # Si se actualizaron datos desde la Web (nombre, cédula o Estado_RUNT), reflejarlo en Google Sheets
            if (nom_limpio and nom_limpio != nom_actual) or (doc_limpio and doc_limpio != doc_actual) or (runt_norm and runt_norm != runt_actual):
                try:
                    ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
                    todas_al = ws.get_all_values()
                    id_mat_ex = str(existente_misma_cat.get("ID_Matricula", "")).strip()
                    est_curso_ex = str(existente_misma_cat.get("Estado", "En curso")).strip() or "En curso"
                    for idx_f in range(2, len(todas_al) + 1):
                        fila_f = todas_al[idx_f - 1]
                        if len(fila_f) >= 4:
                            m_id_f = str(fila_f[0]).strip()
                            d_f = _normalizar_texto(str(fila_f[1]).strip())
                            n_f = _normalizar_texto(str(fila_f[2]).strip())
                            c_f = normalizar_categoria(str(fila_f[3]).strip())
                            est_f = str(fila_f[4]).strip() if len(fila_f) > 4 and str(fila_f[4]).strip() else est_curso_ex
                            if c_f == cat_norm and (
                                (id_mat_ex and m_id_f == id_mat_ex)
                                or (_normalizar_texto(doc_actual) and d_f == _normalizar_texto(doc_actual))
                                or (_normalizar_texto(nom_actual) and n_f == _normalizar_texto(nom_actual))
                            ):
                                ws.update(
                                    values=[[m_id_f or id_mat_ex, doc_limpio, nom_limpio, cat_norm, est_f, runt_final, fecha_runt_final]],
                                    range_name=f"A{idx_f}:G{idx_f}",
                                    value_input_option="USER_ENTERED",
                                )
                                break
                except Exception as exc_up_al:
                    logger.warning("Aviso actualizando datos de alumno existente en Sheets: %s", exc_up_al)
            existente_misma_cat["Documento"] = doc_limpio
            existente_misma_cat["Nombre_Completo"] = nom_limpio
            existente_misma_cat["Categoria"] = cat_norm
            existente_misma_cat["Estado_RUNT"] = runt_final
            existente_misma_cat["Fecha_RUNT"] = fecha_runt_final
            return existente_misma_cat

        # 2. Si está en otra categoría, recuperar su nombre/cédula oficial para crearle la matrícula en la nueva categoría
        existente_otra_cat = self.buscar_alumno(documento) or self.buscar_alumno(nombre_depurado)
        if existente_otra_cat:
            doc_prev = str(existente_otra_cat.get("Documento", "")).strip()
            nom_prev = str(existente_otra_cat.get("Nombre_Completo", "")).strip()
            if (not documento or documento in ("Pendiente", "Desconocido")) and doc_prev:
                documento = doc_prev
            if (not nombre_depurado or nombre_depurado == "Desconocido") and nom_prev:
                nombre_depurado = limpiar_solo_nombre_persona(nom_prev) or nom_prev

        runt_inicial = runt_norm if runt_norm else "No Registrado RUNT"
        fecha_runt_ini = datetime.now(TZ_COLOMBIA).strftime("%Y-%m-%d %H:%M") if runt_inicial == "Cargue RUNT" else ""
        try:
            ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            filas = self._leer_filas_diccionario(ws, HEADERS_ALUMNOS)
            max_num = len(filas)
            for f in filas:
                m_id = re.search(r"(\d+)", str(f.get("ID_Matricula", "")))
                if m_id:
                    max_num = max(max_num, int(m_id.group(1)))
            nuevo_id = f"MAT-{max_num + 1:03d}"
            nueva_fila = [nuevo_id, documento, nombre_depurado, cat_norm, "En curso", runt_inicial, fecha_runt_ini]
            self._agregar_fila_al_final(ws, nueva_fila)
            return {
                "ID_Matricula": nuevo_id,
                "Documento": documento,
                "Nombre_Completo": nombre_depurado,
                "Categoria": cat_norm,
                "Estado": "En curso",
                "Estado_RUNT": runt_inicial,
                "Fecha_RUNT": fecha_runt_ini,
            }
        except Exception as exc:
            logger.warning("No se pudo agregar automáticamente en hoja Alumnos: %s", exc)
            return {
                "ID_Matricula": documento,
                "Documento": documento,
                "Nombre_Completo": nombre_depurado,
                "Categoria": cat_norm,
                "Estado": "En curso",
                "Estado_RUNT": runt_inicial,
                "Fecha_RUNT": fecha_runt_ini,
            }

    def actualizar_estado_runt_alumno(
        self,
        documento_o_nombre: str,
        categoria: str,
        estado_runt: str,
        nombre_alumno: str = "",
    ) -> Dict[str, Any]:
        """
        Actualiza manualmente el Estado_RUNT ('Cargue RUNT' o 'No Registrado RUNT')
        para un alumno y categoría específicos (ej. A2 Moto vs C1 Carro son independientes).
        Si el alumno tenía clases en Registro_Clases pero aún no tenía fila en Alumnos para esa categoría,
        crea la fila y la marca con el estado RUNT indicado.
        """
        cat_norm = normalizar_categoria(categoria)
        es_cargado = "cargue" in str(estado_runt).lower() or "cargado" in str(estado_runt).lower()
        runt_val = "Cargue RUNT" if es_cargado else "No Registrado RUNT"
        fecha_runt = datetime.now(TZ_COLOMBIA).strftime("%Y-%m-%d %H:%M") if es_cargado else ""

        crit_norm = _normalizar_texto(documento_o_nombre)
        nom_norm = _normalizar_texto(nombre_alumno)

        try:
            ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            todas = ws.get_all_values()
            for idx_f in range(2, len(todas) + 1):
                fila = todas[idx_f - 1]
                if not any(str(c).strip() for c in fila):
                    continue
                id_m = _normalizar_texto(str(fila[0]).strip()) if len(fila) > 0 else ""
                doc_f = _normalizar_texto(str(fila[1]).strip()) if len(fila) > 1 else ""
                nom_f = _normalizar_texto(str(fila[2]).strip()) if len(fila) > 2 else ""
                cat_f = normalizar_categoria(str(fila[3]).strip()) if len(fila) > 3 else "B1"

                if cat_f != cat_norm:
                    continue

                coincide = False
                if crit_norm and crit_norm in (id_m, doc_f, nom_f):
                    coincide = True
                elif nom_norm and nom_f and (nom_norm == nom_f or nom_norm in nom_f or nom_f in nom_norm):
                    coincide = True

                if coincide:
                    est_actual = str(fila[5]).strip() if len(fila) > 5 else ""
                    fecha_actual = str(fila[6]).strip() if len(fila) > 6 else ""
                    sin_cambio = (
                        est_actual.lower() == runt_val.lower()
                        and ((es_cargado and bool(fecha_actual)) or (not es_cargado and not fecha_actual))
                    )
                    if not sin_cambio:
                        ws.update(
                            values=[[runt_val, fecha_runt]],
                            range_name=f"F{idx_f}:G{idx_f}",
                            value_input_option="USER_ENTERED",
                        )
                    return {
                        "ok": True,
                        "ID_Matricula": fila[0] if len(fila) > 0 else "",
                        "Documento": fila[1] if len(fila) > 1 else documento_o_nombre,
                        "Nombre_Completo": fila[2] if len(fila) > 2 else (nombre_alumno or documento_o_nombre),
                        "Categoria": cat_norm,
                        "Estado_RUNT": runt_val,
                        "Fecha_RUNT": fecha_actual if sin_cambio else fecha_runt,
                    }
        except Exception as exc:
            logger.warning("Aviso buscando fila para actualizar_estado_runt_alumno: %s", exc)

        # Si no existía fila en hoja Alumnos para esta categoría, crearla con el Estado_RUNT solicitado
        creado = self.registrar_o_actualizar_alumno(
            documento=documento_o_nombre if documento_o_nombre.isdigit() else "",
            nombre_completo=nombre_alumno or documento_o_nombre,
            categoria=cat_norm,
            estado_runt=runt_val,
        )
        creado["ok"] = True
        return creado

    def buscar_matriculas_alumno(self, criterio: str) -> List[Dict[str, Any]]:
        """
        Devuelve TODAS las filas de matrícula que tenga un mismo alumno en la hoja 'Alumnos'
        (por ejemplo, cuando cursa simultáneamente A2 Moto y C1 Carro).
        """
        if not criterio or criterio in ("Desconocido", "Pendiente"):
            return []
        try:
            ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            registros = self._leer_filas_diccionario(ws, HEADERS_ALUMNOS)
        except Exception as exc:
            logger.error("Error leyendo hoja Alumnos: %s", exc)
            return []

        criterio_limpio = _normalizar_texto(criterio)
        if not criterio_limpio:
            return []

        coincidencias: List[Dict[str, Any]] = []
        doc_encontrado = ""
        nom_encontrado = ""

        for fila in registros:
            id_mat = _normalizar_texto(str(fila.get("ID_Matricula", "")).strip())
            doc = _normalizar_texto(str(fila.get("Documento", "")).strip())
            nombre_norm = _normalizar_texto(str(fila.get("Nombre_Completo", "")).strip())

            es_match = False
            if criterio_limpio in (id_mat, doc) and criterio_limpio:
                es_match = True
            elif nombre_norm and (
                criterio_limpio == nombre_norm
                or (len(criterio_limpio.split()) >= 2 and criterio_limpio in nombre_norm)
            ):
                es_match = True

            if es_match:
                coincidencias.append(fila)
                if doc:
                    doc_encontrado = doc
                if nombre_norm:
                    nom_encontrado = nombre_norm

        # Si coincidió por ID_Matricula en una fila, traer también las demás matrículas con la misma cédula o nombre
        if coincidencias and (doc_encontrado or nom_encontrado):
            ids_vistos = {id(f) for f in coincidencias}
            for fila in registros:
                if id(fila) in ids_vistos:
                    continue
                doc = _normalizar_texto(str(fila.get("Documento", "")).strip())
                nombre_norm = _normalizar_texto(str(fila.get("Nombre_Completo", "")).strip())
                if (doc_encontrado and doc == doc_encontrado) or (
                    nom_encontrado and nombre_norm == nom_encontrado
                ):
                    coincidencias.append(fila)
                    ids_vistos.add(id(fila))

        return coincidencias

    def buscar_alumno(
        self,
        criterio: str,
        categoria: Optional[str] = None,
        solo_misma_categoria: bool = False,
    ) -> Optional[Dict[str, Any]]:
        """
        Busca un alumno en la hoja 'Alumnos'.
        - Si se especifica 'categoria' (ej. 'C1'), prioriza la matrícula de ese alumno en esa categoría.
        - Si 'solo_misma_categoria=True', solo devuelve resultado si existe en esa categoría exacta.
        """
        if not criterio or criterio in ("Desconocido", "Pendiente"):
            return None
        try:
            ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            registros = self._leer_filas_diccionario(ws, HEADERS_ALUMNOS)
        except Exception as exc:
            logger.error("Error leyendo hoja Alumnos: %s", exc)
            return None

        criterio_limpio = _normalizar_texto(criterio)
        if not criterio_limpio:
            return None

        cat_buscada = normalizar_categoria(categoria) if categoria else None
        mapa_nombres_cat: Dict[str, Dict[str, Any]] = {}
        mapa_nombres_general: Dict[str, Dict[str, Any]] = {}
        candidato_otra_cat: Optional[Dict[str, Any]] = None

        for fila in registros:
            id_mat = str(fila.get("ID_Matricula", "")).strip()
            doc = str(fila.get("Documento", "")).strip()
            nombre = str(fila.get("Nombre_Completo", "")).strip()
            cat_fila = normalizar_categoria(str(fila.get("Categoria", "B1")))
            misma_cat = (cat_buscada is None) or (cat_fila == cat_buscada)

            if criterio_limpio in (_normalizar_texto(id_mat), _normalizar_texto(doc)):
                if misma_cat:
                    return fila
                if candidato_otra_cat is None:
                    candidato_otra_cat = fila

            nombre_norm = _normalizar_texto(nombre)
            if nombre_norm:
                if criterio_limpio == nombre_norm or (
                    len(criterio_limpio.split()) >= 2 and criterio_limpio in nombre_norm
                ):
                    if misma_cat:
                        return fila
                    if candidato_otra_cat is None:
                        candidato_otra_cat = fila
                if misma_cat:
                    mapa_nombres_cat[nombre_norm] = fila
                mapa_nombres_general[nombre_norm] = fila

        if mapa_nombres_cat and len(criterio_limpio.split()) >= 2:
            coincidencias = difflib.get_close_matches(
                criterio_limpio, list(mapa_nombres_cat.keys()), n=1, cutoff=0.72
            )
            if coincidencias:
                return mapa_nombres_cat[coincidencias[0]]

        if solo_misma_categoria:
            return None

        if candidato_otra_cat is not None:
            return candidato_otra_cat

        if mapa_nombres_general and len(criterio_limpio.split()) >= 2:
            coincidencias = difflib.get_close_matches(
                criterio_limpio, list(mapa_nombres_general.keys()), n=1, cutoff=0.72
            )
            if coincidencias:
                return mapa_nombres_general[coincidencias[0]]
        return None

    def reclasificar_clases_alumno(
        self,
        id_alumno: str,
        nombre_alumno: str,
        cat_origen: str,
        cat_destino: str,
    ) -> int:
        """
        Mueve en la hoja 'Registro_Clases' los registros previos de un alumno que estaban
        en 'cat_origen' (ej. C1 por error) hacia 'cat_destino' (ej. A2 Moto), liberando
        la categoría de origen para registrar las clases del segundo curso.
        """
        modificados = 0
        id_norm = _normalizar_texto(id_alumno)
        nom_norm = _normalizar_texto(nombre_alumno)
        orig_norm = normalizar_categoria(cat_origen)
        dest_norm = normalizar_categoria(cat_destino)
        try:
            ws_reg = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
            filas = ws_reg.get_all_values()
            for idx_fila in range(2, len(filas) + 1):
                fila = filas[idx_fila - 1]
                if not any(str(c).strip() for c in fila):
                    continue
                id_f = _normalizar_texto(str(fila[4]).strip()) if len(fila) > 4 else ""
                nom_f = _normalizar_texto(
                    limpiar_solo_nombre_persona(str(fila[5]).strip() if len(fila) > 5 else "")
                    or (str(fila[5]).strip() if len(fila) > 5 else "")
                )
                cat_f = normalizar_categoria(str(fila[6]).strip()) if len(fila) > 6 else "B1"
                mismo_id = bool(
                    id_norm and id_f and id_norm not in ("pendiente", "desconocido") and id_f == id_norm
                )
                mismo_nom = bool(nom_norm and nom_f and nom_f == nom_norm)
                if (mismo_id or mismo_nom) and cat_f == orig_norm:
                    ws_reg.update_cell(idx_fila, 7, dest_norm)
                    modificados += 1
            self.sincronizar_estados_alumnos_finalizados()
        except Exception as exc:
            logger.warning("No se pudo reclasificar clases del alumno %s: %s", nombre_alumno, exc)
        return modificados

    def verificar_clase_duplicada(
        self,
        id_alumno: str,
        nombre_alumno: str,
        categoria: str,
        clases_reportadas: str,
    ) -> Optional[Dict[str, Any]]:
        """
        Verifica si alguna de las clases que el instructor intenta registrar (ej. '1-2')
        YA FUE REGISTRADA anteriormente para ese mismo alumno EN ESA MISMA CATEGORÍA.
        Si encuentra duplicado, devuelve el detalle del registro previo, la sugerencia de la siguiente clase
        y las demás categorías activas del alumno (por si realiza 2 cursos a la vez como A2 y C1).
        """
        nums_nuevos = set(extraer_numeros_clase(clases_reportadas))
        if not nums_nuevos:
            return None

        try:
            ws_reg = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
            todas_filas = self._leer_filas_diccionario(ws_reg, HEADERS_REGISTRO)
        except Exception as exc:
            logger.warning("No se pudo verificar duplicado de clase: %s", exc)
            return None

        id_norm = _normalizar_texto(id_alumno)
        nom_norm = _normalizar_texto(nombre_alumno)
        cat_norm = normalizar_categoria(categoria)

        todas_clases_alumno: set = set()
        clases_por_otra_cat: Dict[str, set] = {}
        conflicto_info: Optional[Dict[str, Any]] = None

        for fila in todas_filas:
            if _es_registro_inasistencia(fila):
                continue
            id_f = _normalizar_texto(str(fila.get("ID_Alumno", "")))
            nom_f = _normalizar_texto(
                limpiar_solo_nombre_persona(str(fila.get("Nombre_Alumno", "")))
                or str(fila.get("Nombre_Alumno", ""))
            )
            cat_f = normalizar_categoria(str(fila.get("Categoria", "B1")))

            mismo_id = bool(
                id_norm
                and id_f
                and id_norm not in ("pendiente", "desconocido")
                and id_f == id_norm
            )
            mismo_nom = bool(nom_norm and nom_f and nom_f == nom_norm)
            if not (mismo_id or mismo_nom):
                continue

            nums_fila = set(extraer_numeros_clase(str(fila.get("Clases_Reportadas", ""))))
            if cat_f != cat_norm:
                if nums_fila:
                    clases_por_otra_cat.setdefault(cat_f, set()).update(nums_fila)
                continue

            todas_clases_alumno.update(nums_fila)

            repetidas = sorted(nums_nuevos & nums_fila)
            if repetidas and conflicto_info is None:
                conflicto_info = {
                    "clases_duplicadas": formatear_clases(repetidas),
                    "fecha_previa": str(fila.get("Fecha_Reporte", "")).strip() or "Hoy",
                    "hora_previa": str(fila.get("Hora_Reporte", "")).strip() or "--:--",
                    "instructor_previo": str(fila.get("Instructor", "")).strip() or "Instructor",
                }

        if not conflicto_info:
            return None

        cfg = FORMATOS_OFICIALES.get(cat_norm, FORMATOS_OFICIALES["B1"])
        total_cat = cfg.get("total_clases", 20)
        ordenadas_vistas = sorted(todas_clases_alumno)
        conflicto_info["todas_registradas_txt"] = (
            ", ".join(str(c) for c in ordenadas_vistas) if ordenadas_vistas else "Ninguna"
        )

        # Detectar otras categorías matriculadas para este mismo alumno (ej. A2 y C1)
        matriculas = self.buscar_matriculas_alumno(id_alumno) or self.buscar_matriculas_alumno(nombre_alumno)
        otras_cats: List[str] = []
        for m in matriculas:
            c_m = normalizar_categoria(str(m.get("Categoria", "")))
            if c_m and c_m != cat_norm and c_m not in otras_cats:
                otras_cats.append(c_m)
        for c_o in clases_por_otra_cat.keys():
            if c_o != cat_norm and c_o not in otras_cats:
                otras_cats.append(c_o)

        # Si el alumno no tenía segunda fila en Alumnos pero las clases son 1..15 y está en C1/B1, sugerir también A2
        if not otras_cats and cat_norm in ("C1", "B1"):
            otras_cats.append("A2")
        elif not otras_cats and cat_norm in ("A1", "A2"):
            otras_cats.append("C1")

        conflicto_info["otras_categorias"] = otras_cats
        resumen_otras: List[str] = []
        for c_o in otras_cats:
            vistas_o = sorted(clases_por_otra_cat.get(c_o, set()))
            txt_o = ", ".join(str(x) for x in vistas_o) if vistas_o else "Ninguna aún"
            nom_corto_o = FORMATOS_OFICIALES.get(c_o, FORMATOS_OFICIALES["B1"])["nombre_corto"]
            resumen_otras.append(f"• *{nom_corto_o}:* `{txt_o}`")
        conflicto_info["resumen_otras_txt"] = "\n".join(resumen_otras)

        # Calcular cuál es la siguiente clase disponible para sugerirla al instructor
        disponibles = [c for c in range(1, total_cat + 1) if c not in todas_clases_alumno]
        if len(disponibles) >= 2 and disponibles[1] == disponibles[0] + 1:
            conflicto_info["sugerencia_siguiente"] = f"{disponibles[0]}-{disponibles[1]}"
        elif disponibles:
            conflicto_info["sugerencia_siguiente"] = str(disponibles[0])
        else:
            conflicto_info["sugerencia_siguiente"] = "Curso completado"

        return conflicto_info

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
        observacion_instructor: str = "",
    ) -> int:
        ws = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
        return self._agregar_fila_al_final(
            ws,
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
                observacion_instructor,
            ],
        )

    def guardar_observacion_en_fila(
        self,
        fila_idx: int,
        observacion_formateada: str,
    ) -> None:
        """Guarda o agrega una observación del instructor en la columna L (Observaciones_Instructor) de Registro_Clases."""
        try:
            ws = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
            valor_actual = str(ws.cell(fila_idx, 12).value or "").strip()
            nuevo_valor = (
                f"{valor_actual} | {observacion_formateada}"
                if valor_actual
                else observacion_formateada
            )
            ws.update_cell(fila_idx, 12, nuevo_valor)
        except Exception as exc:
            logger.warning("No se pudo guardar observación en fila %s: %s", fila_idx, exc)

    def registrar_cierre_jornada(
        self,
        timestamp: str,
        fecha_jornada: str,
        hora_cierre: str,
        instructor: str,
        documento_instructor: str,
        clases_registradas_bot: int,
        clases_segun_agenda: int,
        estado_cruce: str,
        detalle_alumnos_dia: str,
        url_captura_agenda: str,
    ) -> None:
        ws = self.obtener_o_crear_hoja(SHEET_CIERRES, HEADERS_CIERRES)
        self._agregar_fila_al_final(
            ws,
            [
                timestamp,
                fecha_jornada,
                hora_cierre,
                instructor,
                documento_instructor,
                str(clases_registradas_bot),
                str(clases_segun_agenda),
                estado_cruce,
                detalle_alumnos_dia,
                url_captura_agenda,
            ],
        )

    def obtener_clases_alumno(
        self,
        criterio_alumno: str,
        categoria: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Obtiene los datos del alumno y las filas de 'Registro_Clases'.
        - Si se especifica 'categoria', filtra estrictamente los registros de esa categoría
          para no mezclar las clases de Moto (A2) con las de Carro (B1/C1).
        - También devuelve 'categorias_detectadas' con todas las categorías en las que el alumno
          tiene matrícula o registros de clases.
        """
        cat_filtro = normalizar_categoria(categoria) if categoria else None
        info_alumno = self.buscar_alumno(criterio_alumno, categoria=cat_filtro)
        matriculas = self.buscar_matriculas_alumno(criterio_alumno)
        ws_reg = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
        todas_filas = self._leer_filas_diccionario(ws_reg, HEADERS_REGISTRO)

        criterio_norm = _normalizar_texto(criterio_alumno)
        nombre_oficial_norm = (
            _normalizar_texto(info_alumno.get("Nombre_Completo", ""))
            if info_alumno
            else criterio_norm
        )
        ids_oficiales_norm = {
            _normalizar_texto(str(m.get("ID_Matricula", "")))
            for m in matriculas
            if str(m.get("ID_Matricula", "")).strip()
        }
        if info_alumno and str(info_alumno.get("ID_Matricula", "")).strip():
            ids_oficiales_norm.add(_normalizar_texto(str(info_alumno.get("ID_Matricula", ""))))
        doc_oficial_norm = (
            _normalizar_texto(str(info_alumno.get("Documento", "")))
            if info_alumno
            else ""
        )

        clases_filtradas = []
        categorias_en_registros: List[str] = []
        for m in matriculas:
            c_mat = normalizar_categoria(str(m.get("Categoria", "")))
            if c_mat and c_mat not in categorias_en_registros:
                categorias_en_registros.append(c_mat)

        for fila in todas_filas:
            id_fila = _normalizar_texto(str(fila.get("ID_Alumno", "")))
            nom_fila = _normalizar_texto(str(fila.get("Nombre_Alumno", "")))
            cat_fila = normalizar_categoria(str(fila.get("Categoria", "B1")))

            coincide_id = bool(
                id_fila and (id_fila in ids_oficiales_norm or id_fila in (doc_oficial_norm, criterio_norm))
            )
            coincide_nombre = bool(
                nom_fila
                and (
                    criterio_norm in nom_fila
                    or (nombre_oficial_norm and nombre_oficial_norm in nom_fila)
                    or (nombre_oficial_norm and nom_fila in nombre_oficial_norm)
                )
            )
            if coincide_id or coincide_nombre:
                if cat_fila not in categorias_en_registros:
                    categorias_en_registros.append(cat_fila)
                if cat_filtro is None or cat_fila == cat_filtro:
                    clases_filtradas.append(fila)

        if not info_alumno and clases_filtradas:
            ultima = clases_filtradas[-1]
            info_alumno = {
                "ID_Matricula": ultima.get("ID_Alumno") or "S/N",
                "Documento": ultima.get("ID_Alumno") or "Pendiente",
                "Nombre_Completo": ultima.get("Nombre_Alumno") or criterio_alumno,
                "Categoria": cat_filtro or normalizar_categoria(str(ultima.get("Categoria", "B1"))),
                "Estado": "En curso",
            }
        elif info_alumno and cat_filtro:
            info_alumno = dict(info_alumno)
            info_alumno["Categoria"] = cat_filtro

        return {
            "alumno": info_alumno,
            "registros": clases_filtradas,
            "categorias_detectadas": categorias_en_registros,
        }

    def listar_instructores_registrados(self) -> List[Dict[str, str]]:
        """
        Devuelve la lista consolidada de instructores (de la hoja 'Instructores' y de 'Registro_Clases').
        """
        instructores: List[Dict[str, str]] = []
        vistos_norm: set = set()
        try:
            ws_inst = self.obtener_o_crear_hoja(SHEET_INSTRUCTORES, HEADERS_INSTRUCTORES)
            filas_inst = self._leer_filas_diccionario(ws_inst, HEADERS_INSTRUCTORES)
            for f in filas_inst:
                nom = str(f.get("Nombre_Instructor", "")).strip()
                doc = str(f.get("Documento_Instructor", "")).strip()
                tg_id = str(f.get("Telegram_User_ID", "")).strip()
                if not nom:
                    continue
                norm = _normalizar_texto(nom)
                if norm not in vistos_norm:
                    vistos_norm.add(norm)
                    instructores.append(
                        {"nombre": nom, "documento": doc, "telegram_id": tg_id}
                    )

            ws_reg = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
            filas_reg = self._leer_filas_diccionario(ws_reg, HEADERS_REGISTRO)
            for f in filas_reg:
                nom = str(f.get("Instructor", "")).strip()
                if not nom:
                    continue
                norm = _normalizar_texto(nom)
                if norm not in vistos_norm:
                    vistos_norm.add(norm)
                    instructores.append(
                        {"nombre": nom, "documento": "", "telegram_id": ""}
                    )
        except Exception as exc:
            logger.warning("Error listando instructores: %s", exc)
        return instructores

    @staticmethod
    def obtener_rango_fechas_periodo(
        periodo: str, ahora: Optional[datetime] = None
    ) -> Tuple[str, str, str]:
        """
        Retorna (fecha_inicio_yyyy_mm_dd, fecha_fin_yyyy_mm_dd, etiqueta_legible).
        Períodos soportados: 'dia' (hoy), 'ayer', 'semana' (lun-dom actual),
        'semana_ant' (lun-dom semana anterior), 'mes' (mes actual).
        """
        dt = ahora or datetime.now(TZ_COLOMBIA)
        hoy_date = dt.date()
        per = (periodo or "semana").strip().lower()

        if per in ("dia", "día", "hoy"):
            f_str = hoy_date.strftime("%Y-%m-%d")
            return f_str, f_str, f"Día de Hoy ({f_str})"

        if per == "ayer":
            ayer_date = hoy_date - timedelta(days=1)
            f_str = ayer_date.strftime("%Y-%m-%d")
            return f_str, f_str, f"Día de Ayer ({f_str})"

        if per in ("semana_ant", "semana_anterior", "pasada"):
            inicio_actual = hoy_date - timedelta(days=hoy_date.weekday())
            ini_ant = inicio_actual - timedelta(days=7)
            fin_ant = inicio_actual - timedelta(days=1)
            return (
                ini_ant.strftime("%Y-%m-%d"),
                fin_ant.strftime("%Y-%m-%d"),
                f"Semana Anterior ({ini_ant.strftime('%Y-%m-%d')} al {fin_ant.strftime('%Y-%m-%d')})",
            )

        if per == "mes":
            ini_mes = hoy_date.replace(day=1)
            if hoy_date.month == 12:
                sig_mes = hoy_date.replace(year=hoy_date.year + 1, month=1, day=1)
            else:
                sig_mes = hoy_date.replace(month=hoy_date.month + 1, day=1)
            fin_mes = sig_mes - timedelta(days=1)
            return (
                ini_mes.strftime("%Y-%m-%d"),
                fin_mes.strftime("%Y-%m-%d"),
                f"Mes Actual ({ini_mes.strftime('%Y-%m-%d')} al {fin_mes.strftime('%Y-%m-%d')})",
            )

        if per in ("mes_ant", "mes_anterior"):
            primero_actual = hoy_date.replace(day=1)
            fin_mes_ant = primero_actual - timedelta(days=1)
            ini_mes_ant = fin_mes_ant.replace(day=1)
            return (
                ini_mes_ant.strftime("%Y-%m-%d"),
                fin_mes_ant.strftime("%Y-%m-%d"),
                f"Mes Anterior ({ini_mes_ant.strftime('%Y-%m-%d')} al {fin_mes_ant.strftime('%Y-%m-%d')})",
            )

        if per in ("anio", "año", "anio_actual"):
            y = hoy_date.year
            return f"{y}-01-01", f"{y}-12-31", f"Año Actual ({y})"

        if per in ("anio_ant", "año_anterior", "2025", "anio_2025"):
            return "2025-01-01", "2025-12-31", "Año 2025 (01/01/2025 al 31/12/2025)"

        if per in ("2026", "anio_2026"):
            return "2026-01-01", "2026-12-31", "Año 2026 (01/01/2026 al 31/12/2026)"

        if per in ("todo", "historico", "histórico", "todos"):
            return "2020-01-01", "2099-12-31", "Histórico Completo (Todos los registros)"

        meses_esp = {
            "enero": (1, "Enero"),
            "febrero": (2, "Febrero"),
            "marzo": (3, "Marzo"),
            "abril": (4, "Abril"),
            "mayo": (5, "Mayo"),
            "junio": (6, "Junio"),
            "julio": (7, "Julio"),
            "agosto": (8, "Agosto"),
            "septiembre": (9, "Septiembre"),
            "octubre": (10, "Octubre"),
            "noviembre": (11, "Noviembre"),
            "diciembre": (12, "Diciembre"),
        }
        if per in meses_esp:
            m_num, m_nom = meses_esp[per]
            y = hoy_date.year
            ini_m = hoy_date.replace(year=y, month=m_num, day=1)
            if m_num == 12:
                sig_m = hoy_date.replace(year=y + 1, month=1, day=1)
            else:
                sig_m = hoy_date.replace(year=y, month=m_num + 1, day=1)
            fin_m = sig_m - timedelta(days=1)
            return (
                ini_m.strftime("%Y-%m-%d"),
                fin_m.strftime("%Y-%m-%d"),
                f"Mes de {m_nom}",
            )

        # Por defecto: 'semana' (Lunes a Domingo de la semana actual)
        ini_sem = hoy_date - timedelta(days=hoy_date.weekday())
        fin_sem = ini_sem + timedelta(days=6)
        return (
            ini_sem.strftime("%Y-%m-%d"),
            fin_sem.strftime("%Y-%m-%d"),
            f"Semana Actual ({ini_sem.strftime('%Y-%m-%d')} al {fin_sem.strftime('%Y-%m-%d')})",
        )

    def consultar_clases_por_instructor(
        self, criterio_instructor: str, periodo: str = "semana"
    ) -> Dict[str, Any]:
        """
        Filtra los registros de 'Registro_Clases' por instructor (o 'TODOS') y por período
        ('dia', 'ayer', 'semana', 'semana_ant', 'mes'), calculando el total exacto de clases dictadas
        (contando bloques dobles como 2 clases y bloques individuales como 1 clase).
        """
        fecha_ini, fecha_fin, etiqueta_periodo = self.obtener_rango_fechas_periodo(periodo)
        instructores_lista = self.listar_instructores_registrados()

        es_todos = criterio_instructor.strip().upper() == "TODOS"
        criterio_norm = _normalizar_texto(criterio_instructor)

        nombre_oficial = criterio_instructor.strip()
        doc_oficial = ""
        if not es_todos:
            for inst in instructores_lista:
                nom_i = inst["nombre"]
                doc_i = inst["documento"]
                tg_i = inst["telegram_id"]
                if (
                    criterio_norm == _normalizar_texto(nom_i)
                    or (doc_i and criterio_norm == _normalizar_texto(doc_i))
                    or (tg_i and criterio_norm == _normalizar_texto(tg_i))
                    or (criterio_norm and criterio_norm in _normalizar_texto(nom_i))
                ):
                    nombre_oficial = nom_i
                    doc_oficial = doc_i
                    break

        nombre_oficial_norm = _normalizar_texto(nombre_oficial)
        ws_reg = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
        todas_filas = self._leer_filas_diccionario(ws_reg, HEADERS_REGISTRO)

        registros_periodo: List[Dict[str, Any]] = []
        total_clases = 0
        clases_por_cat: Dict[str, int] = {}
        clases_por_instructor: Dict[str, int] = {}
        en_tiempo_count = 0
        fuera_tiempo_count = 0
        clases_unicas_vistas: set = set()

        for fila in todas_filas:
            id_alu_f = _normalizar_texto(str(fila.get("ID_Alumno", "")))
            nom_alu_f = _normalizar_texto(
                limpiar_solo_nombre_persona(str(fila.get("Nombre_Alumno", "")))
                or str(fila.get("Nombre_Alumno", ""))
            )
            cat_f = normalizar_categoria(str(fila.get("Categoria", "B1")))
            raw_clases_f = str(fila.get("Clases_Reportadas", "")).strip()
            es_ref_o_adic = (
                "refuerzo" in raw_clases_f.lower()
                or "adicional" in raw_clases_f.lower()
            )
            nums_f = tuple(sorted(extraer_numeros_clase(raw_clases_f)))
            clave_alu_f = id_alu_f or nom_alu_f
            if clave_alu_f and nums_f and not es_ref_o_adic:
                clave_dup = (clave_alu_f, cat_f, nums_f)
                if clave_dup in clases_unicas_vistas:
                    # Omitir registro duplicado de la misma clase para el mismo alumno
                    continue
                clases_unicas_vistas.add(clave_dup)

            fecha_fila = str(fila.get("Fecha_Reporte", "")).strip()
            if not fecha_fila and fila.get("Timestamp"):
                fecha_fila = str(fila.get("Timestamp", "")).strip()[:10]
            if not (fecha_ini <= fecha_fila <= fecha_fin):
                continue

            inst_fila = str(fila.get("Instructor", "")).strip() or "Sin nombre"
            inst_fila_norm = _normalizar_texto(inst_fila)

            if not es_todos:
                coincide = (
                    inst_fila_norm == nombre_oficial_norm
                    or (criterio_norm and criterio_norm in inst_fila_norm)
                    or (nombre_oficial_norm and nombre_oficial_norm in inst_fila_norm)
                )
                if not coincide:
                    continue

            raw_clases = str(fila.get("Clases_Reportadas", "")).strip()
            m_adic = re.search(r"adicional\s*\((\d+)\s*cl", raw_clases, flags=re.IGNORECASE)
            if m_adic:
                cant = int(m_adic.group(1))
            else:
                nums = extraer_numeros_clase(raw_clases)
                cant = len(nums) if nums else 1
            cat = normalizar_categoria(str(fila.get("Categoria", "B1")))
            val_hor = str(fila.get("Validacion_Horario", "EN TIEMPO")).strip()

            if "FUERA" in val_hor.upper():
                fuera_tiempo_count += 1
            else:
                en_tiempo_count += 1

            total_clases += cant
            clases_por_cat[cat] = clases_por_cat.get(cat, 0) + cant
            clases_por_instructor[inst_fila] = clases_por_instructor.get(inst_fila, 0) + cant

            item_enriquecido = dict(fila)
            item_enriquecido["Cantidad_Clases_Calc"] = cant
            item_enriquecido["Nombre_Alumno"] = (
                limpiar_solo_nombre_persona(str(fila.get("Nombre_Alumno", "")))
                or str(fila.get("Nombre_Alumno", ""))
            )
            registros_periodo.append(item_enriquecido)

        return {
            "es_todos": es_todos,
            "instructor_nombre": "TODOS LOS INSTRUCTORES" if es_todos else nombre_oficial,
            "instructor_documento": doc_oficial,
            "periodo_codigo": periodo,
            "periodo_etiqueta": etiqueta_periodo,
            "fecha_inicio": fecha_ini,
            "fecha_fin": fecha_fin,
            "total_clases": total_clases,
            "total_reportes": len(registros_periodo),
            "clases_por_categoria": clases_por_cat,
            "clases_por_instructor": clases_por_instructor,
            "en_tiempo": en_tiempo_count,
            "fuera_tiempo": fuera_tiempo_count,
            "registros": registros_periodo,
        }

    def consultar_cierres_jornada(self, periodo: str = "semana") -> Dict[str, Any]:
        """Consulta los reportes de Cierre de Jornada (Agenda práctica) enviados por los instructores."""
        fecha_ini, fecha_fin, etiqueta_periodo = self.obtener_rango_fechas_periodo(periodo)
        cierres_periodo: List[Dict[str, Any]] = []
        try:
            ws_cie = self.obtener_o_crear_hoja(SHEET_CIERRES, HEADERS_CIERRES)
            filas = self._leer_filas_diccionario(ws_cie, HEADERS_CIERRES)
            for f in filas:
                f_rep = str(f.get("Fecha_Jornada", "")).strip()
                if not f_rep and f.get("Timestamp"):
                    f_rep = str(f.get("Timestamp", "")).strip()[:10]
                if fecha_ini <= f_rep <= fecha_fin:
                    cierres_periodo.append(f)
        except Exception as exc:
            logger.warning("Error consultando cierres de jornada: %s", exc)
        return {
            "periodo_codigo": periodo,
            "periodo_etiqueta": etiqueta_periodo,
            "cierres": list(reversed(cierres_periodo)),
        }

    @staticmethod
    def _normalizar_fecha_iso(fecha_raw: str, timestamp_raw: str = "") -> str:
        """
        Convierte fechas en formato DD/MM/YYYY, YYYY-MM-DD, DD-MM-YYYY o nombres de mes
        ('Julio', 'Agosto', 'Septiembre') a formato comparable YYYY-MM-DD.
        """
        txt = str(fecha_raw or "").strip()
        if not txt and timestamp_raw:
            txt = str(timestamp_raw).strip()

        if not txt:
            return ""

        # Caso YYYY-MM-DD
        m_iso = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})", txt)
        if m_iso:
            y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
            return f"{y:04d}-{m:02d}-{d:02d}"

        # Caso DD/MM/YYYY o DD-MM-YYYY
        m_lat = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})", txt)
        if m_lat:
            d, m, y = int(m_lat.group(1)), int(m_lat.group(2)), int(m_lat.group(3))
            return f"{y:04d}-{m:02d}-{d:02d}"

        # Caso nombre de mes (ej. 'Julio', 'Agosto', 'Septiembre')
        meses_map = {
            "enero": "01",
            "febrero": "02",
            "marzo": "03",
            "abril": "04",
            "mayo": "05",
            "junio": "06",
            "julio": "07",
            "agosto": "08",
            "septiembre": "09",
            "octubre": "10",
            "noviembre": "11",
            "diciembre": "12",
        }
        t_low = _normalizar_texto(txt)
        for m_nom, m_cod in meses_map.items():
            if t_low.startswith(m_nom):
                m_anio = re.search(r"\b(20\d{2})\b", t_low)
                if m_anio:
                    anio_uso = m_anio.group(1)
                elif timestamp_raw and re.match(r"^20\d{2}", str(timestamp_raw).strip()):
                    anio_uso = str(timestamp_raw).strip()[:4]
                else:
                    anio_uso = str(datetime.now(TZ_COLOMBIA).year)
                return f"{anio_uso}-{m_cod}-15"

        if timestamp_raw and timestamp_raw != fecha_raw:
            return SheetsService._normalizar_fecha_iso(timestamp_raw, "")
        return ""

    @staticmethod
    def _limpiar_numero_entero(val: Any, default: int = 0) -> int:
        s = re.sub(r"[^\d]", "", str(val or ""))
        return int(s) if s else default

    def registrar_inasistencia(
        self,
        documento_alumno: str,
        nombre_alumno: str,
        categoria: str,
        clases_faltadas: str,
        nombre_instructor: str,
        no_horas: int = 2,
        valor_hora: int = 25000,
        estado: str = "Pendiente Pago",
        fecha: str = "",
        hora: str = "",
    ) -> Dict[str, Any]:
        """
        Registra una nueva inasistencia de alumno en la hoja 'Inasistencias'
        manteniendo compatibilidad con las columnas de AppSheet.
        """
        ahora = datetime.now(TZ_COLOMBIA)
        fecha_uso = fecha or ahora.strftime("%d/%m/%Y")
        hora_uso = hora or ahora.strftime("%H:%M:%S")
        id_nov = uuid.uuid4().hex[:8]

        ws = self.obtener_o_crear_hoja(SHEET_INASISTENCIAS, HEADERS_INASISTENCIAS)
        filas = self._leer_filas_diccionario(ws, HEADERS_INASISTENCIAS)
        max_consec = 0
        for f in filas:
            c_val = self._limpiar_numero_entero(f.get("Consecutivo", "0"), 0)
            if c_val > max_consec:
                max_consec = c_val
        nuevo_consec = max_consec + 1

        nom_limpio = limpiar_solo_nombre_persona(nombre_alumno) or nombre_alumno.strip().title()
        doc_limpio = str(documento_alumno or "").strip()
        cat_limpia = normalizar_categoria(categoria) if categoria else ""

        # Si falta el documento o categoría, intentar enriquecer desde la hoja Alumnos
        if not doc_limpio or not cat_limpia:
            ficha = self.buscar_alumno(doc_limpio or nom_limpio, categoria=cat_limpia or None)
            if ficha:
                if not doc_limpio:
                    doc_limpio = str(ficha.get("Documento", "")).strip()
                if not cat_limpia:
                    cat_limpia = normalizar_categoria(str(ficha.get("Categoria", "B1")))
                if str(ficha.get("Nombre_Completo", "")).strip():
                    nom_limpio = str(ficha.get("Nombre_Completo", "")).strip()

        clases_str = str(clases_faltadas or "No especificada").strip()
        fila_nueva = [
            id_nov,
            str(nuevo_consec),
            fecha_uso,
            hora_uso,
            doc_limpio,
            nom_limpio,
            cat_limpia,
            clases_str,
            nombre_instructor.strip(),
            str(no_horas),
            str(valor_hora),
            estado.strip(),
        ]
        self._agregar_fila_al_final(ws, fila_nueva)
        return {
            "ID_Novedad": id_nov,
            "Consecutivo": str(nuevo_consec),
            "Fecha": fecha_uso,
            "Hora": hora_uso,
            "Documento_Alumno": doc_limpio,
            "Nombre_Alumno": nom_limpio,
            "Categoria": cat_limpia,
            "Clases_Faltadas": clases_str,
            "Nombre_Instructor": nombre_instructor.strip(),
            "No_Horas": int(no_horas),
            "Valor_Hora": int(valor_hora),
            "Valor_Total": int(no_horas) * int(valor_hora),
            "Estado": estado.strip(),
        }

    def actualizar_estado_inasistencia(
        self, consecutivo_o_id: str, nuevo_estado: str
    ) -> Optional[Dict[str, Any]]:
        """
        Actualiza el Estado de una inasistencia ('Pendiente Pago', 'Cancelo Efectivo', 'Cancelo Trasferencia')
        buscando por Consecutivo o ID_Novedad.
        """
        crit = str(consecutivo_o_id or "").strip().lower()
        if not crit:
            return None
        ws = self.obtener_o_crear_hoja(SHEET_INASISTENCIAS, HEADERS_INASISTENCIAS)
        todas = ws.get_all_values()
        for idx_fila in range(2, len(todas) + 1):
            fila = todas[idx_fila - 1]
            if not any(str(c).strip() for c in fila):
                continue
            id_f = str(fila[0]).strip().lower() if len(fila) > 0 else ""
            con_f = str(fila[1]).strip().lower() if len(fila) > 1 else ""
            if crit in (id_f, con_f):
                ws.update_cell(idx_fila, 12, nuevo_estado)
                horas_i = self._limpiar_numero_entero(fila[9] if len(fila) > 9 else 1, 1)
                val_i = self._limpiar_numero_entero(fila[10] if len(fila) > 10 else 0, 0)
                return {
                    "ID_Novedad": fila[0] if len(fila) > 0 else "",
                    "Consecutivo": fila[1] if len(fila) > 1 else "",
                    "Fecha": fila[2] if len(fila) > 2 else "",
                    "Hora": fila[3] if len(fila) > 3 else "",
                    "Documento_Alumno": fila[4] if len(fila) > 4 else "",
                    "Nombre_Alumno": fila[5] if len(fila) > 5 else "",
                    "Categoria": fila[6] if len(fila) > 6 else "",
                    "Clases_Faltadas": fila[7] if len(fila) > 7 else "",
                    "Nombre_Instructor": fila[8] if len(fila) > 8 else "",
                    "No_Horas": horas_i,
                    "Valor_Hora": val_i,
                    "Valor_Total": horas_i * val_i,
                    "Estado": nuevo_estado,
                }
        return None

    def obtener_inasistencias_alumno(
        self,
        criterio_alumno: str,
        nombre_alternativo: str = "",
        categoria: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """
        Obtiene todas las inasistencias registradas para un alumno por Cédula o Nombre.
        Se usa tanto para el Reporte PDF Individual de Inasistencias como para inyectar
        automáticamente las fallas en OBSERVACIONES del Formato Oficial de Clases (F-59/60/61/62).
        """
        crit_norm = _normalizar_texto(criterio_alumno)
        nom_alt_norm = _normalizar_texto(nombre_alternativo)
        if not crit_norm and not nom_alt_norm:
            return []

        # Si solo nos pasaron cédula, buscar también su nombre en Alumnos para cruzar con registros históricos de AppSheet
        if crit_norm.isdigit() and not nom_alt_norm:
            ficha = self.buscar_alumno(criterio_alumno, categoria=categoria)
            if ficha:
                nom_alt_norm = _normalizar_texto(str(ficha.get("Nombre_Completo", "")))

        cat_buscada = normalizar_categoria(categoria) if categoria else None
        resultado: List[Dict[str, Any]] = []
        try:
            self.sembrar_datos_historicos_appsheet()
            ws = self.obtener_o_crear_hoja(SHEET_INASISTENCIAS, HEADERS_INASISTENCIAS)
            filas = self._leer_filas_diccionario(ws, HEADERS_INASISTENCIAS)
            for f in filas:
                doc_f = _normalizar_texto(str(f.get("Documento_Alumno", "")))
                nom_f = _normalizar_texto(str(f.get("Nombre_Alumno", "")))
                cat_f = str(f.get("Categoria", "")).strip()
                cat_f_norm = normalizar_categoria(cat_f) if cat_f else ""

                coincide = False
                if crit_norm and doc_f and crit_norm == doc_f:
                    coincide = True
                elif crit_norm and nom_f and (crit_norm == nom_f or (len(crit_norm) >= 4 and crit_norm in nom_f)):
                    coincide = True
                elif nom_alt_norm and nom_f and (nom_alt_norm == nom_f or nom_alt_norm in nom_f or nom_f in nom_alt_norm):
                    coincide = True

                if not coincide:
                    continue

                if cat_buscada and cat_f_norm and cat_f_norm != cat_buscada:
                    continue

                horas_i = max(1, self._limpiar_numero_entero(f.get("No_Horas", "1"), 1))
                val_i = self._limpiar_numero_entero(f.get("Valor_Hora", "0"), 0)
                item = dict(f)
                item["No_Horas_Int"] = horas_i
                item["Valor_Hora_Int"] = val_i
                item["Valor_Total_Int"] = horas_i * val_i
                item["Fecha_ISO"] = self._normalizar_fecha_iso(str(f.get("Fecha", "")))
                resultado.append(item)
        except Exception as exc:
            logger.warning("Error obteniendo inasistencias del alumno %s: %s", criterio_alumno, exc)
        return resultado

    def consultar_inasistencias_periodo(
        self,
        periodo: str = "mes",
        filtro_estado: str = "TODOS",
        criterio_alumno: str = "",
    ) -> Dict[str, Any]:
        """
        Consulta las inasistencias filtradas por período ('dia', 'semana', 'semana_ant',
        'mes', 'mes_ant', 'anio', 'anio_2025', 'todo') y opcionalmente por estado o alumno.
        """
        self.sembrar_datos_historicos_appsheet()
        fecha_ini, fecha_fin, etiqueta_periodo = self.obtener_rango_fechas_periodo(periodo)
        ws = self.obtener_o_crear_hoja(SHEET_INASISTENCIAS, HEADERS_INASISTENCIAS)
        filas = self._leer_filas_diccionario(ws, HEADERS_INASISTENCIAS)

        crit_alu_norm = _normalizar_texto(criterio_alumno)
        est_filtro_norm = _normalizar_texto(filtro_estado)

        registros_filtrados: List[Dict[str, Any]] = []
        total_horas = 0
        total_valor = 0
        total_efectivo = 0
        total_transferencia = 0
        total_pendiente = 0
        casos_pendientes = 0

        for f in filas:
            fecha_raw = str(f.get("Fecha", "")).strip()
            fecha_iso = self._normalizar_fecha_iso(fecha_raw)
            if periodo.lower() not in ("todo", "historico", "histórico", "todos"):
                if not fecha_iso or not (fecha_ini <= fecha_iso <= fecha_fin):
                    continue

            if crit_alu_norm:
                doc_f = _normalizar_texto(str(f.get("Documento_Alumno", "")))
                nom_f = _normalizar_texto(str(f.get("Nombre_Alumno", "")))
                if crit_alu_norm not in doc_f and crit_alu_norm not in nom_f:
                    continue

            est_raw = str(f.get("Estado", "Pendiente Pago")).strip()
            est_norm = _normalizar_texto(est_raw)
            if est_filtro_norm and est_filtro_norm not in ("todos", "todo"):
                if "pendiente" in est_filtro_norm and "pendiente" not in est_norm:
                    continue
                if "efectivo" in est_filtro_norm and "efectivo" not in est_norm:
                    continue
                if ("transferencia" in est_filtro_norm or "trasferencia" in est_filtro_norm) and not (
                    "transferencia" in est_norm or "trasferencia" in est_norm
                ):
                    continue

            horas_i = max(1, self._limpiar_numero_entero(f.get("No_Horas", "1"), 1))
            val_i = self._limpiar_numero_entero(f.get("Valor_Hora", "0"), 0)
            subtotal = horas_i * val_i

            total_horas += horas_i
            total_valor += subtotal
            if "efectivo" in est_norm:
                total_efectivo += subtotal
            elif "transferencia" in est_norm or "trasferencia" in est_norm:
                total_transferencia += subtotal
            else:
                total_pendiente += subtotal
                casos_pendientes += 1

            item = dict(f)
            item["No_Horas_Int"] = horas_i
            item["Valor_Hora_Int"] = val_i
            item["Valor_Total_Int"] = subtotal
            item["Fecha_ISO"] = fecha_iso
            registros_filtrados.append(item)

        return {
            "periodo_codigo": periodo,
            "periodo_etiqueta": etiqueta_periodo,
            "fecha_inicio": fecha_ini,
            "fecha_fin": fecha_fin,
            "filtro_estado": filtro_estado,
            "total_registros": len(registros_filtrados),
            "total_horas": total_horas,
            "total_valor": total_valor,
            "total_efectivo": total_efectivo,
            "total_transferencia": total_transferencia,
            "total_pendiente": total_pendiente,
            "casos_pendientes": casos_pendientes,
            "registros": registros_filtrados,
        }

    def registrar_mantenimiento(
        self,
        placa: str,
        tipo_mantenimiento: str,
        establecimiento: str,
        nit_cc: str,
        actividad_realizada: str,
        nombre_tecnico: str,
        rol: str,
        observaciones: str,
        kilometraje_actual: str = "0",
        costo: str = "0",
        proximo_cambio: str = "0",
        nombre_instructor: str = "",
        correo: str = "autoorientecrm@gmail.com",
        fecha: str = "",
    ) -> Dict[str, Any]:
        """
        Registra un mantenimiento de vehículo (Preventivo / Correctivo) en la hoja 'Mantenimientos'
        con todos los campos del formulario de AppSheet y del formato oficial SISI-F2-08 (PESV).
        """
        ahora = datetime.now(TZ_COLOMBIA)
        ts_str = ahora.strftime("%Y-%m-%d %H:%M:%S")
        fecha_uso = fecha.strip() if fecha else ahora.strftime("%Y-%m-%d")
        placa_up = placa.strip().upper()
        tipo_clean = tipo_mantenimiento.strip().title() or "Preventivo"
        rol_up = rol.strip().upper() or "MECANICO"
        km_clean = str(self._limpiar_numero_entero(kilometraje_actual, 0))
        costo_clean = str(self._limpiar_numero_entero(costo, 0))
        prox_clean = str(self._limpiar_numero_entero(proximo_cambio, 0))

        ws = self.obtener_o_crear_hoja(SHEET_MANTENIMIENTOS, HEADERS_MANTENIMIENTOS)
        fila_nueva = [
            ts_str,
            fecha_uso,
            correo.strip() or "autoorientecrm@gmail.com",
            nombre_instructor.strip(),
            placa_up,
            tipo_clean,
            establecimiento.strip(),
            nit_cc.strip() or "0",
            actividad_realizada.strip() or "Mantenimiento",
            nombre_tecnico.strip() or establecimiento.strip(),
            rol_up,
            observaciones.strip(),
            km_clean,
            costo_clean,
            prox_clean,
        ]
        self._agregar_fila_al_final(ws, fila_nueva)
        return {
            "Timestamp": ts_str,
            "Fecha": fecha_uso,
            "Correo": correo.strip() or "autoorientecrm@gmail.com",
            "Nombre_Instructor": nombre_instructor.strip(),
            "Placa": placa_up,
            "Tipo_Mantenimiento": tipo_clean,
            "Establecimiento": establecimiento.strip(),
            "NIT_CC": nit_cc.strip() or "0",
            "Actividad_Realizada": actividad_realizada.strip() or "Mantenimiento",
            "Nombre_Tecnico": nombre_tecnico.strip() or establecimiento.strip(),
            "Rol": rol_up,
            "Observaciones": observaciones.strip(),
            "Kilometraje_Actual": int(km_clean),
            "Costo": int(costo_clean),
            "Proximo_Cambio": int(prox_clean),
        }

    def consultar_mantenimientos_periodo(
        self, periodo: str = "todo", placa: str = "TODAS"
    ) -> Dict[str, Any]:
        """
        Consulta los registros de mantenimiento de vehículos por Semana, Mes, Mes específico
        (ej. 'julio', 'agosto', 'septiembre'), Año ('anio', '2025', '2026') o Todo el histórico,
        y opcionalmente filtra por Placa.
        """
        self.sembrar_datos_historicos_appsheet()
        fecha_ini, fecha_fin, etiqueta_periodo = self.obtener_rango_fechas_periodo(periodo)
        ws = self.obtener_o_crear_hoja(SHEET_MANTENIMIENTOS, HEADERS_MANTENIMIENTOS)
        filas = self._leer_filas_diccionario(ws, HEADERS_MANTENIMIENTOS)

        placa_filtro = placa.strip().upper() if placa else "TODAS"
        per_low = (periodo or "todo").strip().lower()
        meses_nombres = (
            "enero", "febrero", "marzo", "abril", "mayo", "junio",
            "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
        )

        registros_filtrados: List[Dict[str, Any]] = []
        total_costo = 0
        por_placa: Dict[str, int] = {}
        preventivos = 0
        correctivos = 0

        for f in filas:
            pl_f = str(f.get("Placa", "")).strip().upper()
            if not pl_f:
                continue
            if placa_filtro not in ("TODAS", "TODOS", "") and pl_f != placa_filtro:
                continue

            fecha_raw = str(f.get("Fecha", "")).strip()
            ts_raw = str(f.get("Timestamp", "")).strip()
            fecha_iso = self._normalizar_fecha_iso(fecha_raw, ts_raw)

            if per_low in meses_nombres:
                # Si filtra por nombre de mes (ej. 'julio', 'agosto', 'septiembre'), coincidir por nombre de mes o mes numérico
                mes_idx = meses_nombres.index(per_low) + 1
                es_mes = (
                    per_low in _normalizar_texto(fecha_raw)
                    or (fecha_iso and len(fecha_iso) >= 7 and fecha_iso[5:7] == f"{mes_idx:02d}")
                )
                if not es_mes:
                    continue
            elif per_low not in ("todo", "historico", "histórico", "todos"):
                if not fecha_iso or not (fecha_ini <= fecha_iso <= fecha_fin):
                    continue

            costo_i = self._limpiar_numero_entero(f.get("Costo", "0"), 0)
            km_i = self._limpiar_numero_entero(f.get("Kilometraje_Actual", "0"), 0)
            prox_i = self._limpiar_numero_entero(f.get("Proximo_Cambio", "0"), 0)
            tipo_m = str(f.get("Tipo_Mantenimiento", "Preventivo")).strip()

            if "correctivo" in tipo_m.lower():
                correctivos += 1
            else:
                preventivos += 1

            total_costo += costo_i
            por_placa[pl_f] = por_placa.get(pl_f, 0) + 1

            item = dict(f)
            item["Costo_Int"] = costo_i
            item["Kilometraje_Int"] = km_i
            item["Proximo_Cambio_Int"] = prox_i
            item["Fecha_ISO"] = fecha_iso
            registros_filtrados.append(item)

        return {
            "periodo_codigo": periodo,
            "periodo_etiqueta": etiqueta_periodo,
            "fecha_inicio": fecha_ini,
            "fecha_fin": fecha_fin,
            "placa_filtro": placa_filtro,
            "total_registros": len(registros_filtrados),
            "preventivos": preventivos,
            "correctivos": correctivos,
            "total_costo": total_costo,
            "por_placa": por_placa,
            "registros": registros_filtrados,
        }

    def listar_vehiculos_activos(self, forzar_recarga: bool = False) -> List[Dict[str, str]]:
        """
        Devuelve el listado dinámico de vehículos activos (Carros y Motos) desde la hoja 'Vehiculos_PESV'.
        Si la hoja aún no tiene vehículos registrados, siembra los 6 vehículos iniciales.
        """
        ahora_ts = time.time()
        cache_veh = getattr(self, "_cache_vehiculos", None)
        cache_ts = getattr(self, "_cache_vehiculos_ts", 0.0)
        if not forzar_recarga and cache_veh and (ahora_ts - cache_ts < 120):
            return list(cache_veh)

        try:
            self.sembrar_datos_historicos_appsheet()
            ws = self.obtener_o_crear_hoja(SHEET_VEHICULOS, HEADERS_VEHICULOS)
            filas = self._leer_filas_diccionario(ws, HEADERS_VEHICULOS)
            activos: List[Dict[str, str]] = []
            placas_vistas = set()
            for f in filas:
                pl = re.sub(r"[^A-Z0-9]", "", str(f.get("Placa", "")).strip().upper())
                if not pl or pl in placas_vistas:
                    continue
                est = str(f.get("Estado", "Activo")).strip().lower()
                if est in ("retirado", "eliminado", "inactivo", "baja"):
                    placas_vistas.add(pl)
                    continue
                tip_raw = str(f.get("Tipo_Vehiculo", "")).strip().lower()
                # Si la placa termina en letra (ej. YNE31F, AZA25G) o dice moto, es moto
                es_moto = "moto" in tip_raw or (not tip_raw and len(pl) == 6 and pl[-1].isalpha())
                icono = "🏍️" if es_moto else "🚗"
                tipo_lbl = f"{icono} {'Moto' if es_moto else 'Carro'}"
                activos.append(
                    {
                        "placa": pl,
                        "tipo": tipo_lbl,
                        "etiqueta": f"{icono} {pl}",
                    }
                )
                placas_vistas.add(pl)

            if activos:
                self._cache_vehiculos = activos
                self._cache_vehiculos_ts = ahora_ts
                return list(activos)
        except Exception as exc:
            logger.warning("No se pudo leer Vehiculos_PESV, usando lista por defecto: %s", exc)

        return list(PLACAS_VEHICULOS_OFICIALES)

    def agregar_vehiculo_flota(self, placa: str, tipo_vehiculo: str = "Carro") -> Dict[str, str]:
        """
        Registra un vehículo nuevo (o reactiva uno previamente retirado) en la hoja 'Vehiculos_PESV'.
        """
        pl_clean = re.sub(r"[^A-Z0-9]", "", (placa or "").strip().upper())
        es_moto = "moto" in (tipo_vehiculo or "").lower() or (
            not tipo_vehiculo and len(pl_clean) == 6 and pl_clean[-1].isalpha()
        )
        tipo_guardar = "Moto" if es_moto else "Carro"
        icono = "🏍️" if es_moto else "🚗"
        fecha_hoy = datetime.now(TZ_COLOMBIA).strftime("%Y-%m-%d")

        try:
            self.sembrar_datos_historicos_appsheet()
            ws = self.obtener_o_crear_hoja(SHEET_VEHICULOS, HEADERS_VEHICULOS)
            todas_filas = ws.get_all_values()
            fila_existente = None
            for idx_f in range(2, len(todas_filas) + 1):
                fila = todas_filas[idx_f - 1]
                pl_f = re.sub(r"[^A-Z0-9]", "", str(fila[0]).strip().upper()) if fila else ""
                if pl_f == pl_clean:
                    fila_existente = idx_f
                    break

            if fila_existente is not None:
                ws.update(
                    values=[[pl_clean, tipo_guardar, "Activo", fecha_hoy]],
                    range_name=f"A{fila_existente}:D{fila_existente}",
                    value_input_option="USER_ENTERED",
                )
            else:
                self._agregar_fila_al_final(ws, [pl_clean, tipo_guardar, "Activo", fecha_hoy])
        except Exception as exc:
            logger.warning("Error registrando vehículo %s en Vehiculos_PESV: %s", pl_clean, exc)

        self._cache_vehiculos = None
        self.listar_vehiculos_activos(forzar_recarga=True)
        return {
            "placa": pl_clean,
            "tipo": f"{icono} {tipo_guardar}",
            "etiqueta": f"{icono} {pl_clean}",
        }

    def eliminar_vehiculo_flota(self, placa: str) -> bool:
        """
        Retira / elimina un vehículo de la flota activa en la hoja 'Vehiculos_PESV'
        marcando su Estado como 'Retirado' para que no aparezca en los botones de placas activas.
        """
        pl_clean = re.sub(r"[^A-Z0-9]", "", (placa or "").strip().upper())
        if not pl_clean:
            return False
        modificado = False
        try:
            self.sembrar_datos_historicos_appsheet()
            ws = self.obtener_o_crear_hoja(SHEET_VEHICULOS, HEADERS_VEHICULOS)
            todas_filas = ws.get_all_values()
            for idx_f in range(2, len(todas_filas) + 1):
                fila = todas_filas[idx_f - 1]
                pl_f = re.sub(r"[^A-Z0-9]", "", str(fila[0]).strip().upper()) if fila else ""
                if pl_f == pl_clean:
                    ws.update_cell(idx_f, 3, "Retirado")
                    modificado = True
        except Exception as exc:
            logger.warning("Error retirando vehículo %s de Vehiculos_PESV: %s", pl_clean, exc)

        self._cache_vehiculos = None
        self.listar_vehiculos_activos(forzar_recarga=True)
        return modificado

    def obtener_panorama_web(
        self,
        nomina_periodo: str = "semana",
        nomina_instructor: str = "TODOS",
        ina_periodo: str = "todo",
        ina_estado: str = "TODOS",
        mant_periodo: str = "todo",
        mant_placa: str = "TODAS",
        hallazgos_cartera: Optional[Dict[str, Any]] = None,
        periodo_nomina: Optional[str] = None,
        instructor_nomina: Optional[str] = None,
        periodo_inasistencias: Optional[str] = None,
        estado_inasistencias: Optional[str] = None,
        periodo_mantenimientos: Optional[str] = None,
        placa_mantenimientos: Optional[str] = None,
    ) -> Dict[str, Any]:
        """
        Consolida en una sola respuesta todos los módulos del Bot de Clases (Google Sheets)
        y los cruza en tiempo real con la Auditoría de Cartera (Excel) por Cédula y Nombre.
        """
        if periodo_nomina is not None:
            nomina_periodo = periodo_nomina
        if instructor_nomina is not None:
            nomina_instructor = instructor_nomina
        if periodo_inasistencias is not None:
            ina_periodo = periodo_inasistencias
        if estado_inasistencias is not None:
            ina_estado = estado_inasistencias
        if periodo_mantenimientos is not None:
            mant_periodo = periodo_mantenimientos
        if placa_mantenimientos is not None:
            mant_placa = placa_mantenimientos

        if mant_periodo.startswith("mes_") and len(mant_periodo) == 6 and mant_periodo[4:].isdigit():
            meses_idx_map = {
                "01": "enero", "02": "febrero", "03": "marzo", "04": "abril",
                "05": "mayo", "06": "junio", "07": "julio", "08": "agosto",
                "09": "septiembre", "10": "octubre", "11": "noviembre", "12": "diciembre",
            }
            mant_periodo = meses_idx_map.get(mant_periodo[4:], mant_periodo)

        self.sembrar_datos_historicos_appsheet()
        alumnos_todos = self.listar_alumnos_activos(modo="todos")

        # Incluir también alumnos que tengan clases en Registro_Clases pero aún no figuren en hoja Alumnos
        # y construir el mapa de primera y última fecha de clase práctica por (Alumno, Categoría) para el Semáforo de 90 Días
        vistos_alu_cat = {
            (
                _normalizar_texto(str(a.get("Documento", ""))) or _normalizar_texto(str(a.get("Nombre_Completo", ""))),
                normalizar_categoria(str(a.get("Categoria", "B1"))),
            )
            for a in alumnos_todos
        }
        primera_clase_por_doc: Dict[Tuple[str, str], str] = {}
        primera_clase_por_nom: Dict[Tuple[str, str], str] = {}
        ultima_clase_por_doc: Dict[Tuple[str, str], str] = {}
        ultima_clase_por_nom: Dict[Tuple[str, str], str] = {}
        try:
            por_doc, por_nom = self._calcular_mapa_clases_por_alumno()
            ws_reg = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
            filas_reg = self._leer_filas_diccionario(ws_reg, HEADERS_REGISTRO)
            for r_f in filas_reg:
                if _es_registro_inasistencia(r_f):
                    continue
                f_iso_r = self._normalizar_fecha_iso(str(r_f.get("Fecha_Reporte", "")), str(r_f.get("Timestamp", "")))
                if not f_iso_r or len(f_iso_r) < 10:
                    continue
                f_iso_10 = f_iso_r[:10]
                doc_rf = _normalizar_texto(str(r_f.get("ID_Alumno", "")))
                nom_rf = _normalizar_texto(
                    limpiar_solo_nombre_persona(str(r_f.get("Nombre_Alumno", "")))
                    or str(r_f.get("Nombre_Alumno", ""))
                )
                cat_rf = normalizar_categoria(str(r_f.get("Categoria", "B1")))
                if doc_rf and doc_rf not in ("pendiente", "desconocido", "s/n"):
                    kd = (doc_rf, cat_rf)
                    if kd not in primera_clase_por_doc or f_iso_10 < primera_clase_por_doc[kd]:
                        primera_clase_por_doc[kd] = f_iso_10
                    if kd not in ultima_clase_por_doc or f_iso_10 > ultima_clase_por_doc[kd]:
                        ultima_clase_por_doc[kd] = f_iso_10
                if nom_rf:
                    kn = (nom_rf, cat_rf)
                    if kn not in primera_clase_por_nom or f_iso_10 < primera_clase_por_nom[kn]:
                        primera_clase_por_nom[kn] = f_iso_10
                    if kn not in ultima_clase_por_nom or f_iso_10 > ultima_clase_por_nom[kn]:
                        ultima_clase_por_nom[kn] = f_iso_10

            for r in reversed(filas_reg):
                if _es_registro_inasistencia(r):
                    continue
                doc_r = str(r.get("ID_Alumno", "")).strip()
                nom_r = limpiar_solo_nombre_persona(str(r.get("Nombre_Alumno", ""))) or str(r.get("Nombre_Alumno", "")).strip()
                if not nom_r:
                    continue
                cat_r = normalizar_categoria(str(r.get("Categoria", "B1")))
                clave = (_normalizar_texto(doc_r) or _normalizar_texto(nom_r), cat_r)
                if clave in vistos_alu_cat:
                    continue
                vistos_alu_cat.add(clave)
                cfg = FORMATOS_OFICIALES.get(cat_r, FORMATOS_OFICIALES["B1"])
                total_req = cfg.get("total_clases", 20)
                clases_v = set()
                if _normalizar_texto(doc_r) and (_normalizar_texto(doc_r), cat_r) in por_doc:
                    clases_v.update(por_doc[(_normalizar_texto(doc_r), cat_r)])
                if _normalizar_texto(nom_r) and (_normalizar_texto(nom_r), cat_r) in por_nom:
                    clases_v.update(por_nom[(_normalizar_texto(nom_r), cat_r)])
                clases_val = {c for c in clases_v if 1 <= c <= total_req}
                cant_c = len(clases_val)
                ya_fin = cant_c >= total_req
                alumnos_todos.append({
                    "ID_Matricula": doc_r or "REG-BOT",
                    "Documento": doc_r,
                    "Nombre_Completo": nom_r,
                    "Categoria": cat_r,
                    "Estado": "Finalizado" if ya_fin else "En curso",
                    "Estado_Real": "Finalizado" if ya_fin else "En curso",
                    "Estado_RUNT": "No Registrado RUNT",
                    "Fecha_RUNT": "",
                    "Clases_Completadas": cant_c,
                    "Clases_Requeridas": total_req,
                    "Curso_Completado": ya_fin,
                    "Clases_Vistas_Lista": sorted(clases_val),
                })
        except Exception as exc:
            logger.warning("Aviso cruzando Registro_Clases en obtener_panorama_web: %s", exc)

        # Mapa de cartera por Cédula, por (Cédula, Categoría) y por Nombre para cruce instantáneo (incluyendo Paz y Salvo de todas_matriculas)
        mapa_cartera_doc: Dict[str, Dict[str, Any]] = {}
        mapa_cartera_doc_cat: Dict[Tuple[str, str], Dict[str, Any]] = {}
        mapa_cartera_nom: Dict[str, Dict[str, Any]] = {}
        mapa_cartera_nom_cat: Dict[Tuple[str, str], Dict[str, Any]] = {}
        lista_cartera_tokens: List[Tuple[set, Dict[str, Any]]] = []
        registros_cartera_unificados: List[Dict[str, Any]] = []
        if hallazgos_cartera and isinstance(hallazgos_cartera, dict):
            for key_lista in (
                "todas_matriculas",
                "registros",
                "jt_cartera",
                "abono_cero_real",
                "abono_cero_multicurso",
                "cartera_critica_90",
                "mora_media_60",
                "mora_temprana_30",
            ):
                sub_l = hallazgos_cartera.get(key_lista)
                if isinstance(sub_l, list):
                    registros_cartera_unificados.extend(sub_l)

            for reg_c in registros_cartera_unificados:
                doc_c = _normalizar_texto(str(reg_c.get("NUMERODOCUMENTO", "")))
                nom_c = _normalizar_texto(str(reg_c.get("ALUMNO", "")))
                cat_c = normalizar_categoria(str(reg_c.get("CATEGORIA", "B1")))
                saldo_cli = max(0.0, float(reg_c.get("SALDO_REAL_CLIENTE", reg_c.get("VALORPENDIENTE", 0)) or 0))
                saldo_comp = max(0.0, float(reg_c.get("SALDO_COMPENSADO", reg_c.get("VALORPENDIENTE", saldo_cli)) or 0))
                abonado_cli = float(reg_c.get("TOTAL_ABONADO_CLIENTE", reg_c.get("ABONOS", 0)) or 0)
                abonado_mat = float(reg_c.get("ABONO_COMPENSADO", reg_c.get("ABONOS", 0)) or 0)
                contratado_cli = float(reg_c.get("TOTAL_CONTRATADO_CLIENTE", reg_c.get("VALORCURSO", 0)) or 0)
                contratado_mat = float(reg_c.get("VALORCURSO", 0) or 0)
                info_c = {
                    "saldo_cliente": saldo_cli,
                    "saldo_curso": saldo_comp if saldo_cli > 0 else 0.0,
                    "abonado_cliente": abonado_cli,
                    "abonado_curso": abonado_mat,
                    "contratado_cliente": contratado_cli,
                    "contratado_curso": contratado_mat,
                    "es_multicurso": bool(reg_c.get("ES_MULTICURSO", False)),
                    "categorias_cliente": str(reg_c.get("CATEGORIAS_CLIENTE", cat_c)),
                    "convenio": str(reg_c.get("CONVENIO_NORM", "")),
                    "dias_mora": int(reg_c.get("DIAS_ANTIGUEDAD", 0) or 0),
                    "riesgo_cero": str(reg_c.get("TIPO_RIESGO_CERO", "")),
                    "fecha_inscripcion": str(reg_c.get("FECHAINSCRIPCION", "")).strip(),
                }
                if doc_c:
                    mapa_cartera_doc[doc_c] = info_c
                    mapa_cartera_doc_cat[(doc_c, cat_c)] = info_c
                if nom_c:
                    mapa_cartera_nom[nom_c] = info_c
                    mapa_cartera_nom_cat[(nom_c, cat_c)] = info_c
                    toks = {t for t in nom_c.split() if len(t) >= 3}
                    if len(toks) >= 2:
                        lista_cartera_tokens.append((toks, info_c))

        if not alumnos_todos and not self.conectado_en_vivo and registros_cartera_unificados:
            vistos_seed_alu = set()
            for reg_c in registros_cartera_unificados:
                doc_s = str(reg_c.get("NUMERODOCUMENTO", "")).strip()
                nom_s = str(reg_c.get("ALUMNO", "")).strip()
                cat_c = normalizar_categoria(str(reg_c.get("CATEGORIA", "B1")))
                k_seed = (doc_s or nom_s, cat_c)
                if not nom_s or k_seed in vistos_seed_alu:
                    continue
                vistos_seed_alu.add(k_seed)
                cfg_c = FORMATOS_OFICIALES.get(cat_c, FORMATOS_OFICIALES["B1"])
                tot_req = cfg_c.get("total_clases", 20)
                alumnos_todos.append({
                    "ID_Matricula": str(reg_c.get("NUMEROMATRICULA", "")),
                    "Documento": doc_s,
                    "Nombre_Completo": nom_s,
                    "Categoria": cat_c,
                    "Estado": "En curso",
                    "Estado_Real": "En curso",
                    "Estado_RUNT": "No Registrado RUNT",
                    "Fecha_RUNT": "",
                    "Clases_Completadas": 0,
                    "Clases_Requeridas": tot_req,
                    "Curso_Completado": False,
                    "Clases_Vistas_Lista": [],
                })
                if len(alumnos_todos) >= 120:
                    break

        alertas_cruce: List[Dict[str, Any]] = []
        alumnos_activos_cnt = 0
        alumnos_finalizados_cnt = 0

        for alu in alumnos_todos:
            if alu.get("Curso_Completado"):
                alumnos_finalizados_cnt += 1
            else:
                alumnos_activos_cnt += 1

            doc_a = _normalizar_texto(str(alu.get("Documento", "")))
            nom_a = _normalizar_texto(str(alu.get("Nombre_Completo", "")))
            cat_a_norm = normalizar_categoria(str(alu.get("Categoria", "B1")))
            c_info = (
                mapa_cartera_doc_cat.get((doc_a, cat_a_norm))
                or mapa_cartera_nom_cat.get((nom_a, cat_a_norm))
                or mapa_cartera_doc.get(doc_a)
                or mapa_cartera_nom.get(nom_a)
            )
            if not c_info and nom_a:
                toks_a = {t for t in nom_a.split() if len(t) >= 3 and t not in ("teoria", "clase", "practica")}
                if len(toks_a) >= 2:
                    for toks_c, cand_info in lista_cartera_tokens:
                        if toks_a.issubset(toks_c) or toks_c.issubset(toks_a):
                            c_info = cand_info
                            break
            if c_info:
                alu["Cartera_Cruzada"] = True
                alu["Saldo_Cartera"] = c_info["saldo_cliente"]
                alu["Saldo_Curso_Cartera"] = c_info.get("saldo_curso", c_info["saldo_cliente"])
                alu["Abonado_Cartera"] = c_info["abonado_cliente"]
                alu["Abonado_Curso_Cartera"] = c_info.get("abonado_curso", c_info["abonado_cliente"])
                alu["Contratado_Cartera"] = c_info["contratado_cliente"]
                alu["Contratado_Curso_Cartera"] = c_info.get("contratado_curso", c_info["contratado_cliente"])
                alu["Convenio_Cartera"] = c_info["convenio"]
                alu["Dias_Mora_Cartera"] = c_info["dias_mora"]
                alu["Fecha_Inscripcion_Cartera"] = c_info.get("fecha_inscripcion", "")
                if c_info["riesgo_cero"] == "RIESGO_CERO_REAL":
                    alu["Estado_Cartera"] = "RIESGO $0 REAL"
                elif c_info["saldo_cliente"] > 0 and c_info["dias_mora"] > 90:
                    alu["Estado_Cartera"] = f"MORA CRÍTICA ({c_info['dias_mora']}d)"
                elif c_info["saldo_cliente"] > 0:
                    alu["Estado_Cartera"] = "CON SALDO"
                else:
                    alu["Estado_Cartera"] = "PAZ Y SALVO"

                if alu.get("Clases_Completadas", 0) > 0 and (
                    c_info["riesgo_cero"] == "RIESGO_CERO_REAL"
                    or (c_info["saldo_cliente"] > 0 and c_info["dias_mora"] > 60)
                ):
                    alertas_cruce.append({
                        "Documento": alu.get("Documento", ""),
                        "Nombre_Completo": alu.get("Nombre_Completo", ""),
                        "Categoria": alu.get("Categoria", ""),
                        "Clases_Completadas": alu.get("Clases_Completadas", 0),
                        "Clases_Requeridas": alu.get("Clases_Requeridas", 20),
                        "Saldo_Cartera": c_info["saldo_cliente"],
                        "Abonado_Cartera": c_info["abonado_cliente"],
                        "Estado_Cartera": alu["Estado_Cartera"],
                        "Convenio": c_info["convenio"],
                    })
            else:
                alu["Cartera_Cruzada"] = False
                alu["Saldo_Cartera"] = 0
                alu["Abonado_Cartera"] = 0
                alu["Convenio_Cartera"] = "-"
                alu["Estado_Cartera"] = "No cruzado"

        instructores = self.listar_instructores_registrados()
        if not instructores:
            inst_vistos = set()
            for row_ina in SEED_INASISTENCIAS_HISTORICAS:
                nom_i = str(row_ina[8]).strip() if len(row_ina) > 8 else ""
                if nom_i and _normalizar_texto(nom_i) not in inst_vistos:
                    inst_vistos.add(_normalizar_texto(nom_i))
                    instructores.append({"nombre": nom_i, "documento": "", "telegram_id": ""})
        nomina = self.consultar_clases_por_instructor(nomina_instructor, nomina_periodo)
        cierres = self.consultar_cierres_jornada(nomina_periodo)
        inasistencias = self.consultar_inasistencias_periodo(ina_periodo, ina_estado)
        mantenimientos = self.consultar_mantenimientos_periodo(mant_periodo, mant_placa)
        vehiculos = self.listar_vehiculos_activos()

        # Normalización enriquecida para el Dashboard Web Central (Módulos 2, 3, 4 y 5)
        alumnos_web_list: List[Dict[str, Any]] = []
        mora_en_clase_cnt = 0
        cero_en_clase_cnt = 0
        runt_cargados_cnt = 0
        runt_pendientes_cnt = 0
        semaforo_rojo_cnt = 0
        semaforo_amarillo_cnt = 0
        semaforo_verde_cnt = 0
        hoy_date = datetime.now(TZ_COLOMBIA).date()

        for alu in alumnos_todos:
            cat_a = normalizar_categoria(str(alu.get("Categoria", "B1")))
            cfg_a = FORMATOS_OFICIALES.get(cat_a, FORMATOS_OFICIALES["B1"])
            comp_a = int(alu.get("Clases_Completadas", 0) or 0)
            req_a = int(alu.get("Clases_Requeridas", cfg_a.get("total_clases", 20)) or 20)
            pct_a = min(100, int(round((comp_a / max(1, req_a)) * 100)))
            clases_faltantes_a = max(0, req_a - comp_a)
            pct_faltante_a = max(0, 100 - pct_a)
            saldo_c = float(alu.get("Saldo_Cartera", 0) or 0)
            abonado_c = float(alu.get("Abonado_Cartera", 0) or 0)
            abonado_curso_c = float(alu.get("Abonado_Curso_Cartera", abonado_c) or 0)
            contratado_curso_c = float(alu.get("Contratado_Curso_Cartera", 0) or 0)
            est_c = str(alu.get("Estado_Cartera", "No cruzado"))
            raw_runt_a = str(alu.get("Estado_RUNT", "No Registrado RUNT")).strip()
            es_runt_cargado = "cargue" in raw_runt_a.lower() or "cargado" in raw_runt_a.lower()
            est_runt_a = "Cargue RUNT" if es_runt_cargado else "No Registrado RUNT"
            fecha_runt_a = str(alu.get("Fecha_RUNT", "")).strip()
            if es_runt_cargado:
                runt_cargados_cnt += 1
            else:
                runt_pendientes_cnt += 1

            # Cálculo del Semáforo de Vigencia de 90 Días desde la 1ra Clase Práctica
            doc_norm_a = _normalizar_texto(str(alu.get("Documento", "")))
            nom_norm_a = _normalizar_texto(str(alu.get("Nombre_Completo", "")))
            fecha_primera_iso = (
                primera_clase_por_doc.get((doc_norm_a, cat_a))
                or primera_clase_por_nom.get((nom_norm_a, cat_a))
                or ""
            )
            fecha_ultima_iso = (
                ultima_clase_por_doc.get((doc_norm_a, cat_a))
                or ultima_clase_por_nom.get((nom_norm_a, cat_a))
                or ""
            )
            origen_fecha_90 = "1ra Clase" if fecha_primera_iso else ""
            if not fecha_primera_iso:
                f_insc_c = str(alu.get("Fecha_Inscripcion_Cartera", "")).strip()
                if f_insc_c and len(f_insc_c) >= 10:
                    fecha_primera_iso = f_insc_c[:10]
                    origen_fecha_90 = "Matrícula"

            fecha_venc_iso = ""
            dias_transcurridos_90 = 0
            dias_restantes_90 = 90
            ritmo_clases_dia = 0.0
            semaforo_codigo = "SIN_INICIAR"
            semaforo_color = "AZUL"
            semaforo_etiqueta = "⚪ Sin Fecha 1ra Clase"
            semaforo_detalle = f"Faltan {clases_faltantes_a} clases ({pct_faltante_a}%)"
            urgencia_score = 0.0

            if alu.get("Curso_Completado") or comp_a >= req_a:
                semaforo_codigo = "COMPLETADO"
                semaforo_color = "VERDE"
                semaforo_etiqueta = "✅ 100% Terminado"
                semaforo_detalle = f"Cumplió {comp_a}/{req_a} clases dentro del plazo"
                urgencia_score = -10.0
                if fecha_primera_iso:
                    try:
                        d0 = datetime.strptime(fecha_primera_iso[:10], "%Y-%m-%d").date()
                        fecha_venc_iso = (d0 + timedelta(days=90)).strftime("%Y-%m-%d")
                        dias_transcurridos_90 = max(0, (hoy_date - d0).days)
                        dias_restantes_90 = 90 - dias_transcurridos_90
                    except Exception:
                        pass
            elif fecha_primera_iso:
                try:
                    d0 = datetime.strptime(fecha_primera_iso[:10], "%Y-%m-%d").date()
                    d_venc = d0 + timedelta(days=90)
                    fecha_venc_iso = d_venc.strftime("%Y-%m-%d")
                    dias_transcurridos_90 = max(0, (hoy_date - d0).days)
                    dias_restantes_90 = 90 - dias_transcurridos_90
                    ritmo_clases_dia = round(clases_faltantes_a / max(1, dias_restantes_90), 2) if dias_restantes_90 > 0 else float(clases_faltantes_a)

                    # Reglas de Semáforo (Plazo 90 Días vs % de Clases Faltantes)
                    es_vencido = dias_restantes_90 <= 0
                    es_critico = (
                        es_vencido
                        or dias_restantes_90 <= 15
                        or (dias_restantes_90 <= 25 and pct_faltante_a >= 60)
                        or (dias_restantes_90 <= 35 and pct_faltante_a >= 80)
                        or ritmo_clases_dia >= 1.0
                    )
                    es_alerta = (
                        not es_critico
                        and (
                            dias_restantes_90 <= 30
                            or (dias_restantes_90 <= 45 and pct_faltante_a >= 50)
                            or (dias_restantes_90 <= 60 and pct_faltante_a >= 75)
                            or ritmo_clases_dia >= 0.45
                        )
                    )

                    if es_vencido:
                        semaforo_codigo = "VENCIDO"
                        semaforo_color = "ROJO"
                        semaforo_etiqueta = f"🚨 VENCIDO (+{abs(dias_restantes_90)}d)"
                        semaforo_detalle = f"Venció el {fecha_venc_iso} • Faltan {clases_faltantes_a} clases ({pct_faltante_a}%)"
                        urgencia_score = 1000.0 + abs(dias_restantes_90) + pct_faltante_a
                        if not es_runt_cargado:
                            semaforo_rojo_cnt += 1
                    elif es_critico:
                        semaforo_codigo = "CRITICO"
                        semaforo_color = "ROJO"
                        semaforo_etiqueta = f"🔴 CRÍTICO ({dias_restantes_90}d rest.)"
                        semaforo_detalle = f"Falta {pct_faltante_a}% ({clases_faltantes_a} cl.) • Req. {ritmo_clases_dia} cl/día"
                        urgencia_score = 500.0 + (90 - dias_restantes_90) + pct_faltante_a
                        if not es_runt_cargado:
                            semaforo_rojo_cnt += 1
                    elif es_alerta:
                        semaforo_codigo = "ALERTA"
                        semaforo_color = "AMARILLO"
                        semaforo_etiqueta = f"🟡 ALERTA ({dias_restantes_90}d rest.)"
                        semaforo_detalle = f"Falta {pct_faltante_a}% ({clases_faltantes_a} cl.) • Vence {fecha_venc_iso}"
                        urgencia_score = 200.0 + (90 - dias_restantes_90) + (pct_faltante_a * 0.5)
                        if not es_runt_cargado:
                            semaforo_amarillo_cnt += 1
                    else:
                        semaforo_codigo = "A_TIEMPO"
                        semaforo_color = "VERDE"
                        semaforo_etiqueta = f"🟢 A TIEMPO ({dias_restantes_90}d rest.)"
                        semaforo_detalle = f"Falta {pct_faltante_a}% ({clases_faltantes_a} cl.) • Vence {fecha_venc_iso}"
                        urgencia_score = float(90 - dias_restantes_90)
                        if not es_runt_cargado:
                            semaforo_verde_cnt += 1
                except Exception:
                    pass

            alerta_code = "AL_DIA"
            alerta_txt = (
                f"✅ PAZ Y SALVO (Abonó ${int(abonado_curso_c):,})"
                if abonado_curso_c > 0
                else (f"✅ PAZ Y SALVO (Pagó ${int(abonado_c):,})" if abonado_c > 0 else "✅ PAZ Y SALVO")
            )
            if not alu.get("Cartera_Cruzada"):
                alerta_code = "NO_CRUZADO"
                alerta_txt = "Sin cruce en Excel actual"
            elif est_c == "RIESGO $0 REAL":
                alerta_code = "RIESGO_CERO_REAL"
                alerta_txt = f"🚨 ABONO $0 REAL (Debe ${int(saldo_c):,})"
                if not alu.get("Curso_Completado"):
                    cero_en_clase_cnt += 1
                    mora_en_clase_cnt += 1
            elif "MORA CRÍTICA" in est_c:
                alerta_code = "MORA_CRITICA_90"
                alerta_txt = f"⏳ {est_c} (Abonó ${int(abonado_c):,} • Debe ${int(saldo_c):,})"
                if not alu.get("Curso_Completado"):
                    mora_en_clase_cnt += 1
            elif saldo_c > 0:
                alerta_code = "CON_SALDO"
                alerta_txt = (
                    f"⚠️ Saldo: ${int(saldo_c):,} (Abonó ${int(abonado_c):,})"
                    if abonado_c > 0
                    else f"⚠️ Saldo Pendiente (${int(saldo_c):,})"
                )
                if not alu.get("Curso_Completado"):
                    mora_en_clase_cnt += 1

            alumnos_web_list.append({
                "id_matricula": str(alu.get("ID_Matricula", "")),
                "cedula": str(alu.get("Documento", "")),
                "nombre": str(alu.get("Nombre_Completo", "")),
                "categoria": cat_a,
                "formato_codigo": cfg_a.get("codigo_formato", "A.O.C.E.A - F - 61"),
                "clases_completadas": comp_a,
                "total_requerido": req_a,
                "clases_faltantes": clases_faltantes_a,
                "porcentaje_avance": pct_a,
                "porcentaje_faltante": pct_faltante_a,
                "estado_curso": "FINALIZADO" if alu.get("Curso_Completado") else "EN CURSO",
                "estado_runt": est_runt_a,
                "fecha_runt": fecha_runt_a,
                "fecha_primera_clase": fecha_primera_iso,
                "fecha_ultima_clase": fecha_ultima_iso,
                "origen_fecha_90": origen_fecha_90,
                "fecha_vencimiento_90": fecha_venc_iso,
                "dias_transcurridos_90": dias_transcurridos_90,
                "dias_restantes_90": dias_restantes_90,
                "ritmo_clases_dia": ritmo_clases_dia,
                "semaforo_codigo": semaforo_codigo,
                "semaforo_color": semaforo_color,
                "semaforo_etiqueta": semaforo_etiqueta,
                "semaforo_detalle": semaforo_detalle,
                "urgencia_score": urgencia_score,
                "alerta_cartera": alerta_code,
                "alerta_texto": alerta_txt,
                "saldo_cartera": saldo_c,
                "abonado_cartera": abonado_c,
                "abonado_curso": abonado_curso_c,
                "contratado_curso": contratado_curso_c,
                "convenio": str(alu.get("Convenio_Cartera", "-")),
            })

        # Resumen por instructor para Nómina Web
        resumen_inst_map: Dict[str, Dict[str, Any]] = {}
        detalle_nomina_web: List[Dict[str, Any]] = []
        for r_nom in nomina.get("registros", []):
            inst_n = str(r_nom.get("Instructor", "Sin nombre")).strip() or "Sin nombre"
            cat_n = normalizar_categoria(str(r_nom.get("Categoria", "B1")))
            cant_n = int(r_nom.get("Cantidad_Clases_Calc", 1) or 1)
            val_h = str(r_nom.get("Validacion_Horario", "EN TIEMPO")).strip()
            es_fuera = "FUERA" in val_h.upper()
            if inst_n not in resumen_inst_map:
                resumen_inst_map[inst_n] = {
                    "instructor": inst_n,
                    "sesiones": 0,
                    "clases_dictadas": 0,
                    "a_tiempo": 0,
                    "fuera_horario": 0,
                    "categorias": {},
                }
            st_i = resumen_inst_map[inst_n]
            st_i["sesiones"] += 1
            st_i["clases_dictadas"] += cant_n
            if es_fuera:
                st_i["fuera_horario"] += 1
            else:
                st_i["a_tiempo"] += 1
            st_i["categorias"][cat_n] = st_i["categorias"].get(cat_n, 0) + cant_n

            detalle_nomina_web.append({
                "fecha": str(r_nom.get("Fecha_Reporte", "")),
                "hora": str(r_nom.get("Hora_Reporte", "")),
                "instructor": inst_n,
                "cedula": str(r_nom.get("ID_Alumno", "")),
                "alumno": str(r_nom.get("Nombre_Alumno", "")),
                "categoria": cat_n,
                "clases": str(r_nom.get("Clases_Reportadas", "")),
                "num_clases": cant_n,
                "estado_validacion": "⚠️ FUERA DE HORARIO" if es_fuera else "✅ A TIEMPO",
                "detalle_validacion": str(r_nom.get("Observaciones_Instructor", "")) or val_h,
            })

        nomina["titulo_periodo"] = nomina.get("periodo_etiqueta", "Periodo")
        nomina["instructor_filtro"] = nomina.get("instructor_nombre", "TODOS")
        nomina["total_sesiones"] = nomina.get("total_reportes", 0)
        nomina["a_tiempo"] = nomina.get("en_tiempo", 0)
        nomina["fuera_horario"] = nomina.get("fuera_tiempo", 0)
        nomina["instructores_disponibles"] = [i["nombre"] for i in instructores if i.get("nombre")]
        nomina["resumen_instructores"] = sorted(
            resumen_inst_map.values(), key=lambda x: x["clases_dictadas"], reverse=True
        )
        nomina["detalle_clases"] = detalle_nomina_web

        # Normalización de Inasistencias para Web
        inas_reg_raw = inasistencias.get("registros", [])
        cnt_efec = 0
        cnt_transf = 0
        cnt_pend = 0
        inas_web_rows: List[Dict[str, Any]] = []
        for idx_in, r_in in enumerate(inas_reg_raw, start=1):
            est_p = str(r_in.get("Estado", "Pendiente Pago")).strip()
            est_l = est_p.lower()
            if "efectivo" in est_l:
                cnt_efec += 1
            elif "transf" in est_l or "trasf" in est_l:
                cnt_transf += 1
            else:
                cnt_pend += 1
            consec_val = str(r_in.get("Consecutivo", idx_in)).strip() or str(idx_in)
            inas_web_rows.append({
                "fila": consec_val,
                "id_novedad": str(r_in.get("ID_Novedad", "")),
                "fecha": str(r_in.get("Fecha", "")),
                "hora": str(r_in.get("Hora", "")),
                "cedula": str(r_in.get("Documento_Alumno", "")),
                "alumno": str(r_in.get("Nombre_Alumno", "")),
                "categoria": str(r_in.get("Categoria", "B1")),
                "clases_faltadas": str(r_in.get("Clases_Faltadas", "")),
                "instructor": str(r_in.get("Nombre_Instructor", "")),
                "estado_pago": est_p,
            })
        inasistencias["titulo_periodo"] = inasistencias.get("periodo_etiqueta", "Histórico")
        inasistencias["kpis"] = {
            "total": len(inas_web_rows),
            "pendientes": cnt_pend,
            "efectivo": cnt_efec,
            "transferencia": cnt_transf,
        }
        inasistencias["registros_raw"] = inas_reg_raw
        inasistencias["registros"] = inas_web_rows

        # Normalización de Mantenimientos y Flota PESV para Web
        flota_web: List[Dict[str, Any]] = []
        mapa_tipo_placa: Dict[str, str] = {}
        for v in vehiculos:
            pl_v = str(v.get("placa", "")).strip().upper()
            tip_v = str(v.get("tipo", "🚗 Carro"))
            es_m = "moto" in tip_v.lower() or "🏍" in tip_v
            icono_v = "🏍️" if es_m else "🚗"
            cat_v = "A2" if es_m else ("C1/B1" if pl_v in ("UDT984", "FSP335", "DUX342") else "B1")
            marca_v = "AKT / YAMAHA" if es_m else "RENAULT / CHEVROLET"
            mapa_tipo_placa[pl_v] = "MOTO" if es_m else "CARRO"
            flota_web.append({
                "placa": pl_v,
                "icono": icono_v,
                "tipo": "MOTO" if es_m else "CARRO",
                "categoria": cat_v,
                "marca": marca_v,
            })

        mant_raw = mantenimientos.get("registros", [])
        mant_web_rows: List[Dict[str, Any]] = []
        for idx_m, r_m in enumerate(mant_raw, start=1):
            pl_m = str(r_m.get("Placa", "")).strip().upper()
            mant_web_rows.append({
                "id": f"MANT-{idx_m:03d}",
                "fecha": str(r_m.get("Fecha", "")),
                "placa": pl_m,
                "tipo_vehiculo": mapa_tipo_placa.get(pl_m, "VEHÍCULO CEA"),
                "tipo_mantenimiento": str(r_m.get("Tipo_Mantenimiento", "PREVENTIVO")).upper(),
                "kilometraje": str(r_m.get("Kilometraje_Actual", "0")),
                "actividad": str(r_m.get("Actividad_Realizada", "")),
                "taller": str(r_m.get("Establecimiento", "")),
                "mecanico": str(r_m.get("Nombre_Tecnico", "")),
            })

        mantenimientos["titulo_periodo"] = mantenimientos.get("periodo_etiqueta", "Histórico")
        mantenimientos["flota_activa"] = flota_web
        mantenimientos["kpis"] = {
            "total": len(mant_web_rows),
            "preventivos": int(mantenimientos.get("preventivos", 0)),
            "correctivos": int(mantenimientos.get("correctivos", 0)),
            "vehiculos_intervenidos": len(mantenimientos.get("por_placa", {})),
        }
        mantenimientos["registros_raw"] = mant_raw
        mantenimientos["registros"] = mant_web_rows

        cierres_web = [
            {
                "fecha": str(c.get("Fecha_Jornada", "")),
                "hora": str(c.get("Hora_Cierre", "")),
                "instructor": str(c.get("Instructor", "")),
                "estado": str(c.get("Estado_Cruce", "CERRADA")),
                "drive_link": str(c.get("URL_Captura_Agenda", "")),
            }
            for c in cierres.get("cierres", [])
        ]

        formatos_resumen = {
            k: {
                "nombre_corto": v["nombre_corto"],
                "codigo_formato": v["codigo_formato"],
                "total_clases": v["total_clases"],
                "filas": [
                    {
                        "clases": f["clases"],
                        "etapa": f["etapa"],
                        "resumen": f["resumen"],
                        "tema": f["tema"],
                    }
                    for f in v.get("filas", [])
                ],
            }
            for k, v in FORMATOS_OFICIALES.items()
        }

        return {
            "ok": True,
            "conectado_google_sheets": self.conectado_en_vivo,
            "sheets_conectado": self.conectado_en_vivo,
            "sheet_id": GOOGLE_SHEET_ID,
            "ultimo_error": self._ultimo_error_msg,
            "clases": {
                "kpis": {
                    "total_alumnos": len(alumnos_web_list),
                    "en_curso": alumnos_activos_cnt,
                    "finalizados": alumnos_finalizados_cnt,
                    "runt_pendientes": runt_pendientes_cnt,
                    "runt_cargados": runt_cargados_cnt,
                    "semaforo_rojo": semaforo_rojo_cnt,
                    "semaforo_amarillo": semaforo_amarillo_cnt,
                    "semaforo_verde": semaforo_verde_cnt,
                    "total_sesiones_registradas": nomina.get("total_reportes", 0),
                    "alumnos_con_mora_en_clase": mora_en_clase_cnt,
                    "alumnos_abono_cero_en_clase": cero_en_clase_cnt,
                },
                "alumnos": alumnos_web_list,
            },
            "alumnos": {
                "total": len(alumnos_todos),
                "activos_count": alumnos_activos_cnt,
                "finalizados_count": alumnos_finalizados_cnt,
                "alertas_cruce_cartera": alertas_cruce,
                "registros": alumnos_todos,
            },
            "instructores": instructores,
            "nomina": nomina,
            "cierres": cierres,
            "cierres_jornada": cierres_web,
            "inasistencias": inasistencias,
            "mantenimientos": mantenimientos,
            "vehiculos": vehiculos,
            "talleres_pesv": TALLERES_FRECUENTES_PESV,
            "formatos_oficiales": formatos_resumen,
        }

    # Adaptadores directos para endpoints /api/bot/* del Dashboard Web
    def obtener_historial_alumno(self, criterio: str, categoria: Optional[str] = None) -> Dict[str, Any]:
        res = self.obtener_clases_alumno(criterio, categoria=categoria)
        alu = res.get("alumno") or {}
        return {
            "alumno": {
                "ID_Matricula": alu.get("ID_Matricula", criterio),
                "Documento": alu.get("Documento", criterio),
                "Nombre_Completo": alu.get("Nombre_Completo", criterio or "ALUMNO CEA"),
                "Categoria": alu.get("Categoria", categoria or "B1"),
                "Estado": alu.get("Estado", "En curso"),
                "cedula": alu.get("Documento", criterio),
                "nombre": alu.get("Nombre_Completo", criterio or "ALUMNO CEA"),
                "categoria": alu.get("Categoria", categoria or "B1"),
            },
            "registros": res.get("registros") or [],
        }

    def obtener_clases_por_instructor_periodo(self, instructor: str = "TODOS", periodo: str = "mes") -> Dict[str, Any]:
        return self.consultar_clases_por_instructor(instructor, periodo)

    def obtener_inasistencias_periodo(self, periodo: str = "todo", estado: str = "TODOS") -> Dict[str, Any]:
        return self.consultar_inasistencias_periodo(periodo, estado)

    def buscar_inasistencias_alumno(self, criterio_alumno: str) -> Dict[str, Any]:
        lista = self.obtener_inasistencias_alumno(criterio_alumno)
        ficha = self.buscar_alumno(criterio_alumno) or {}
        if not ficha and lista:
            primera = lista[0]
            ficha = {
                "Nombre_Completo": primera.get("Nombre_Alumno", criterio_alumno),
                "Documento": primera.get("Documento_Alumno", "Registrado en sistema"),
                "Categoria": primera.get("Categoria", "B1"),
            }
        elif not ficha:
            ficha = {
                "Nombre_Completo": criterio_alumno,
                "Documento": criterio_alumno,
                "Categoria": "B1",
            }
        return {"alumno": ficha, "inasistencias": lista}

    def obtener_mantenimientos_periodo(self, periodo: str = "todo", placa: str = "TODAS") -> Dict[str, Any]:
        if periodo.startswith("mes_") and len(periodo) == 6 and periodo[4:].isdigit():
            meses_idx_map = {
                "01": "enero", "02": "febrero", "03": "marzo", "04": "abril",
                "05": "mayo", "06": "junio", "07": "julio", "08": "agosto",
                "09": "septiembre", "10": "octubre", "11": "noviembre", "12": "diciembre",
            }
            periodo = meses_idx_map.get(periodo[4:], periodo)
        return self.consultar_mantenimientos_periodo(periodo, placa)

    def asegurar_alumno(
        self,
        cedula: str,
        nombre: str,
        categoria: str = "A2",
        telefono: str = "",
        estado_runt: Optional[str] = None,
    ) -> Dict[str, Any]:
        return self.registrar_o_actualizar_alumno(cedula, nombre, categoria, estado_runt=estado_runt)

    def editar_alumno_completo(
        self,
        cedula: str,
        nombre: str,
        categoria: str = "A2",
        estado: str = "",
        estado_runt: Optional[str] = None,
        fecha_runt: str = "",
        fecha_primera_clase: str = "",
        telefono: str = "",
        orig_cedula: str = "",
        orig_nombre: str = "",
        orig_categoria: str = "",
    ) -> Dict[str, Any]:
        """
        Edita CUALQUIER dato de un alumno en la hoja 'Alumnos' (Cédula, Nombre, Categoría,
        Estado del curso, Estado_RUNT, Fecha_RUNT) y propaga cambios de Cédula/Nombre/Categoría
        o Fecha de 1ra Clase en 'Registro_Clases' para que nunca se pierda su historial.
        """
        doc_nuevo = str(cedula or "").strip()
        nom_nuevo = limpiar_solo_nombre_persona(str(nombre or "")) or str(nombre or "").strip()
        cat_nueva = normalizar_categoria(categoria or "A2")
        est_nuevo = str(estado or "").strip()
        runt_norm = ""
        if estado_runt:
            runt_norm = "Cargue RUNT" if ("cargue" in str(estado_runt).lower() or "cargado" in str(estado_runt).lower()) else "No Registrado RUNT"
        fecha_runt_in = str(fecha_runt or "").strip()
        fecha_primera_in = str(fecha_primera_clase or "").strip()

        orig_doc_n = _normalizar_texto(orig_cedula or doc_nuevo)
        orig_nom_n = _normalizar_texto(orig_nombre or nom_nuevo)
        orig_cat_n = normalizar_categoria(orig_categoria or cat_nueva)

        ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
        todas = ws.get_all_values()
        fila_encontrada = 0
        fila_prev: List[str] = []

        for idx_f in range(2, len(todas) + 1):
            fila = todas[idx_f - 1]
            if not any(str(c).strip() for c in fila):
                continue
            d_f = _normalizar_texto(str(fila[1]).strip()) if len(fila) > 1 else ""
            n_f = _normalizar_texto(str(fila[2]).strip()) if len(fila) > 2 else ""
            c_f = normalizar_categoria(str(fila[3]).strip()) if len(fila) > 3 else "B1"
            if c_f != orig_cat_n:
                continue
            if (orig_doc_n and d_f == orig_doc_n) or (orig_nom_n and n_f == orig_nom_n):
                fila_encontrada = idx_f
                fila_prev = fila
                break

        id_mat_final = str(fila_prev[0]).strip() if len(fila_prev) > 0 and str(fila_prev[0]).strip() else ""
        est_final = est_nuevo or (str(fila_prev[4]).strip() if len(fila_prev) > 4 and str(fila_prev[4]).strip() else "En curso")
        runt_final = runt_norm or (str(fila_prev[5]).strip() if len(fila_prev) > 5 and str(fila_prev[5]).strip() else "No Registrado RUNT")
        if fecha_runt_in:
            fecha_runt_final = fecha_runt_in
        elif runt_final == "Cargue RUNT":
            prev_fr = str(fila_prev[6]).strip() if len(fila_prev) > 6 else ""
            fecha_runt_final = prev_fr or datetime.now(TZ_COLOMBIA).strftime("%Y-%m-%d %H:%M")
        else:
            fecha_runt_final = ""

        if fila_encontrada >= 2:
            ws.update(
                values=[[id_mat_final or f"MAT-{fila_encontrada:03d}", doc_nuevo, nom_nuevo, cat_nueva, est_final, runt_final, fecha_runt_final]],
                range_name=f"A{fila_encontrada}:G{fila_encontrada}",
                value_input_option="USER_ENTERED",
            )
        else:
            max_num = max(0, len(todas) - 1)
            for f in todas[1:]:
                if f:
                    m_id = re.search(r"(\d+)", str(f[0]))
                    if m_id:
                        max_num = max(max_num, int(m_id.group(1)))
            id_mat_final = f"MAT-{max_num + 1:03d}"
            self._agregar_fila_al_final(ws, [id_mat_final, doc_nuevo, nom_nuevo, cat_nueva, est_final, runt_final, fecha_runt_final])

        # Propagar cambios de Cédula / Nombre / Categoría / Fecha 1ra Clase en Registro_Clases
        clases_actualizadas = 0
        cambio_identidad = (
            (_normalizar_texto(doc_nuevo) != orig_doc_n)
            or (_normalizar_texto(nom_nuevo) != orig_nom_n)
            or (cat_nueva != orig_cat_n)
        )
        if cambio_identidad or fecha_primera_in:
            try:
                ws_reg = self.obtener_o_crear_hoja(SHEET_REGISTRO, HEADERS_REGISTRO)
                filas_reg = ws_reg.get_all_values()
                matches_reg: List[Tuple[int, str, List[str]]] = []
                for idx_r in range(2, len(filas_reg) + 1):
                    fr = filas_reg[idx_r - 1]
                    if len(fr) < 7:
                        continue
                    d_r = _normalizar_texto(str(fr[4]).strip())
                    n_r = _normalizar_texto(limpiar_solo_nombre_persona(str(fr[5])) or str(fr[5]))
                    c_r = normalizar_categoria(str(fr[6]).strip())
                    if c_r != orig_cat_n:
                        continue
                    if (orig_doc_n and d_r == orig_doc_n) or (orig_nom_n and n_r == orig_nom_n):
                        f_iso = self._normalizar_fecha_iso(str(fr[1]), str(fr[0])) or "9999-99-99"
                        matches_reg.append((idx_r, f_iso, fr))

                if cambio_identidad:
                    for idx_r, _, _ in matches_reg:
                        ws_reg.update(
                            values=[[doc_nuevo, nom_nuevo, cat_nueva]],
                            range_name=f"E{idx_r}:G{idx_r}",
                            value_input_option="USER_ENTERED",
                        )
                        clases_actualizadas += 1

                if fecha_primera_in and matches_reg:
                    matches_reg.sort(key=lambda x: x[1])
                    idx_primera, iso_actual, _ = matches_reg[0]
                    if iso_actual[:10] != fecha_primera_in[:10]:
                        ws_reg.update_cell(idx_primera, 2, fecha_primera_in[:10])
                        clases_actualizadas += 1
            except Exception as exc_reg:
                logger.warning("Aviso propagando cambios en Registro_Clases: %s", exc_reg)

        return {
            "ok": True,
            "ID_Matricula": id_mat_final,
            "Documento": doc_nuevo,
            "Nombre_Completo": nom_nuevo,
            "Categoria": cat_nueva,
            "Estado": est_final,
            "Estado_RUNT": runt_final,
            "Fecha_RUNT": fecha_runt_final,
            "clases_actualizadas": clases_actualizadas,
        }

    def agregar_vehiculo_pesv(self, placa: str, tipo: str = "CARRO", categoria: str = "B1", marca: str = "") -> Dict[str, Any]:
        res = self.agregar_vehiculo_flota(placa, tipo)
        res["ok"] = True
        return res

    def retirar_vehiculo_pesv(self, placa: str) -> Dict[str, Any]:
        ok = self.eliminar_vehiculo_flota(placa)
        return {"ok": True, "modificado": ok, "placa": placa}


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
        if _es_registro_inasistencia(fila):
            continue
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

    # Agregar filas de Clases Adicionales / de Refuerzo al final de la tabla si existen
    # (Cuidando de no confundir el tema oficial de la clase 14 'Refuerzo de pruebas CALE' con una clase de refuerzo adicional,
    #  y excluyendo cualquier reporte de inasistencia).
    for fila_r in registros:
        if _es_registro_inasistencia(fila_r):
            continue
        cl_raw = str(fila_r.get("Clases_Reportadas", "")).strip()
        tm_raw = str(fila_r.get("Tema_Registrado", "")).strip()
        es_fila_refuerzo = (
            "(refuerzo)" in cl_raw.lower()
            or "adicional" in cl_raw.lower()
            or tm_raw.upper().startswith("[CLASE DE REFUERZO]")
        )
        if es_fila_refuerzo:
            inst_r = str(fila_r.get("Instructor", "")).strip()
            fech_r = str(fila_r.get("Fecha_Reporte", "")).strip()
            hor_r = str(fila_r.get("Hora_Reporte", "")).strip()
            etiqueta_adic = cl_raw.replace("(Refuerzo)", "").replace("Adicional", "Adic.").strip() or "Adic."
            filas_tabla.append(
                [
                    Paragraph("<b>Clase Adicional /<br/>Refuerzo</b>", estilo_celda_centro),
                    Paragraph(f"<b>{etiqueta_adic}</b>", estilo_celda_centro),
                    Paragraph(tm_raw.replace("\n", "<br/>"), estilo_celda),
                    Paragraph(
                        f"<b>{inst_r}</b><br/><font size=6>{fech_r} {hor_r}</font>",
                        estilo_celda_centro,
                    ),
                    Paragraph(f"{nombre_alumno}", estilo_celda_centro),
                ]
            )

    tabla_clases = Table(
        filas_tabla,
        colWidths=[2.3 * cm, 1.3 * cm, 9.5 * cm, 3.2 * cm, 3.2 * cm],
        repeatRows=2,
    )
    tabla_clases.setStyle(TableStyle(estilos_tabla))
    elementos.append(tabla_clases)
    elementos.append(Spacer(1, 0.15 * cm))

    # Recopilar todas las observaciones realizadas por los instructores por clase dictada
    lineas_observaciones: List[str] = []
    inasistencias_vistas_clave = set()

    for fila_r in registros:
        if _es_registro_inasistencia(fila_r):
            continue
        obs_txt = str(fila_r.get("Observaciones_Instructor", "")).strip()
        if obs_txt:
            cl_r = str(fila_r.get("Clases_Reportadas", "")).strip()
            fech_r = str(fila_r.get("Fecha_Reporte", "")).strip()
            hor_r = str(fila_r.get("Hora_Reporte", "")).strip()
            inst_r = str(fila_r.get("Instructor", "")).strip()
            lineas_observaciones.append(
                f"• <b>Clase {cl_r} ({fech_r} {hor_r}):</b> {obs_txt} "
                f"— <i>Observación realizada por Instructor: <b>{inst_r}</b></i>"
            )

    # Incluir en OBSERVACIONES únicamente: Nombre del profesor, fecha y hora y qué clase faltó
    # (Sin registrar la clase en el cuadro de control de clases dictadas).
    for fila_r in registros:
        if not _es_registro_inasistencia(fila_r):
            continue
        nums_f = extraer_numeros_clase(str(fila_r.get("Clases_Reportadas", "")))
        cl_falt_r = formatear_clases(nums_f) if nums_f else str(fila_r.get("Clases_Reportadas", "")).strip()
        cl_falt_r = re.sub(r"\(?no\s+asisti[oó]?\)?", "", cl_falt_r, flags=re.IGNORECASE).strip() or "programada"
        fech_r = str(fila_r.get("Fecha_Reporte", "")).strip()
        hor_r = str(fila_r.get("Hora_Reporte", "")).strip()
        inst_r = str(fila_r.get("Instructor", "")).strip()
        clave_ina = (fech_r, cl_falt_r)
        if clave_ina not in inasistencias_vistas_clave:
            inasistencias_vistas_clave.add(clave_ina)
            lineas_observaciones.append(
                f"• El alumno no asistió a la clase número <b>{cl_falt_r}</b> — "
                f"<b>Profesor:</b> {inst_r} | <b>Fecha y Hora:</b> {fech_r} {hor_r}"
            )

    inasistencias_alu = sheets_service.obtener_inasistencias_alumno(
        criterio_alumno=documento_alumno or nombre_alumno,
        nombre_alternativo=nombre_alumno,
        categoria=categoria_key,
    )
    if inasistencias_alu:
        for ina in inasistencias_alu:
            fech_i = str(ina.get("Fecha", "")).strip()
            hor_i = str(ina.get("Hora", "")).strip()
            cl_falt = str(ina.get("Clases_Faltadas", "")).strip()
            if not cl_falt or cl_falt.lower() in ("no especificada", "s/n", "ninguna"):
                cl_falt_limpia = "programada"
            else:
                cl_falt_limpia = cl_falt
            inst_i = str(ina.get("Nombre_Instructor", "")).strip()
            clave_ina = (fech_i, cl_falt_limpia)
            if clave_ina not in inasistencias_vistas_clave:
                inasistencias_vistas_clave.add(clave_ina)
                lineas_observaciones.append(
                    f"• El alumno no asistió a la clase número <b>{cl_falt_limpia}</b> — "
                    f"<b>Profesor:</b> {inst_i} | <b>Fecha y Hora:</b> {fech_i} {hor_i}"
                )

    contenido_observaciones = (
        "<br/>".join(lineas_observaciones) if lineas_observaciones else "&nbsp;"
    )

    # Determinar el último instructor que dictó la clase final del curso (y su cédula oficial)
    ultimo_instructor = ""
    registros_asistidos = [r for r in registros if not _es_registro_inasistencia(r)]
    if registros_asistidos:
        ultimo_instructor = str(registros_asistidos[-1].get("Instructor", "")).strip()
    if not ultimo_instructor and mapa_clases:
        max_clase_num = max(mapa_clases.keys())
        ultimo_instructor = str(mapa_clases[max_clase_num].get("instructor", "")).strip()
    doc_ultimo_instructor = (
        sheets_service.obtener_documento_por_nombre_instructor(ultimo_instructor)
        if ultimo_instructor
        else ""
    )

    filas_pie = [
        [Paragraph("<b>OBSERVACIONES AL PROCESO DE APRENDIZAJE:</b>", estilo_celda), "", ""],
        [Paragraph(contenido_observaciones, estilo_celda), "", ""],
        [
            Paragraph(
                "Concluida las clases se realizara evaluación práctica, calificada y aprobada por SICOV.",
                estilo_celda,
            ),
            "",
            "",
        ],
        [
            Paragraph(
                "En constancia de haber realizado las clases de Técnicas de Conducción por el alumno e instructor abajo identificados,",
                estilo_celda,
            ),
            "",
            "",
        ],
        [
            Paragraph(
                f"<b>Nombre Alumno:</b> {nombre_alumno}<br/><br/>"
                f"<b>Firma del Alumno:</b> ____________________________<br/>"
                f"<b>CC, TI, CE.:</b> {documento_alumno}",
                estilo_celda,
            ),
            Paragraph(
                f"<b>Nombre Instructor:</b> {ultimo_instructor or '______________________'}<br/><br/>"
                f"<b>Firma del Instructor:</b> ____________________________<br/>"
                f"<b>CC. Instructor:</b> {doc_ultimo_instructor or '______________________'}",
                estilo_celda,
            ),
            Paragraph(
                "<br/><br/>____________________________<br/><b>Recibido de Archivo</b>",
                estilo_celda_centro,
            ),
        ],
    ]

    tabla_pie = Table(filas_pie, colWidths=[6.8 * cm, 6.8 * cm, 5.9 * cm])
    tabla_pie.setStyle(
        TableStyle(
            [
                ("BOX", (0, 0), (-1, 1), 0.6, colors.black),
                ("LINEBELOW", (0, 0), (-1, 0), 0.4, colors.grey),
                ("SPAN", (0, 0), (2, 0)),
                ("SPAN", (0, 1), (2, 1)),
                ("SPAN", (0, 2), (2, 2)),
                ("SPAN", (0, 3), (2, 3)),
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


def generar_pdf_liquidacion_instructor(datos_consulta: Dict[str, Any]) -> io.BytesIO:
    """
    Genera el PDF Oficial de Relación / Liquidación de Clases Dictadas por Instructor
    (por Día, Semana o Mes) para control de pago semanal en CEA Auto Oriente.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=1.0 * cm,
        rightMargin=1.0 * cm,
        topMargin=0.9 * cm,
        bottomMargin=0.9 * cm,
        title="Relacion de Clases Dictadas por Instructor - CEA Auto Oriente",
    )

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "TituloLiq",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=12.0,
        alignment=TA_CENTER,
    )
    estilo_meta = ParagraphStyle(
        "MetaLiq",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.8,
        leading=10.0,
        alignment=TA_LEFT,
    )
    estilo_cab = ParagraphStyle(
        "CabLiq",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.2,
        leading=9.0,
        alignment=TA_CENTER,
    )
    estilo_celda = ParagraphStyle(
        "CeldaLiq",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.0,
        leading=8.8,
        alignment=TA_LEFT,
    )
    estilo_celda_c = ParagraphStyle(
        "CeldaLiqC",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.0,
        leading=8.8,
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

    ahora_col = datetime.now(TZ_COLOMBIA).strftime("%Y-%m-%d %H:%M")
    filas_enc = [
        [
            celda_logo,
            Paragraph(
                "<b>CEA AUTO ORIENTE<br/>RELACIÓN Y LIQUIDACIÓN DE CLASES PRÁCTICAS DICTADAS</b>",
                estilo_titulo,
            ),
        ]
    ]
    t_enc = Table(filas_enc, colWidths=[5.5 * cm, 14.0 * cm])
    t_enc.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    elementos.append(t_enc)
    elementos.append(Spacer(1, 0.2 * cm))

    nom_inst = datos_consulta["instructor_nombre"]
    doc_inst = datos_consulta.get("instructor_documento") or "Registrado en sistema"
    desglose_cat = (
        ", ".join(
            f"{k}: {v} clase(s)"
            for k, v in sorted(datos_consulta["clases_por_categoria"].items())
        )
        or "Sin clases en este período"
    )

    filas_resumen = [
        [
            Paragraph(f"<b>Instructor:</b> {nom_inst}", estilo_meta),
            Paragraph(f"<b>Documento:</b> {doc_inst}", estilo_meta),
        ],
        [
            Paragraph(f"<b>Período Consultado:</b> {datos_consulta['periodo_etiqueta']}", estilo_meta),
            Paragraph(f"<b>Fecha Generación:</b> {ahora_col}", estilo_meta),
        ],
        [
            Paragraph(
                f"<b>TOTAL CLASES DICTADAS (PARA PAGO):</b> <font size=9><b>{datos_consulta['total_clases']}</b></font> clase(s)",
                estilo_meta,
            ),
            Paragraph(f"<b>Desglose por Categoría:</b> {desglose_cat}", estilo_meta),
        ],
        [
            Paragraph(
                f"<b>Evidencias En Tiempo:</b> {datos_consulta['en_tiempo']}", estilo_meta
            ),
            Paragraph(
                f"<b>Evidencias Fuera de Tiempo:</b> {datos_consulta['fuera_tiempo']}",
                estilo_meta,
            ),
        ],
    ]
    t_res = Table(filas_resumen, colWidths=[9.75 * cm, 9.75 * cm])
    t_res.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#E8F5E9")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    elementos.append(t_res)
    elementos.append(Spacer(1, 0.25 * cm))

    # Si es consolidado de TODOS LOS INSTRUCTORES, agregar tabla resumen por instructor
    if datos_consulta.get("es_todos") and datos_consulta.get("clases_por_instructor"):
        filas_cons = [
            [
                Paragraph("<b>INSTRUCTOR</b>", estilo_cab),
                Paragraph("<b>TOTAL CLASES DICTADAS EN EL PERÍODO</b>", estilo_cab),
            ]
        ]
        for inst_k, cant_v in sorted(
            datos_consulta["clases_por_instructor"].items(),
            key=lambda x: x[1],
            reverse=True,
        ):
            filas_cons.append(
                [
                    Paragraph(f"<b>{inst_k}</b>", estilo_celda),
                    Paragraph(f"<b>{cant_v} clase(s)</b>", estilo_celda_c),
                ]
            )
        t_cons = Table(filas_cons, colWidths=[13.0 * cm, 6.5 * cm])
        t_cons.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAEAEA")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ]
            )
        )
        elementos.append(t_cons)
        elementos.append(Spacer(1, 0.25 * cm))

    # Tabla detallada de clases dictadas
    filas_det = [
        [
            Paragraph("<b>#</b>", estilo_cab),
            Paragraph("<b>Fecha / Hora</b>", estilo_cab),
            Paragraph("<b>Instructor</b>", estilo_cab),
            Paragraph("<b>Alumno</b>", estilo_cab),
            Paragraph("<b>Cédula Alumno</b>", estilo_cab),
            Paragraph("<b>Cat.</b>", estilo_cab),
            Paragraph("<b>Clase(s)</b>", estilo_cab),
            Paragraph("<b>Cant.</b>", estilo_cab),
            Paragraph("<b>Horario</b>", estilo_cab),
        ]
    ]

    for i, reg in enumerate(datos_consulta["registros"], start=1):
        fecha_h = f"{reg.get('Fecha_Reporte', '')} {reg.get('Hora_Reporte', '')}".strip()
        inst_r = str(reg.get("Instructor", "")).strip()
        alu_r = str(reg.get("Nombre_Alumno", "")).strip()
        doc_r = str(reg.get("ID_Alumno", "")).strip()
        cat_r = normalizar_categoria(str(reg.get("Categoria", "B1")))
        clases_r = str(reg.get("Clases_Reportadas", "")).strip()
        cant_r = reg.get("Cantidad_Clases_Calc", 1)
        hor_r = str(reg.get("Validacion_Horario", "EN TIEMPO")).strip()
        hor_corto = "Fuera de tiempo" if "FUERA" in hor_r.upper() else "En tiempo"

        filas_det.append(
            [
                Paragraph(str(i), estilo_celda_c),
                Paragraph(fecha_h, estilo_celda_c),
                Paragraph(inst_r, estilo_celda),
                Paragraph(alu_r, estilo_celda),
                Paragraph(doc_r, estilo_celda_c),
                Paragraph(cat_r, estilo_celda_c),
                Paragraph(clases_r, estilo_celda_c),
                Paragraph(f"<b>{cant_r}</b>", estilo_celda_c),
                Paragraph(hor_corto, estilo_celda_c),
            ]
        )

    if len(filas_det) == 1:
        filas_det.append(
            [
                Paragraph("-", estilo_celda_c),
                Paragraph("Sin registros en este período", estilo_celda_c),
                "",
                "",
                "",
                "",
                "",
                Paragraph("0", estilo_celda_c),
                "",
            ]
        )

    t_det = Table(
        filas_det,
        colWidths=[
            0.8 * cm,
            2.5 * cm,
            3.3 * cm,
            4.2 * cm,
            2.3 * cm,
            1.1 * cm,
            1.5 * cm,
            1.2 * cm,
            2.6 * cm,
        ],
        repeatRows=1,
    )
    t_det.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAEAEA")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ]
        )
    )
    elementos.append(t_det)
    elementos.append(Spacer(1, 0.6 * cm))

    filas_firmas = [
        [
            Paragraph(
                "________________________________________<br/><b>Firma del Instructor</b>",
                estilo_celda_c,
            ),
            Paragraph(
                "________________________________________<br/><b>Revisado / Liquidación Administración CEA Auto Oriente</b>",
                estilo_celda_c,
            ),
        ]
    ]
    t_firmas = Table(filas_firmas, colWidths=[9.75 * cm, 9.75 * cm])
    elementos.append(t_firmas)

    doc.build(elementos)
    buffer.seek(0)
    return buffer


def generar_pdf_inasistencias_individual(
    datos_alumno: Dict[str, Any], inasistencias: Optional[List[Dict[str, Any]]] = None
) -> io.BytesIO:
    """
    Genera el Reporte Individual en PDF de Control de Inasistencias de un Alumno,
    detallando cuántas veces faltó, qué clases/horas perdió, instructor programado y estado de pago.
    """
    if inasistencias is None and isinstance(datos_alumno, dict) and "inasistencias" in datos_alumno:
        inasistencias = datos_alumno.get("inasistencias") or []
        datos_alumno = datos_alumno.get("alumno") or {}
    inasistencias = inasistencias or []
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=1.0 * cm,
        rightMargin=1.0 * cm,
        topMargin=0.9 * cm,
        bottomMargin=0.9 * cm,
        title="Reporte Individual de Inasistencias - CEA Auto Oriente",
    )

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "TituloInaInd",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=12.0,
        alignment=TA_CENTER,
    )
    estilo_meta = ParagraphStyle(
        "MetaInaInd",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.0,
        leading=10.2,
        alignment=TA_LEFT,
    )
    estilo_cab = ParagraphStyle(
        "CabInaInd",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.2,
        leading=9.0,
        alignment=TA_CENTER,
    )
    estilo_celda = ParagraphStyle(
        "CeldaInaInd",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.2,
        leading=9.0,
        alignment=TA_LEFT,
    )
    estilo_celda_c = ParagraphStyle(
        "CeldaInaIndC",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.2,
        leading=9.0,
        alignment=TA_CENTER,
    )

    elementos = []
    logo_path = obtener_ruta_logo()
    if logo_path:
        celda_logo = Image(
            str(logo_path), width=4.6 * cm, height=1.25 * cm, kind="proportional"
        )
    else:
        celda_logo = Paragraph("<b>CEA AUTO ORIENTE S.A.S.</b>", estilo_titulo)

    ahora_col = datetime.now(TZ_COLOMBIA).strftime("%Y-%m-%d %H:%M")
    t_enc = Table(
        [
            [
                celda_logo,
                Paragraph(
                    "<b>CEA AUTO ORIENTE S.A.S.<br/>INFORME INDIVIDUAL DE CONTROL DE INASISTENCIAS A CLASES PRÁCTICAS</b>",
                    estilo_titulo,
                ),
            ]
        ],
        colWidths=[5.5 * cm, 14.0 * cm],
    )
    t_enc.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    elementos.append(t_enc)
    elementos.append(Spacer(1, 0.25 * cm))

    nom_alu = str(datos_alumno.get("Nombre_Completo") or datos_alumno.get("Nombre_Alumno") or "Alumno").strip()
    doc_alu = str(datos_alumno.get("Documento") or datos_alumno.get("Documento_Alumno") or "Registrado en sistema").strip()
    cat_alu = str(datos_alumno.get("Categoria") or "General").strip()

    total_veces = len(inasistencias)
    total_horas = sum(int(x.get("No_Horas_Int", 1)) for x in inasistencias)
    total_valor = sum(int(x.get("Valor_Total_Int", 0)) for x in inasistencias)
    total_pagado = sum(
        int(x.get("Valor_Total_Int", 0))
        for x in inasistencias
        if "cancelo" in _normalizar_texto(str(x.get("Estado", "")))
    )
    total_pendiente = total_valor - total_pagado

    filas_res = [
        [
            Paragraph(f"<b>Nombre del Alumno:</b> {nom_alu}", estilo_meta),
            Paragraph(f"<b>Documento (CC/TI):</b> {doc_alu}", estilo_meta),
        ],
        [
            Paragraph(f"<b>Categoría:</b> {cat_alu}", estilo_meta),
            Paragraph(f"<b>Fecha de Emisión:</b> {ahora_col}", estilo_meta),
        ],
        [
            Paragraph(
                f"<b>TOTAL VECES QUE NO ASISTIÓ:</b> <font size=9><b>{total_veces} vez/veces</b></font>",
                estilo_meta,
            ),
            Paragraph(
                f"<b>TOTAL HORAS PERDIDAS:</b> <font size=9><b>{total_horas} hora(s)</b></font>",
                estilo_meta,
            ),
        ],
        [
            Paragraph(f"<b>Valor Cancelado:</b> ${total_pagado:,}", estilo_meta),
            Paragraph(f"<b>Saldo Pendiente de Pago:</b> <b>${total_pendiente:,}</b>", estilo_meta),
        ],
    ]
    t_res = Table(filas_res, colWidths=[9.75 * cm, 9.75 * cm])
    t_res.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 2), (-1, 2), colors.HexColor("#FFF3E0")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 3.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
            ]
        )
    )
    elementos.append(t_res)
    elementos.append(Spacer(1, 0.25 * cm))

    filas_det = [
        [
            Paragraph("<b># / Cons.</b>", estilo_cab),
            Paragraph("<b>Fecha</b>", estilo_cab),
            Paragraph("<b>Hora</b>", estilo_cab),
            Paragraph("<b>Clase(s) Faltada(s)</b>", estilo_cab),
            Paragraph("<b>Instructor Programado</b>", estilo_cab),
            Paragraph("<b>No. Horas</b>", estilo_cab),
            Paragraph("<b>Valor Hora</b>", estilo_cab),
            Paragraph("<b>Subtotal</b>", estilo_cab),
            Paragraph("<b>Estado</b>", estilo_cab),
        ]
    ]

    for idx, ina in enumerate(inasistencias, start=1):
        consec = str(ina.get("Consecutivo", idx)).strip()
        fech = str(ina.get("Fecha", "")).strip()
        hor = str(ina.get("Hora", "")).strip()
        cl_f = str(ina.get("Clases_Faltadas", "No especificada")).strip() or "No especificada"
        inst = str(ina.get("Nombre_Instructor", "")).strip()
        horas_i = int(ina.get("No_Horas_Int", 1))
        val_i = int(ina.get("Valor_Hora_Int", 0))
        sub_i = int(ina.get("Valor_Total_Int", horas_i * val_i))
        est = str(ina.get("Estado", "Pendiente Pago")).strip()

        filas_det.append(
            [
                Paragraph(f"{idx} (#{consec})", estilo_celda_c),
                Paragraph(fech, estilo_celda_c),
                Paragraph(hor, estilo_celda_c),
                Paragraph(cl_f, estilo_celda_c),
                Paragraph(inst, estilo_celda),
                Paragraph(str(horas_i), estilo_celda_c),
                Paragraph(f"${val_i:,}", estilo_celda_c),
                Paragraph(f"<b>${sub_i:,}</b>", estilo_celda_c),
                Paragraph(est, estilo_celda_c),
            ]
        )

    if len(filas_det) == 1:
        filas_det.append(
            [
                Paragraph("-", estilo_celda_c),
                Paragraph("Sin inasistencias registradas", estilo_celda_c),
                "",
                "",
                "",
                Paragraph("0", estilo_celda_c),
                "$0",
                "$0",
                "",
            ]
        )

    t_det = Table(
        filas_det,
        colWidths=[
            1.6 * cm,
            2.1 * cm,
            1.7 * cm,
            2.5 * cm,
            4.5 * cm,
            1.4 * cm,
            1.8 * cm,
            1.8 * cm,
            2.1 * cm,
        ],
        repeatRows=1,
    )
    t_det.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAEAEA")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    elementos.append(t_det)
    elementos.append(Spacer(1, 0.7 * cm))

    t_firmas = Table(
        [
            [
                Paragraph(
                    f"________________________________________<br/><b>Firma del Alumno</b><br/>{nom_alu}",
                    estilo_celda_c,
                ),
                Paragraph(
                    "________________________________________<br/><b>Control de Clases Prácticas — CEA Auto Oriente S.A.S.</b>",
                    estilo_celda_c,
                ),
            ]
        ],
        colWidths=[9.75 * cm, 9.75 * cm],
    )
    elementos.append(t_firmas)

    doc.build(elementos)
    buffer.seek(0)
    return buffer


def generar_pdf_inasistencias_periodo(datos_consulta: Dict[str, Any]) -> io.BytesIO:
    """
    Genera el Reporte Consolidado en PDF de Inasistencias por Semana, Mes o Año
    para CEA Auto Oriente S.A.S.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        leftMargin=0.9 * cm,
        rightMargin=0.9 * cm,
        topMargin=0.9 * cm,
        bottomMargin=0.9 * cm,
        title="Reporte Consolidado de Inasistencias - CEA Auto Oriente",
    )

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "TituloInaPer",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=12.0,
        alignment=TA_CENTER,
    )
    estilo_meta = ParagraphStyle(
        "MetaInaPer",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.8,
        leading=10.0,
        alignment=TA_LEFT,
    )
    estilo_cab = ParagraphStyle(
        "CabInaPer",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=6.8,
        leading=8.5,
        alignment=TA_CENTER,
    )
    estilo_celda = ParagraphStyle(
        "CeldaInaPer",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.7,
        leading=8.3,
        alignment=TA_LEFT,
    )
    estilo_celda_c = ParagraphStyle(
        "CeldaInaPerC",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6.7,
        leading=8.3,
        alignment=TA_CENTER,
    )

    elementos = []
    logo_path = obtener_ruta_logo()
    if logo_path:
        celda_logo = Image(
            str(logo_path), width=4.6 * cm, height=1.25 * cm, kind="proportional"
        )
    else:
        celda_logo = Paragraph("<b>CEA AUTO ORIENTE S.A.S.</b>", estilo_titulo)

    ahora_col = datetime.now(TZ_COLOMBIA).strftime("%Y-%m-%d %H:%M")
    t_enc = Table(
        [
            [
                celda_logo,
                Paragraph(
                    "<b>CEA AUTO ORIENTE S.A.S.<br/>REPORTE CONSOLIDADO DE INASISTENCIAS A CLASES PRÁCTICAS</b>",
                    estilo_titulo,
                ),
            ]
        ],
        colWidths=[5.5 * cm, 14.2 * cm],
    )
    t_enc.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    elementos.append(t_enc)
    elementos.append(Spacer(1, 0.2 * cm))

    filas_res = [
        [
            Paragraph(f"<b>Período Consultado:</b> {datos_consulta['periodo_etiqueta']}", estilo_meta),
            Paragraph(f"<b>Fecha Generación:</b> {ahora_col}", estilo_meta),
        ],
        [
            Paragraph(
                f"<b>Total Inasistencias Registradas:</b> <b>{datos_consulta['total_registros']}</b> caso(s)",
                estilo_meta,
            ),
            Paragraph(
                f"<b>Total Horas No Asistidas:</b> <b>{datos_consulta['total_horas']}</b> hora(s)",
                estilo_meta,
            ),
        ],
        [
            Paragraph(
                f"<b>Cancelado en Efectivo:</b> ${datos_consulta['total_efectivo']:,}",
                estilo_meta,
            ),
            Paragraph(
                f"<b>Cancelado por Transferencia:</b> ${datos_consulta['total_transferencia']:,}",
                estilo_meta,
            ),
        ],
        [
            Paragraph(
                f"<b>Pendiente de Pago ({datos_consulta['casos_pendientes']} casos):</b> <b>${datos_consulta['total_pendiente']:,}</b>",
                estilo_meta,
            ),
            Paragraph(
                f"<b>Valor Total Novedades:</b> <b>${datos_consulta['total_valor']:,}</b>",
                estilo_meta,
            ),
        ],
    ]
    t_res = Table(filas_res, colWidths=[9.85 * cm, 9.85 * cm])
    t_res.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#FFF3E0")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    elementos.append(t_res)
    elementos.append(Spacer(1, 0.25 * cm))

    filas_det = [
        [
            Paragraph("<b>Cons.</b>", estilo_cab),
            Paragraph("<b>Fecha</b>", estilo_cab),
            Paragraph("<b>Hora</b>", estilo_cab),
            Paragraph("<b>Nombre Apellido Alumno</b>", estilo_cab),
            Paragraph("<b>Clase(s)</b>", estilo_cab),
            Paragraph("<b>Nombre Instructor</b>", estilo_cab),
            Paragraph("<b>Horas</b>", estilo_cab),
            Paragraph("<b>Valor Hora</b>", estilo_cab),
            Paragraph("<b>Total</b>", estilo_cab),
            Paragraph("<b>Estado</b>", estilo_cab),
        ]
    ]

    for reg in datos_consulta.get("registros", []):
        consec = str(reg.get("Consecutivo", "")).strip()
        fech = str(reg.get("Fecha", "")).strip()
        hor = str(reg.get("Hora", "")).strip()
        alu = str(reg.get("Nombre_Alumno", "")).strip()
        cl_f = str(reg.get("Clases_Faltadas", "")).strip()
        if not cl_f or cl_f.lower() == "no especificada":
            cl_f = "-"
        inst = str(reg.get("Nombre_Instructor", "")).strip()
        horas_i = int(reg.get("No_Horas_Int", 1))
        val_i = int(reg.get("Valor_Hora_Int", 0))
        sub_i = int(reg.get("Valor_Total_Int", horas_i * val_i))
        est = str(reg.get("Estado", "")).strip()

        filas_det.append(
            [
                Paragraph(consec, estilo_celda_c),
                Paragraph(fech, estilo_celda_c),
                Paragraph(hor, estilo_celda_c),
                Paragraph(alu, estilo_celda),
                Paragraph(cl_f, estilo_celda_c),
                Paragraph(inst, estilo_celda),
                Paragraph(str(horas_i), estilo_celda_c),
                Paragraph(f"${val_i:,}", estilo_celda_c),
                Paragraph(f"<b>${sub_i:,}</b>", estilo_celda_c),
                Paragraph(est, estilo_celda_c),
            ]
        )

    if len(filas_det) == 1:
        filas_det.append(
            [
                Paragraph("-", estilo_celda_c),
                Paragraph("Sin registros", estilo_celda_c),
                "",
                "",
                "",
                "",
                "0",
                "$0",
                "$0",
                "",
            ]
        )

    t_det = Table(
        filas_det,
        colWidths=[
            1.1 * cm,
            1.8 * cm,
            1.4 * cm,
            3.8 * cm,
            1.5 * cm,
            3.9 * cm,
            1.0 * cm,
            1.6 * cm,
            1.6 * cm,
            2.0 * cm,
        ],
        repeatRows=1,
    )
    t_det.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAEAEA")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
            ]
        )
    )
    elementos.append(t_det)
    elementos.append(Spacer(1, 0.5 * cm))

    img_firma_rep = obtener_imagen_firma_rep_legal(width_cm=3.2, height_cm=1.2)
    t_firmas = Table(
        [
            [img_firma_rep, ""],
            [
                Paragraph(
                    "________________________________________<br/><b>SEGUNDO MURILLO</b><br/>REPRESENTANTE LEGAL",
                    estilo_celda_c,
                ),
                Paragraph(
                    "________________________________________<br/><b>Control Administrativo — CEA Auto Oriente S.A.S.</b>",
                    estilo_celda_c,
                ),
            ],
        ],
        colWidths=[9.85 * cm, 9.85 * cm],
    )
    t_firmas.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "BOTTOM")]))
    elementos.append(t_firmas)

    doc.build(elementos)
    buffer.seek(0)
    return buffer


def generar_pdf_mantenimientos_pesv(datos_consulta: Dict[str, Any]) -> io.BytesIO:
    """
    Genera el PDF Oficial del Plan Estratégico de Seguridad Vial (PESV):
    Código: SISI- F2-08 | Versión: 1 | Fecha: 23/02/2023
    Cea Auto Oriente S.A.S. — RELACION MANTENIMIENTOS PREVENTIVOS
    Sistema de Gestión en Plan Estratégico de Seguridad Vial
    Con las 9 columnas oficiales requeridas por la plataforma y las firmas de
    SEGUNDO MURILLO (REPRESENTANTE LEGAL) y MARLY ALFONSO (REPRESENTANTE PESV).
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        leftMargin=1.0 * cm,
        rightMargin=1.0 * cm,
        topMargin=0.8 * cm,
        bottomMargin=0.8 * cm,
        title="SISI-F2-08 Relacion Mantenimientos Preventivos PESV - CEA Auto Oriente",
    )

    styles = getSampleStyleSheet()
    estilo_titulo = ParagraphStyle(
        "TituloPESV",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10.0,
        leading=12.5,
        alignment=TA_CENTER,
    )
    estilo_meta_izq = ParagraphStyle(
        "MetaPESVIzq",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8.0,
        leading=10.0,
        alignment=TA_LEFT,
    )
    estilo_meta_val = ParagraphStyle(
        "MetaPESVVal",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.0,
        leading=10.0,
        alignment=TA_CENTER,
    )
    estilo_cab = ParagraphStyle(
        "CabPESV",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.2,
        leading=9.0,
        alignment=TA_CENTER,
    )
    estilo_celda = ParagraphStyle(
        "CeldaPESV",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.2,
        leading=9.0,
        alignment=TA_LEFT,
    )
    estilo_celda_c = ParagraphStyle(
        "CeldaPESVC",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.2,
        leading=9.0,
        alignment=TA_CENTER,
    )

    elementos = []
    logo_path = obtener_ruta_logo()
    if logo_path:
        celda_logo = Image(
            str(logo_path), width=5.2 * cm, height=1.45 * cm, kind="proportional"
        )
    else:
        celda_logo = Paragraph("<b>AUTO ORIENTE</b>", estilo_titulo)

    # Encabezado idéntico al formato oficial SISI- F2-08
    filas_enc = [
        [
            Paragraph("<b>Código:</b>", estilo_meta_izq),
            Paragraph("SISI- F2-08", estilo_meta_val),
            Paragraph("<b>Cea Auto Oriente S.A.S.</b>", estilo_titulo),
            celda_logo,
        ],
        [
            Paragraph("<b>Versión:</b>", estilo_meta_izq),
            Paragraph("1", estilo_meta_val),
            Paragraph("<b>RELACION MANTENIMIENTOS PREVENTIVOS</b>", estilo_titulo),
            "",
        ],
        [
            Paragraph("<b>Fecha:</b>", estilo_meta_izq),
            Paragraph("23/02/2023", estilo_meta_val),
            Paragraph(
                "<b>Sistema de Gestión en Plan Estratégico de Seguridad Vial</b>",
                estilo_titulo,
            ),
            "",
        ],
    ]
    t_enc = Table(filas_enc, colWidths=[2.3 * cm, 3.2 * cm, 14.9 * cm, 5.5 * cm])
    t_enc.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.7, colors.black),
                ("SPAN", (3, 0), (3, 2)),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (3, 0), (3, 2), "CENTER"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    elementos.append(t_enc)
    elementos.append(Spacer(1, 0.25 * cm))

    # Subtítulo discreto de filtro aplicado si se consultó una placa o período específico
    placa_f = datos_consulta.get("placa_filtro", "TODAS")
    per_etq = datos_consulta.get("periodo_etiqueta", "")
    if placa_f not in ("TODAS", "TODOS", "") or datos_consulta.get("periodo_codigo") not in ("todo", "todos"):
        txt_filtro = f"<b>Período del Reporte:</b> {per_etq}"
        if placa_f not in ("TODAS", "TODOS", ""):
            txt_filtro += f" &nbsp;&nbsp;|&nbsp;&nbsp; <b>Vehículo (Placa):</b> {placa_f}"
        txt_filtro += f" &nbsp;&nbsp;|&nbsp;&nbsp; <b>Total Registros:</b> {datos_consulta.get('total_registros', 0)}"
        if datos_consulta.get("total_costo", 0) > 0:
            txt_filtro += f" &nbsp;&nbsp;|&nbsp;&nbsp; <b>Inversión Registrada:</b> ${datos_consulta['total_costo']:,}"
        elementos.append(Paragraph(txt_filtro, estilo_celda))
        elementos.append(Spacer(1, 0.15 * cm))

    # Tabla con las 9 columnas exactas del formato oficial SISI- F2-08
    filas_tabla = [
        [
            Paragraph("<b>FECHA</b>", estilo_cab),
            Paragraph("<b>PLACA</b>", estilo_cab),
            Paragraph("<b>TIPO<br/>MANTENIMIENTO</b>", estilo_cab),
            Paragraph("<b>ESTABLECIMIENTO</b>", estilo_cab),
            Paragraph("<b>NIT</b>", estilo_cab),
            Paragraph("<b>NOMBRE REAL.<br/>ACTIVIDAD</b>", estilo_cab),
            Paragraph("<b>Nombre</b>", estilo_cab),
            Paragraph("<b>ROL</b>", estilo_cab),
            Paragraph("<b>OBSERVACIONES</b>", estilo_cab),
        ]
    ]

    for reg in datos_consulta.get("registros", []):
        fech_r = str(reg.get("Fecha", "")).strip()
        placa_r = str(reg.get("Placa", "")).strip().upper()
        tipo_r = str(reg.get("Tipo_Mantenimiento", "Preventivo")).strip()
        estab_r = str(reg.get("Establecimiento", "")).strip()
        nit_r = str(reg.get("NIT_CC", "")).strip()
        act_r = str(reg.get("Actividad_Realizada", "Mantenimiento")).strip()
        nom_r = str(reg.get("Nombre_Tecnico", "")).strip()
        rol_r = str(reg.get("Rol", "Mecanico")).strip().title()
        obs_r = str(reg.get("Observaciones", "")).strip()

        # Si el registro tiene kilometraje, próximo cambio o costo > 0, anexarlo de forma limpia en observaciones si no está ya
        detalles_extra = []
        km_i = int(reg.get("Kilometraje_Int", 0))
        prox_i = int(reg.get("Proximo_Cambio_Int", 0))
        costo_i = int(reg.get("Costo_Int", 0))
        if km_i > 0:
            detalles_extra.append(f"Km: {km_i:,}")
        if prox_i > 0:
            detalles_extra.append(f"Próx. cambio: {prox_i:,} km")
        if costo_i > 0:
            detalles_extra.append(f"Costo: ${costo_i:,}")
        if detalles_extra:
            extra_str = " (" + " | ".join(detalles_extra) + ")"
            if extra_str not in obs_r:
                obs_r = (obs_r + extra_str).strip()

        filas_tabla.append(
            [
                Paragraph(fech_r, estilo_celda_c),
                Paragraph(f"<b>{placa_r}</b>", estilo_celda_c),
                Paragraph(tipo_r, estilo_celda_c),
                Paragraph(estab_r, estilo_celda_c),
                Paragraph(nit_r, estilo_celda_c),
                Paragraph(act_r, estilo_celda_c),
                Paragraph(nom_r, estilo_celda_c),
                Paragraph(rol_r, estilo_celda_c),
                Paragraph(obs_r, estilo_celda),
            ]
        )

    if len(filas_tabla) == 1:
        filas_tabla.append(
            [
                Paragraph("-", estilo_celda_c),
                Paragraph("-", estilo_celda_c),
                Paragraph("-", estilo_celda_c),
                Paragraph("Sin mantenimientos en este período", estilo_celda_c),
                Paragraph("-", estilo_celda_c),
                Paragraph("-", estilo_celda_c),
                Paragraph("-", estilo_celda_c),
                Paragraph("-", estilo_celda_c),
                Paragraph("-", estilo_celda_c),
            ]
        )

    t_mant = Table(
        filas_tabla,
        colWidths=[
            2.2 * cm,
            1.8 * cm,
            2.6 * cm,
            3.6 * cm,
            2.2 * cm,
            2.8 * cm,
            3.4 * cm,
            2.0 * cm,
            5.3 * cm,
        ],
        repeatRows=1,
    )
    t_mant.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.6, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAEAEA")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    elementos.append(t_mant)
    elementos.append(Spacer(1, 0.55 * cm))

    # Firmas oficiales al pie: SEGUNDO MURILLO (REPRESENTANTE LEGAL) y MARLY ALFONSO (REPRESENTANTE PESV)
    img_firma_segundo = obtener_imagen_firma_rep_legal(width_cm=3.5, height_cm=1.35)
    img_firma_marly = obtener_imagen_firma_pesv(width_cm=3.8, height_cm=1.35)

    filas_firmas = [
        [img_firma_segundo, img_firma_marly],
        [
            Paragraph(
                "<b>SEGUNDO MURILLO</b><br/><br/><b>REPRESENTANTE LEGAL</b>",
                estilo_celda_c,
            ),
            Paragraph(
                "<b>MARLY ALFONSO</b><br/><br/><b>REPRESENTANTE PESV</b>",
                estilo_celda_c,
            ),
        ],
    ]
    t_firmas = Table(filas_firmas, colWidths=[12.95 * cm, 12.95 * cm])
    t_firmas.setStyle(
        TableStyle(
            [
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "BOTTOM"),
                ("TOPPADDING", (0, 0), (-1, -1), 2),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
            ]
        )
    )
    elementos.append(t_firmas)

    doc.build(elementos)
    buffer.seek(0)
    return buffer

