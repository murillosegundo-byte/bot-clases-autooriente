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
SHEET_CIERRES = "Cierres_Jornada"
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
# IDs de usuarios administradores autorizados (incluye al director/admin principal)
ADMIN_USER_IDS: set = {6843014700}
# Lista de chats privados de administradores suscritos a alertas fuera de horario (/alertas)
ADMIN_CHATS_ALERTA: set = {6843014700}
# Lista de chats privados de administradores que desean copia en vivo de cada clase (/supervisar)
ADMIN_CHATS_SUPERVISION: set = set()
# Grupos donde opera el bot para verificar si quien habla por privado es Administrador del grupo
GRUPOS_CONOCIDOS: set = set()
if ADMIN_CHAT_ID_ENV.lstrip("-").isdigit():
    _admin_id_env = int(ADMIN_CHAT_ID_ENV)
    ADMIN_CHATS_ALERTA.add(_admin_id_env)
    if _admin_id_env > 0:
        ADMIN_USER_IDS.add(_admin_id_env)
    elif _admin_id_env < 0:
        GRUPOS_CONOCIDOS.add(_admin_id_env)


async def es_usuario_administrador(update: Update, context: ContextTypes.DEFAULT_TYPE) -> bool:
    """
    Verifica si quien ejecuta el comando es Administrador:
    1. Está en la lista de ADMIN_USER_IDS.
    2. Tiene rango de Creator / Administrator en el grupo actual o en cualquiera de los GRUPOS_CONOCIDOS.
    3. Su nombre en Telegram corresponde a Segundo / Administración.
    """
    user = update.effective_user
    chat = update.effective_chat
    if not user:
        return False
    if chat and chat.type in ("group", "supergroup"):
        GRUPOS_CONOCIDOS.add(chat.id)
    if user.id in ADMIN_USER_IDS:
        return True
    nombre_tg = (user.full_name or "").lower()
    if "segundo" in nombre_tg or "murillo" in nombre_tg or "autooriente" in nombre_tg:
        ADMIN_USER_IDS.add(user.id)
        return True
    if chat and chat.type in ("group", "supergroup"):
        try:
            miembro = await context.bot.get_chat_member(chat.id, user.id)
            if miembro.status in ("creator", "administrator"):
                ADMIN_USER_IDS.add(user.id)
                return True
        except Exception:
            pass
    elif chat and chat.type == "private":
        for gid in list(GRUPOS_CONOCIDOS):
            try:
                miembro = await context.bot.get_chat_member(gid, user.id)
                if miembro.status in ("creator", "administrator"):
                    ADMIN_USER_IDS.add(user.id)
                    return True
            except Exception:
                pass
    return False


async def _enviar_mensaje_seguro(
    bot,
    chat_id: int,
    text: str,
    reply_markup: Optional[InlineKeyboardMarkup] = None,
    reply_to_message_id: Optional[int] = None,
):
    """Envía mensaje con Markdown y reintenta sin formato si algún nombre tiene caracteres especiales."""
    try:
        return await bot.send_message(
            chat_id=chat_id,
            text=text,
            reply_markup=reply_markup,
            reply_to_message_id=reply_to_message_id,
            parse_mode="Markdown",
        )
    except Exception:
        texto_plano = text.replace("*", "").replace("`", "").replace("_", "")
        return await bot.send_message(
            chat_id=chat_id,
            text=texto_plano,
            reply_markup=reply_markup,
            reply_to_message_id=reply_to_message_id,
        )


async def _editar_mensaje_seguro(
    mensaje,
    text: str,
    reply_markup: Optional[InlineKeyboardMarkup] = None,
):
    """Edita mensaje con Markdown y reintenta en texto plano si falla el parseo."""
    try:
        return await mensaje.edit_text(
            text=text,
            reply_markup=reply_markup,
            parse_mode="Markdown",
        )
    except Exception as exc:
        if "Message is not modified" in str(exc):
            return mensaje
        texto_plano = text.replace("*", "").replace("`", "").replace("_", "")
        try:
            return await mensaje.edit_text(
                text=texto_plano,
                reply_markup=reply_markup,
            )
        except Exception:
            return mensaje


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
    r"motos?|motocicletas?|carros?|veh[ií]culos?|instructor(?:es)?|profes?(?:or)?|notas?"
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
        self.limpiar_nombres_contaminados()

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
                nombre = str(f.get("Nombre_Completo", "")).strip()
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

                # Sincronizar el estado en Sheets de acuerdo exclusivamente a las clases de ESTA categoría
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

                item_alu = dict(f)
                item_alu["Clases_Completadas"] = cant_completadas
                item_alu["Clases_Requeridas"] = total_req
                item_alu["Curso_Completado"] = ya_cumplio_minimo

                if modo == "refuerzo":
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
        self, documento: str, nombre_completo: str, categoria: str
    ) -> Dict[str, Any]:
        """
        Registra o recupera la matrícula de un alumno para la CATEGORÍA indicada.
        - Si el alumno ya está matriculado en esa misma categoría, devuelve esa matrícula.
        - Si el alumno ya existe en otra categoría (ej. A2 Moto) y ahora inicia un segundo curso
          (ej. C1 Carro), crea automáticamente una nueva fila de matrícula (MAT-XXX) para la
          nueva categoría sin sobrescribir ni bloquear el otro curso.
        """
        cat_norm = normalizar_categoria(categoria)
        nombre_depurado = limpiar_solo_nombre_persona(nombre_completo) or nombre_completo.strip().title()

        # 1. Buscar si ya tiene matrícula en ESTA MISMA categoría
        existente_misma_cat = (
            self.buscar_alumno(documento, categoria=cat_norm, solo_misma_categoria=True)
            or self.buscar_alumno(nombre_depurado, categoria=cat_norm, solo_misma_categoria=True)
        )
        if existente_misma_cat:
            nom_actual = str(existente_misma_cat.get("Nombre_Completo", "")).strip()
            nom_limpio = limpiar_solo_nombre_persona(nom_actual) or nom_actual
            existente_misma_cat["Nombre_Completo"] = nom_limpio
            existente_misma_cat["Categoria"] = cat_norm
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

        try:
            ws = self.obtener_o_crear_hoja(SHEET_ALUMNOS, HEADERS_ALUMNOS)
            filas = self._leer_filas_diccionario(ws, HEADERS_ALUMNOS)
            max_num = len(filas)
            for f in filas:
                m_id = re.search(r"(\d+)", str(f.get("ID_Matricula", "")))
                if m_id:
                    max_num = max(max_num, int(m_id.group(1)))
            nuevo_id = f"MAT-{max_num + 1:03d}"
            nueva_fila = [nuevo_id, documento, nombre_depurado, cat_norm, "En curso"]
            self._agregar_fila_al_final(ws, nueva_fila)
            return {
                "ID_Matricula": nuevo_id,
                "Documento": documento,
                "Nombre_Completo": nombre_depurado,
                "Categoria": cat_norm,
                "Estado": "En curso",
            }
        except Exception as exc:
            logger.warning("No se pudo agregar automáticamente en hoja Alumnos: %s", exc)
            return {
                "ID_Matricula": documento,
                "Documento": documento,
                "Nombre_Completo": nombre_depurado,
                "Categoria": cat_norm,
                "Estado": "En curso",
            }

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

    # Agregar filas de Clases Adicionales / de Refuerzo al final de la tabla si existen
    for fila_r in registros:
        cl_raw = str(fila_r.get("Clases_Reportadas", "")).strip()
        tm_raw = str(fila_r.get("Tema_Registrado", "")).strip()
        if "refuerzo" in cl_raw.lower() or "adicional" in cl_raw.lower() or "refuerzo" in tm_raw.lower():
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
    for fila_r in registros:
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

    contenido_observaciones = (
        "<br/>".join(lineas_observaciones) if lineas_observaciones else "&nbsp;"
    )

    # Determinar el último instructor que dictó la clase final del curso (y su cédula oficial)
    ultimo_instructor = ""
    if registros:
        ultimo_instructor = str(registros[-1].get("Instructor", "")).strip()
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
            InlineKeyboardButton(
                "📋 Cierre de Jornada (Pantallazo Agenda del Día)",
                callback_data=f"cierre_jornada|{owner_id}",
            ),
        ],
        [
            InlineKeyboardButton(
                "📊 Consultar mis clases (Día/Semana/Mes)",
                callback_data=f"ver_mis_clases|semana|{owner_id}",
            ),
        ],
        [
            InlineKeyboardButton(
                "👨‍🏫 Cambiar / Registrar otro Instructor",
                callback_data=f"cambiar_inst|{owner_id}",
            ),
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


def teclado_seleccion_alumno(
    categoria: str, owner_id: int, pagina: int = 0, filtro_cat: Optional[str] = None
) -> InlineKeyboardMarkup:
    cat_consulta = filtro_cat or categoria
    es_modo_refuerzo = cat_consulta == "REFUERZO"
    if es_modo_refuerzo:
        alumnos = sheets_service.listar_alumnos_activos(categoria=categoria, modo="refuerzo")
        if not alumnos:
            alumnos = sheets_service.listar_alumnos_activos(categoria="TODAS", modo="refuerzo")
    else:
        alumnos = sheets_service.listar_alumnos_activos(categoria=cat_consulta, modo="activos")

    por_pagina = 25
    total_alumnos = len(alumnos)
    total_paginas = max(1, (total_alumnos + por_pagina - 1) // por_pagina)
    pagina = max(0, min(pagina, total_paginas - 1))

    inicio = pagina * por_pagina
    fin = inicio + por_pagina
    alumnos_pagina = alumnos[inicio:fin]

    botones = [
        [
            InlineKeyboardButton(
                "➕ Alumno Nuevo (Escribir Nombre y Cédula)",
                callback_data=f"nuevo_alu|{categoria}|{owner_id}",
            )
        ]
    ]

    for f in alumnos_pagina:
        nombre = str(f.get("Nombre_Completo", "")).strip()
        doc = str(f.get("Documento", "")).strip()
        cat_alu = str(f.get("Categoria", "")).strip()
        cl_comp = f.get("Clases_Completadas", 0)
        cl_req = f.get("Clases_Requeridas", 20)
        if nombre and doc:
            etiqueta_cat = f" [{cat_alu}]" if (cat_consulta in ("TODAS", "REFUERZO") and cat_alu) else ""
            if es_modo_refuerzo:
                texto_btn = f"🔄 {nombre}{etiqueta_cat} ({cl_comp}/{cl_req} Fin.) — CC: {doc}"
                cb_alu = f"alu_ref|{doc}|{owner_id}"
            else:
                texto_btn = f"👤 {nombre}{etiqueta_cat} ({cl_comp}/{cl_req} cl.) — CC: {doc}"
                cb_alu = f"alu|{doc}|{owner_id}"
            botones.append(
                [
                    InlineKeyboardButton(
                        texto_btn,
                        callback_data=cb_alu,
                    )
                ]
            )

    if total_paginas > 1:
        fila_nav = []
        if pagina > 0:
            fila_nav.append(
                InlineKeyboardButton(
                    "⬅️ Anterior",
                    callback_data=f"alu_pag|{cat_consulta}|{pagina - 1}|{owner_id}",
                )
            )
        fila_nav.append(
            InlineKeyboardButton(
                f"📄 Pág {pagina + 1}/{total_paginas} ({total_alumnos})",
                callback_data=f"alu_pag|{cat_consulta}|{pagina}|{owner_id}",
            )
        )
        if pagina < total_paginas - 1:
            fila_nav.append(
                InlineKeyboardButton(
                    "Siguiente ➡️",
                    callback_data=f"alu_pag|{cat_consulta}|{pagina + 1}|{owner_id}",
                )
            )
        botones.append(fila_nav)

    if es_modo_refuerzo:
        botones.append(
            [
                InlineKeyboardButton(
                    "⬅️ Volver a Alumnos en Curso (Pendientes de horas)",
                    callback_data=f"alu_pag|{categoria}|0|{owner_id}",
                )
            ]
        )
    else:
        if cat_consulta != "TODAS":
            botones.append(
                [
                    InlineKeyboardButton(
                        "👥 Ver alumnos en curso de todas las categorías",
                        callback_data=f"alu_pag|TODAS|0|{owner_id}",
                    )
                ]
            )
        else:
            botones.append(
                [
                    InlineKeyboardButton(
                        f"🏷 Filtrar solo categoría {categoria}",
                        callback_data=f"alu_pag|{categoria}|0|{owner_id}",
                    )
                ]
            )
        botones.append(
            [
                InlineKeyboardButton(
                    "🔄 Clase de Refuerzo (Buscar Alumno Finalizado)",
                    callback_data=f"alu_pag|REFUERZO|0|{owner_id}",
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


def _recuperar_sesion_desde_texto_mensaje(sesion: Dict[str, Any], texto_msg: str) -> None:
    """
    Si el servidor en Render se reinició mientras el instructor tenía el menú abierto en Telegram,
    recupera la Categoría, Clase(s) y Tema directamente del texto visible del mensaje del bot
    para evitar cualquier error de 'NoneType'.
    """
    if not texto_msg:
        return

    if not sesion.get("categoria"):
        m_cat = re.search(
            r"Categor[ií]a:\s*([^\n|]+)", texto_msg, flags=re.IGNORECASE
        )
        if m_cat:
            sesion["categoria"] = normalizar_categoria(m_cat.group(1))

    if not sesion.get("clases_texto"):
        m_cl = re.search(
            r"Clase\(s\)(?:\s+marcada\(s\))?:\s*([0-9,\-\s]+)",
            texto_msg,
            flags=re.IGNORECASE,
        )
        if m_cl:
            nums = extraer_numeros_clase(m_cl.group(1))
            if nums:
                cat_k = sesion.get("categoria") or "B1"
                cfg = FORMATOS_OFICIALES.get(cat_k, FORMATOS_OFICIALES["B1"])
                temas_rec: List[str] = []
                for fila in cfg["filas"]:
                    if any(c in nums for c in fila["clases"]):
                        t_l = fila["tema"].replace("\n", " ")
                        if t_l not in temas_rec:
                            temas_rec.append(t_l)
                sesion["clases_texto"] = formatear_clases(nums)
                sesion["cantidad_clases"] = len(nums)
                sesion["tema_oficial"] = (
                    " / ".join(temas_rec) if temas_rec else "Práctica de conducción"
                )


def construir_vista_cierre_jornada(
    sesion: Dict[str, Any], owner_id: int
) -> Tuple[str, InlineKeyboardMarkup]:
    """
    Construye la pantalla interactiva cuando el instructor envía al finalizar la jornada
    el pantallazo de 'Agenda práctica' con el listado de clases dictadas en el día.
    Cruza automáticamente con las fotos individuales registradas hoy por ese instructor.
    """
    instructor_nom = sesion.get("instructor_nombre") or "Instructor"
    datos_hoy = sheets_service.consultar_clases_por_instructor(
        instructor_nom, periodo="dia"
    )
    total_hoy = datos_hoy["total_clases"]
    registros_hoy = datos_hoy["registros"]

    if registros_hoy:
        lineas_hoy = []
        for idx, r in enumerate(registros_hoy[:20], start=1):
            hora_c = str(r.get("Hora_Reporte", ""))
            alu = str(r.get("Nombre_Alumno", ""))
            cat = normalizar_categoria(str(r.get("Categoria", "B1")))
            cl_rep = str(r.get("Clases_Reportadas", ""))
            cant = r.get("Cantidad_Clases_Calc", 1)
            lineas_hoy.append(
                f"`{idx:02d}.` `{hora_c}` — *{alu}* ({cat} Cl.{cl_rep} = *{cant} cl.*)"
            )
        detalle_txt = "\n".join(lineas_hoy)
    else:
        detalle_txt = "_Aún no figuran fotos individuales registradas hoy con tu nombre._"

    texto = (
        f"📋 *REPORTE DE CIERRE DE JORNADA (AGENDA PRÁCTICA)*\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"{_encabezado_instructor(sesion)}"
        f"📅 *Fecha y Hora de Cierre:* {sesion['fecha_reporte']} ({sesion['hora_reporte']})\n"
        f"📸 *Clases registradas hoy con foto en el bot:* `{total_hoy}` clase(s)\n\n"
        f"📝 *Tus clases con foto registradas hoy:*\n"
        f"{detalle_txt}\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👇 *Para finalizar tu Cierre de Jornada, toca cuántas clases/horas aparecen como \"Finalizada Ok\" en tu pantallazo de Agenda Práctica de hoy:*"
    )

    botones = []
    if total_hoy > 0:
        botones.append(
            [
                InlineKeyboardButton(
                    f"✅ Confirmar {total_hoy} clase(s) (Coincide exacto con Agenda)",
                    callback_data=f"guardar_cierre|{total_hoy}|{owner_id}",
                )
            ]
        )

    botones.extend(
        [
            [
                InlineKeyboardButton("2 cl.", callback_data=f"guardar_cierre|2|{owner_id}"),
                InlineKeyboardButton("4 cl.", callback_data=f"guardar_cierre|4|{owner_id}"),
                InlineKeyboardButton("6 cl.", callback_data=f"guardar_cierre|6|{owner_id}"),
                InlineKeyboardButton("8 cl.", callback_data=f"guardar_cierre|8|{owner_id}"),
            ],
            [
                InlineKeyboardButton("10 cl.", callback_data=f"guardar_cierre|10|{owner_id}"),
                InlineKeyboardButton("12 cl.", callback_data=f"guardar_cierre|12|{owner_id}"),
                InlineKeyboardButton("14 cl.", callback_data=f"guardar_cierre|14|{owner_id}"),
                InlineKeyboardButton("16 cl.", callback_data=f"guardar_cierre|16|{owner_id}"),
            ],
            [
                InlineKeyboardButton("1 cl.", callback_data=f"guardar_cierre|1|{owner_id}"),
                InlineKeyboardButton("3 cl.", callback_data=f"guardar_cierre|3|{owner_id}"),
                InlineKeyboardButton("5 cl.", callback_data=f"guardar_cierre|5|{owner_id}"),
                InlineKeyboardButton("7 cl.", callback_data=f"guardar_cierre|7|{owner_id}"),
                InlineKeyboardButton("9 cl.", callback_data=f"guardar_cierre|9|{owner_id}"),
                InlineKeyboardButton("11 cl.", callback_data=f"guardar_cierre|11|{owner_id}"),
            ],
            [
                InlineKeyboardButton(
                    "⬅️ Volver (Es foto de una clase individual)",
                    callback_data=f"volver_cat|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar|{owner_id}"),
            ],
        ]
    )
    return texto, InlineKeyboardMarkup(botones)


async def guardar_cierre_jornada_en_sheets(
    chat_id: int,
    user_id: int,
    clases_agenda: int,
    context: ContextTypes.DEFAULT_TYPE,
    mensaje_respuesta,
) -> None:
    """
    Guarda el pantallazo de 'Agenda práctica' al finalizar la jornada en la hoja 'Cierres_Jornada',
    compara las clases de la Agenda con las fotos individuales enviadas en el día,
    y envía copia privada del cierre al Administrador.
    """
    sesion = SESIONES_REPORTE.pop((chat_id, user_id), None)
    if not sesion:
        return

    menu_msg_id = sesion.get("menu_msg_id")
    if menu_msg_id:
        PROPIETARIO_MENSAJE_BOT.pop(menu_msg_id, None)

    try:
        instructor = sesion.get("instructor_nombre") or "Instructor"
        doc_instructor = sesion.get("instructor_documento") or ""

        datos_hoy = sheets_service.consultar_clases_por_instructor(
            instructor, periodo="dia"
        )
        clases_bot = datos_hoy["total_clases"]
        registros_hoy = datos_hoy["registros"]

        resumen_alumnos = [
            f"{r.get('Nombre_Alumno', '')} ({normalizar_categoria(str(r.get('Categoria', 'B1')))} Cl.{r.get('Clases_Reportadas', '')})"
            for r in registros_hoy
        ]
        detalle_alumnos_str = (
            ", ".join(resumen_alumnos) if resumen_alumnos else "Sin fotos individuales hoy"
        )

        if clases_agenda == clases_bot:
            estado_cruce = f"COINCIDE EXACTO ({clases_agenda} clases)"
            alerta_cruce_txt = (
                f"\n\n✅ *Cruce de Jornada Exitoso:* Las **{clases_agenda} clase(s)** de tu "
                f"Agenda Práctica coinciden exactamente con tus fotos registradas hoy."
            )
        else:
            estado_cruce = (
                f"DIFERENCIA (Agenda: {clases_agenda} cl. vs Fotos Bot: {clases_bot} cl.)"
            )
            alerta_cruce_txt = (
                f"\n\n⚠️ *Novedad en Cruce de Jornada:*\n"
                f"En tu captura de Agenda Práctica reportas **{clases_agenda} clase(s)** y en el bot "
                f"figuran **{clases_bot} clase(s)** registradas con foto hoy.\n"
                f"📩 *Se adjunta tu pantallazo de Agenda al administrador para validación.*"
            )

        url_captura = subir_foto_evidencia(
            imagen_bytes=sesion.get("foto_bytes"),
            documento_alumno=doc_instructor or "AGENDA",
            nombre_alumno=f"Cierre_{instructor}",
            clases_texto=f"{clases_agenda}cl",
            fecha_reporte=sesion["fecha_reporte"],
            fallback_file_id=sesion.get("file_id", ""),
            fallback_url=sesion.get("file_url", ""),
        )

        sheets_service.registrar_cierre_jornada(
            timestamp=sesion["timestamp"],
            fecha_jornada=sesion["fecha_reporte"],
            hora_cierre=sesion["hora_reporte"],
            instructor=instructor,
            documento_instructor=doc_instructor,
            clases_registradas_bot=clases_bot,
            clases_segun_agenda=clases_agenda,
            estado_cruce=estado_cruce,
            detalle_alumnos_dia=detalle_alumnos_str,
            url_captura_agenda=url_captura,
        )

        doc_inst_txt = f" (CC: `{doc_instructor}`)" if doc_instructor else ""
        texto_exito = (
            f"✅ *CIERRE DE JORNADA REGISTRADO EXITOSAMENTE*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"👨‍🏫 *Instructor:* {instructor}{doc_inst_txt}\n"
            f"📅 *Fecha y Hora de Cierre:* {sesion['fecha_reporte']} {sesion['hora_reporte']}\n"
            f"📋 *Clases según Agenda Práctica:* `{clases_agenda}` clase(s)\n"
            f"📸 *Clases con foto hoy en el Bot:* `{clases_bot}` clase(s)\n"
            f"👥 *Alumnos registrados hoy:* {detalle_alumnos_str}"
            f"{alerta_cruce_txt}"
        )
        await mensaje_respuesta.edit_text(texto_exito, parse_mode="Markdown")

        # Enviar el pantallazo de Cierre de Jornada (Agenda Práctica) en PRIVADO al Administrador
        destinos_admin = set(ADMIN_CHATS_ALERTA) | set(ADMIN_USER_IDS) | set(ADMIN_CHATS_SUPERVISION)
        for admin_chat in list(destinos_admin):
            if admin_chat != chat_id and admin_chat > 0:
                try:
                    if sesion.get("file_id"):
                        await context.bot.send_photo(
                            chat_id=admin_chat,
                            photo=sesion["file_id"],
                            caption=f"📋 *SOPORTE DE CIERRE DE JORNADA (AGENDA PRÁCTICA)*\n{texto_exito}",
                            parse_mode="Markdown",
                        )
                    else:
                        await context.bot.send_message(
                            chat_id=admin_chat,
                            text=f"📋 *SOPORTE DE CIERRE DE JORNADA (AGENDA PRÁCTICA)*\n{texto_exito}",
                            parse_mode="Markdown",
                        )
                except Exception:
                    pass

    except Exception as exc:
        logger.exception("Error guardando cierre de jornada: %s", exc)
        await mensaje_respuesta.edit_text(f"⚠️ Error al guardar el cierre de jornada: {exc}")


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
        # Recuperar categoría y clases del texto del mensaje si el servidor en Render se reinició
        _recuperar_sesion_desde_texto_mensaje(
            sesion, getattr(mensaje_respuesta, "text", "") or ""
        )

        categoria = sesion.get("categoria")
        clases_texto = sesion.get("clases_texto")

        if not categoria or not clases_texto:
            # Si tras un reinicio de servidor faltan la categoría o las clases, pedir seleccionarlas sin dar error
            SESIONES_REPORTE[(chat_id, user_id)] = sesion
            await mensaje_respuesta.edit_text(
                f"{_encabezado_instructor(sesion)}"
                f"🔄 *El servidor se actualizó hace un momento.*\n"
                f"Por favor vuelve a seleccionar la **Categoría** y la **Clase** en los botones de abajo para guardar tu reporte:",
                reply_markup=teclado_categorias(user_id),
                parse_mode="Markdown",
            )
            PROPIETARIO_MENSAJE_BOT[mensaje_respuesta.message_id] = user_id
            sesion["menu_msg_id"] = mensaje_respuesta.message_id
            return

        nombre_alumno = sesion.get("nombre_alumno_limpio") or "Desconocido"
        documento_alumno = sesion.get("documento_detectado") or "Pendiente"
        cantidad_clases = sesion.get("cantidad_clases") or 2
        tema_oficial = sesion.get("tema_oficial") or "Práctica de conducción"
        minuto_envio = sesion.get("minuto_envio", 45)
        doc_instructor = sesion.get("instructor_documento", "")
        es_refuerzo = bool(sesion.get("es_refuerzo", False))

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

        # 2.5 VALIDACIÓN ANTI-DUPLICADOS: No permitir registrar la misma clase dos veces salvo que sea Clase de Refuerzo
        if not es_refuerzo:
            duplicado = sheets_service.verificar_clase_duplicada(
                id_alumno=id_alumno_final,
                nombre_alumno=nombre_final,
                categoria=categoria,
                clases_reportadas=clases_texto,
            )
            if duplicado:
                # Restaurar la sesión para que el instructor pueda corregir la clase, elegir refuerzo o cancelar
                sesion["nombre_alumno_limpio"] = nombre_final
                sesion["documento_detectado"] = id_alumno_final
                SESIONES_REPORTE[(chat_id, user_id)] = sesion
                if menu_msg_id:
                    PROPIETARIO_MENSAJE_BOT[menu_msg_id] = user_id

                cfg = FORMATOS_OFICIALES.get(categoria, FORMATOS_OFICIALES["B1"])
                bloque_otras_cats = (
                    f"\n🔀 *Progreso en otros cursos del mismo alumno:*\n{duplicado['resumen_otras_txt']}\n"
                    if duplicado.get("resumen_otras_txt")
                    else ""
                )
                texto_dup = (
                    f"⛔ *CLASE DUPLICADA — NO SE PERMITE REGISTRAR DOS VECES LA MISMA CLASE*\n"
                    f"━━━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 *Alumno:* {nombre_final} (CC: `{id_alumno_final}`)\n"
                    f"🏷 *Categoría:* {cfg['nombre_corto']}\n"
                    f"⚠️ *Clase(s) que intentaste registrar:* `{clases_texto}`\n"
                    f"🚨 *Motivo del bloqueo:* La(s) clase(s) `{duplicado['clases_duplicadas']}` **YA FUE(RON) REGISTRADA(S)** para este alumno en *{cfg['nombre_corto']}* el `{duplicado['fecha_previa']}` a las `{duplicado['hora_previa']}` por *{duplicado['instructor_previo']}*.\n\n"
                    f"📚 *Clases ya registradas en {categoria}:* `{duplicado['todas_registradas_txt']}`\n"
                    f"💡 *Siguiente clase disponible en {categoria}:* `{duplicado['sugerencia_siguiente']}`\n"
                    f"{bloque_otras_cats}\n"
                    f"👇 *Selecciona qué deseas hacer:*"
                )
                botones_dup = [
                    [
                        InlineKeyboardButton(
                            f"📚 Cambiar Número de Clase en {categoria} (Sugerida: {duplicado['sugerencia_siguiente']})",
                            callback_data=f"cat|{categoria}|{user_id}",
                        )
                    ],
                ]
                for otra_cat in duplicado.get("otras_categorias", [])[:2]:
                    nom_otra = FORMATOS_OFICIALES.get(otra_cat, FORMATOS_OFICIALES["B1"])["nombre_corto"]
                    botones_dup.append(
                        [
                            InlineKeyboardButton(
                                f"🔀 Pasar clases previas a {nom_otra} y guardar esta en {categoria}",
                                callback_data=f"reclasificar_dup|{categoria}|{otra_cat}|{user_id}",
                            )
                        ]
                    )
                    botones_dup.append(
                        [
                            InlineKeyboardButton(
                                f"🏷 Registrar esta foto en {nom_otra} (Clase {clases_texto})",
                                callback_data=f"cambiar_cat_dup|{otra_cat}|{user_id}",
                            )
                        ]
                    )
                botones_dup.extend(
                    [
                        [
                            InlineKeyboardButton(
                                "🔄 Registrar como Clase de Refuerzo Adicional",
                                callback_data=f"confirmar_refuerzo|{user_id}",
                            )
                        ],
                        [
                            InlineKeyboardButton(
                                "👤 Cambiar de Alumno",
                                callback_data=f"cambiar_alu_dup|{categoria}|{user_id}",
                            )
                        ],
                        [
                            InlineKeyboardButton(
                                "✅ Finalizar / Cancelar registro duplicado",
                                callback_data=f"cancelar|{user_id}",
                            )
                        ],
                    ]
                )
                await _editar_mensaje_seguro(
                    mensaje_respuesta,
                    texto_dup,
                    reply_markup=InlineKeyboardMarkup(botones_dup),
                )
                PROPIETARIO_MENSAJE_BOT[mensaje_respuesta.message_id] = user_id
                sesion["menu_msg_id"] = mensaje_respuesta.message_id
                return

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

        clases_guardar = f"{clases_texto} (Refuerzo)" if es_refuerzo else clases_texto
        tema_guardar = f"[CLASE DE REFUERZO] {tema_oficial}" if es_refuerzo else tema_oficial

        # 5. Guardar en 'Registro_Clases'
        sheets_service.registrar_clase(
            timestamp=sesion["timestamp"],
            fecha_reporte=sesion["fecha_reporte"],
            hora_reporte=sesion["hora_reporte"],
            instructor=instructor,
            id_alumno=id_alumno_final,
            nombre_alumno=nombre_final,
            categoria=categoria,
            clases_reportadas=clases_guardar,
            tema_registrado=tema_guardar,
            url_foto_drive=url_foto_drive,
            validacion_horario=estado_horario_sheets,
        )

        # 5.5 Verificar progreso total del alumno EN ESTA CATEGORÍA y si acaba de cumplir el mínimo requerido
        cfg = FORMATOS_OFICIALES.get(categoria, FORMATOS_OFICIALES["B1"])
        total_requerido = cfg.get("total_clases", 20)
        resultado_alu = sheets_service.obtener_clases_alumno(
            id_alumno_final, categoria=categoria
        )
        registros_alu = resultado_alu.get("registros", [])
        mapa_clases_alu = mapear_clases_registradas(registros_alu)
        clases_completadas = len({c for c in mapa_clases_alu.keys() if 1 <= c <= total_requerido})
        curso_completado = clases_completadas >= total_requerido

        if curso_completado:
            sheets_service.actualizar_estado_alumno(
                id_alumno_final, "Finalizado", categoria=categoria
            )

        doc_inst_txt = f" (CC: `{doc_instructor}`)" if doc_instructor else ""
        etiqueta_ref = " 🔄 *(CLASE DE REFUERZO)*" if es_refuerzo else ""
        aviso_graduacion = (
            f"\n🎓 *¡CURSO PRÁCTICO COMPLETADO ({clases_completadas}/{total_requerido} CLASES)!*\n"
            f"El alumno ya cumplió el mínimo de horas reglamentarias y fue archivado como *Finalizado* "
            f"(se oculta del listado diario y se envía el PDF oficial para impresión y firmas de Alumno e Instructor)."
            if curso_completado
            else ""
        )
        texto_exito = (
            f"✅ *CLASE REGISTRADA EXITOSAMENTE*{etiqueta_ref}\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"👤 *Alumno:* {nombre_final}\n"
            f"🪪 *Documento (CC/TI):* `{id_alumno_final}`\n"
            f"🏷 *Categoría:* {cfg['nombre_corto']}\n"
            f"📚 *Clase(s) registrada(s):* `{clases_guardar}`\n"
            f"📊 *Progreso del curso:* `{clases_completadas}/{total_requerido}` clases\n"
            f"📝 *Tema Oficial:* {tema_guardar}\n"
            f"👨‍🏫 *Instructor:* {instructor}{doc_inst_txt}\n"
            f"🕒 *Fecha y Hora de envío:* {sesion['fecha_reporte']} {sesion['hora_reporte']}\n"
            f"{aviso_graduacion}"
            f"{alerta_horario_txt}"
        )
        await _editar_mensaje_seguro(mensaje_respuesta, texto_exito)

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

        # 5.8 SI EL ALUMNO CUMPLIÓ EL MÍNIMO DE CLASES (O RECIBIÓ REFUERZO POST-CURSO):
        # Generar y enviar automáticamente el PDF Oficial al Administrador y al Instructor para imprimir y firmar
        if curso_completado:
            try:
                datos_alumno_pdf = resultado_alu.get("alumno") or {
                    "ID_Matricula": id_alumno_final,
                    "Documento": id_alumno_final,
                    "Nombre_Completo": nombre_final,
                    "Categoria": categoria,
                    "Estado": "Finalizado",
                }
                datos_alumno_pdf["Categoria"] = categoria
                pdf_bytes = generar_pdf_seguimiento(datos_alumno_pdf, registros_alu).getvalue()
                codigo_fmt = cfg["codigo_formato"].replace(" ", "")
                nombre_arch = nombre_final.replace(" ", "_")
                filename_pdf = f"{codigo_fmt}_{categoria}_{nombre_arch}_FINALIZADO.pdf"
                caption_pdf = (
                    f"🎓 FORMATO OFICIAL LISTO PARA IMPRIMIR Y FIRMAR ({codigo_fmt})\n"
                    f"━━━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 Alumno: {nombre_final} (CC: {id_alumno_final})\n"
                    f"🏷 Categoría: {cfg['nombre_corto']} ({clases_completadas}/{total_requerido} clases completadas)\n"
                    f"👨‍🏫 Último Instructor: {instructor}{f' (CC: {doc_instructor})' if doc_instructor else ''}\n"
                    f"✍️ Incluye casillas de firma del Alumno, del Instructor y Recibido de Archivo."
                )

                # Enviar el PDF en el chat donde el instructor registró la última clase
                await context.bot.send_document(
                    chat_id=chat_id,
                    document=io.BytesIO(pdf_bytes),
                    filename=filename_pdf,
                    caption=caption_pdf,
                )

                # Enviar también al chat privado del Instructor (si el registro fue en grupo)
                if chat_id != user_id:
                    try:
                        await context.bot.send_document(
                            chat_id=user_id,
                            document=io.BytesIO(pdf_bytes),
                            filename=filename_pdf,
                            caption=caption_pdf,
                        )
                    except Exception:
                        pass

                # Enviar copia automática al chat privado de todos los Administradores
                destinos_admin = set(ADMIN_USER_IDS) | set(ADMIN_CHATS_ALERTA)
                for adm_id in destinos_admin:
                    if adm_id > 0 and adm_id not in (chat_id, user_id):
                        try:
                            await context.bot.send_document(
                                chat_id=adm_id,
                                document=io.BytesIO(pdf_bytes),
                                filename=filename_pdf,
                                caption=caption_pdf,
                            )
                        except Exception:
                            pass
            except Exception as exc_pdf:
                logger.warning("No se pudo enviar automáticamente el PDF de curso completado: %s", exc_pdf)

        # 6. Si el envío fue fuera de los minutos estipulados, notificar en PRIVADO a los administradores
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

        # 7. Si el administrador activó /supervisar en su chat privado, enviarle copia en vivo de la evidencia
        if ADMIN_CHATS_SUPERVISION:
            for sup_chat in list(ADMIN_CHATS_SUPERVISION):
                if sup_chat != chat_id:
                    try:
                        if sesion.get("file_id"):
                            await context.bot.send_photo(
                                chat_id=sup_chat,
                                photo=sesion["file_id"],
                                caption=f"👁️ *SUPERVISIÓN EN VIVO*\n{texto_exito}",
                                parse_mode="Markdown",
                            )
                        else:
                            await context.bot.send_message(
                                chat_id=sup_chat,
                                text=f"👁️ *SUPERVISIÓN EN VIVO*\n{texto_exito}",
                                parse_mode="Markdown",
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
    if nombre_cap and len(nombre_cap.split()) >= 2 and not doc_cap:
        ficha_cap = sheets_service.buscar_alumno(nombre_cap)
        if ficha_cap and str(ficha_cap.get("Documento", "")).strip():
            nombre_cap = str(ficha_cap.get("Nombre_Completo", nombre_cap)).strip()
            doc_cap = str(ficha_cap.get("Documento", "")).strip()

    tiene_datos_completos = bool(
        nombre_cap and len(nombre_cap.split()) >= 2 and doc_cap
    )
    es_captura_cierre = bool(
        re.search(
            r"\b(agenda|cierre|jornada|finalizada\s*ok|resumen)\b",
            caption,
            flags=re.IGNORECASE,
        )
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

    # Si el pie de foto indica que es el pantallazo de Agenda Práctica / Cierre de Jornada, abrir directo la vista de Cierre
    if es_captura_cierre:
        texto_cierre, teclado_cierre = construir_vista_cierre_jornada(sesion, user.id)
        msg_bot = await msg.reply_text(
            texto_cierre,
            reply_markup=teclado_cierre,
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


def _formatear_texto_liquidacion(datos: Dict[str, Any]) -> str:
    desglose_cat = (
        " | ".join(
            f"*{k}:* `{v}`" for k, v in sorted(datos["clases_por_categoria"].items())
        )
        or "Ninguna"
    )
    doc_txt = (
        f" (CC: `{datos['instructor_documento']}`)"
        if datos.get("instructor_documento")
        else ""
    )
    lineas = [
        "📊 *REPORTE DE CLASES DICTADAS (CONTROL SEMANAL / PAGO)*",
        "━━━━━━━━━━━━━━━━━━━━━━",
        f"👨‍🏫 *Instructor:* {datos['instructor_nombre']}{doc_txt}",
        f"📅 *Período:* {datos['periodo_etiqueta']}",
        f"✅ *TOTAL CLASES DICTADAS:* `{datos['total_clases']}` clase(s)",
        f"🏷 *Por Categoría:* {desglose_cat}",
        f"⏱️ *Horario:* En tiempo `{datos['en_tiempo']}` | Fuera de tiempo `{datos['fuera_tiempo']}`",
        "━━━━━━━━━━━━━━━━━━━━━━",
    ]

    if datos.get("es_todos") and datos.get("clases_por_instructor"):
        lineas.append("👥 *Consolidado por Instructor:*")
        for inst_n, cant_c in sorted(
            datos["clases_por_instructor"].items(), key=lambda x: x[1], reverse=True
        ):
            lineas.append(f"• *{inst_n}:* `{cant_c}` clase(s)")
        lineas.append("━━━━━━━━━━━━━━━━━━━━━━")

    registros = datos.get("registros", [])
    if not registros:
        lineas.append("ℹ️ _No hay clases registradas en este período._")
    else:
        lineas.append(f"📋 *Detalle de clases ({len(registros)} reporte(s)):*")
        for idx, r in enumerate(registros[:25], start=1):
            fecha_corta = str(r.get("Fecha_Reporte", ""))[-5:]
            hora_c = str(r.get("Hora_Reporte", ""))
            alu = str(r.get("Nombre_Alumno", ""))
            cat = normalizar_categoria(str(r.get("Categoria", "B1")))
            clases_rep = str(r.get("Clases_Reportadas", ""))
            cant = r.get("Cantidad_Clases_Calc", 1)
            inst_tag = f" [{r.get('Instructor', '')}]" if datos.get("es_todos") else ""
            lineas.append(
                f"`{idx:02d}.` `{fecha_corta} {hora_c}` — *{alu}* ({cat} Cl.{clases_rep} = *{cant} cl.*){inst_tag}"
            )
        if len(registros) > 25:
            lineas.append(
                f"_...y {len(registros) - 25} registro(s) más (descarga el PDF para ver la tabla completa)._"
            )

    return "\n".join(lineas)


def teclado_periodos_liquidacion(
    criterio_inst: str,
    periodo_actual: str,
    owner_id: int,
    es_admin: bool = False,
) -> InlineKeyboardMarkup:
    """Botones para cambiar entre Día (Hoy/Ayer), Semana, Semana Anterior, Mes y Descargar PDF."""
    criterio_seguro = criterio_inst[:24].replace("|", "")

    def _marca(p: str, texto: str) -> str:
        return f"🔘 {texto}" if periodo_actual == p else texto

    botones = [
        [
            InlineKeyboardButton(
                _marca("dia", "📅 Hoy"),
                callback_data=f"liq|{criterio_seguro}|dia|{owner_id}",
            ),
            InlineKeyboardButton(
                _marca("ayer", "🗓️ Ayer"),
                callback_data=f"liq|{criterio_seguro}|ayer|{owner_id}",
            ),
        ],
        [
            InlineKeyboardButton(
                _marca("semana", "📆 Esta Semana"),
                callback_data=f"liq|{criterio_seguro}|semana|{owner_id}",
            ),
            InlineKeyboardButton(
                _marca("semana_ant", "⏪ Semana Pasada"),
                callback_data=f"liq|{criterio_seguro}|semana_ant|{owner_id}",
            ),
        ],
        [
            InlineKeyboardButton(
                _marca("mes", "🗓️ Este Mes"),
                callback_data=f"liq|{criterio_seguro}|mes|{owner_id}",
            ),
        ],
        [
            InlineKeyboardButton(
                "📄 Descargar Relación en PDF",
                callback_data=f"liq_pdf|{criterio_seguro}|{periodo_actual}|{owner_id}",
            ),
        ],
    ]
    if es_admin:
        botones.append(
            [
                InlineKeyboardButton(
                    "👨‍🏫 Elegir otro Instructor",
                    callback_data=f"menu_nomina|{owner_id}",
                )
            ]
        )
    botones.append(
        [
            InlineKeyboardButton(
                "✅ Finalizar / Cerrar Consulta",
                callback_data=f"cerrar_consulta|{owner_id}",
            )
        ]
    )
    return InlineKeyboardMarkup(botones)


def teclado_panel_admin(owner_id: int) -> InlineKeyboardMarkup:
    """Panel principal interactivo con botones para el Administrador en su chat privado."""
    sup_activo = owner_id in ADMIN_CHATS_SUPERVISION
    alerta_activa = owner_id in ADMIN_CHATS_ALERTA
    txt_sup = "👁️ Supervisión en Vivo: ACTIVADA (Toca para apagar)" if sup_activo else "👁️ Activar Supervisión en Vivo (/supervisar)"
    txt_ale = "🔔 Alertas de Horario: ACTIVADAS" if alerta_activa else "🔔 Activar Alertas Fuera de Horario (/alertas)"
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    "📊 Liquidación / Clases por Instructor",
                    callback_data=f"menu_nomina|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    "👥 Consolidado Semanal (Todos los Instructores)",
                    callback_data=f"liq|TODOS|semana|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    "📋 Ver Cierres de Jornada (Agenda Práctica)",
                    callback_data=f"ver_cierres|semana|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    "📄 Consultar Alumno / Generar PDF Oficial",
                    callback_data=f"pedir_pdf_alumno|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    txt_sup,
                    callback_data=f"toggle_supervisar|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    txt_ale,
                    callback_data=f"toggle_alertas|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    "✅ Finalizar / Cerrar Panel",
                    callback_data=f"cerrar_consulta|{owner_id}",
                )
            ],
        ]
    )


def _formatear_texto_cierres_jornada(datos: Dict[str, Any]) -> str:
    cierres = datos.get("cierres", [])
    lineas = [
        "📋 *REPORTE DE CIERRES DE JORNADA (AGENDA PRÁCTICA)*",
        "━━━━━━━━━━━━━━━━━━━━━━",
        f"📅 *Período:* {datos.get('periodo_etiqueta', '')}",
        f"✅ *Total cierres recibidos:* `{len(cierres)}`",
        "━━━━━━━━━━━━━━━━━━━━━━",
    ]
    if not cierres:
        lineas.append("ℹ️ _No se han enviado pantallazos de Cierre de Jornada en este período._")
    else:
        for idx, c in enumerate(cierres[:20], start=1):
            f_j = str(c.get("Fecha_Jornada", ""))[-5:]
            h_j = str(c.get("Hora_Cierre", ""))
            inst = str(c.get("Instructor", ""))
            c_ag = str(c.get("Clases_En_Agenda", "0"))
            c_bot = str(c.get("Clases_Con_Foto_Bot", "0"))
            est = str(c.get("Estado_Cruce", ""))
            lineas.append(
                f"`{idx:02d}.` `{f_j} {h_j}` — *{inst}* | Agenda: `{c_ag}` vs Fotos: `{c_bot}` ({est})"
            )
    return "\n".join(lineas)


def teclado_periodos_cierres(periodo_actual: str, owner_id: int) -> InlineKeyboardMarkup:
    def _m(p: str, t: str) -> str:
        return f"🔘 {t}" if periodo_actual == p else t

    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(_m("dia", "📅 Hoy"), callback_data=f"ver_cierres|dia|{owner_id}"),
                InlineKeyboardButton(_m("ayer", "🗓️ Ayer"), callback_data=f"ver_cierres|ayer|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_m("semana", "📆 Esta Semana"), callback_data=f"ver_cierres|semana|{owner_id}"),
                InlineKeyboardButton(_m("semana_ant", "⏪ Semana Pasada"), callback_data=f"ver_cierres|semana_ant|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_m("mes", "🗓️ Este Mes"), callback_data=f"ver_cierres|mes|{owner_id}"),
            ],
            [
                InlineKeyboardButton("👑 Volver al Panel de Admin", callback_data=f"panel_admin|{owner_id}"),
            ],
            [
                InlineKeyboardButton("✅ Finalizar / Cerrar Consulta", callback_data=f"cerrar_consulta|{owner_id}"),
            ],
        ]
    )


def teclado_lista_instructores_admin(owner_id: int) -> InlineKeyboardMarkup:
    instructores = sheets_service.listar_instructores_registrados()
    botones = [
        [
            InlineKeyboardButton(
                "👥 TODOS LOS INSTRUCTORES (Consolidado Semanal)",
                callback_data=f"liq|TODOS|semana|{owner_id}",
            )
        ]
    ]
    for inst in instructores[:18]:
        nom = inst["nombre"]
        doc = inst.get("documento", "")
        criterio_corto = (doc if doc else nom)[:24].replace("|", "")
        etiqueta_btn = f"👨‍🏫 {nom}" + (f" ({doc})" if doc else "")
        botones.append(
            [
                InlineKeyboardButton(
                    etiqueta_btn,
                    callback_data=f"liq|{criterio_corto}|semana|{owner_id}",
                )
            ]
        )
    botones.append(
        [
            InlineKeyboardButton(
                "👑 Volver al Panel Principal Admin",
                callback_data=f"panel_admin|{owner_id}",
            )
        ]
    )
    botones.append(
        [
            InlineKeyboardButton(
                "✅ Finalizar / Cerrar Consulta",
                callback_data=f"cerrar_consulta|{owner_id}",
            )
        ]
    )
    return InlineKeyboardMarkup(botones)


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

    es_admin = await es_usuario_administrador(update, context)

    # 1. CONTROL DE TURNOS EN GRUPO: Validar que quien presiona el botón sea el instructor dueño de la foto (o el Administrador)
    owner_id = _extraer_owner_id_de_callback(data, msg_id)
    if owner_id is not None and user_id != owner_id and not es_admin:
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

    # Si un administrador interactúa con el menú de un instructor, mantener la sesión del instructor dueño de la foto
    sesion_uid = (
        owner_id
        if (owner_id is not None and (chat_id, owner_id) in SESIONES_REPORTE)
        else user_id
    )

    # Manejadores de consulta de clases dictadas y Panel de Administrador
    if accion_limpia == "cerrar_consulta":
        sesion_activa = SESIONES_REPORTE.get((chat_id, sesion_uid))
        if sesion_activa and sesion_activa.get("foto_bytes"):
            await _editar_mensaje_seguro(
                query.message,
                _texto_menu_paso_1(sesion_activa),
                reply_markup=teclado_categorias(sesion_uid),
            )
        else:
            PROPIETARIO_MENSAJE_BOT.pop(msg_id, None)
            await _editar_mensaje_seguro(
                query.message,
                "✅ *Consulta finalizada.*\n*(Escribe `/panel` o `/nomina` cuando desees abrir una nueva consulta)*.",
            )
        return

    if accion_limpia == "panel_admin":
        ADMIN_USER_IDS.add(user_id)
        await _editar_mensaje_seguro(
            query.message,
            "👑 *PANEL DE ADMINISTRACIÓN — CEA AUTO ORIENTE*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Selecciona la consulta o función administrativa que deseas ejecutar:",
            reply_markup=teclado_panel_admin(user_id),
        )
        return

    if accion_limpia == "menu_nomina":
        ADMIN_USER_IDS.add(user_id)
        await _editar_mensaje_seguro(
            query.message,
            "👑 *PANEL DE LIQUIDACIÓN Y CONTROL DE CLASES POR INSTRUCTOR*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Selecciona el instructor que deseas consultar (por Día, Semana o Mes):",
            reply_markup=teclado_lista_instructores_admin(user_id),
        )
        return

    if accion_limpia.startswith("ver_cierres|"):
        periodo = accion_limpia.split("|", 1)[1]
        datos_cie = sheets_service.consultar_cierres_jornada(periodo=periodo)
        await _editar_mensaje_seguro(
            query.message,
            _formatear_texto_cierres_jornada(datos_cie),
            reply_markup=teclado_periodos_cierres(periodo, user_id),
        )
        return

    if accion_limpia == "pedir_pdf_alumno":
        await _editar_mensaje_seguro(
            query.message,
            "📄 *CONSULTAR ALUMNO Y GENERAR FORMATO PDF OFICIAL*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Escribe ahora mismo aquí en el chat la **Cédula** o el **Nombre del Alumno** "
            "(o usa `/reporte 1095920107`):\n\n"
            "📌 *Ejemplo:*\n"
            "`Juan Manuel Melo` o `1095920107`\n\n"
            "*(Apenas lo envíes, el bot generará y te enviará el PDF oficial de seguimiento).* 👇",
            reply_markup=InlineKeyboardMarkup(
                [
                    [InlineKeyboardButton("👑 Volver al Panel de Admin", callback_data=f"panel_admin|{user_id}")],
                    [InlineKeyboardButton("✅ Finalizar / Cerrar Consulta", callback_data=f"cerrar_consulta|{user_id}")],
                ]
            ),
        )
        return

    if accion_limpia == "toggle_supervisar":
        ADMIN_USER_IDS.add(user_id)
        if user_id in ADMIN_CHATS_SUPERVISION:
            ADMIN_CHATS_SUPERVISION.discard(user_id)
            aviso_sup = "🔕 *Supervisión en vivo DESACTIVADA.*"
        else:
            ADMIN_CHATS_SUPERVISION.add(user_id)
            aviso_sup = (
                "👁️ *Supervisión en vivo ACTIVADA en este chat privado.*\n"
                "Recibirás copia en vivo de cada foto y clase que registren los instructores."
            )
        await _editar_mensaje_seguro(
            query.message,
            f"{aviso_sup}\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"👑 *PANEL DE ADMINISTRACIÓN — CEA AUTO ORIENTE*",
            reply_markup=teclado_panel_admin(user_id),
        )
        return

    if accion_limpia == "toggle_alertas":
        ADMIN_USER_IDS.add(user_id)
        ADMIN_CHATS_ALERTA.add(user_id)
        await _editar_mensaje_seguro(
            query.message,
            "🔔 *Alertas Fuera de Horario ACTIVADAS en este chat privado.*\n"
            "Recibirás aviso inmediato cuando un instructor envíe foto fuera del minuto 40-50 (2 clases) o 18-28 (1 clase).\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "👑 *PANEL DE ADMINISTRACIÓN — CEA AUTO ORIENTE*",
            reply_markup=teclado_panel_admin(user_id),
        )
        return

    if accion_limpia.startswith("ver_mis_clases|"):
        periodo = accion_limpia.split("|", 1)[1]
        perfil_inst = sheets_service.obtener_instructor_registrado(user_id)
        criterio_inst = (
            perfil_inst["nombre"]
            if perfil_inst
            else (query.from_user.full_name or "Instructor")
        )
        datos_liq = sheets_service.consultar_clases_por_instructor(
            criterio_inst, periodo=periodo
        )
        await _editar_mensaje_seguro(
            query.message,
            _formatear_texto_liquidacion(datos_liq),
            reply_markup=teclado_periodos_liquidacion(
                criterio_inst, periodo, user_id, es_admin=es_admin
            ),
        )
        return

    if accion_limpia.startswith("liq|"):
        _, criterio_inst, periodo = accion_limpia.split("|", 2)
        datos_liq = sheets_service.consultar_clases_por_instructor(
            criterio_inst, periodo=periodo
        )
        await _editar_mensaje_seguro(
            query.message,
            _formatear_texto_liquidacion(datos_liq),
            reply_markup=teclado_periodos_liquidacion(
                criterio_inst, periodo, user_id, es_admin=True if query.message.chat.type == "private" else es_admin
            ),
        )
        return

    if accion_limpia.startswith("liq_pdf|"):
        _, criterio_inst, periodo = accion_limpia.split("|", 2)
        datos_liq = sheets_service.consultar_clases_por_instructor(
            criterio_inst, periodo=periodo
        )
        pdf_buf = generar_pdf_liquidacion_instructor(datos_liq)
        nom_archivo_inst = datos_liq["instructor_nombre"].replace(" ", "_")
        filename = f"Liquidacion_Clases_{nom_archivo_inst}_{datos_liq['fecha_inicio']}_a_{datos_liq['fecha_fin']}.pdf"
        await context.bot.send_document(
            chat_id=chat_id if query.message.chat.type == "private" else user_id,
            document=pdf_buf,
            filename=filename,
            caption=(
                f"📄 Relación de Clases Dictadas — CEA Auto Oriente\n"
                f"👨‍🏫 Instructor: {datos_liq['instructor_nombre']}\n"
                f"📅 Período: {datos_liq['periodo_etiqueta']}\n"
                f"✅ Total Clases Dictadas: {datos_liq['total_clases']}"
            ),
        )
        return

    sesion = SESIONES_REPORTE.get((chat_id, sesion_uid))
    if not sesion:
        ahora = datetime.now(TZ_COLOMBIA)
        perfil_inst = sheets_service.obtener_instructor_registrado(sesion_uid)
        sesion = {
            "owner_id": sesion_uid,
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
        SESIONES_REPORTE[(chat_id, sesion_uid)] = sesion
        PROPIETARIO_MENSAJE_BOT[msg_id] = sesion_uid

    # Si la sesión perdió categoría o clases por reinicio de Render, recuperarla del texto visible del mensaje
    _recuperar_sesion_desde_texto_mensaje(
        sesion, getattr(query.message, "text", "") or ""
    )

    if accion_limpia == "cierre_jornada":
        texto_cierre, teclado_cierre = construir_vista_cierre_jornada(sesion, sesion_uid)
        await query.message.edit_text(
            texto_cierre,
            reply_markup=teclado_cierre,
            parse_mode="Markdown",
        )
        return

    if accion_limpia.startswith("guardar_cierre|"):
        cant_str = accion_limpia.split("|", 1)[1]
        clases_agenda = int(cant_str) if cant_str.isdigit() else 0
        await guardar_cierre_jornada_en_sheets(
            chat_id, sesion_uid, clases_agenda, context, query.message
        )
        return

    if accion_limpia == "cancelar":
        SESIONES_REPORTE.pop((chat_id, sesion_uid), None)
        PROPIETARIO_MENSAJE_BOT.pop(msg_id, None)
        await query.message.edit_text(
            f"❌ Registro cancelado por *{sesion.get('instructor_nombre', 'el instructor')}*.",
            parse_mode="Markdown",
        )
        return

    if accion_limpia == "cambiar_inst":
        sesion["esperando_registro_instructor"] = True
        sesion["esperando_nombre_texto"] = False
        sesion["menu_msg_id"] = msg_id
        await query.message.edit_text(
            "👨‍🏫 *Registrar / Cambiar Instructor Oficial*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Escribe ahora en el chat el **NOMBRE COMPLETO (Nombre y Apellido)** y el **NÚMERO DE CÉDULA** del instructor:\n\n"
            "📌 *Ejemplo:*\n"
            "`Jhoan Sebastián Carreño Sierra 1099739620`\n\n"
            "*(Se agregará en una fila nueva de la hoja `Instructores` sin borrar a los instructores anteriores).* 👇",
            parse_mode="Markdown",
        )
        return

    if accion_limpia == "volver_cat":
        sesion["indices_clases"] = []
        sesion["esperando_nombre_texto"] = False
        await query.message.edit_text(
            f"{_encabezado_instructor(sesion)}"
            f"👇 *Paso 1:* Selecciona la **Categoría** del curso:",
            reply_markup=teclado_categorias(sesion_uid),
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
            reply_markup=teclado_clases_por_categoria(cat_key, [], sesion_uid),
            parse_mode="Markdown",
        )
        return

    if accion_limpia.startswith("limpiar|"):
        cat_key = accion_limpia.split("|", 1)[1]
        sesion["indices_clases"] = []
        await query.message.edit_text(
            _instrucciones_submenu_clases(sesion, cat_key, []),
            reply_markup=teclado_clases_por_categoria(cat_key, [], sesion_uid),
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
                cat_key, sesion["indices_clases"], sesion_uid
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
            await guardar_en_sheets_y_confirmar(chat_id, sesion_uid, context, query.message)
        else:
            sesion["esperando_nombre_texto"] = True
            await query.message.edit_text(
                mensaje_indicaciones_alumno(sesion, cat_key, clases_txt, tema_txt),
                reply_markup=teclado_seleccion_alumno(cat_key, sesion_uid, pagina=0),
                parse_mode="Markdown",
            )
        return

    if accion_limpia == "confirmar_refuerzo":
        sesion["es_refuerzo"] = True
        sesion["esperando_nombre_texto"] = False
        await guardar_en_sheets_y_confirmar(chat_id, sesion_uid, context, query.message)
        return

    if accion_limpia.startswith("reclasificar_dup|"):
        _, cat_origen, cat_destino = accion_limpia.split("|", 2)
        id_alu = sesion.get("documento_detectado") or ""
        nom_alu = sesion.get("nombre_alumno_limpio") or ""
        # Asegurar que el alumno tenga matrícula en ambas categorías (ej. A2 y C1)
        if id_alu and nom_alu:
            sheets_service.registrar_o_actualizar_alumno(id_alu, nom_alu, cat_destino)
            sheets_service.registrar_o_actualizar_alumno(id_alu, nom_alu, cat_origen)
        sheets_service.reclasificar_clases_alumno(
            id_alumno=id_alu,
            nombre_alumno=nom_alu,
            cat_origen=cat_origen,
            cat_destino=cat_destino,
        )
        sesion["categoria"] = cat_origen
        sesion["es_refuerzo"] = False
        sesion["esperando_nombre_texto"] = False
        await guardar_en_sheets_y_confirmar(chat_id, sesion_uid, context, query.message)
        return

    if accion_limpia.startswith("cambiar_cat_dup|"):
        otra_cat = normalizar_categoria(accion_limpia.split("|", 1)[1])
        sesion["categoria"] = otra_cat
        nums_actuales = extraer_numeros_clase(sesion.get("clases_texto") or "")
        cfg_otra = FORMATOS_OFICIALES.get(otra_cat, FORMATOS_OFICIALES["B1"])
        temas_nuevos: List[str] = []
        for fila_f in cfg_otra["filas"]:
            if any(c in nums_actuales for c in fila_f["clases"]):
                t_limp = fila_f["tema"].replace("\n", " ")
                if t_limp not in temas_nuevos:
                    temas_nuevos.append(t_limp)
        if temas_nuevos:
            sesion["tema_oficial"] = " / ".join(temas_nuevos)
        sesion["es_refuerzo"] = False
        sesion["esperando_nombre_texto"] = False
        await guardar_en_sheets_y_confirmar(chat_id, sesion_uid, context, query.message)
        return

    if accion_limpia.startswith("cambiar_alu_dup|"):
        cat_key = accion_limpia.split("|", 1)[1]
        sesion["nombre_alumno_limpio"] = None
        sesion["documento_detectado"] = None
        sesion["es_refuerzo"] = False
        sesion["esperando_nombre_texto"] = True
        clases_txt = sesion.get("clases_texto") or "1-2"
        tema_txt = sesion.get("tema_oficial") or ""
        await _editar_mensaje_seguro(
            query.message,
            mensaje_indicaciones_alumno(sesion, cat_key, clases_txt, tema_txt),
            reply_markup=teclado_seleccion_alumno(cat_key, sesion_uid, pagina=0),
        )
        return

    if accion_limpia.startswith("alu_pag|"):
        _, filtro_cat, pag_str = accion_limpia.split("|", 2)
        cat_key = sesion.get("categoria") or (
            filtro_cat if filtro_cat not in ("TODAS", "REFUERZO") else "B1"
        )
        clases_txt = sesion.get("clases_texto") or "1-2"
        tema_txt = sesion.get("tema_oficial") or ""
        sesion["esperando_nombre_texto"] = True
        sesion["es_refuerzo"] = filtro_cat == "REFUERZO"
        encabezado_extra = (
            "\n🔄 *MODO CLASE DE REFUERZO:* Mostrando alumnos que ya cumplieron el mínimo de horas para asignarles clase de refuerzo adicional."
            if filtro_cat == "REFUERZO"
            else ""
        )
        await _editar_mensaje_seguro(
            query.message,
            mensaje_indicaciones_alumno(sesion, cat_key, clases_txt, tema_txt) + encabezado_extra,
            reply_markup=teclado_seleccion_alumno(
                cat_key, sesion_uid, pagina=int(pag_str), filtro_cat=filtro_cat
            ),
        )
        return

    if accion_limpia.startswith("nuevo_alu|"):
        cat_key = accion_limpia.split("|", 1)[1]
        sesion["esperando_nombre_texto"] = True
        sesion["es_refuerzo"] = False
        clases_txt = sesion.get("clases_texto") or "1-2"
        cfg = FORMATOS_OFICIALES.get(cat_key, FORMATOS_OFICIALES["B1"])
        await _editar_mensaje_seguro(
            query.message,
            f"{_encabezado_instructor(sesion)}"
            f"✅ *Categoría:* {cfg['nombre_corto']} | *Clase(s):* `{clases_txt}`\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"➕ *REGISTRAR ALUMNO NUEVO Y FINALIZAR CLASE*\n\n"
            f"Escribe ahora mismo abajo en la barra de mensaje de Telegram el **NOMBRE COMPLETO** y el **NÚMERO DE CÉDULA** del alumno y envíalo:\n\n"
            f"📌 *Ejemplo:*\n"
            f"`Juan Manuel Melo Hortua 1095920107`\n\n"
            f"*(Apenas envíes el mensaje con el nombre y cédula, el bot registrará al alumno en la hoja Alumnos y guardará la clase automáticamente).* 👇",
            reply_markup=InlineKeyboardMarkup(
                [
                    [
                        InlineKeyboardButton(
                            "⬅️ Volver al listado de Alumnos",
                            callback_data=f"alu_pag|{cat_key}|0|{sesion_uid}",
                        )
                    ],
                    [
                        InlineKeyboardButton(
                            "❌ Cancelar", callback_data=f"cancelar|{sesion_uid}"
                        )
                    ],
                ]
            ),
        )
        return

    if accion_limpia.startswith("alu_ref|") or accion_limpia.startswith("alu|"):
        es_ref_btn = accion_limpia.startswith("alu_ref|")
        doc_alu = accion_limpia.split("|", 1)[1]
        ficha = sheets_service.buscar_alumno(
            doc_alu, categoria=sesion.get("categoria")
        )
        if ficha:
            sesion["nombre_alumno_limpio"] = str(ficha.get("Nombre_Completo", "")).strip()
            sesion["documento_detectado"] = str(ficha.get("Documento", doc_alu)).strip()
        else:
            sesion["nombre_alumno_limpio"] = doc_alu
            sesion["documento_detectado"] = doc_alu

        sesion["es_refuerzo"] = es_ref_btn
        sesion["esperando_nombre_texto"] = False
        await guardar_en_sheets_y_confirmar(chat_id, sesion_uid, context, query.message)
        return


async def _generar_y_enviar_pdf_alumno(
    criterio: str,
    destino_chat_id: int,
    user_id: int,
    context: ContextTypes.DEFAULT_TYPE,
) -> bool:
    """
    Busca un alumno por cédula o nombre y envía su(s) formato(s) PDF oficial(es) al chat indicado.
    Si el alumno está realizando más de un curso (ej. A2 Moto y C1 Carro), genera y envía
    automáticamente el PDF independiente de cada categoría.
    """
    aviso = await _enviar_mensaje_seguro(
        context.bot,
        chat_id=destino_chat_id,
        text=f"⏳ Consultando y generando formato PDF para *{criterio}*...",
    )
    try:
        # Detectar si el criterio incluye explícitamente una categoría (ej. '1098826670 C1' o 'Jeridson Gonzalez A2')
        m_cat_exp = re.search(r"\b(A1|A2|B1|C1|RC1|RECAT_C1)\b", criterio, flags=re.IGNORECASE)
        cat_explicita = normalizar_categoria(m_cat_exp.group(1)) if m_cat_exp else None
        criterio_limpio = (
            re.sub(r"\b(A1|A2|B1|C1|RC1|RECAT_C1)\b", "", criterio, flags=re.IGNORECASE).strip()
            if cat_explicita
            else criterio.strip()
        )

        resultado_general = sheets_service.obtener_clases_alumno(
            criterio_limpio, categoria=cat_explicita
        )
        datos_base = resultado_general["alumno"]
        registros_base = resultado_general["registros"]
        cats_detectadas = (
            [cat_explicita]
            if cat_explicita
            else (resultado_general.get("categorias_detectadas") or [])
        )

        if not datos_base and not registros_base:
            await _editar_mensaje_seguro(
                aviso,
                f"❌ No se encontraron registros de alumno para: *{criterio}*.",
                reply_markup=teclado_panel_admin(user_id),
            )
            return False

        if not cats_detectadas:
            cats_detectadas = [
                normalizar_categoria(str((datos_base or {}).get("Categoria", "B1")))
            ]

        enviados = 0
        for cat_item in cats_detectadas:
            res_cat = sheets_service.obtener_clases_alumno(
                criterio_limpio, categoria=cat_item
            )
            alu_cat = res_cat["alumno"] or dict(datos_base or {})
            alu_cat["Categoria"] = cat_item
            regs_cat = res_cat["registros"]

            # Si tiene múltiples matrículas y una aún no tiene clases ni fue pedida explícitamente, igual enviar su formato o priorizar las activas
            pdf_buffer = generar_pdf_seguimiento(alu_cat, regs_cat)
            codigo_fmt = FORMATOS_OFICIALES.get(cat_item, FORMATOS_OFICIALES["B1"])[
                "codigo_formato"
            ].replace(" ", "")
            nombre_limpio = (
                str(alu_cat.get("Nombre_Completo", criterio_limpio))
                .strip()
                .replace(" ", "_")
            )
            filename = f"{codigo_fmt}_{cat_item}_{nombre_limpio}.pdf"

            await context.bot.send_document(
                chat_id=destino_chat_id,
                document=pdf_buffer,
                filename=filename,
                caption=(
                    f"📄 Formato Oficial {codigo_fmt}\n"
                    f"👤 Alumno: {alu_cat.get('Nombre_Completo')}\n"
                    f"🪪 Documento: {alu_cat.get('Documento')}\n"
                    f"🏷 Categoría: {cat_item}\n"
                    f"📊 Reportes registrados en {cat_item}: {len(regs_cat)}"
                ),
            )
            enviados += 1

        try:
            await aviso.delete()
        except Exception:
            pass
        return enviados > 0
    except Exception as exc:
        logger.exception("Error generando reporte PDF: %s", exc)
        await _editar_mensaje_seguro(aviso, f"⚠️ Error al generar el PDF: {exc}")
        return False


async def procesar_texto_chat(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not msg.text:
        return

    texto = msg.text.strip()
    if texto.startswith("/"):
        return

    es_admin = await es_usuario_administrador(update, context)
    sesion_uid = user.id
    sesion = SESIONES_REPORTE.get((msg.chat_id, sesion_uid))

    # Si quien escribe no tiene sesión propia, verificar si respondió al mensaje del bot o si hay una sesión esperando nombre de alumno
    if not sesion:
        if msg.reply_to_message:
            owner_reply = PROPIETARIO_MENSAJE_BOT.get(msg.reply_to_message.message_id)
            if owner_reply and (owner_reply == user.id or es_admin):
                sesion_uid = owner_reply
                sesion = SESIONES_REPORTE.get((msg.chat_id, sesion_uid))

        if not sesion:
            nom_prueba, doc_prueba = extraer_nombre_y_cedula(texto)
            if nom_prueba and len(nom_prueba.split()) >= 2 and doc_prueba:
                # Buscar la sesión activa más reciente en este chat que esté esperando el nombre del alumno
                for (c_id, u_id), ses_cand in reversed(list(SESIONES_REPORTE.items())):
                    if c_id == msg.chat_id and ses_cand.get("esperando_nombre_texto"):
                        if es_admin or u_id == user.id:
                            sesion_uid = u_id
                            sesion = ses_cand
                            break

    # SI ESTÁ EN EL CHAT PRIVADO CON EL BOT Y NO TIENE UNA FOTO ABIERTA:
    # Permitir consultas directas de Administrador o Instructor sin quedarse en silencio jamás
    if not sesion:
        if chat and chat.type == "private":
            t_low = _normalizar_texto(texto)
            if t_low in ("nomina", "liquidacion", "instructores", "pago", "pagos"):
                await cmd_nomina_admin(update, context)
                return
            if t_low in ("supervisar", "supervision"):
                await cmd_supervisar_admin(update, context)
                return
            if t_low in ("alertas", "alerta"):
                await cmd_alertas_admin(update, context)
                return
            if t_low in ("cierres", "cierre", "agenda", "agenda practica"):
                await cmd_cierres_admin(update, context)
                return
            if t_low in ("clases", "mis clases"):
                if es_admin:
                    await cmd_nomina_admin(update, context)
                else:
                    await cmd_mis_clases(update, context)
                return

            # Verificar si el texto escrito coincide con un Instructor registrado
            instructores_reg = sheets_service.listar_instructores_registrados()
            for inst in instructores_reg:
                nom_i = _normalizar_texto(inst.get("nombre", ""))
                doc_i = str(inst.get("documento", "")).strip()
                if (t_low and len(t_low) >= 3 and t_low in nom_i) or (doc_i and texto == doc_i):
                    datos_liq = sheets_service.consultar_clases_por_instructor(
                        inst["nombre"], periodo="semana"
                    )
                    msg_resp = await _enviar_mensaje_seguro(
                        context.bot,
                        chat_id=msg.chat_id,
                        text=_formatear_texto_liquidacion(datos_liq),
                        reply_markup=teclado_periodos_liquidacion(
                            inst["nombre"], "semana", user.id, es_admin=True
                        ),
                        reply_to_message_id=msg.message_id,
                    )
                    PROPIETARIO_MENSAJE_BOT[msg_resp.message_id] = user.id
                    return

            # Verificar si el texto escrito coincide con un Alumno (Cédula o Nombre)
            ficha_alu = sheets_service.buscar_alumno(texto)
            if ficha_alu:
                await _generar_y_enviar_pdf_alumno(
                    criterio=str(ficha_alu.get("Documento") or texto),
                    destino_chat_id=msg.chat_id,
                    user_id=user.id,
                    context=context,
                )
                return

            # En cualquier otro caso en el chat privado, responder siempre con el Panel Interactivo
            ADMIN_USER_IDS.add(user.id)
            msg_panel = await _enviar_mensaje_seguro(
                context.bot,
                chat_id=msg.chat_id,
                text=(
                    "👑 *PANEL DE CONTROL Y CONSULTAS — CEA AUTO ORIENTE*\n"
                    "━━━━━━━━━━━━━━━━━━━━━━\n"
                    "Selecciona con un toque lo que deseas consultar o escribe directamente "
                    "el **Nombre/Cédula de un Alumno** o **Instructor**:"
                ),
                reply_markup=teclado_panel_admin(user.id),
                reply_to_message_id=msg.message_id,
            )
            PROPIETARIO_MENSAJE_BOT[msg_panel.message_id] = user.id
        return

    # CASO A: El instructor está registrando por primera vez su propio Nombre Oficial + Cédula
    if sesion.get("esperando_registro_instructor"):
        nombre_inst, doc_inst = extraer_nombre_y_cedula(texto)
        if not nombre_inst or len(nombre_inst.split()) < 2 or not doc_inst:
            await _enviar_mensaje_seguro(
                context.bot,
                chat_id=msg.chat_id,
                text=(
                    "⚠️ *Para registrarte como Instructor Oficial, envía tu Nombre Completo (Nombre y Apellido) y tu Cédula:*\n\n"
                    "📌 *Ejemplo:*\n"
                    "`Segundo Murillo 13959739`"
                ),
                reply_to_message_id=msg.message_id,
            )
            return

        info = sheets_service.guardar_instructor_oficial(
            telegram_user_id=sesion_uid,
            nombre_oficial=nombre_inst,
            documento_oficial=doc_inst,
        )
        sesion["instructor_nombre"] = info["nombre"]
        sesion["instructor_documento"] = info["documento"]
        sesion["esperando_registro_instructor"] = False

        try:
            await msg.delete()
        except Exception:
            pass

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
                    reply_markup=teclado_categorias(sesion_uid),
                    parse_mode="Markdown",
                )
                return
            except Exception:
                pass

        nuevo_menu = await _enviar_mensaje_seguro(
            context.bot,
            chat_id=msg.chat_id,
            text=texto_paso1,
            reply_markup=teclado_categorias(sesion_uid),
            reply_to_message_id=sesion.get("foto_msg_id"),
        )
        sesion["menu_msg_id"] = nuevo_menu.message_id
        PROPIETARIO_MENSAJE_BOT[nuevo_menu.message_id] = sesion_uid
        return

    # CASO B: El instructor (o administrador) está enviando el Nombre Completo y Cédula del Alumno para su foto
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
                try:
                    await msg.delete()
                except Exception:
                    pass
                aviso = await context.bot.send_message(
                    chat_id=msg.chat_id,
                    text="⏳ Guardando registro en Google Sheets...",
                    reply_to_message_id=sesion.get("foto_msg_id"),
                )
                await guardar_en_sheets_y_confirmar(msg.chat_id, sesion_uid, context, aviso)
                return

        if not nombre or len(nombre.split()) < 2 or not documento:
            faltantes = []
            if not nombre or len(nombre.split()) < 2:
                faltantes.append("• **Nombre y Apellido completos** (ej. *Sergio Andrés Pérez*)")
            if not documento:
                faltantes.append("• **Número de Cédula / Documento** (sin puntos, ej. *1098765432*)")

            lista_faltantes = "\n".join(faltantes)
            await _enviar_mensaje_seguro(
                context.bot,
                chat_id=msg.chat_id,
                text=(
                    f"👨‍🏫 *{sesion.get('instructor_nombre', 'Instructor')}, faltan datos del alumno:*\n"
                    f"{lista_faltantes}\n\n"
                    f"✍️ *Por favor envíalo completo en un solo mensaje así:*\n"
                    f"`Sergio Andrés Pérez 1098765432`"
                ),
                reply_to_message_id=msg.message_id,
            )
            return

        sesion["nombre_alumno_limpio"] = nombre
        sesion["documento_detectado"] = documento
        sesion["esperando_nombre_texto"] = False
        try:
            await msg.delete()
        except Exception:
            pass
        aviso = await context.bot.send_message(
            chat_id=msg.chat_id,
            text="⏳ Guardando alumno, instructor y clase en Google Sheets...",
            reply_to_message_id=sesion.get("foto_msg_id"),
        )
        await guardar_en_sheets_y_confirmar(msg.chat_id, sesion_uid, context, aviso)
        return


async def cmd_panel_admin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Abre el Panel Interactivo de Administración en el chat privado."""
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    es_admin = await es_usuario_administrador(update, context)
    if chat.type in ("group", "supergroup"):
        try:
            await msg.delete()
        except Exception:
            pass
        if not es_admin:
            return
    else:
        ADMIN_USER_IDS.add(user.id)

    destino = msg.chat_id if chat.type == "private" else user.id
    msg_resp = await _enviar_mensaje_seguro(
        context.bot,
        chat_id=destino,
        text=(
            "👑 *PANEL DE ADMINISTRACIÓN — CEA AUTO ORIENTE*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Selecciona la consulta o función administrativa que deseas ejecutar:"
        ),
        reply_markup=teclado_panel_admin(user.id),
    )
    PROPIETARIO_MENSAJE_BOT[msg_resp.message_id] = user.id


async def cmd_cierres_admin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Muestra los reportes de Cierre de Jornada (Agenda práctica) enviados por los instructores."""
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    es_admin = await es_usuario_administrador(update, context)
    if chat.type in ("group", "supergroup"):
        try:
            await msg.delete()
        except Exception:
            pass
        if not es_admin:
            return
    else:
        ADMIN_USER_IDS.add(user.id)

    destino = msg.chat_id if chat.type == "private" else user.id
    datos_cie = sheets_service.consultar_cierres_jornada(periodo="semana")
    msg_resp = await _enviar_mensaje_seguro(
        context.bot,
        chat_id=destino,
        text=_formatear_texto_cierres_jornada(datos_cie),
        reply_markup=teclado_periodos_cierres("semana", user.id),
    )
    PROPIETARIO_MENSAJE_BOT[msg_resp.message_id] = user.id


async def cmd_alertas_admin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Activa las notificaciones fuera de horario en el chat PRIVADO del administrador.
    """
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    es_admin = await es_usuario_administrador(update, context)
    if chat.type in ("group", "supergroup"):
        try:
            await msg.delete()
        except Exception:
            pass
        if not es_admin:
            return

    ADMIN_USER_IDS.add(user.id)
    ADMIN_CHATS_ALERTA.add(user.id)
    destino = msg.chat_id if chat.type == "private" else user.id
    msg_resp = await _enviar_mensaje_seguro(
        context.bot,
        chat_id=destino,
        text=(
            "🔔 *Alertas Privadas de Administrador ACTIVADAS.*\n"
            "Recibirás aquí en tu chat privado un aviso automático cada vez que un instructor envíe "
            "una foto fuera de los minutos estipulados (minuto 40 al 50 para 2 clases / minuto 18 al 28 para 1 clase)."
        ),
        reply_markup=teclado_panel_admin(user.id),
    )
    PROPIETARIO_MENSAJE_BOT[msg_resp.message_id] = user.id


async def cmd_supervisar_admin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Activa o desactiva en el chat privado del administrador la recepción en vivo
    de cada foto y clase registrada por cualquier instructor.
    """
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    es_admin = await es_usuario_administrador(update, context)
    if chat.type in ("group", "supergroup"):
        try:
            await msg.delete()
        except Exception:
            pass
        if not es_admin:
            return

    ADMIN_USER_IDS.add(user.id)
    if user.id in ADMIN_CHATS_SUPERVISION:
        ADMIN_CHATS_SUPERVISION.discard(user.id)
        estado_txt = "🔕 *Supervisión en vivo DESACTIVADA.* (Seguirás recibiendo las alertas de horario con `/alertas`)."
    else:
        ADMIN_CHATS_SUPERVISION.add(user.id)
        estado_txt = (
            "👁️ *Supervisión en vivo ACTIVADA en tu chat privado.*\n"
            "Cada vez que un instructor registre una clase en el grupo o por privado con el bot, "
            "recibirás aquí mismo una copia de la foto y los datos registrados."
        )

    destino = msg.chat_id if chat.type == "private" else user.id
    msg_resp = await _enviar_mensaje_seguro(
        context.bot,
        chat_id=destino,
        text=estado_txt,
        reply_markup=teclado_panel_admin(user.id),
    )
    PROPIETARIO_MENSAJE_BOT[msg_resp.message_id] = user.id


async def cmd_mis_clases(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Permite a cada instructor consultar sus propias clases dictadas (por Día, Semana o Mes)
    y descargar su PDF de relación de clases.
    """
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    es_admin = await es_usuario_administrador(update, context)
    perfil_inst = sheets_service.obtener_instructor_registrado(user.id)
    criterio_inst = (
        perfil_inst["nombre"]
        if perfil_inst
        else (user.full_name or "Instructor")
    )

    periodo = "semana"
    if context.args:
        arg0 = context.args[0].lower().strip()
        if arg0 in ("dia", "día", "hoy"):
            periodo = "dia"
        elif arg0 == "ayer":
            periodo = "ayer"
        elif arg0 in ("pasada", "anterior", "semana_ant"):
            periodo = "semana_ant"
        elif arg0 == "mes":
            periodo = "mes"

    datos_liq = sheets_service.consultar_clases_por_instructor(
        criterio_inst, periodo=periodo
    )
    msg_resp = await _enviar_mensaje_seguro(
        context.bot,
        chat_id=msg.chat_id,
        text=_formatear_texto_liquidacion(datos_liq),
        reply_markup=teclado_periodos_liquidacion(
            criterio_inst, periodo, user.id, es_admin=es_admin or (chat.type == "private")
        ),
        reply_to_message_id=msg.message_id,
    )
    PROPIETARIO_MENSAJE_BOT[msg_resp.message_id] = user.id


async def cmd_nomina_admin(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Exclusivo para Administradores (/nomina, /liquidacion, /clases_instructor):
    Muestra el panel para consultar las clases dictadas por cualquier instructor (o todos)
    por Día, Semana, Semana Anterior o Mes y generar el PDF de pago semanal.
    Si se escribe en un grupo, borra el comando y responde en el chat privado del Administrador.
    """
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    es_admin = await es_usuario_administrador(update, context)
    if chat.type in ("group", "supergroup"):
        try:
            await msg.delete()
        except Exception:
            pass
        if not es_admin:
            await cmd_mis_clases(update, context)
            return
    else:
        # En chat privado siempre permitir abrir el panel de liquidación / consulta
        ADMIN_USER_IDS.add(user.id)

    destino_privado = msg.chat_id if chat.type == "private" else user.id

    try:
        if context.args:
            criterio_inst = " ".join(context.args).strip()
            datos_liq = sheets_service.consultar_clases_por_instructor(
                criterio_inst, periodo="semana"
            )
            msg_resp = await _enviar_mensaje_seguro(
                context.bot,
                chat_id=destino_privado,
                text=_formatear_texto_liquidacion(datos_liq),
                reply_markup=teclado_periodos_liquidacion(
                    criterio_inst, "semana", user.id, es_admin=True
                ),
            )
            PROPIETARIO_MENSAJE_BOT[msg_resp.message_id] = user.id
            return

        msg_resp = await _enviar_mensaje_seguro(
            context.bot,
            chat_id=destino_privado,
            text=(
                "👑 *PANEL DE LIQUIDACIÓN Y CLASES POR INSTRUCTOR*\n"
                "━━━━━━━━━━━━━━━━━━━━━━\n"
                "Selecciona el instructor que deseas consultar (por Día, Semana o Mes) para su liquidación:"
            ),
            reply_markup=teclado_lista_instructores_admin(user.id),
        )
        PROPIETARIO_MENSAJE_BOT[msg_resp.message_id] = user.id
    except Exception as exc:
        logger.exception("Error en cmd_nomina_admin: %s", exc)
        await _enviar_mensaje_seguro(
            context.bot,
            chat_id=destino_privado,
            text=f"⚠️ Ocurrió un error al consultar Google Sheets: {exc}",
        )


async def cmd_ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    es_admin = await es_usuario_administrador(update, context)
    es_privado = chat.type == "private"

    bloque_admin = (
        "\n👑 *Comandos Exclusivos de Administración (Privados):*\n"
        "• `/panel` o `/admin` — Abre el Panel de Administración con botones.\n"
        "• `/nomina` o `/liquidacion` — Consultar clases dictadas por cualquier instructor (Día/Semana/Mes) y descargar PDF de pago.\n"
        "• `/cierres` — Ver pantallazos de Cierre de Jornada (Agenda Práctica).\n"
        "• `/reporte [Cédula o Nombre]` — Genera el PDF oficial del alumno y te lo envía por privado.\n"
        "• `/alertas` — Recibe alertas privadas cuando envíen fotos fuera de horario.\n"
        "• `/supervisar` — Recibe en tu privado copia en vivo de todas las clases.\n"
        "• `/inicializar` — Verifica y limpia las hojas en Google Sheets."
        if (es_admin or es_privado)
        else ""
    )

    texto = (
        "🚗 *Bot de Seguimiento de Clases — CEA Auto Oriente*\n\n"
        "👨‍🏫 *Vinculación Oficial del Instructor (1 sola vez):*\n"
        "`/soy [Tu Nombre Completo] [Tu Cédula]`\n"
        "*(Si aún no lo has hecho, el bot te lo pedirá automáticamente en tu primera foto).* \n\n"
        "📸 *Pasos para el Instructor:*\n"
        "1️⃣ Envía la **foto de la clase** directamente al chat con el bot.\n"
        "   • *Si son 2 clases:* enviar entre el **minuto 40 y el minuto 50**.\n"
        "   • *Si es 1 sola clase:* enviar entre el **minuto 18 y el minuto 28**.\n"
        "2️⃣ Toca la **Categoría** (`🏍️ A1 Moto`, `🏍️ A2 Moto`, `🚗 B1 Carro Particular`, `🚕 C1 Servicio Público`, `🔄 Recategorización C1`).\n"
        "3️⃣ Marca **1 clase** o **2 clases** (`✅`) y pulsa **`💾 CONFIRMAR CLASE(S)`**.\n"
        "4️⃣ Escribe el **Nombre Completo y Cédula** del alumno (ej. `Sergio Andrés Pérez 1098765432`).\n\n"
        "📊 *Consultar mis Clases Dictadas (Día / Semana / Mes):*\n"
        "`/mis_clases` _(Permite ver tus clases por día, semana o mes y descargar tu PDF soporte)_.\n"
        f"{bloque_admin}"
    )
    teclado_inicio = teclado_panel_admin(user.id) if (es_privado and es_admin) else None
    msg_resp = await _enviar_mensaje_seguro(
        context.bot,
        chat_id=msg.chat_id,
        text=texto,
        reply_markup=teclado_inicio,
        reply_to_message_id=msg.message_id,
    )
    if teclado_inicio:
        PROPIETARIO_MENSAJE_BOT[msg_resp.message_id] = user.id


async def cmd_inicializar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    es_admin = await es_usuario_administrador(update, context)
    if chat.type in ("group", "supergroup"):
        try:
            await msg.delete()
        except Exception:
            pass
        if not es_admin:
            return
    else:
        ADMIN_USER_IDS.add(user.id)

    destino = msg.chat_id if chat.type == "private" else user.id
    try:
        sheets_service.inicializar_estructura()
        sheets_service.limpiar_nombres_contaminados()
        await _enviar_mensaje_seguro(
            context.bot,
            chat_id=destino,
            text="✅ Hojas `Alumnos`, `Instructores`, `Registro_Clases` y `Cierres_Jornada` verificadas y depuradas en Google Sheets.",
            reply_markup=teclado_panel_admin(user.id),
        )
    except Exception as exc:
        await _enviar_mensaje_seguro(
            context.bot,
            chat_id=destino,
            text=f"❌ Error conectando con Google Sheets: {exc}",
        )


async def cmd_registrar_alumno(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Comando /alumno o /matricular para agregar un alumno directamente al listado de 'Alumnos'
    (ej: /alumno Juan Manuel Melo Hortua B1 1095920107).
    """
    msg = update.effective_message
    user = update.effective_user
    if not msg or not user:
        return

    texto_args = " ".join(context.args).strip() if context.args else ""
    if not texto_args:
        await _enviar_mensaje_seguro(
            context.bot,
            chat_id=msg.chat_id,
            text=(
                "👤 *Matricular / Agregar Alumno al Listado*\n"
                "━━━━━━━━━━━━━━━━━━━━━━\n"
                "Escribe el comando seguido del Nombre Completo, Categoría y Cédula:\n"
                "`/alumno Juan Manuel Melo Hortua B1 1095920107`"
            ),
            reply_to_message_id=msg.message_id,
        )
        return

    m_cat = re.search(r"\b(A1|A2|B1|C1|C2|RC1)\b", texto_args, flags=re.IGNORECASE)
    cat_detectada = normalizar_categoria(m_cat.group(1)) if m_cat else "B1"
    nombre_alu, doc_alu = extraer_nombre_y_cedula(texto_args)

    if not nombre_alu or len(nombre_alu.split()) < 2 or not doc_alu:
        await _enviar_mensaje_seguro(
            context.bot,
            chat_id=msg.chat_id,
            text=(
                "⚠️ Debes indicar el **Nombre Completo** y la **Cédula** del alumno.\n"
                "📌 *Ejemplo:*\n"
                "`/alumno Juan Manuel Melo Hortua B1 1095920107`"
            ),
            reply_to_message_id=msg.message_id,
        )
        return

    ficha = sheets_service.registrar_o_actualizar_alumno(
        documento=doc_alu,
        nombre_completo=nombre_alu,
        categoria=cat_detectada,
    )
    await _enviar_mensaje_seguro(
        context.bot,
        chat_id=msg.chat_id,
        text=(
            f"✅ *ALUMNO AGREGADO AL LISTADO EXITOSAMENTE*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"👤 *Alumno:* {ficha.get('Nombre_Completo')}\n"
            f"🪪 *Documento:* `{ficha.get('Documento')}`\n"
            f"🏷 *Categoría:* `{ficha.get('Categoria')}`\n\n"
            f"Ya aparece de primero en los botones para escogerlo al registrar la clase."
        ),
        reply_to_message_id=msg.message_id,
    )


async def cmd_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Exclusivo para Administradores:
    - Si se pide desde un grupo, borra el comando del grupo para que no aparezca en el chat general
      y envía el PDF directamente al chat privado del Administrador.
    """
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    es_admin = await es_usuario_administrador(update, context)

    if chat.type in ("group", "supergroup"):
        try:
            await msg.delete()
        except Exception:
            pass
        if not es_admin:
            try:
                await _enviar_mensaje_seguro(
                    context.bot,
                    chat_id=user.id,
                    text="⛔ *Acceso restringido:* Solo el administrador de CEA Auto Oriente puede generar los formatos PDF oficiales.",
                )
            except Exception:
                pass
            return

    ADMIN_USER_IDS.add(user.id)
    destino_privado = msg.chat_id if chat.type == "private" else user.id

    if not context.args:
        await _enviar_mensaje_seguro(
            context.bot,
            chat_id=destino_privado,
            text=(
                "📄 *GENERAR REPORTE PDF OFICIAL DE ALUMNO*\n"
                "━━━━━━━━━━━━━━━━━━━━━━\n"
                "Escribe el nombre o cédula del alumno (o envíalo directamente como mensaje aquí en el chat privado).\n"
                "📌 *Ejemplo:* `/reporte 1095304561` o escribe `1095304561`"
            ),
            reply_markup=teclado_panel_admin(user.id),
        )
        return

    criterio = " ".join(context.args).strip()
    await _generar_y_enviar_pdf_alumno(
        criterio=criterio,
        destino_chat_id=destino_privado,
        user_id=user.id,
        context=context,
    )


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
    threading.Thread(target=sheets_service.limpiar_nombres_contaminados, daemon=True).start()

    app = ApplicationBuilder().token(TELEGRAM_BOT_TOKEN).build()
    app.add_handler(CommandHandler(["start", "ayuda", "help"], cmd_ayuda))
    app.add_handler(CommandHandler(["panel", "admin", "administrador"], cmd_panel_admin))
    app.add_handler(CommandHandler(["cierres", "agenda"], cmd_cierres_admin))
    app.add_handler(CommandHandler(["soy", "instructor", "mi_nombre"], cmd_registrar_instructor))
    app.add_handler(CommandHandler(["alumno", "matricular"], cmd_registrar_alumno))
    app.add_handler(CommandHandler(["clase", "menu", "registrar"], cmd_menu_clase))
    app.add_handler(CommandHandler(["mis_clases", "misclases", "clases"], cmd_mis_clases))
    app.add_handler(CommandHandler(["nomina", "liquidacion", "clases_instructor"], cmd_nomina_admin))
    app.add_handler(CommandHandler("alertas", cmd_alertas_admin))
    app.add_handler(CommandHandler("supervisar", cmd_supervisar_admin))
    app.add_handler(CommandHandler("inicializar", cmd_inicializar))
    app.add_handler(CommandHandler("reporte", cmd_reporte))

    app.add_handler(CallbackQueryHandler(manejar_botones_callback))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.IMAGE, procesar_mensaje_imagen))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), procesar_texto_chat))

    logger.info("🚗 Bot CEA Auto Oriente iniciado...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
