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
SHEET_INASISTENCIAS = "Inasistencias"
SHEET_MANTENIMIENTOS = "Mantenimientos"
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

PLACAS_VEHICULOS_OFICIALES: List[Dict[str, str]] = [
    {"placa": "UDT984", "tipo": "🚗 Carro", "etiqueta": "🚗 UDT984"},
    {"placa": "FSP335", "tipo": "🚗 Carro", "etiqueta": "🚗 FSP335"},
    {"placa": "DUX342", "tipo": "🚗 Carro", "etiqueta": "🚗 DUX342"},
    {"placa": "YNE31F", "tipo": "🏍️ Moto", "etiqueta": "🏍️ YNE31F"},
    {"placa": "YNE30F", "tipo": "🏍️ Moto", "etiqueta": "🏍️ YNE30F"},
    {"placa": "AZA25G", "tipo": "🏍️ Moto", "etiqueta": "🏍️ AZA25G"},
]

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
        self.sembrar_datos_historicos_appsheet()
        self.limpiar_nombres_contaminados()

    def sembrar_datos_historicos_appsheet(self) -> None:
        """
        Si las hojas 'Inasistencias' o 'Mantenimientos' están recién creadas (sin datos),
        carga automáticamente el historial proveniente de la base de datos de AppSheet.
        """
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

    # Incluir automáticamente en OBSERVACIONES cuando el alumno no asiste y pierde la clase
    inasistencias_alu = sheets_service.obtener_inasistencias_alumno(
        criterio_alumno=documento_alumno or nombre_alumno,
        nombre_alternativo=nombre_alumno,
        categoria=categoria_key,
    )
    if inasistencias_alu:
        total_veces_falto = len(inasistencias_alu)
        total_horas_perdidas = sum(int(x.get("No_Horas_Int", 1)) for x in inasistencias_alu)
        lineas_observaciones.append(
            f"<b>⚠️ CONTROL DE INASISTENCIAS (CLASES PERDIDAS):</b> El alumno no asistió en "
            f"<b>{total_veces_falto} ocasión(es)</b> (Total: <b>{total_horas_perdidas} hora(s) perdida(s)</b>):"
        )
        for idx_ina, ina in enumerate(inasistencias_alu, start=1):
            fech_i = str(ina.get("Fecha", "")).strip()
            hor_i = str(ina.get("Hora", "")).strip()
            cl_falt = str(ina.get("Clases_Faltadas", "")).strip()
            if not cl_falt or cl_falt.lower() in ("no especificada", "s/n", "ninguna"):
                cl_falt_txt = f"{ina.get('No_Horas_Int', 1)} hora(s) práctica(s)"
            else:
                cl_falt_txt = f"Clase(s) {cl_falt} ({ina.get('No_Horas_Int', 1)} h)"
            inst_i = str(ina.get("Nombre_Instructor", "")).strip()
            est_i = str(ina.get("Estado", "Pendiente Pago")).strip()
            lineas_observaciones.append(
                f"  {idx_ina}. <b>Fecha {fech_i} ({hor_i}):</b> Faltó a <b>{cl_falt_txt}</b> "
                f"— Instructor programado: <b>{inst_i}</b> | Estado: <i>{est_i}</i>"
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


def generar_pdf_inasistencias_individual(
    datos_alumno: Dict[str, Any], inasistencias: List[Dict[str, Any]]
) -> io.BytesIO:
    """
    Genera el Reporte Individual en PDF de Control de Inasistencias de un Alumno,
    detallando cuántas veces faltó, qué clases/horas perdió, instructor programado y estado de pago.
    """
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
                "🚫 Reportar Inasistencia de Alumno (Perdió Clase)",
                callback_data=f"ina_nuevo_ini|{owner_id}",
            ),
        ],
        [
            InlineKeyboardButton(
                "🔧 Registrar Mantenimiento de Vehículo (PESV)",
                callback_data=f"mant_nuevo_ini|{owner_id}",
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
                    "🚫 Reportar Inasistencia de Alumno",
                    callback_data=f"ina_nuevo_ini|{owner_id}",
                ),
                InlineKeyboardButton(
                    "📊 Reportes PDF Inasistencias",
                    callback_data=f"ina_ver|todo|TODOS|{owner_id}",
                ),
            ],
            [
                InlineKeyboardButton(
                    "🔧 Registrar Mantenimiento Vehículo",
                    callback_data=f"mant_nuevo_ini|{owner_id}",
                ),
                InlineKeyboardButton(
                    "🛠️ Reporte PDF Mantenimientos (SISI-F2-08)",
                    callback_data=f"mant_ver|todo|TODAS|{owner_id}",
                ),
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


def _formatear_texto_inasistencias(datos: Dict[str, Any]) -> str:
    est_filtro = datos.get("filtro_estado", "TODOS")
    etiqueta_est = "Todos los estados" if est_filtro.upper() == "TODOS" else est_filtro
    lineas = [
        "🚫 *CONTROL Y REPORTE DE INASISTENCIAS DE ALUMNOS*",
        "━━━━━━━━━━━━━━━━━━━━━━",
        f"📅 *Período:* {datos.get('periodo_etiqueta', '')}",
        f"🏷 *Filtro Estado:* `{etiqueta_est}`",
        f"📌 *Total Inasistencias:* `{datos.get('total_registros', 0)}` caso(s) (`{datos.get('total_horas', 0)}` h perdidas)",
        f"💵 *Cancelado Efectivo:* `${datos.get('total_efectivo', 0):,}`",
        f"🏦 *Cancelado Transferencia:* `${datos.get('total_transferencia', 0):,}`",
        f"⏳ *Pendiente de Pago ({datos.get('casos_pendientes', 0)}):* `${datos.get('total_pendiente', 0):,}`",
        "━━━━━━━━━━━━━━━━━━━━━━",
    ]
    registros = datos.get("registros", [])
    if not registros:
        lineas.append("ℹ️ _No hay inasistencias registradas con este filtro en este período._")
    else:
        lineas.append(f"📋 *Últimos registros ({len(registros)} total):*")
        for r in list(reversed(registros))[:18]:
            cons = str(r.get("Consecutivo", "")).strip()
            fech = str(r.get("Fecha", "")).strip()
            alu = str(r.get("Nombre_Alumno", "")).strip()
            inst = str(r.get("Nombre_Instructor", "")).strip()
            hrs = r.get("No_Horas_Int", 1)
            sub = r.get("Valor_Total_Int", 0)
            est = str(r.get("Estado", "")).strip()
            icono_est = "⏳" if "pendiente" in est.lower() else ("💵" if "efectivo" in est.lower() else "🏦")
            cl_f = str(r.get("Clases_Faltadas", "")).strip()
            cl_tag = f" [Cl.{cl_f}]" if (cl_f and cl_f.lower() != "no especificada") else ""
            lineas.append(
                f"`#{cons}` `{fech}` — *{alu}*{cl_tag} ({hrs}h = `${sub:,}`) {icono_est} _{est}_ | 👨‍🏫 {inst[:18]}"
            )
        if len(registros) > 18:
            lineas.append(
                f"_...y {len(registros) - 18} registro(s) más (pulsa Descargar PDF para ver el informe completo)._"
            )
    return "\n".join(lineas)


def teclado_consulta_inasistencias(
    periodo_actual: str, filtro_estado: str, owner_id: int
) -> InlineKeyboardMarkup:
    def _m(p: str, t: str) -> str:
        return f"🔘 {t}" if periodo_actual == p else t

    def _me(e: str, t: str) -> str:
        return f"🔘 {t}" if filtro_estado.upper() == e.upper() else t

    est_s = filtro_estado[:15].replace("|", "")
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(_m("semana", "📆 Esta Semana"), callback_data=f"ina_ver|semana|{est_s}|{owner_id}"),
                InlineKeyboardButton(_m("semana_ant", "⏪ Sem. Pasada"), callback_data=f"ina_ver|semana_ant|{est_s}|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_m("mes", "🗓️ Este Mes"), callback_data=f"ina_ver|mes|{est_s}|{owner_id}"),
                InlineKeyboardButton(_m("mes_ant", "🗓️ Mes Anterior"), callback_data=f"ina_ver|mes_ant|{est_s}|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_m("2026", "📅 Año 2026"), callback_data=f"ina_ver|2026|{est_s}|{owner_id}"),
                InlineKeyboardButton(_m("2025", "📅 Año 2025"), callback_data=f"ina_ver|2025|{est_s}|{owner_id}"),
                InlineKeyboardButton(_m("todo", "📚 Todo"), callback_data=f"ina_ver|todo|{est_s}|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_me("TODOS", "📋 Todos"), callback_data=f"ina_ver|{periodo_actual}|TODOS|{owner_id}"),
                InlineKeyboardButton(_me("Pendiente", "⏳ Solo Pendientes"), callback_data=f"ina_ver|{periodo_actual}|Pendiente|{owner_id}"),
            ],
            [
                InlineKeyboardButton(
                    "📄 Descargar Reporte Consolidado en PDF",
                    callback_data=f"ina_pdf_per|{periodo_actual}|{est_s}|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    "👤 Reporte PDF Individual por Alumno",
                    callback_data=f"ina_pedir_alu|{owner_id}",
                ),
                InlineKeyboardButton(
                    "💳 Cambiar Estado Pago (#Cons)",
                    callback_data=f"ina_pedir_estado|{owner_id}",
                ),
            ],
            [
                InlineKeyboardButton(
                    "➕ Registrar Nueva Inasistencia",
                    callback_data=f"ina_nuevo_ini|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton("👑 Volver al Panel Admin", callback_data=f"panel_admin|{owner_id}"),
                InlineKeyboardButton("✅ Cerrar", callback_data=f"cerrar_consulta|{owner_id}"),
            ],
        ]
    )


def _formatear_texto_mantenimientos(datos: Dict[str, Any]) -> str:
    placa_f = datos.get("placa_filtro", "TODAS")
    desglose_placas = (
        " | ".join(f"*{k}:* `{v}`" for k, v in sorted(datos.get("por_placa", {}).items()))
        or "Ninguno"
    )
    lineas = [
        "🛠️ *CONTROL DE MANTENIMIENTO DE VEHÍCULOS (PESV — SISI-F2-08)*",
        "━━━━━━━━━━━━━━━━━━━━━━",
        f"📅 *Período:* {datos.get('periodo_etiqueta', '')}",
        f"🚗 *Vehículo (Placa):* `{placa_f}`",
        f"✅ *Total Mantenimientos:* `{datos.get('total_registros', 0)}` (`Preventivos: {datos.get('preventivos', 0)}` | `Correctivos: {datos.get('correctivos', 0)}`)",
        f"🔢 *Por Placa:* {desglose_placas}",
    ]
    if datos.get("total_costo", 0) > 0:
        lineas.append(f"💰 *Costo Total Registrado:* `${datos.get('total_costo', 0):,}`")
    lineas.append("━━━━━━━━━━━━━━━━━━━━━━")

    registros = datos.get("registros", [])
    if not registros:
        lineas.append("ℹ️ _No hay mantenimientos registrados en este período/placa._")
    else:
        lineas.append(f"📋 *Relación de Mantenimientos ({len(registros)} registro(s)):*")
        for idx, r in enumerate(list(reversed(registros))[:18], start=1):
            fech = str(r.get("Fecha", "")).strip()
            pl = str(r.get("Placa", "")).strip().upper()
            tip = str(r.get("Tipo_Mantenimiento", "Preventivo")).strip()
            estab = str(r.get("Establecimiento", "")).strip()
            obs = str(r.get("Observaciones", "")).strip()
            lineas.append(
                f"`{idx:02d}.` `{fech}` — *{pl}* ({tip}) | 🏪 *{estab}*: _{obs[:45]}_"
            )
        if len(registros) > 18:
            lineas.append(
                f"_...y {len(registros) - 18} registro(s) más (descarga el PDF Oficial SISI-F2-08 para ver la tabla completa)._"
            )
    return "\n".join(lineas)


def teclado_consulta_mantenimientos(
    periodo_actual: str, placa_actual: str, owner_id: int
) -> InlineKeyboardMarkup:
    def _m(p: str, t: str) -> str:
        return f"🔘 {t}" if periodo_actual.lower() == p.lower() else t

    def _mp(pl: str, t: str) -> str:
        return f"🔘 {t}" if placa_actual.upper() == pl.upper() else t

    pl_s = placa_actual.upper()[:10].replace("|", "")
    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(_m("semana", "📆 Esta Semana"), callback_data=f"mant_ver|semana|{pl_s}|{owner_id}"),
                InlineKeyboardButton(_m("mes", "🗓️ Este Mes"), callback_data=f"mant_ver|mes|{pl_s}|{owner_id}"),
                InlineKeyboardButton(_m("mes_ant", "🗓️ Mes Ant."), callback_data=f"mant_ver|mes_ant|{pl_s}|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_m("julio", "🗓️ Julio"), callback_data=f"mant_ver|julio|{pl_s}|{owner_id}"),
                InlineKeyboardButton(_m("agosto", "🗓️ Agosto"), callback_data=f"mant_ver|agosto|{pl_s}|{owner_id}"),
                InlineKeyboardButton(_m("septiembre", "🗓️ Septiembre"), callback_data=f"mant_ver|septiembre|{pl_s}|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_m("2026", "📅 Año 2026"), callback_data=f"mant_ver|2026|{pl_s}|{owner_id}"),
                InlineKeyboardButton(_m("2025", "📅 Año 2025"), callback_data=f"mant_ver|2025|{pl_s}|{owner_id}"),
                InlineKeyboardButton(_m("todo", "📚 Todo Histórico"), callback_data=f"mant_ver|todo|{pl_s}|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_mp("TODAS", "🚘 Todas"), callback_data=f"mant_ver|{periodo_actual}|TODAS|{owner_id}"),
                InlineKeyboardButton(_mp("UDT984", "🚗 UDT984"), callback_data=f"mant_ver|{periodo_actual}|UDT984|{owner_id}"),
                InlineKeyboardButton(_mp("FSP335", "🚗 FSP335"), callback_data=f"mant_ver|{periodo_actual}|FSP335|{owner_id}"),
                InlineKeyboardButton(_mp("DUX342", "🚗 DUX342"), callback_data=f"mant_ver|{periodo_actual}|DUX342|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_mp("YNE31F", "🏍️ YNE31F"), callback_data=f"mant_ver|{periodo_actual}|YNE31F|{owner_id}"),
                InlineKeyboardButton(_mp("YNE30F", "🏍️ YNE30F"), callback_data=f"mant_ver|{periodo_actual}|YNE30F|{owner_id}"),
                InlineKeyboardButton(_mp("AZA25G", "🏍️ AZA25G"), callback_data=f"mant_ver|{periodo_actual}|AZA25G|{owner_id}"),
            ],
            [
                InlineKeyboardButton(
                    "📄 Descargar PDF Oficial SISI-F2-08 (PESV)",
                    callback_data=f"mant_pdf|{periodo_actual}|{pl_s}|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    "🔧 Registrar Nuevo Mantenimiento",
                    callback_data=f"mant_nuevo_ini|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton("👑 Volver al Panel Admin", callback_data=f"panel_admin|{owner_id}"),
                InlineKeyboardButton("✅ Cerrar", callback_data=f"cerrar_consulta|{owner_id}"),
            ],
        ]
    )


def teclado_config_nueva_inasistencia(sesion: Dict[str, Any], owner_id: int) -> InlineKeyboardMarkup:
    cat_sel = sesion.get("ina_categoria") or "B1"
    horas_sel = int(sesion.get("ina_horas", 2))
    valor_sel = int(sesion.get("ina_valor_hora", 25000))
    est_sel = sesion.get("ina_estado") or "Pendiente Pago"
    cl_sel = sesion.get("ina_clases_faltadas") or "No especificada"

    def _mc(c: str, t: str) -> str:
        return f"✅ {t}" if cat_sel == c else t

    def _mh(h: int) -> str:
        return f"✅ {h}h" if horas_sel == h else f"{h}h"

    def _mv(v: int) -> str:
        etq = f"${v // 1000}k"
        return f"✅ {etq}" if valor_sel == v else etq

    def _me(e: str, t: str) -> str:
        return f"✅ {t}" if est_sel == e else t

    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(_mc("A1", "🏍️ A1"), callback_data=f"ina_cfg|cat|A1|{owner_id}"),
                InlineKeyboardButton(_mc("A2", "🏍️ A2"), callback_data=f"ina_cfg|cat|A2|{owner_id}"),
                InlineKeyboardButton(_mc("B1", "🚗 B1"), callback_data=f"ina_cfg|cat|B1|{owner_id}"),
                InlineKeyboardButton(_mc("C1", "🚕 C1"), callback_data=f"ina_cfg|cat|C1|{owner_id}"),
            ],
            [
                InlineKeyboardButton(
                    f"📚 Clase(s) a la que faltó: {cl_sel} (Toca para elegir)",
                    callback_data=f"ina_elegir_clases|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton(_mh(1), callback_data=f"ina_cfg|hrs|1|{owner_id}"),
                InlineKeyboardButton(_mh(2), callback_data=f"ina_cfg|hrs|2|{owner_id}"),
                InlineKeyboardButton(_mh(3), callback_data=f"ina_cfg|hrs|3|{owner_id}"),
                InlineKeyboardButton(_mh(4), callback_data=f"ina_cfg|hrs|4|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_mv(11000), callback_data=f"ina_cfg|val|11000|{owner_id}"),
                InlineKeyboardButton(_mv(15000), callback_data=f"ina_cfg|val|15000|{owner_id}"),
                InlineKeyboardButton(_mv(20000), callback_data=f"ina_cfg|val|20000|{owner_id}"),
                InlineKeyboardButton(_mv(25000), callback_data=f"ina_cfg|val|25000|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_me("Pendiente Pago", "⏳ Pendiente Pago"), callback_data=f"ina_cfg|est|Pendiente Pago|{owner_id}"),
                InlineKeyboardButton(_me("Cancelo Efectivo", "💵 Efectivo"), callback_data=f"ina_cfg|est|Cancelo Efectivo|{owner_id}"),
                InlineKeyboardButton(_me("Cancelo Trasferencia", "🏦 Transferencia"), callback_data=f"ina_cfg|est|Cancelo Trasferencia|{owner_id}"),
            ],
            [
                InlineKeyboardButton(
                    "👤 PASO FINAL: Seleccionar Alumno que Faltó ➡️",
                    callback_data=f"ina_sel_alu_pag|0|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton("⬅️ Volver", callback_data=f"volver_cat|{owner_id}"),
                InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar|{owner_id}"),
            ],
        ]
    )


def _texto_config_nueva_inasistencia(sesion: Dict[str, Any]) -> str:
    cat_sel = sesion.get("ina_categoria") or "B1"
    horas_sel = int(sesion.get("ina_horas", 2))
    valor_sel = int(sesion.get("ina_valor_hora", 25000))
    est_sel = sesion.get("ina_estado") or "Pendiente Pago"
    cl_sel = sesion.get("ina_clases_faltadas") or "No especificada"
    subtotal = horas_sel * valor_sel
    return (
        f"🚫 *REGISTRAR INASISTENCIA DE ALUMNO (PERDIÓ CLASE)*\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"{_encabezado_instructor(sesion)}"
        f"🏷 *Categoría:* `{cat_sel}`\n"
        f"📚 *Clase(s) a la que faltó:* `{cl_sel}`\n"
        f"⏱️ *No. Horas perdidas:* `{horas_sel}` hora(s)\n"
        f"💲 *Valor por Hora:* `${valor_sel:,}` *(Total: `${subtotal:,}`)*\n"
        f"💳 *Estado de Pago:* `{est_sel}`\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👇 *Ajusta con un toque Categoría, Clase faltada, Horas, Valor o Estado y luego pulsa* "
        f"**`👤 PASO FINAL: Seleccionar Alumno que Faltó ➡️`** *(o escribe directamente el Nombre y Cédula del alumno en el chat)*:"
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

    # -------------------------------------------------------------------------
    # MANEJADORES DE INASISTENCIAS Y MANTENIMIENTOS (APPSHEET -> BOT CENTRALIZADO)
    # -------------------------------------------------------------------------
    if accion_limpia.startswith("ina_ver|"):
        _, per_ina, est_ina = accion_limpia.split("|", 2)
        datos_ina = sheets_service.consultar_inasistencias_periodo(
            periodo=per_ina, filtro_estado=est_ina
        )
        await _editar_mensaje_seguro(
            query.message,
            _formatear_texto_inasistencias(datos_ina),
            reply_markup=teclado_consulta_inasistencias(per_ina, est_ina, user_id),
        )
        return

    if accion_limpia.startswith("ina_pdf_per|"):
        _, per_ina, est_ina = accion_limpia.split("|", 2)
        datos_ina = sheets_service.consultar_inasistencias_periodo(
            periodo=per_ina, filtro_estado=est_ina
        )
        pdf_buf = generar_pdf_inasistencias_periodo(datos_ina)
        filename = f"Reporte_Inasistencias_{per_ina}_{datos_ina['fecha_inicio']}_a_{datos_ina['fecha_fin']}.pdf"
        await context.bot.send_document(
            chat_id=chat_id if query.message.chat.type == "private" else user_id,
            document=pdf_buf,
            filename=filename,
            caption=(
                f"📄 Reporte Consolidado de Inasistencias — CEA Auto Oriente\n"
                f"📅 Período: {datos_ina['periodo_etiqueta']}\n"
                f"🚫 Total casos: {datos_ina['total_registros']} ({datos_ina['total_horas']} h)\n"
                f"💵 Efectivo: ${datos_ina['total_efectivo']:,} | 🏦 Transf.: ${datos_ina['total_transferencia']:,}\n"
                f"⏳ Pendiente Pago: ${datos_ina['total_pendiente']:,}"
            ),
        )
        return

    if accion_limpia == "ina_pedir_alu":
        sesion_tmp = SESIONES_REPORTE.setdefault((chat_id, user_id), {"owner_id": user_id})
        sesion_tmp["esperando_alu_pdf_inasistencia"] = True
        PROPIETARIO_MENSAJE_BOT[msg_id] = user_id
        await _editar_mensaje_seguro(
            query.message,
            "👤 *REPORTE PDF INDIVIDUAL DE INASISTENCIAS POR ALUMNO*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Escribe ahora mismo aquí en el chat el **Nombre o Cédula del Alumno** "
            "(ej. `Nicolás Vera`, `Keiner Ospina` o su número de cédula):\n\n"
            "*(El bot generará su PDF Individual de Inasistencias y también su Formato Oficial de Control de Clases con las fallas incluidas en Observaciones).* 👇",
            reply_markup=InlineKeyboardMarkup(
                [
                    [InlineKeyboardButton("⬅️ Volver a Reportes de Inasistencias", callback_data=f"ina_ver|todo|TODOS|{user_id}")],
                    [InlineKeyboardButton("✅ Cancelar", callback_data=f"cerrar_consulta|{user_id}")],
                ]
            ),
        )
        return

    if accion_limpia == "ina_pedir_estado":
        sesion_tmp = SESIONES_REPORTE.setdefault((chat_id, user_id), {"owner_id": user_id})
        sesion_tmp["esperando_consecutivo_estado_ina"] = True
        PROPIETARIO_MENSAJE_BOT[msg_id] = user_id
        datos_pend = sheets_service.consultar_inasistencias_periodo(
            periodo="todo", filtro_estado="Pendiente"
        )
        botones_pend = []
        for r_p in list(reversed(datos_pend.get("registros", [])))[:10]:
            cons_p = str(r_p.get("Consecutivo", "")).strip()
            alu_p = str(r_p.get("Nombre_Alumno", "")).strip()[:18]
            sub_p = r_p.get("Valor_Total_Int", 0)
            botones_pend.append(
                [
                    InlineKeyboardButton(
                        f"💵 #{cons_p} {alu_p} (${sub_p:,}) -> Efectivo",
                        callback_data=f"ina_set_est|{cons_p}|Cancelo Efectivo|{user_id}",
                    ),
                    InlineKeyboardButton(
                        f"🏦 #{cons_p} -> Transf.",
                        callback_data=f"ina_set_est|{cons_p}|Cancelo Trasferencia|{user_id}",
                    ),
                ]
            )
        botones_pend.append(
            [InlineKeyboardButton("⬅️ Volver a Inasistencias", callback_data=f"ina_ver|todo|TODOS|{user_id}")]
        )
        await _editar_mensaje_seguro(
            query.message,
            "💳 *ACTUALIZAR ESTADO DE PAGO DE INASISTENCIA*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Toca uno de los casos pendientes abajo para marcarlo como **Canceló Efectivo** o **Canceló Transferencia**, "
            "o escribe en el chat el número de consecutivo seguido del estado (ej. `152 efectivo` o `152 transferencia`):",
            reply_markup=InlineKeyboardMarkup(botones_pend),
        )
        return

    if accion_limpia.startswith("ina_set_est|"):
        _, cons_id, nuevo_est = accion_limpia.split("|", 2)
        actualizado = sheets_service.actualizar_estado_inasistencia(cons_id, nuevo_est)
        if actualizado:
            await _editar_mensaje_seguro(
                query.message,
                f"✅ *ESTADO DE INASISTENCIA ACTUALIZADO*\n"
                f"━━━━━━━━━━━━━━━━━━━━━━\n"
                f"🔢 *Consecutivo:* `#{actualizado['Consecutivo']}` (ID: `{actualizado['ID_Novedad']}`)\n"
                f"👤 *Alumno:* {actualizado['Nombre_Alumno']}\n"
                f"📅 *Fecha:* {actualizado['Fecha']} ({actualizado['Hora']})\n"
                f"💳 *Nuevo Estado:* *{actualizado['Estado']}*",
                reply_markup=teclado_consulta_inasistencias("todo", "TODOS", user_id),
            )
        else:
            await _editar_mensaje_seguro(
                query.message,
                f"⚠️ No se encontró la inasistencia `#{cons_id}`.",
                reply_markup=teclado_consulta_inasistencias("todo", "TODOS", user_id),
            )
        return

    if accion_limpia.startswith("ina_pdf_alu|"):
        crit_alu = accion_limpia.split("|", 1)[1]
        await _enviar_pdf_inasistencia_individual_alumno(crit_alu, chat_id, user_id, context)
        return

    if accion_limpia.startswith("mant_ver|"):
        _, per_m, placa_m = accion_limpia.split("|", 2)
        datos_m = sheets_service.consultar_mantenimientos_periodo(periodo=per_m, placa=placa_m)
        await _editar_mensaje_seguro(
            query.message,
            _formatear_texto_mantenimientos(datos_m),
            reply_markup=teclado_consulta_mantenimientos(per_m, placa_m, user_id),
        )
        return

    if accion_limpia.startswith("mant_pdf|"):
        _, per_m, placa_m = accion_limpia.split("|", 2)
        datos_m = sheets_service.consultar_mantenimientos_periodo(periodo=per_m, placa=placa_m)
        pdf_buf = generar_pdf_mantenimientos_pesv(datos_m)
        sufijo_placa = f"_{placa_m}" if placa_m not in ("TODAS", "TODOS", "") else "_TODOS"
        filename = f"SISI-F2-08_Mantenimientos_PESV_{per_m}{sufijo_placa}.pdf"
        await context.bot.send_document(
            chat_id=chat_id if query.message.chat.type == "private" else user_id,
            document=pdf_buf,
            filename=filename,
            caption=(
                f"🛠️ Formato Oficial SISI-F2-08 — RELACIÓN MANTENIMIENTOS PREVENTIVOS (PESV)\n"
                f"📅 Período: {datos_m['periodo_etiqueta']}\n"
                f"🚗 Placa: {datos_m['placa_filtro']}\n"
                f"✅ Total Mantenimientos: {datos_m['total_registros']}\n"
                f"✍️ Firmado por: SEGUNDO MURILLO (Rep. Legal) y MARLY ALFONSO (Rep. PESV)"
            ),
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

    if accion_limpia == "ina_nuevo_ini":
        sesion["ina_categoria"] = sesion.get("categoria") or "B1"
        sesion["ina_horas"] = sesion.get("ina_horas", 2)
        sesion["ina_valor_hora"] = sesion.get("ina_valor_hora", 25000)
        sesion["ina_estado"] = sesion.get("ina_estado", "Pendiente Pago")
        sesion["ina_clases_faltadas"] = sesion.get("clases_texto") or "No especificada"
        sesion["esperando_alumno_inasistencia"] = True
        sesion["esperando_nombre_texto"] = False
        sesion["menu_msg_id"] = msg_id
        PROPIETARIO_MENSAJE_BOT[msg_id] = sesion_uid
        await _editar_mensaje_seguro(
            query.message,
            _texto_config_nueva_inasistencia(sesion),
            reply_markup=teclado_config_nueva_inasistencia(sesion, sesion_uid),
        )
        return

    if accion_limpia.startswith("ina_cfg|"):
        _, campo_c, val_c = accion_limpia.split("|", 2)
        if campo_c == "cat":
            sesion["ina_categoria"] = normalizar_categoria(val_c)
        elif campo_c == "hrs":
            sesion["ina_horas"] = int(val_c)
        elif campo_c == "val":
            sesion["ina_valor_hora"] = int(val_c)
        elif campo_c == "est":
            sesion["ina_estado"] = val_c
        elif campo_c == "cl":
            sesion["ina_clases_faltadas"] = val_c
        sesion["esperando_alumno_inasistencia"] = True
        await _editar_mensaje_seguro(
            query.message,
            _texto_config_nueva_inasistencia(sesion),
            reply_markup=teclado_config_nueva_inasistencia(sesion, sesion_uid),
        )
        return

    if accion_limpia == "ina_elegir_clases":
        cat_ina = sesion.get("ina_categoria") or "B1"
        cfg_ina = FORMATOS_OFICIALES.get(cat_ina, FORMATOS_OFICIALES["B1"])
        botones_cl = [
            [
                InlineKeyboardButton(
                    "⏭️ No especificar número de clase",
                    callback_data=f"ina_cfg|cl|No especificada|{sesion_uid}",
                )
            ]
        ]
        fila_par = []
        for idx_f, fila_f in enumerate(cfg_ina["filas"]):
            etq_f = fila_f["etiqueta"].replace(".0", "")
            nums_f = formatear_clases(fila_f["clases"])
            fila_par.append(
                InlineKeyboardButton(
                    f"Clase {nums_f} ({fila_f['resumen'][:18]})",
                    callback_data=f"ina_cfg|cl|{nums_f}|{sesion_uid}",
                )
            )
            if len(fila_par) == 2:
                botones_cl.append(fila_par)
                fila_par = []
        if fila_par:
            botones_cl.append(fila_par)
        botones_cl.append(
            [InlineKeyboardButton("⬅️ Volver a Configuración de Inasistencia", callback_data=f"ina_nuevo_ini|{sesion_uid}")]
        )
        await _editar_mensaje_seguro(
            query.message,
            f"📚 *SELECCIONA LA CLASE A LA QUE FALTÓ EL ALUMNO ({cfg_ina['nombre_corto']}):*\n"
            f"*(Esta clase saldrá en las OBSERVACIONES del formato oficial de control de clases del alumno)*:",
            reply_markup=InlineKeyboardMarkup(botones_cl),
        )
        return

    if accion_limpia.startswith("ina_sel_alu_pag|"):
        pag_ina = int(accion_limpia.split("|", 1)[1])
        cat_ina = sesion.get("ina_categoria") or "B1"
        sesion["esperando_alumno_inasistencia"] = True
        alumnos_ina = sheets_service.listar_alumnos_activos(categoria=cat_ina, modo="activos")
        if not alumnos_ina:
            alumnos_ina = sheets_service.listar_alumnos_activos(categoria="TODAS", modo="activos")
        por_pag = 20
        tot_a = len(alumnos_ina)
        tot_p = max(1, (tot_a + por_pag - 1) // por_pag)
        pag_ina = max(0, min(pag_ina, tot_p - 1))
        sub_alu = alumnos_ina[pag_ina * por_pag : (pag_ina + 1) * por_pag]

        botones_alu_ina = []
        for fa in sub_alu:
            nom_a = str(fa.get("Nombre_Completo", "")).strip()
            doc_a = str(fa.get("Documento", "")).strip()
            cat_a = str(fa.get("Categoria", "")).strip()
            if nom_a and doc_a:
                botones_alu_ina.append(
                    [
                        InlineKeyboardButton(
                            f"👤 {nom_a} [{cat_a}] — CC: {doc_a}",
                            callback_data=f"ina_guardar_alu|{doc_a}|{sesion_uid}",
                        )
                    ]
                )
        if tot_p > 1:
            nav_i = []
            if pag_ina > 0:
                nav_i.append(InlineKeyboardButton("⬅️ Ant.", callback_data=f"ina_sel_alu_pag|{pag_ina - 1}|{sesion_uid}"))
            nav_i.append(InlineKeyboardButton(f"Pág {pag_ina + 1}/{tot_p}", callback_data=f"ina_sel_alu_pag|{pag_ina}|{sesion_uid}"))
            if pag_ina < tot_p - 1:
                nav_i.append(InlineKeyboardButton("Sig. ➡️", callback_data=f"ina_sel_alu_pag|{pag_ina + 1}|{sesion_uid}"))
            botones_alu_ina.append(nav_i)

        botones_alu_ina.append(
            [InlineKeyboardButton("⬅️ Volver a Configurar Horas/Estado", callback_data=f"ina_nuevo_ini|{sesion_uid}")]
        )
        botones_alu_ina.append(
            [InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar|{sesion_uid}")]
        )
        await _editar_mensaje_seguro(
            query.message,
            f"🚫 *SELECCIONA EL ALUMNO QUE NO ASISTIÓ A CLASE ({cat_ina})*\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"Toca el nombre del alumno en los botones de abajo, o **escribe directamente en el chat el Nombre y Apellido (y Cédula opcional)** del alumno:",
            reply_markup=InlineKeyboardMarkup(botones_alu_ina),
        )
        return

    if accion_limpia.startswith("ina_guardar_alu|"):
        doc_alu_sel = accion_limpia.split("|", 1)[1]
        cat_ina = sesion.get("ina_categoria") or "B1"
        ficha_a = sheets_service.buscar_alumno(doc_alu_sel, categoria=cat_ina)
        nom_alu_sel = str(ficha_a.get("Nombre_Completo", doc_alu_sel)).strip() if ficha_a else doc_alu_sel
        await _finalizar_registro_inasistencia_en_bot(
            chat_id=chat_id,
            user_id=sesion_uid,
            documento_alumno=doc_alu_sel,
            nombre_alumno=nom_alu_sel,
            context=context,
            mensaje_editar=query.message,
        )
        return

    # FLUJO DE REGISTRO DE MANTENIMIENTO DE VEHÍCULO (PESV / APPSHEET)
    if accion_limpia == "mant_nuevo_ini":
        sesion["mant_placa"] = sesion.get("mant_placa") or "UDT984"
        sesion["mant_tipo"] = sesion.get("mant_tipo") or "Preventivo"
        sesion["mant_rol"] = sesion.get("mant_rol") or "MECANICO"
        sesion["mant_establecimiento"] = sesion.get("mant_establecimiento") or "Motollantas Richard"
        sesion["mant_nit"] = sesion.get("mant_nit") or "13721566"
        sesion["mant_actividad"] = sesion.get("mant_actividad") or "Mantenimiento"
        sesion["mant_tecnico"] = sesion.get("mant_tecnico") or "Motollantas Richard"
        sesion["esperando_obs_mantenimiento"] = False
        sesion["menu_msg_id"] = msg_id
        PROPIETARIO_MENSAJE_BOT[msg_id] = sesion_uid
        await _editar_mensaje_seguro(
            query.message,
            _texto_config_mantenimiento(sesion),
            reply_markup=teclado_config_mantenimiento(sesion, sesion_uid),
        )
        return

    if accion_limpia.startswith("mant_cfg|"):
        _, campo_m, val_m = accion_limpia.split("|", 2)
        if campo_m == "pl":
            sesion["mant_placa"] = val_m.upper()
        elif campo_m == "tip":
            sesion["mant_tipo"] = val_m
        elif campo_m == "rol":
            sesion["mant_rol"] = val_m.upper()
        elif campo_m == "tal":
            idx_t = int(val_m)
            if 0 <= idx_t < len(TALLERES_FRECUENTES_PESV):
                t_info = TALLERES_FRECUENTES_PESV[idx_t]
                sesion["mant_establecimiento"] = t_info["nombre"]
                sesion["mant_tecnico"] = t_info["nombre"]
                sesion["mant_nit"] = t_info["nit"]
        await _editar_mensaje_seguro(
            query.message,
            _texto_config_mantenimiento(sesion),
            reply_markup=teclado_config_mantenimiento(sesion, sesion_uid),
        )
        return

    if accion_limpia == "mant_elegir_taller":
        botones_t = []
        for idx_t, t_info in enumerate(TALLERES_FRECUENTES_PESV):
            botones_t.append(
                [
                    InlineKeyboardButton(
                        f"🏪 {t_info['nombre']} (NIT: {t_info['nit']})",
                        callback_data=f"mant_cfg|tal|{idx_t}|{sesion_uid}",
                    )
                ]
            )
        botones_t.append(
            [InlineKeyboardButton("✍️ Escribir otro Establecimiento y NIT", callback_data=f"mant_escribir_taller|{sesion_uid}")]
        )
        botones_t.append(
            [InlineKeyboardButton("⬅️ Volver al Formulario de Mantenimiento", callback_data=f"mant_nuevo_ini|{sesion_uid}")]
        )
        await _editar_mensaje_seguro(
            query.message,
            "🏪 *SELECCIONA EL ESTABLECIMIENTO / TALLER QUE REALIZA EL MANTENIMIENTO:*",
            reply_markup=InlineKeyboardMarkup(botones_t),
        )
        return

    if accion_limpia == "mant_escribir_taller":
        sesion["esperando_taller_mantenimiento"] = True
        await _editar_mensaje_seguro(
            query.message,
            "✍️ *ESCRIBIR ESTABLECIMIENTO Y NIT / CC*\n"
            "━━━━━━━━━━━━━━━━━━━━━━\n"
            "Escribe ahora en el chat el **Nombre del Taller/Establecimiento** y su **NIT o Cédula**\n"
            "📌 *Ejemplo:* `Lubricentro Oriental 901234567`",
            reply_markup=InlineKeyboardMarkup(
                [[InlineKeyboardButton("⬅️ Volver", callback_data=f"mant_nuevo_ini|{sesion_uid}")]]
            ),
        )
        return

    if accion_limpia == "mant_paso_obs":
        sesion["esperando_obs_mantenimiento"] = True
        pl_m = sesion.get("mant_placa", "UDT984")
        estab_m = sesion.get("mant_establecimiento", "Motollantas Richard")
        botones_obs_rapidas = [
            [
                InlineKeyboardButton("🛢️ Cambio de aceite", callback_data=f"mant_guardar_rap|Cambio de aceite|{sesion_uid}"),
                InlineKeyboardButton("⛓️ Aceite y tensión cadena", callback_data=f"mant_guardar_rap|Cambio de aceite y tensión de cadena|{sesion_uid}"),
            ],
            [
                InlineKeyboardButton("🛑 Mantenimiento frenos", callback_data=f"mant_guardar_rap|Mantenimiento de frenos y cambio de pastillas|{sesion_uid}"),
                InlineKeyboardButton("🛞 Cambio/Calibración llantas", callback_data=f"mant_guardar_rap|Cambio y calibración de llantas|{sesion_uid}"),
            ],
            [
                InlineKeyboardButton("🔋 Cambio de Batería", callback_data=f"mant_guardar_rap|Cambio Batería|{sesion_uid}"),
                InlineKeyboardButton("⚖️ Balanceo y Alineación", callback_data=f"mant_guardar_rap|Balanceo y alineación|{sesion_uid}"),
            ],
            [
                InlineKeyboardButton("⬅️ Volver a Configurar Placa/Taller", callback_data=f"mant_nuevo_ini|{sesion_uid}"),
                InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar|{sesion_uid}"),
            ],
        ]
        await _editar_mensaje_seguro(
            query.message,
            f"🔧 *PASO FINAL: ¿QUÉ MANTENIMIENTO / OBSERVACIÓN SE REALIZÓ EN {pl_m}?*\n"
            f"🏪 *Taller:* {estab_m} (NIT: `{sesion.get('mant_nit', '0')}`)\n"
            f"━━━━━━━━━━━━━━━━━━━━━━\n"
            f"👉 Toca uno de los mantenimientos frecuentes abajo para guardarlo de inmediato, "
            f"**o escribe en el chat la descripción de lo que se realizó**\n"
            f"*(Opcional: puedes agregar Kilometraje, Costo y Próximo cambio separados por coma, ej:*\n"
            f"`Cambio de aceite y filtros, km 45000, costo 180000, proximo 50000` *)*:",
            reply_markup=InlineKeyboardMarkup(botones_obs_rapidas),
        )
        return

    if accion_limpia.startswith("mant_guardar_rap|"):
        obs_rap = accion_limpia.split("|", 1)[1]
        await _finalizar_registro_mantenimiento_en_bot(
            chat_id=chat_id,
            user_id=sesion_uid,
            observaciones_texto=obs_rap,
            context=context,
            mensaje_editar=query.message,
        )
        return

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


def _texto_config_mantenimiento(sesion: Dict[str, Any]) -> str:
    pl = sesion.get("mant_placa") or "UDT984"
    tip = sesion.get("mant_tipo") or "Preventivo"
    rol = sesion.get("mant_rol") or "MECANICO"
    estab = sesion.get("mant_establecimiento") or "Motollantas Richard"
    nit = sesion.get("mant_nit") or "13721566"
    return (
        f"🔧 *FORMULARIO DE REGISTRO DE MANTENIMIENTO DE VEHÍCULO (PESV)*\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"{_encabezado_instructor(sesion)}"
        f"🚗 *Placa Seleccionada:* `{pl}`\n"
        f"🛠️ *Tipo de Mantenimiento:* `{tip}`\n"
        f"🏪 *Establecimiento / Taller:* *{estab}* *(NIT: `{nit}`)*\n"
        f"👨‍🔧 *Rol del Técnico:* `{rol}`\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👇 *Selecciona la Placa, Tipo, Rol o Taller y luego toca* **`📝 PASO FINAL: Actividad / Observaciones ➡️`**:"
    )


def teclado_config_mantenimiento(sesion: Dict[str, Any], owner_id: int) -> InlineKeyboardMarkup:
    pl_sel = (sesion.get("mant_placa") or "UDT984").upper()
    tip_sel = sesion.get("mant_tipo") or "Preventivo"
    rol_sel = (sesion.get("mant_rol") or "MECANICO").upper()
    estab_sel = sesion.get("mant_establecimiento") or "Motollantas Richard"

    def _mp(p: str, t: str) -> str:
        return f"✅ {t}" if pl_sel == p else t

    def _mt(tp: str) -> str:
        return f"✅ {tp}" if tip_sel.lower() == tp.lower() else tp

    def _mr(r: str) -> str:
        return f"✅ {r}" if rol_sel == r else r

    return InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(_mp("UDT984", "🚗 UDT984"), callback_data=f"mant_cfg|pl|UDT984|{owner_id}"),
                InlineKeyboardButton(_mp("FSP335", "🚗 FSP335"), callback_data=f"mant_cfg|pl|FSP335|{owner_id}"),
                InlineKeyboardButton(_mp("DUX342", "🚗 DUX342"), callback_data=f"mant_cfg|pl|DUX342|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_mp("YNE31F", "🏍️ YNE31F"), callback_data=f"mant_cfg|pl|YNE31F|{owner_id}"),
                InlineKeyboardButton(_mp("YNE30F", "🏍️ YNE30F"), callback_data=f"mant_cfg|pl|YNE30F|{owner_id}"),
                InlineKeyboardButton(_mp("AZA25G", "🏍️ AZA25G"), callback_data=f"mant_cfg|pl|AZA25G|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_mt("Preventivo"), callback_data=f"mant_cfg|tip|Preventivo|{owner_id}"),
                InlineKeyboardButton(_mt("Correctivo"), callback_data=f"mant_cfg|tip|Correctivo|{owner_id}"),
            ],
            [
                InlineKeyboardButton(_mr("MECANICO"), callback_data=f"mant_cfg|rol|MECANICO|{owner_id}"),
                InlineKeyboardButton(_mr("ELECTRICO"), callback_data=f"mant_cfg|rol|ELECTRICO|{owner_id}"),
                InlineKeyboardButton(_mr("PINTOR"), callback_data=f"mant_cfg|rol|PINTOR|{owner_id}"),
            ],
            [
                InlineKeyboardButton(
                    f"🏪 Taller: {estab_sel} (Toca para cambiar)",
                    callback_data=f"mant_elegir_taller|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    "📝 PASO FINAL: Actividad / Observaciones ➡️",
                    callback_data=f"mant_paso_obs|{owner_id}",
                )
            ],
            [
                InlineKeyboardButton("⬅️ Volver", callback_data=f"volver_cat|{owner_id}"),
                InlineKeyboardButton("❌ Cancelar", callback_data=f"cancelar|{owner_id}"),
            ],
        ]
    )


async def _enviar_pdf_inasistencia_individual_alumno(
    criterio: str,
    destino_chat_id: int,
    user_id: int,
    context: ContextTypes.DEFAULT_TYPE,
) -> bool:
    """
    Busca las inasistencias de un alumno y envía su Informe Individual de Inasistencias en PDF,
    además de su Formato Oficial de Seguimiento de Clases (si está matriculado) con las inasistencias
    en el recuadro de OBSERVACIONES AL PROCESO DE APRENDIZAJE.
    """
    ficha_alu = sheets_service.buscar_alumno(criterio)
    nom_busq = str(ficha_alu.get("Nombre_Completo", "")).strip() if ficha_alu else criterio.strip()
    doc_busq = str(ficha_alu.get("Documento", "")).strip() if ficha_alu else ""

    inasistencias = sheets_service.obtener_inasistencias_alumno(
        criterio_alumno=doc_busq or criterio,
        nombre_alternativo=nom_busq,
    )
    if not inasistencias and not ficha_alu:
        await _enviar_mensaje_seguro(
            context.bot,
            chat_id=destino_chat_id,
            text=f"❌ No se encontraron inasistencias ni matrícula para: *{criterio}*.",
            reply_markup=teclado_consulta_inasistencias("todo", "TODOS", user_id),
        )
        return False

    datos_alu_pdf = ficha_alu or {
        "Nombre_Completo": (inasistencias[0].get("Nombre_Alumno") if inasistencias else nom_busq),
        "Documento": (inasistencias[0].get("Documento_Alumno") if inasistencias else doc_busq) or "Registrado",
        "Categoria": (inasistencias[0].get("Categoria") if inasistencias else "General") or "General",
    }

    pdf_ind = generar_pdf_inasistencias_individual(datos_alu_pdf, inasistencias)
    nom_limpio = str(datos_alu_pdf.get("Nombre_Completo", criterio)).strip().replace(" ", "_")
    total_horas = sum(int(x.get("No_Horas_Int", 1)) for x in inasistencias)

    await context.bot.send_document(
        chat_id=destino_chat_id,
        document=pdf_ind,
        filename=f"Informe_Individual_Inasistencias_{nom_limpio}.pdf",
        caption=(
            f"🚫 Informe Individual de Inasistencias — CEA Auto Oriente\n"
            f"👤 Alumno: {datos_alu_pdf.get('Nombre_Completo')}\n"
            f"📌 Veces que no asistió: {len(inasistencias)} ({total_horas} hora(s) perdida(s))"
        ),
    )

    # Si el alumno tiene ficha o clases en Registro_Clases, enviar también el Formato Oficial de Clases con la observación
    if ficha_alu:
        await _generar_y_enviar_pdf_alumno(
            criterio=doc_busq or nom_busq,
            destino_chat_id=destino_chat_id,
            user_id=user_id,
            context=context,
        )
    return True


async def _finalizar_registro_inasistencia_en_bot(
    chat_id: int,
    user_id: int,
    documento_alumno: str,
    nombre_alumno: str,
    context: ContextTypes.DEFAULT_TYPE,
    mensaje_editar=None,
) -> None:
    sesion = SESIONES_REPORTE.pop((chat_id, user_id), {}) or {}
    cat_ina = sesion.get("ina_categoria") or sesion.get("categoria") or "B1"
    cl_falt = sesion.get("ina_clases_faltadas") or sesion.get("clases_texto") or "No especificada"
    horas_ina = int(sesion.get("ina_horas", 2))
    val_ina = int(sesion.get("ina_valor_hora", 25000))
    est_ina = sesion.get("ina_estado") or "Pendiente Pago"
    inst_nom = sesion.get("instructor_nombre") or "Instructor"

    reg = sheets_service.registrar_inasistencia(
        documento_alumno=documento_alumno,
        nombre_alumno=nombre_alumno,
        categoria=cat_ina,
        clases_faltadas=cl_falt,
        nombre_instructor=inst_nom,
        no_horas=horas_ina,
        valor_hora=val_ina,
        estado=est_ina,
    )
    crit_btn = (reg["Documento_Alumno"] or reg["Nombre_Alumno"])[:24].replace("|", "")
    cons_id = reg["Consecutivo"]

    texto_ok = (
        f"🚫 *INASISTENCIA REGISTRADA EXITOSAMENTE*\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"🔢 *Consecutivo:* `#{reg['Consecutivo']}` *(ID: `{reg['ID_Novedad']}`)*\n"
        f"📅 *Fecha y Hora:* `{reg['Fecha']} {reg['Hora']}`\n"
        f"👤 *Alumno:* {reg['Nombre_Alumno']}{f' (CC: `{reg["Documento_Alumno"]}`)' if reg['Documento_Alumno'] else ''}\n"
        f"🏷 *Categoría:* `{reg['Categoria'] or cat_ina}` | *Clase(s) Faltada(s):* `{reg['Clases_Faltadas']}`\n"
        f"👨‍🏫 *Instructor:* {reg['Nombre_Instructor']}\n"
        f"⏱️ *No. Horas:* `{reg['No_Horas']}` | *Valor Hora:* `${reg['Valor_Hora']:,}` *(Total: `${reg['Valor_Total']:,}`)*\n"
        f"💳 *Estado:* *{reg['Estado']}*\n\n"
        f"✅ *Esta inasistencia ya quedó anotada automáticamente en las OBSERVACIONES del formato oficial de control de clases del alumno.*"
    )
    teclado_ok = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    "📄 Descargar PDF Individual + Formato de Clases",
                    callback_data=f"ina_pdf_alu|{crit_btn}|{user_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    "💵 Marcar Canceló Efectivo",
                    callback_data=f"ina_set_est|{cons_id}|Cancelo Efectivo|{user_id}",
                ),
                InlineKeyboardButton(
                    "🏦 Marcar Transferencia",
                    callback_data=f"ina_set_est|{cons_id}|Cancelo Trasferencia|{user_id}",
                ),
            ],
            [
                InlineKeyboardButton(
                    "📊 Ver Reportes de Inasistencias (Semana/Mes/Año)",
                    callback_data=f"ina_ver|mes|TODOS|{user_id}",
                )
            ],
        ]
    )
    if mensaje_editar:
        await _editar_mensaje_seguro(mensaje_editar, texto_ok, reply_markup=teclado_ok)
    else:
        await _enviar_mensaje_seguro(context.bot, chat_id=chat_id, text=texto_ok, reply_markup=teclado_ok)


async def _finalizar_registro_mantenimiento_en_bot(
    chat_id: int,
    user_id: int,
    observaciones_texto: str,
    context: ContextTypes.DEFAULT_TYPE,
    mensaje_editar=None,
) -> None:
    sesion = SESIONES_REPORTE.pop((chat_id, user_id), {}) or {}
    pl_m = sesion.get("mant_placa") or "UDT984"
    tip_m = sesion.get("mant_tipo") or "Preventivo"
    rol_m = sesion.get("mant_rol") or "MECANICO"
    estab_m = sesion.get("mant_establecimiento") or "Motollantas Richard"
    nit_m = sesion.get("mant_nit") or "13721566"
    inst_m = sesion.get("instructor_nombre") or "CEA Auto Oriente"

    # Extraer opcionalmente km, costo y proximo cambio si vienen en el texto
    km_val = "0"
    costo_val = "0"
    prox_val = "0"
    obs_limpia = observaciones_texto.strip()

    m_km = re.search(r"\bkm[:\s]*([\d\.]+)", obs_limpia, flags=re.IGNORECASE)
    if m_km:
        km_val = re.sub(r"[^\d]", "", m_km.group(1))
    m_costo = re.search(r"\b(?:costo|valor|\$)[:\s]*([\d\.]+)", obs_limpia, flags=re.IGNORECASE)
    if m_costo:
        costo_val = re.sub(r"[^\d]", "", m_costo.group(1))
    m_prox = re.search(r"\b(?:proximo|próximo)[:\s]*([\d\.]+)", obs_limpia, flags=re.IGNORECASE)
    if m_prox:
        prox_val = re.sub(r"[^\d]", "", m_prox.group(1))

    reg_m = sheets_service.registrar_mantenimiento(
        placa=pl_m,
        tipo_mantenimiento=tip_m,
        establecimiento=estab_m,
        nit_cc=nit_m,
        actividad_realizada="Mantenimiento",
        nombre_tecnico=estab_m,
        rol=rol_m,
        observaciones=obs_limpia,
        kilometraje_actual=km_val,
        costo=costo_val,
        proximo_cambio=prox_val,
        nombre_instructor=inst_m,
    )

    texto_ok = (
        f"✅ *MANTENIMIENTO DE VEHÍCULO REGISTRADO (PESV)*\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"📅 *Fecha:* `{reg_m['Fecha']}`\n"
        f"🚗 *Placa:* `{reg_m['Placa']}` | *Tipo:* `{reg_m['Tipo_Mantenimiento']}`\n"
        f"🏪 *Establecimiento:* {reg_m['Establecimiento']} *(NIT: `{reg_m['NIT_CC']}`)*\n"
        f"👨‍🔧 *Rol:* `{reg_m['Rol']}` | *Reportado por:* {reg_m['Nombre_Instructor']}\n"
        f"📝 *Observaciones / Actividad:* {reg_m['Observaciones']}\n"
    )
    if reg_m["Kilometraje_Actual"] > 0 or reg_m["Costo"] > 0 or reg_m["Proximo_Cambio"] > 0:
        texto_ok += (
            f"🔢 *Km Actual:* `{reg_m['Kilometraje_Actual']:,}` | "
            f"*Próx. Cambio:* `{reg_m['Proximo_Cambio']:,}` | "
            f"*Costo:* `${reg_m['Costo']:,}`\n"
        )

    teclado_ok = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    "📄 Descargar PDF Oficial SISI-F2-08 (Este Mes)",
                    callback_data=f"mant_pdf|mes|{reg_m['Placa']}|{user_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    "📄 Descargar PDF Oficial SISI-F2-08 (Todo el Histórico)",
                    callback_data=f"mant_pdf|todo|TODAS|{user_id}",
                )
            ],
            [
                InlineKeyboardButton(
                    "🛠️ Abrir Panel de Mantenimientos (Semana/Mes/Año)",
                    callback_data=f"mant_ver|todo|TODAS|{user_id}",
                )
            ],
        ]
    )
    if mensaje_editar:
        await _editar_mensaje_seguro(mensaje_editar, texto_ok, reply_markup=teclado_ok)
    else:
        await _enviar_mensaje_seguro(context.bot, chat_id=chat_id, text=texto_ok, reply_markup=teclado_ok)


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
            if t_low in ("inasistencia", "inasistencias", "falla", "fallas", "no asistio"):
                await cmd_inasistencias(update, context)
                return
            if t_low in ("mantenimiento", "mantenimientos", "taller", "pesv", "carros", "vehiculos"):
                await cmd_mantenimientos(update, context)
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

    # CASO ESPECIAL 1: El usuario está consultando el PDF individual de inasistencias de un alumno
    if sesion.get("esperando_alu_pdf_inasistencia"):
        sesion["esperando_alu_pdf_inasistencia"] = False
        SESIONES_REPORTE.pop((msg.chat_id, sesion_uid), None)
        await _enviar_pdf_inasistencia_individual_alumno(
            criterio=texto,
            destino_chat_id=msg.chat_id,
            user_id=user.id,
            context=context,
        )
        return

    # CASO ESPECIAL 2: El usuario está actualizando el estado de pago de una inasistencia por consecutivo
    if sesion.get("esperando_consecutivo_estado_ina"):
        sesion["esperando_consecutivo_estado_ina"] = False
        SESIONES_REPORTE.pop((msg.chat_id, sesion_uid), None)
        m_num = re.search(r"\b(\d{1,5})\b", texto)
        cons_id = m_num.group(1) if m_num else texto.split()[0]
        t_low = _normalizar_texto(texto)
        if "transf" in t_low or "trasf" in t_low:
            nuevo_est = "Cancelo Trasferencia"
        elif "pend" in t_low:
            nuevo_est = "Pendiente Pago"
        else:
            nuevo_est = "Cancelo Efectivo"
        act = sheets_service.actualizar_estado_inasistencia(cons_id, nuevo_est)
        if act:
            await _enviar_mensaje_seguro(
                context.bot,
                chat_id=msg.chat_id,
                text=(
                    f"✅ *Estado de Inasistencia #{act['Consecutivo']} actualizado a:* `{act['Estado']}`\n"
                    f"👤 *Alumno:* {act['Nombre_Alumno']}"
                ),
                reply_markup=teclado_consulta_inasistencias("todo", "TODOS", user.id),
            )
        else:
            await _enviar_mensaje_seguro(
                context.bot,
                chat_id=msg.chat_id,
                text=f"⚠️ No se encontró el consecutivo `#{cons_id}` en la hoja Inasistencias.",
                reply_markup=teclado_consulta_inasistencias("todo", "TODOS", user.id),
            )
        return

    # CASO ESPECIAL 3: El instructor/admin está escribiendo un taller nuevo y su NIT para Mantenimiento
    if sesion.get("esperando_taller_mantenimiento"):
        sesion["esperando_taller_mantenimiento"] = False
        nom_t, nit_t = extraer_nombre_y_cedula(texto)
        sesion["mant_establecimiento"] = nom_t or re.sub(r"\d+", "", texto).strip() or texto.strip()
        sesion["mant_tecnico"] = sesion["mant_establecimiento"]
        sesion["mant_nit"] = nit_t or "0"
        await _enviar_mensaje_seguro(
            context.bot,
            chat_id=msg.chat_id,
            text=_texto_config_mantenimiento(sesion),
            reply_markup=teclado_config_mantenimiento(sesion, sesion_uid),
            reply_to_message_id=msg.message_id,
        )
        return

    # CASO ESPECIAL 4: El instructor/admin está escribiendo las observaciones/actividad del Mantenimiento
    if sesion.get("esperando_obs_mantenimiento"):
        sesion["esperando_obs_mantenimiento"] = False
        await _finalizar_registro_mantenimiento_en_bot(
            chat_id=msg.chat_id,
            user_id=sesion_uid,
            observaciones_texto=texto,
            context=context,
        )
        return

    # CASO ESPECIAL 5: El instructor/admin está escribiendo el nombre/cédula del alumno que faltó a clase
    if sesion.get("esperando_alumno_inasistencia"):
        sesion["esperando_alumno_inasistencia"] = False
        nom_in, doc_in = extraer_nombre_y_cedula(texto)
        ficha_in = sheets_service.buscar_alumno(doc_in or nom_in or texto, categoria=sesion.get("ina_categoria"))
        if ficha_in:
            nom_final_in = str(ficha_in.get("Nombre_Completo", "")).strip()
            doc_final_in = str(ficha_in.get("Documento", "")).strip()
        else:
            nom_final_in = nom_in or texto.strip().title()
            doc_final_in = doc_in or ""
        await _finalizar_registro_inasistencia_en_bot(
            chat_id=msg.chat_id,
            user_id=sesion_uid,
            documento_alumno=doc_final_in,
            nombre_alumno=nom_final_in,
            context=context,
        )
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


async def cmd_inasistencias(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Comando /inasistencias, /inasistencia o /falla:
    Permite registrar una inasistencia o consultar los reportes PDF por Semana, Mes, Año o Individual por Alumno.
    """
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    if context.args:
        criterio_arg = " ".join(context.args).strip()
        await _enviar_pdf_inasistencia_individual_alumno(
            criterio=criterio_arg,
            destino_chat_id=msg.chat_id,
            user_id=user.id,
            context=context,
        )
        return

    datos_ina = sheets_service.consultar_inasistencias_periodo(periodo="todo", filtro_estado="TODOS")
    msg_resp = await _enviar_mensaje_seguro(
        context.bot,
        chat_id=msg.chat_id,
        text=_formatear_texto_inasistencias(datos_ina),
        reply_markup=teclado_consulta_inasistencias("todo", "TODOS", user.id),
        reply_to_message_id=msg.message_id,
    )
    PROPIETARIO_MENSAJE_BOT[msg_resp.message_id] = user.id


async def cmd_mantenimientos(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """
    Comando /mantenimiento, /mantenimientos, /taller o /pesv:
    Permite registrar un mantenimiento de vehículo o descargar el PDF Oficial SISI-F2-08 por Semana, Mes, Año o Placa.
    """
    msg = update.effective_message
    user = update.effective_user
    chat = update.effective_chat
    if not msg or not user or not chat:
        return

    placa_arg = "TODAS"
    periodo_arg = "todo"
    if context.args:
        arg0 = context.args[0].strip().upper()
        if any(p["placa"] == arg0 for p in PLACAS_VEHICULOS_OFICIALES):
            placa_arg = arg0
        else:
            periodo_arg = context.args[0].strip().lower()

    datos_m = sheets_service.consultar_mantenimientos_periodo(periodo=periodo_arg, placa=placa_arg)
    msg_resp = await _enviar_mensaje_seguro(
        context.bot,
        chat_id=msg.chat_id,
        text=_formatear_texto_mantenimientos(datos_m),
        reply_markup=teclado_consulta_mantenimientos(periodo_arg, placa_arg, user.id),
        reply_to_message_id=msg.message_id,
    )
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
    app.add_handler(CommandHandler(["inasistencia", "inasistencias", "falla", "fallas"], cmd_inasistencias))
    app.add_handler(CommandHandler(["mantenimiento", "mantenimientos", "taller", "pesv"], cmd_mantenimientos))
    app.add_handler(CommandHandler("inicializar", cmd_inicializar))
    app.add_handler(CommandHandler("reporte", cmd_reporte))

    app.add_handler(CallbackQueryHandler(manejar_botones_callback))
    app.add_handler(MessageHandler(filters.PHOTO | filters.Document.IMAGE, procesar_mensaje_imagen))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), procesar_texto_chat))

    logger.info("🚗 Bot CEA Auto Oriente iniciado...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
