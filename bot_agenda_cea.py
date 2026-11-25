"""
=============================================================================
BOT DE AGENDA GOOGLE CALENDAR + BAÚL DIGITAL + DASHBOARD WEB — CEA AUTO ORIENTE
=============================================================================
Integración 100% bidireccional entre:
1. Google Calendar API (`token.json` / `token_calendar.json`).
2. Bot Interactivo de Telegram (`/start`, `/menu`, `/baul`, Agenda Hoy, 7 Días,
   Crear Evento Puntual, Crear Evento Diario, Recordatorios Automáticos 10 min antes
   y Baúl Digital interactivo).
3. Dashboard Web CEA Auto Oriente (Módulo 7: Agenda Google Calendar + Baúl Digital).
"""

import datetime
import json
import os
import re
import threading
import time
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

BASE_DIR = Path(__file__).resolve().parent
TOKEN_CALENDAR_PATH = BASE_DIR / "token_calendar.json"
TOKEN_FALLBACK_PATH = BASE_DIR / "token.json"
CREDS_CALENDAR_PATH = BASE_DIR / "credentials_calendar_oauth.json"
AGENDA_LOCAL_PATH = BASE_DIR / "agenda_eventos_cea.json"
BAUL_LOCAL_PATH = BASE_DIR / "baul_digital_cea.json"
AGENDA_CFG_PATH = BASE_DIR / "agenda_config_cea.json"

SCOPES = ["https://www.googleapis.com/auth/calendar"]
TIMEZONE_STR = "America/Bogota"
LOCAL_TZ = datetime.timezone(datetime.timedelta(hours=-5))

EVENTOS_NOTIFICADOS = set()
_DAEMON_STARTED = False
_DAEMON_LOCK = threading.Lock()
_CONVERSATION_STATE: Dict[str, Dict[str, Any]] = {}


# Seguridad: el token OAuth de Google Calendar se carga desde GOOGLE_CALENDAR_TOKEN_JSON (env) o token_calendar.json local (nunca hardcodeado en Git)
DEFAULT_CALENDAR_TOKEN_JSON = ""


def _cargar_config_agenda() -> Dict[str, str]:
    cfg = {
        "tg_token": (
            os.getenv("TELEGRAM_AGENDA_TOKEN", "").strip()
            or os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
        ),
        "chat_id": (
            os.getenv("AGENDA_CHAT_ID", "").strip()
            or os.getenv("ADMIN_CHAT_ID", "").strip()
            or "6843014700"
        ),
    }
    for p in (AGENDA_CFG_PATH, Path("/tmp/agenda_config_cea.json")):
        if p.exists():
            try:
                saved = json.loads(p.read_text(encoding="utf-8"))
                if isinstance(saved, dict):
                    if saved.get("tg_token"):
                        cfg["tg_token"] = str(saved["tg_token"]).strip()
                    if saved.get("chat_id"):
                        cfg["chat_id"] = str(saved["chat_id"]).strip()
            except Exception:
                pass
    return cfg


def _guardar_config_agenda(tg_token: str = "", chat_id: str = "", calendar_token_json: str = "") -> Dict[str, str]:
    cfg = _cargar_config_agenda()
    if tg_token.strip():
        cfg["tg_token"] = tg_token.strip()
        os.environ["TELEGRAM_AGENDA_TOKEN"] = cfg["tg_token"]
    if chat_id.strip():
        cfg["chat_id"] = chat_id.strip()
        os.environ["AGENDA_CHAT_ID"] = cfg["chat_id"]

    for p in (AGENDA_CFG_PATH, Path("/tmp/agenda_config_cea.json")):
        try:
            p.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    if calendar_token_json.strip():
        raw_tk = calendar_token_json.strip()
        # Si el usuario pegó un bloque que contiene varios JSON concatenados, extraer el de token OAuth
        m_tk = re.search(r'(\{\s*"token"\s*:.*?"refresh_token"\s*:.*?\})', raw_tk, flags=re.DOTALL)
        if m_tk:
            raw_tk = m_tk.group(1)
        for tp in (TOKEN_CALENDAR_PATH, TOKEN_FALLBACK_PATH, Path("/tmp/token_calendar.json")):
            try:
                tp.write_text(raw_tk, encoding="utf-8")
            except Exception:
                pass
    return cfg


def obtener_token_telegram_agenda() -> str:
    cfg = _cargar_config_agenda()
    tk = cfg.get("tg_token", "").strip()
    if tk and tk != "TU_TELEGRAM_BOT_TOKEN":
        return tk
    return ""


def obtener_chat_id_agenda(override: str = "") -> str:
    if override and str(override).strip():
        return str(override).strip()
    cfg = _cargar_config_agenda()
    return cfg.get("chat_id", "6843014700").strip() or "6843014700"


def _extraer_hora_de_titulo(titulo: str, hora_actual: str) -> str:
    """
    Si el usuario escribe 'reunion a las 4 pm' o 'cita 3:30 pm' en el título y dejó
    la hora por defecto ('09:00'), extrae automáticamente la hora en formato 24h ('16:00').
    """
    if hora_actual and hora_actual != "09:00":
        return hora_actual
    txt = (titulo or "").lower()
    m = re.search(r"\b(?:a\s+las\s+)?(\d{1,2})(?::(\d{2}))?\s*(am|pm|a\.m\.|p\.m\.)\b", txt)
    if m:
        hh = int(m.group(1))
        mm = int(m.group(2) or 0)
        ampm = m.group(3).replace(".", "")
        if ampm == "pm" and hh < 12:
            hh += 12
        elif ampm == "am" and hh == 12:
            hh = 0
        if 0 <= hh <= 23 and 0 <= mm <= 59:
            return f"{hh:02d}:{mm:02d}"
    m24 = re.search(r"\b(?:a\s+las\s+)(\d{1,2}):(\d{2})\b", txt)
    if m24:
        hh, mm = int(m24.group(1)), int(m24.group(2))
        if 0 <= hh <= 23 and 0 <= mm <= 59:
            return f"{hh:02d}:{mm:02d}"
    return hora_actual or "09:00"


# =============================================================================
# 1. ALMACÉN DE AGENDA Y BAÚL DIGITAL
# =============================================================================
def _cargar_eventos_locales() -> List[Dict[str, Any]]:
    if AGENDA_LOCAL_PATH.exists():
        try:
            data = json.loads(AGENDA_LOCAL_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
        except Exception:
            pass
    hoy = datetime.datetime.now(LOCAL_TZ)
    semilla = [
        {
            "id": "ev-cea-01",
            "titulo": "Auditoría de Cartera y Revisión Facturación DIAN (FV-2)",
            "fecha": hoy.strftime("%Y-%m-%d"),
            "hora": "08:30",
            "duracion_min": 60,
            "tipo": "DIARIO",
            "responsable": "Gerencia CEA Auto Oriente",
            "estado": "PROGRAMADO",
            "origen": "Google Calendar / Sistema CEA",
        },
        {
            "id": "ev-cea-02",
            "titulo": "Campaña de Remarketing WhatsApp IA (Clientes Pendientes)",
            "fecha": hoy.strftime("%Y-%m-%d"),
            "hora": "10:30",
            "duracion_min": 45,
            "tipo": "DIARIO",
            "responsable": "Comercial / Admisiones",
            "estado": "PROGRAMADO",
            "origen": "Google Calendar / Sistema CEA",
        },
        {
            "id": "ev-cea-03",
            "titulo": "Verificación Cierre de Jornada Instructores y Clases Prácticas",
            "fecha": hoy.strftime("%Y-%m-%d"),
            "hora": "18:00",
            "duracion_min": 30,
            "tipo": "DIARIO",
            "responsable": "Coordinación Académica",
            "estado": "PROGRAMADO",
            "origen": "Google Calendar / Sistema CEA",
        },
    ]
    try:
        AGENDA_LOCAL_PATH.write_text(json.dumps(semilla, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return semilla


def _guardar_eventos_locales(eventos: List[Dict[str, Any]]) -> None:
    try:
        AGENDA_LOCAL_PATH.write_text(json.dumps(eventos, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _cargar_baul_local() -> List[Dict[str, Any]]:
    if BAUL_LOCAL_PATH.exists():
        try:
            data = json.loads(BAUL_LOCAL_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
        except Exception:
            pass
    ahora_str = datetime.datetime.now(LOCAL_TZ).strftime("%Y-%m-%d %H:%M")
    semilla_baul = [
        {
            "id": "baul-01",
            "categoria": "ACCESO / LINK",
            "titulo": "Accesos Rápidos SISEC / RUNT / HqRunt CEA",
            "contenido": "Portal SISEC Enseñanza, Liquidación PIN, RUNT y CONFE-AFILIACIONES listos para consulta rápida.",
            "fecha": ahora_str,
            "origen": "Baúl Bot Telegram + Web",
        },
        {
            "id": "baul-02",
            "categoria": "CONTABILIDAD",
            "titulo": "Cuentas Contables Auxiliar Siigo (4175 - 4180)",
            "contenido": "Cuenta 41800101: Servicios de Enseñanza (FV-2) | Cuenta 41750502: Devoluciones / Notas Crédito (NC-2).",
            "fecha": ahora_str,
            "origen": "Baúl Bot Telegram + Web",
        },
        {
            "id": "baul-03",
            "categoria": "NOTA / PENDIENTE",
            "titulo": "Recordatorio Operativo Examen Médico CRC La 27",
            "contenido": "Verificar diariamente los alumnos reportados en el grupo Exámenes Autooriente - CRC LA 27 y cruzar con admisiones.",
            "fecha": ahora_str,
            "origen": "Baúl Bot Telegram + Web",
        },
    ]
    try:
        BAUL_LOCAL_PATH.write_text(json.dumps(semilla_baul, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return semilla_baul


def _guardar_baul_local(items: List[Dict[str, Any]]) -> None:
    try:
        BAUL_LOCAL_PATH.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def guardar_item_baul(
    titulo: str,
    contenido: str,
    categoria: str = "NOTA",
    origen: str = "Dashboard Web",
    notificar_tg: bool = True,
    chat_id: str = "",
) -> Tuple[bool, str]:
    if not titulo.strip() and not contenido.strip():
        return False, "Escribe un título o contenido para guardar en el Baúl."
    ahora = datetime.datetime.now(LOCAL_TZ)
    items = _cargar_baul_local()
    tit_clean = titulo.strip() or contenido.strip()[:45]
    nuevo = {
        "id": f"baul-{int(ahora.timestamp() * 1000)}",
        "categoria": (categoria or "NOTA").strip().upper(),
        "titulo": tit_clean,
        "contenido": contenido.strip() or tit_clean,
        "fecha": ahora.strftime("%Y-%m-%d %H:%M"),
        "origen": origen,
    }
    items.insert(0, nuevo)
    _guardar_baul_local(items)

    if notificar_tg:
        cid = obtener_chat_id_agenda(chat_id)
        _enviar_mensaje_telegram_raw(
            cid,
            (
                f"🗄️ *NUEVO ELEMENTO GUARDADO EN TU BAÚL DIGITAL*\n\n"
                f"🏷️ *Categoría:* {nuevo['categoria']}\n"
                f"📌 *Título:* {nuevo['titulo']}\n"
                f"📝 *Detalle:* {nuevo['contenido']}\n"
                f"🕒 *Fecha:* {nuevo['fecha']} ({origen})"
            ),
        )
    return True, f"Guardado en el Baúl Digital: '{tit_clean}'."


def eliminar_item_baul(item_id: str) -> Tuple[bool, str]:
    items = _cargar_baul_local()
    nuevos = [x for x in items if str(x.get("id")) != str(item_id)]
    _guardar_baul_local(nuevos)
    return True, "Elemento eliminado del Baúl Digital."


def enviar_item_baul_telegram(item_id: str, chat_id: str = "") -> Tuple[bool, str]:
    items = _cargar_baul_local()
    target = next((x for x in items if str(x.get("id")) == str(item_id)), None)
    if not target:
        return False, "No se encontró el elemento en el Baúl."
    cid = obtener_chat_id_agenda(chat_id)
    texto = (
        f"🗄️ *BAÚL DIGITAL CEA AUTO ORIENTE*\n\n"
        f"🏷️ *Categoría:* {target.get('categoria', 'NOTA')}\n"
        f"📌 *{target.get('titulo', '')}*\n\n"
        f"{target.get('contenido', '')}\n\n"
        f"🕒 _Guardado: {target.get('fecha', '')}_"
    )
    return _enviar_mensaje_telegram_raw(cid, texto)


# =============================================================================
# 2. CONEXIÓN CON GOOGLE CALENDAR API Y SINCRONIZACIÓN BIDIRECCIONAL
# =============================================================================
def get_calendar_service():
    """Autentica con Google Calendar usando token_calendar.json, token.json o GOOGLE_CALENDAR_TOKEN_JSON."""
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    env_token = os.getenv("GOOGLE_CALENDAR_TOKEN_JSON", "").strip() or DEFAULT_CALENDAR_TOKEN_JSON
    if env_token and env_token.startswith("{") and not TOKEN_CALENDAR_PATH.exists():
        for dest_p in (TOKEN_CALENDAR_PATH, Path("/tmp/token_calendar.json")):
            try:
                dest_p.write_text(env_token, encoding="utf-8")
            except Exception:
                pass

    creds = None
    for cand in (TOKEN_CALENDAR_PATH, TOKEN_FALLBACK_PATH, Path("/tmp/token_calendar.json")):
        if cand.exists():
            try:
                creds = Credentials.from_authorized_user_file(str(cand), SCOPES)
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                    cand.write_text(creds.to_json(), encoding="utf-8")
                if creds and creds.valid:
                    break
            except Exception:
                creds = None

    if not creds or not creds.valid:
        raise RuntimeError("Token de Google Calendar no configurado aún. Pégalo en 'Conectar Google Calendar y Bot'.")
    return build("calendar", "v3", credentials=creds)


def sincronizar_eventos_locales_a_gcal() -> Tuple[int, str]:
    """
    Sube a Google Calendar todos los eventos creados en el Dashboard que aún estén
    en estado 'PROGRAMADO' (como el evento de las 4:00 PM / 16:00).
    """
    try:
        service = get_calendar_service()
    except Exception as exc:
        return 0, str(exc)

    evs = _cargar_eventos_locales()
    subidos = 0
    for ev in evs:
        if ev.get("estado") == "SINCRONIZADO GCAL" or str(ev.get("id", "")).startswith("gcal-"):
            continue
        try:
            titulo = str(ev.get("titulo", "Evento CEA")).strip()
            fecha_usar = str(ev.get("fecha", datetime.datetime.now(LOCAL_TZ).strftime("%Y-%m-%d"))).strip()
            hora_usar = _extraer_hora_de_titulo(titulo, str(ev.get("hora", "09:00")).strip())
            ev["hora"] = hora_usar
            duracion_min = int(ev.get("duracion_min") or 60)
            es_diario = ev.get("tipo") == "DIARIO"
            responsable = str(ev.get("responsable", "Gerencia CEA"))

            h, m = map(int, hora_usar.split(":")[:2])
            inicio_dt = datetime.datetime.strptime(f"{fecha_usar} {h:02d}:{m:02d}", "%Y-%m-%d %H:%M").replace(tzinfo=LOCAL_TZ)
            fin_dt = inicio_dt + datetime.timedelta(minutes=duracion_min)
            cuerpo: Dict[str, Any] = {
                "summary": titulo,
                "description": f"Responsable: {responsable} | Sincronizado desde Dashboard CEA",
                "start": {"dateTime": inicio_dt.isoformat(), "timeZone": TIMEZONE_STR},
                "end": {"dateTime": fin_dt.isoformat(), "timeZone": TIMEZONE_STR},
            }
            if es_diario:
                cuerpo["recurrence"] = ["RRULE:FREQ=DAILY"]
            created = service.events().insert(calendarId="primary", body=cuerpo).execute()
            ev["gcal_id"] = created.get("id", "")
            ev["estado"] = "SINCRONIZADO GCAL"
            subidos += 1
        except Exception:
            continue

    if subidos > 0:
        _guardar_eventos_locales(evs)
    return subidos, f"{subidos} evento(s) sincronizados con Google Calendar."


def configurar_credenciales_agenda_en_vivo(
    tg_token: str = "",
    chat_id: str = "",
    calendar_token_json: str = "",
) -> Dict[str, Any]:
    """
    Guarda el Token del Bot de Agenda de Telegram, el Chat ID y el token.json de Google Calendar,
    sincroniza inmediatamente los eventos locales (incluido el de las 16:00) con Google Calendar,
    enciende el Bot interactivo en segundo plano y envía confirmación a Telegram.
    """
    cfg = _guardar_config_agenda(tg_token=tg_token, chat_id=chat_id, calendar_token_json=calendar_token_json)
    mensajes: List[str] = []
    gcal_ok = False
    tg_ok = False

    try:
        subidos, _ = sincronizar_eventos_locales_a_gcal()
        get_calendar_service()
        gcal_ok = True
        mensajes.append(f"Google Calendar conectado ({subidos} eventos pendientes sincronizados)")
    except Exception as exc:
        mensajes.append(f"Calendar local ({exc})")

    if obtener_token_telegram_agenda():
        iniciar_daemon_agenda_telegram()
        cid = obtener_chat_id_agenda(chat_id)
        ok_send, msg_send = _enviar_mensaje_telegram_raw(
            cid,
            (
                "✅ *BOT DE AGENDA + GOOGLE CALENDAR + BAÚL DIGITAL CONECTADOS*\n\n"
                "🔗 Tu Dashboard Web de *CEA Auto Oriente* y este Bot de Telegram ya están 100% sincronizados en tiempo real.\n"
                "📅 Todos los eventos que crees en la Web o con `/start` se sincronizan y notifican aquí.\n"
                "🗄️ Usa el botón *Mi Baúl Digital* en `/menu` o escribe `/baul` para consultar y guardar notas/archivos."
            ),
            reply_markup=_teclado_principal_dict(),
        )
        tg_ok = ok_send
        mensajes.append(msg_send)
    else:
        mensajes.append("Falta ingresar el Token de Telegram del Bot de Agenda.")

    return {
        "ok": gcal_ok or tg_ok,
        "calendar_conectado": gcal_ok,
        "telegram_conectado": tg_ok,
        "chat_id": cfg.get("chat_id", "1567561963"),
        "message": " • ".join(mensajes),
    }


def obtener_panorama_agenda_web() -> Dict[str, Any]:
    """
    Obtiene los eventos de Hoy, Próximos 7 Días y el Baúl Digital combinando Google Calendar
    en vivo con el almacén persistente del Dashboard CEA.
    """
    iniciar_daemon_agenda_telegram()
    ahora = datetime.datetime.now(LOCAL_TZ)
    hoy_str = ahora.strftime("%Y-%m-%d")
    limite_7d = (ahora + datetime.timedelta(days=7)).strftime("%Y-%m-%d")

    eventos_locales = _cargar_eventos_locales()
    # Corregir automáticamente horas en eventos locales si decían "reunion a las 4 pm" con 09:00
    modificado_local = False
    for ev_loc in eventos_locales:
        h_corr = _extraer_hora_de_titulo(str(ev_loc.get("titulo", "")), str(ev_loc.get("hora", "09:00")))
        if h_corr != ev_loc.get("hora"):
            ev_loc["hora"] = h_corr
            modificado_local = True
    if modificado_local:
        _guardar_eventos_locales(eventos_locales)

    calendar_conectado = False
    eventos_gcal: List[Dict[str, Any]] = []

    try:
        sincronizar_eventos_locales_a_gcal()
        service = get_calendar_service()
        inicio_dia = ahora.replace(hour=0, minute=0, second=0, microsecond=0).isoformat()
        fin_7d = (ahora + datetime.timedelta(days=7)).replace(hour=23, minute=59, second=59).isoformat()
        res = (
            service.events()
            .list(
                calendarId="primary",
                timeMin=inicio_dia,
                timeMax=fin_7d,
                singleEvents=True,
                orderBy="startTime",
            )
            .execute()
        )
        calendar_conectado = True
        for ev in res.get("items", []):
            start_raw = ev.get("start", {}).get("dateTime") or ev.get("start", {}).get("date") or ""
            if "T" in start_raw:
                dt = datetime.datetime.fromisoformat(start_raw)
                f_str = dt.strftime("%Y-%m-%d")
                h_str = dt.strftime("%H:%M")
            else:
                f_str = start_raw[:10]
                h_str = "Todo el día"
            eventos_gcal.append({
                "id": ev.get("id", ""),
                "titulo": ev.get("summary", "Sin título"),
                "fecha": f_str,
                "hora": h_str,
                "duracion_min": 60,
                "tipo": "DIARIO" if ev.get("recurringEventId") else "PUNTUAL",
                "responsable": "Google Calendar",
                "estado": "SINCRONIZADO GCAL",
                "origen": "Google Calendar API",
            })
    except Exception:
        calendar_conectado = False

    vistos = set()
    todos: List[Dict[str, Any]] = []
    for ev in eventos_gcal + eventos_locales:
        ev_copy = dict(ev)
        if ev_copy.get("tipo") == "DIARIO" and ev_copy.get("fecha", "") < hoy_str:
            ev_copy["fecha"] = hoy_str
        clave = (ev_copy.get("titulo", "").strip().upper(), ev_copy.get("fecha", ""), ev_copy.get("hora", ""))
        if clave in vistos:
            continue
        vistos.add(clave)
        todos.append(ev_copy)

    todos.sort(key=lambda x: (x.get("fecha", ""), x.get("hora", "")))
    eventos_hoy = [e for e in todos if e.get("fecha") == hoy_str or e.get("tipo") == "DIARIO"]
    eventos_semana = [e for e in todos if hoy_str <= e.get("fecha", "") <= limite_7d or e.get("tipo") == "DIARIO"]
    baul_items = _cargar_baul_local()
    tk_cfg = obtener_token_telegram_agenda()

    return {
        "ok": True,
        "calendar_conectado": calendar_conectado,
        "bot_token_configurado": bool(tk_cfg),
        "bot_polling_activo": _DAEMON_STARTED and bool(tk_cfg),
        "chat_id_actual": obtener_chat_id_agenda(),
        "fecha_hoy": ahora.strftime("%d/%m/%Y"),
        "kpis": {
            "total_hoy": len(eventos_hoy),
            "total_semana": len(eventos_semana),
            "recurrentes_diarios": sum(1 for e in todos if e.get("tipo") == "DIARIO"),
            "puntuales": sum(1 for e in todos if e.get("tipo") != "DIARIO"),
            "total_baul": len(baul_items),
        },
        "eventos_hoy": eventos_hoy,
        "eventos_semana": eventos_semana,
        "baul_items": baul_items,
    }


def crear_evento_agenda(
    titulo: str,
    fecha: str,
    hora: str,
    duracion_min: int = 60,
    es_diario: bool = False,
    responsable: str = "Gerencia CEA",
    notificar_telegram: bool = True,
    chat_id_override: str = "",
) -> Tuple[bool, str]:
    ahora = datetime.datetime.now(LOCAL_TZ)
    fecha_usar = fecha.strip() if fecha else ahora.strftime("%Y-%m-%d")
    if "/" in fecha_usar:
        try:
            fecha_usar = datetime.datetime.strptime(fecha_usar, "%d/%m/%Y").strftime("%Y-%m-%d")
        except Exception:
            pass
    hora_usar = _extraer_hora_de_titulo(titulo, hora.strip() if hora else "09:00")

    if chat_id_override and str(chat_id_override).strip():
        _guardar_config_agenda(chat_id=str(chat_id_override).strip())

    nuevo_ev = {
        "id": f"ev-{int(ahora.timestamp() * 1000)}",
        "titulo": titulo.strip(),
        "fecha": fecha_usar,
        "hora": hora_usar,
        "duracion_min": int(duracion_min or 60),
        "tipo": "DIARIO" if es_diario else "PUNTUAL",
        "responsable": responsable.strip() or "Gerencia CEA",
        "estado": "PROGRAMADO",
        "origen": "Dashboard Web + Bot Agenda",
    }

    evs = _cargar_eventos_locales()
    evs.insert(0, nuevo_ev)
    _guardar_eventos_locales(evs)

    gcal_msg = "Guardado en Agenda CEA"
    try:
        service = get_calendar_service()
        h, m = map(int, hora_usar.split(":")[:2])
        inicio_dt = datetime.datetime.strptime(f"{fecha_usar} {h:02d}:{m:02d}", "%Y-%m-%d %H:%M").replace(tzinfo=LOCAL_TZ)
        fin_dt = inicio_dt + datetime.timedelta(minutes=int(duracion_min or 60))
        cuerpo: Dict[str, Any] = {
            "summary": titulo.strip(),
            "description": f"Responsable: {responsable}",
            "start": {"dateTime": inicio_dt.isoformat(), "timeZone": TIMEZONE_STR},
            "end": {"dateTime": fin_dt.isoformat(), "timeZone": TIMEZONE_STR},
        }
        if es_diario:
            cuerpo["recurrence"] = ["RRULE:FREQ=DAILY"]
        created = service.events().insert(calendarId="primary", body=cuerpo).execute()
        nuevo_ev["gcal_id"] = created.get("id", "")
        nuevo_ev["estado"] = "SINCRONIZADO GCAL"
        _guardar_eventos_locales(evs)
        gcal_msg = "Sincronizado en Google Calendar"
    except Exception:
        pass

    tg_msg = ""
    if notificar_telegram:
        ok_tg, info_tg = enviar_recordatorio_telegram_agenda(
            titulo=titulo,
            fecha=fecha_usar,
            hora=hora_usar,
            tipo="EVENTO AGENDADO (DIARIO)" if es_diario else "EVENTO AGENDADO",
            responsable=responsable,
            chat_id=chat_id_override,
        )
        tg_msg = f" • {'Notificado en Telegram' if ok_tg else info_tg}"

    return True, f"Evento '{titulo}' ({'Diario' if es_diario else fecha_usar} a las {hora_usar}) creado — {gcal_msg}{tg_msg}."


def eliminar_evento_agenda(ev_id: str) -> Tuple[bool, str]:
    evs = _cargar_eventos_locales()
    target = next((e for e in evs if str(e.get("id")) == str(ev_id)), None)
    nuevos = [e for e in evs if str(e.get("id")) != str(ev_id)]
    _guardar_eventos_locales(nuevos)
    try:
        service = get_calendar_service()
        gcal_id = (target or {}).get("gcal_id") or ev_id
        if gcal_id and not str(gcal_id).startswith("ev-"):
            service.events().delete(calendarId="primary", eventId=str(gcal_id)).execute()
    except Exception:
        pass
    return True, "Evento eliminado de la agenda."


# =============================================================================
# 3. ENVÍO Y BOT INTERACTIVO DE TELEGRAM (AGENDA + BAÚL DIGITAL + RECORDATORIOS)
# =============================================================================
def _teclado_principal_dict() -> Dict[str, Any]:
    return {
        "inline_keyboard": [
            [{"text": "☀️ Ver Agenda de Hoy", "callback_data": "btn_hoy"}],
            [{"text": "📅 Ver Próximos 7 Días", "callback_data": "btn_semana"}],
            [
                {"text": "➕ Nuevo Evento Puntual", "callback_data": "btn_crear"},
                {"text": "🔁 Nuevo Evento Diario", "callback_data": "btn_diario"},
            ],
            [
                {"text": "🗄️ Ver Mi Baúl Digital", "callback_data": "btn_baul"},
                {"text": "📌 Guardar Nota en Baúl", "callback_data": "btn_baul_nuevo"},
            ],
        ]
    }


def _llamar_telegram_api(metodo: str, payload: Dict[str, Any], token_override: str = "") -> Tuple[bool, Dict[str, Any]]:
    candidatos = [
        token_override.strip(),
        obtener_token_telegram_agenda(),
        os.getenv("TELEGRAM_BOT_TOKEN", "").strip(),
    ]
    vistos = set()
    for tk in candidatos:
        if not tk or tk == "TU_TELEGRAM_BOT_TOKEN" or tk in vistos:
            continue
        vistos.add(tk)
        try:
            url = f"https://api.telegram.org/bot{tk}/{metodo}"
            raw = json.dumps(payload).encode("utf-8")
            req = urllib.request.Request(url, data=raw, headers={"Content-Type": "application/json"}, method="POST")
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8", errors="ignore"))
                if data.get("ok"):
                    return True, data
        except Exception:
            continue
    return False, {}


def _enviar_mensaje_telegram_raw(
    chat_id: str,
    texto: str,
    reply_markup: Optional[Dict[str, Any]] = None,
) -> Tuple[bool, str]:
    cid = obtener_chat_id_agenda(chat_id)
    if not cid:
        return False, "Ingresa tu Chat ID de Telegram (ej. 1567561963)."
    payload: Dict[str, Any] = {"chat_id": cid, "text": texto, "parse_mode": "Markdown"}
    if reply_markup:
        payload["reply_markup"] = reply_markup
    ok, _ = _llamar_telegram_api("sendMessage", payload)
    if ok:
        return True, f"Enviado a Telegram (Chat {cid})."
    return False, (
        "Falta conectar el Token del Bot de Agenda en la franja superior de este módulo "
        "(pega tu TELEGRAM_TOKEN y haz clic en 'Conectar Calendar y Bot Telegram')."
    )


def enviar_recordatorio_telegram_agenda(
    titulo: str,
    fecha: str,
    hora: str,
    tipo: str = "RECORDATORIO",
    responsable: str = "CEA Auto Oriente",
    chat_id: str = "",
) -> Tuple[bool, str]:
    cid = obtener_chat_id_agenda(chat_id)
    texto = (
        f"⏰ *AGENDA CEA AUTO ORIENTE — {tipo}*\n\n"
        f"📌 *Actividad:* {titulo}\n"
        f"📅 *Fecha:* {fecha}\n"
        f"🕒 *Hora:* {hora}\n"
        f"👤 *Responsable:* {responsable}"
    )
    return _enviar_mensaje_telegram_raw(cid, texto, reply_markup=_teclado_principal_dict())


def _formatear_texto_agenda_hoy() -> str:
    pan = obtener_panorama_agenda_web()
    evs = pan.get("eventos_hoy", [])
    if not evs:
        return "🌴 No tienes actividades agendadas para hoy."
    lineas = [f"📅 *Tu Agenda para Hoy ({pan.get('fecha_hoy', '')}):*\n"]
    for e in evs:
        tag = "🔁" if e.get("tipo") == "DIARIO" else "▫️"
        lineas.append(f"{tag} *{e.get('hora', '')}* — {e.get('titulo', 'Sin título')} _({e.get('responsable', 'CEA')})_")
    return "\n".join(lineas)


def _formatear_texto_agenda_semana() -> str:
    pan = obtener_panorama_agenda_web()
    evs = pan.get("eventos_semana", [])
    if not evs:
        return "No tienes eventos en los próximos 7 días."
    lineas = ["📅 *Eventos para los próximos 7 días:*\n"]
    for e in evs:
        lineas.append(f"• *{e.get('fecha', '')} ({e.get('hora', '')})* : {e.get('titulo', 'Sin título')}")
    return "\n".join(lineas)


def _formatear_texto_baul() -> str:
    items = _cargar_baul_local()
    if not items:
        return "🗄️ Tu Baúl Digital está vacío. Usa '📌 Guardar Nota en Baúl' o envíame cualquier nota."
    lineas = ["🗄️ *TU BAÚL DIGITAL (SINCRONIZADO CON EL DASHBOARD WEB):*\n"]
    for idx, it in enumerate(items[:12], start=1):
        lineas.append(
            f"*{idx}. [{it.get('categoria', 'NOTA')}] {it.get('titulo', '')}*\n"
            f"   _{it.get('contenido', '')}_\n"
            f"   🕒 `{it.get('fecha', '')}`"
        )
    return "\n\n".join(lineas)


def _procesar_update_telegram(upd: Dict[str, Any]) -> None:
    tk = obtener_token_telegram_agenda()
    if not tk:
        return

    # 1. Botones Inline (CallbackQuery)
    if "callback_query" in upd:
        cq = upd["callback_query"]
        cq_id = cq.get("id")
        data = str(cq.get("data", ""))
        msg = cq.get("message") or {}
        chat_id = str((msg.get("chat") or {}).get("id", ""))
        msg_id = msg.get("message_id")
        if chat_id:
            _guardar_config_agenda(chat_id=chat_id)
        _llamar_telegram_api("answerCallbackQuery", {"callback_query_id": cq_id}, token_override=tk)

        btn_volver = {"inline_keyboard": [[{"text": "⬅️ Volver al Menú", "callback_data": "btn_volver"}]]}

        if data == "btn_volver":
            _CONVERSATION_STATE.pop(chat_id, None)
            _llamar_telegram_api(
                "editMessageText",
                {
                    "chat_id": chat_id,
                    "message_id": msg_id,
                    "text": "👋 *Asistente de Agenda + Baúl Digital CEA*\n\nSelecciona una opción del menú:",
                    "parse_mode": "Markdown",
                    "reply_markup": _teclado_principal_dict(),
                },
                token_override=tk,
            )
            return

        if data == "btn_hoy":
            _llamar_telegram_api(
                "editMessageText",
                {
                    "chat_id": chat_id,
                    "message_id": msg_id,
                    "text": _formatear_texto_agenda_hoy(),
                    "parse_mode": "Markdown",
                    "reply_markup": btn_volver,
                },
                token_override=tk,
            )
            return

        if data == "btn_semana":
            _llamar_telegram_api(
                "editMessageText",
                {
                    "chat_id": chat_id,
                    "message_id": msg_id,
                    "text": _formatear_texto_agenda_semana(),
                    "parse_mode": "Markdown",
                    "reply_markup": btn_volver,
                },
                token_override=tk,
            )
            return

        if data == "btn_baul":
            _llamar_telegram_api(
                "editMessageText",
                {
                    "chat_id": chat_id,
                    "message_id": msg_id,
                    "text": _formatear_texto_baul(),
                    "parse_mode": "Markdown",
                    "reply_markup": {
                        "inline_keyboard": [
                            [{"text": "📌 Guardar Nueva Nota en el Baúl", "callback_data": "btn_baul_nuevo"}],
                            [{"text": "⬅️ Volver al Menú", "callback_data": "btn_volver"}],
                        ]
                    },
                },
                token_override=tk,
            )
            return

        if data == "btn_baul_nuevo":
            _CONVERSATION_STATE[chat_id] = {"step": "BAUL_TEXTO"}
            _llamar_telegram_api(
                "sendMessage",
                {
                    "chat_id": chat_id,
                    "text": (
                        "🗄️ *Guardar en tu Baúl Digital*\n\n"
                        "Escribe la nota, recordatorio, enlace o apunte que deseas guardar "
                        "(usa el formato `Título: Detalle` o escribe directamente el texto; envía /cancelar para salir):"
                    ),
                    "parse_mode": "Markdown",
                },
                token_override=tk,
            )
            return

        if data in ("btn_crear", "btn_diario"):
            es_diario = data == "btn_diario"
            _CONVERSATION_STATE[chat_id] = {"step": "TITULO", "es_diario": es_diario}
            tipo_txt = "recurrente diario" if es_diario else "puntual"
            _llamar_telegram_api(
                "sendMessage",
                {
                    "chat_id": chat_id,
                    "text": (
                        f"✍️ Vamos a crear un evento *{tipo_txt}*.\n\n"
                        "Escribe el *título o nombre* del evento (o envía /cancelar para salir):"
                    ),
                    "parse_mode": "Markdown",
                },
                token_override=tk,
            )
            return

        if data.startswith("fecha_"):
            st = _CONVERSATION_STATE.get(chat_id) or {}
            st["fecha"] = data.replace("fecha_", "")
            st["step"] = "HORA"
            _CONVERSATION_STATE[chat_id] = st
            _llamar_telegram_api(
                "sendMessage",
                {
                    "chat_id": chat_id,
                    "text": "🕒 ¿A qué hora comienza el evento?\n\nEjemplos: `08:00`, `16:00`, `19:30` (formato 24 horas):",
                    "parse_mode": "Markdown",
                },
                token_override=tk,
            )
            return

        if data.startswith("dur_"):
            st = _CONVERSATION_STATE.pop(chat_id, None) or {}
            dur_str = data.replace("dur_", "")
            dur_min = 480 if dur_str == "allday" else int(dur_str or 60)
            titulo = st.get("titulo", "Actividad CEA")
            fecha_ev = st.get("fecha", datetime.datetime.now(LOCAL_TZ).strftime("%Y-%m-%d"))
            hora_ev = st.get("hora", "09:00")
            es_diario = bool(st.get("es_diario", False))

            crear_evento_agenda(
                titulo=titulo,
                fecha=fecha_ev,
                hora=hora_ev,
                duracion_min=dur_min,
                es_diario=es_diario,
                responsable="Bot Telegram Agenda",
                notificar_telegram=False,
                chat_id_override=chat_id,
            )
            _llamar_telegram_api(
                "editMessageText",
                {
                    "chat_id": chat_id,
                    "message_id": msg_id,
                    "text": (
                        f"✅ *Evento Agendado y Sincronizado con el Dashboard Web*\n\n"
                        f"📌 *{titulo}*\n"
                        f"📅 {fecha_ev if not es_diario else 'Todos los días'}\n"
                        f"🕒 {hora_ev} ({dur_min} min)"
                    ),
                    "parse_mode": "Markdown",
                    "reply_markup": _teclado_principal_dict(),
                },
                token_override=tk,
            )
            return

    # 2. Mensajes de Texto o Comandos
    if "message" in upd:
        msg = upd["message"]
        chat_id = str((msg.get("chat") or {}).get("id", ""))
        texto = str(msg.get("text") or msg.get("caption") or "").strip()
        if not chat_id or not texto:
            return
        _guardar_config_agenda(chat_id=chat_id)

        if texto.lower() in ("/start", "/menu", "menu", "menú", "hola"):
            _CONVERSATION_STATE.pop(chat_id, None)
            _llamar_telegram_api(
                "sendMessage",
                {
                    "chat_id": chat_id,
                    "text": (
                        "👋 *Bienvenido a tu Asistente Personal de Agenda y Baúl Digital — CEA Auto Oriente*\n\n"
                        "Conectado en tiempo real con tu Dashboard Web y Google Calendar. Selecciona una opción:"
                    ),
                    "parse_mode": "Markdown",
                    "reply_markup": _teclado_principal_dict(),
                },
                token_override=tk,
            )
            return

        if texto.lower() == "/cancelar":
            _CONVERSATION_STATE.pop(chat_id, None)
            _llamar_telegram_api(
                "sendMessage",
                {
                    "chat_id": chat_id,
                    "text": "Operación cancelada.",
                    "reply_markup": _teclado_principal_dict(),
                },
                token_override=tk,
            )
            return

        if texto.lower().startswith("/baul"):
            resto = texto[5:].strip()
            if resto:
                partes = resto.split(":", 1)
                tit = partes[0].strip()
                cont = partes[1].strip() if len(partes) > 1 else resto
                guardar_item_baul(tit, cont, categoria="NOTA TELEGRAM", origen="Bot Telegram", notificar_tg=False)
                _llamar_telegram_api(
                    "sendMessage",
                    {
                        "chat_id": chat_id,
                        "text": f"🗄️ ✅ Guardado en tu Baúl Digital y sincronizado en el Dashboard Web:\n*{tit}*",
                        "parse_mode": "Markdown",
                        "reply_markup": _teclado_principal_dict(),
                    },
                    token_override=tk,
                )
            else:
                _llamar_telegram_api(
                    "sendMessage",
                    {
                        "chat_id": chat_id,
                        "text": _formatear_texto_baul(),
                        "parse_mode": "Markdown",
                        "reply_markup": _teclado_principal_dict(),
                    },
                    token_override=tk,
                )
            return

        st = _CONVERSATION_STATE.get(chat_id)
        if st:
            step = st.get("step")
            if step == "BAUL_TEXTO":
                _CONVERSATION_STATE.pop(chat_id, None)
                partes = texto.split(":", 1)
                tit = partes[0].strip() if len(partes) > 1 else texto[:40]
                cont = partes[1].strip() if len(partes) > 1 else texto
                guardar_item_baul(tit, cont, categoria="NOTA TELEGRAM", origen="Bot Telegram", notificar_tg=False)
                _llamar_telegram_api(
                    "sendMessage",
                    {
                        "chat_id": chat_id,
                        "text": f"🗄️ ✅ *Guardado en tu Baúl Digital (Web + Telegram):*\n\n📌 *{tit}*\n📝 {cont}",
                        "parse_mode": "Markdown",
                        "reply_markup": _teclado_principal_dict(),
                    },
                    token_override=tk,
                )
                return

            if step == "TITULO":
                st["titulo"] = texto
                if st.get("es_diario"):
                    st["step"] = "HORA"
                    _llamar_telegram_api(
                        "sendMessage",
                        {
                            "chat_id": chat_id,
                            "text": "🕒 ¿A qué hora debe comenzar todos los días?\n\nEjemplos: `07:30`, `14:00`, `16:00`:",
                            "parse_mode": "Markdown",
                        },
                        token_override=tk,
                    )
                else:
                    st["step"] = "FECHA"
                    ahora = datetime.datetime.now(LOCAL_TZ)
                    hoy_str = ahora.strftime("%Y-%m-%d")
                    man_str = (ahora + datetime.timedelta(days=1)).strftime("%Y-%m-%d")
                    _llamar_telegram_api(
                        "sendMessage",
                        {
                            "chat_id": chat_id,
                            "text": "📅 Elige la fecha o escríbela en formato `DD/MM/AAAA` (ej: `14/11/2026`):",
                            "parse_mode": "Markdown",
                            "reply_markup": {
                                "inline_keyboard": [
                                    [{"text": f"Hoy ({ahora.strftime('%d/%m')})", "callback_data": f"fecha_{hoy_str}"}],
                                    [{"text": f"Mañana ({(ahora + datetime.timedelta(days=1)).strftime('%d/%m')})", "callback_data": f"fecha_{man_str}"}],
                                ]
                            },
                        },
                        token_override=tk,
                    )
                return

            if step == "FECHA":
                try:
                    dt = datetime.datetime.strptime(texto, "%d/%m/%Y")
                    st["fecha"] = dt.strftime("%Y-%m-%d")
                    st["step"] = "HORA"
                    _llamar_telegram_api(
                        "sendMessage",
                        {
                            "chat_id": chat_id,
                            "text": "🕒 ¿A qué hora comienza el evento?\n\nEjemplos: `08:00`, `16:00`, `19:00`:",
                            "parse_mode": "Markdown",
                        },
                        token_override=tk,
                    )
                except ValueError:
                    _llamar_telegram_api(
                        "sendMessage",
                        {
                            "chat_id": chat_id,
                            "text": "⚠️ Fecha inválida. Usa el formato `DD/MM/AAAA` (ej: `14/11/2026`):",
                            "parse_mode": "Markdown",
                        },
                        token_override=tk,
                    )
                return

            if step == "HORA":
                hora_norm = _extraer_hora_de_titulo(f"a las {texto}", "09:00") if ("am" in texto.lower() or "pm" in texto.lower()) else texto
                try:
                    datetime.datetime.strptime(hora_norm, "%H:%M")
                    st["hora"] = hora_norm
                    st["step"] = "DURACION"
                    _llamar_telegram_api(
                        "sendMessage",
                        {
                            "chat_id": chat_id,
                            "text": "⏳ ¿Cuánto dura la actividad?",
                            "reply_markup": {
                                "inline_keyboard": [
                                    [
                                        {"text": "30 min", "callback_data": "dur_30"},
                                        {"text": "45 min", "callback_data": "dur_45"},
                                    ],
                                    [
                                        {"text": "1 hora", "callback_data": "dur_60"},
                                        {"text": "2 horas", "callback_data": "dur_120"},
                                    ],
                                    [{"text": "Todo el día", "callback_data": "dur_allday"}],
                                ]
                            },
                        },
                        token_override=tk,
                    )
                except ValueError:
                    _llamar_telegram_api(
                        "sendMessage",
                        {
                            "chat_id": chat_id,
                            "text": "⚠️ Hora no válida. Escríbela en formato HH:MM (ej: `08:30` o `16:00`):",
                            "parse_mode": "Markdown",
                        },
                        token_override=tk,
                    )
                return


def _verificar_recordatorios_proximos() -> None:
    """
    Revisa cada 60 segundos los eventos de Hoy (tanto en Google Calendar como en la Agenda CEA)
    y envía alerta automática a Telegram cuando faltan entre 0 y 10 minutos para iniciar.
    """
    tk = obtener_token_telegram_agenda()
    cid = obtener_chat_id_agenda()
    if not tk or not cid:
        return

    ahora = datetime.datetime.now(LOCAL_TZ)
    hoy_str = ahora.strftime("%Y-%m-%d")
    pan = obtener_panorama_agenda_web()
    for ev in pan.get("eventos_hoy", []):
        ev_id = f"{ev.get('id', '')}_{hoy_str}_{ev.get('hora', '')}"
        if ev_id in EVENTOS_NOTIFICADOS:
            continue
        hora_str = str(ev.get("hora", "")).strip()
        if ":" not in hora_str:
            continue
        try:
            h, m = map(int, hora_str.split(":")[:2])
            dt_ev = ahora.replace(hour=h, minute=m, second=0, microsecond=0)
            minutos = (dt_ev - ahora).total_seconds() / 60.0
            if 0 <= minutos <= 10:
                EVENTOS_NOTIFICADOS.add(ev_id)
                _enviar_mensaje_telegram_raw(
                    cid,
                    (
                        f"⏰ *RECORDATORIO AUTOMÁTICO DE TU AGENDA*\n\n"
                        f"📌 *{ev.get('titulo', 'Actividad')}*\n"
                        f"🕒 Comienza en *{int(round(minutos))} minutos* ({hora_str})\n"
                        f"👤 Responsable: {ev.get('responsable', 'CEA Auto Oriente')}"
                    ),
                )
        except Exception:
            continue


def _loop_daemon_agenda_telegram() -> None:
    offset = 0
    ultimo_check_recordatorios = 0.0
    while True:
        try:
            ahora_ts = time.time()
            if ahora_ts - ultimo_check_recordatorios >= 55.0:
                ultimo_check_recordatorios = ahora_ts
                _verificar_recordatorios_proximos()

            tk = obtener_token_telegram_agenda()
            tk_main_clases = os.getenv("TELEGRAM_BOT_TOKEN", "").strip()
            # Si el token de agenda es el mismo del bot principal de clases (que ya hace polling en main.py),
            # no llamar getUpdates aquí para evitar conflicto 409; main.py procesa los comandos y botones de agenda directamente.
            if not tk or (tk_main_clases and tk == tk_main_clases and os.getenv("UNIFIED_MAIN_BOT_ACTIVE") == "1"):
                time.sleep(5)
                continue

            ok, data = _llamar_telegram_api(
                "getUpdates",
                {"offset": offset, "timeout": 10, "allowed_updates": ["message", "callback_query"]},
                token_override=tk,
            )
            if ok and isinstance(data.get("result"), list):
                for upd in data["result"]:
                    offset = max(offset, int(upd.get("update_id", 0)) + 1)
                    try:
                        _procesar_update_telegram(upd)
                    except Exception:
                        pass
            else:
                time.sleep(4)
        except Exception:
            time.sleep(5)


def iniciar_daemon_agenda_telegram() -> None:
    global _DAEMON_STARTED
    with _DAEMON_LOCK:
        if _DAEMON_STARTED:
            return
        _DAEMON_STARTED = True
        t = threading.Thread(target=_loop_daemon_agenda_telegram, name="AgendaTelegramDaemon", daemon=True)
        t.start()
