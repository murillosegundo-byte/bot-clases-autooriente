"""
=============================================================================
SERVIDOR WEB + BOT DE TELEGRAM INTEGRADO - AUDITORÍA CONTABLE CEA AUTO ORIENTE
=============================================================================
Incluye:
- Motor de Consolidación Multi-Curso por Cédula (231 clientes con 2+ cursos simultáneos).
- Separación inteligente entre Riesgo Cero Real ($0 en todos sus cursos) y Combos Multi-Curso
  (2da matrícula en $0 pero con abono en su 1ra matrícula).
- Carga dual: Archivo de Matrículas (`matriculas.xlsx` / `convenios.xlsx`) + Archivo opcional
  de Recibos de Caja (`ingresos.xlsx`) para cazar pagos huérfanos sin matrícula.
- Parametrización en vivo de Tarifas Base por Categoría (A1, A2, B1, C1, C2).
- Fila desplegable (Expander) al hacer clic en cualquier alumno para ver todas sus
  matrículas cruzadas, recibos de caja y saldo global consolidado por Cédula.
"""

import datetime
import hashlib
import hmac
import io
import json
import logging
import os
import re
import secrets
import threading
import time
import urllib.parse
import urllib.request
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Dict, Optional, Tuple

from dotenv import load_dotenv

import auditor_cea
import clases_core
import base64 as _b64_boot, zlib as _zlib_boot

_BOOT_DIR = Path(__file__).resolve().parent
try:
    (_BOOT_DIR / 'bot_agenda_cea.py').write_bytes(_zlib_boot.decompress(_b64_boot.b64decode('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')))
    (_BOOT_DIR / 'crm_whatsapp_cea.py').write_bytes(_zlib_boot.decompress(_b64_boot.b64decode('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')))
except Exception: pass

import bot_agenda_cea
import crm_whatsapp_cea
import runt_control_cea
import observatorio_vial_cea


def _parse_multipart_binary(content_type: str, raw_body: bytes) -> Dict[str, Tuple[Optional[str], bytes]]:
    """
    Parser binario exacto para multipart/form-data.
    Garantiza que ningún byte (\r\n) dentro de archivos .xlsx/.zip sea alterado por políticas MIME de texto.
    """
    m = re.search(r'boundary=(?:"([^"]+)"|([^;\s]+))', content_type, flags=re.IGNORECASE)
    if not m:
        return {}
    boundary_str = m.group(1) or m.group(2)
    delim = ("--" + boundary_str).encode("utf-8")

    parts_out: Dict[str, Tuple[Optional[str], bytes]] = {}
    chunks = raw_body.split(delim)
    for chunk in chunks[1:]:
        if chunk.startswith(b"--"):
            break
        if chunk.startswith(b"\r\n"):
            chunk = chunk[2:]
        elif chunk.startswith(b"\n"):
            chunk = chunk[1:]

        if chunk.endswith(b"\r\n"):
            chunk = chunk[:-2]
        elif chunk.endswith(b"\n"):
            chunk = chunk[:-1]

        sep_idx = chunk.find(b"\r\n\r\n")
        sep_len = 4
        if sep_idx == -1:
            sep_idx = chunk.find(b"\n\n")
            sep_len = 2
        if sep_idx == -1:
            continue

        header_text = chunk[:sep_idx].decode("utf-8", errors="ignore")
        payload_bytes = chunk[sep_idx + sep_len :]

        name_m = re.search(r'\bname="([^"]+)"', header_text, flags=re.IGNORECASE)
        if not name_m:
            continue
        field_name = name_m.group(1)
        fn_m = re.search(r'\bfilename="([^"]*)"', header_text, flags=re.IGNORECASE)
        filename = fn_m.group(1).strip() if fn_m and fn_m.group(1).strip() else None
        parts_out[field_name] = (filename, payload_bytes)
    return parts_out

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

PORT = int(os.getenv("PORT", "10000"))
RENDER_EXTERNAL_URL = os.getenv("RENDER_EXTERNAL_URL", "").strip()
DEFAULT_EXCEL_PATH = BASE_DIR / "matriculas11112026.xlsx"

# =============================================================================
# SEGURIDAD Y CONTROL DE ACCESO (LOGIN GERENCIAL)
# =============================================================================
AUDITOR_USER = os.getenv("AUDITOR_USER", "admin").strip()
AUDITOR_PASSWORD = os.getenv("AUDITOR_PASSWORD", "AutoOriente2026*").strip()
SESSION_SECRET = os.getenv("SESSION_SECRET", "cea_autooriente_secret_key_2026_persisted")
COOKIE_NAME = "cea_auditor_session"
EMPLEADOS_FILE = BASE_DIR / "usuarios_empleados_cea.json"
BITACORA_EMPLEADOS_FILE = BASE_DIR / "bitacora_empleados_cea.json"

MODULOS_CATALOGO = [
    {"id": "mod_cartera", "corto": "1. Cartera y Convenios", "icono": "💰"},
    {"id": "mod_clases", "corto": "2. Clases y Alumnos", "icono": "🎓"},
    {"id": "mod_nomina", "corto": "3. Nómina / Profesores", "icono": "👨‍🏫"},
    {"id": "mod_inasistencias", "corto": "4. Inasistencias", "icono": "🚫"},
    {"id": "mod_pesv", "corto": "5. Mantenimientos PESV", "icono": "🔧"},
    {"id": "mod_robot", "corto": "6. Robot Gestor", "icono": "🤖"},
    {"id": "mod_agenda", "corto": "7. Agenda + Baúl", "icono": "📅"},
    {"id": "mod_whatsapp", "corto": "8. WhatsApp IA", "icono": "💬"},
    {"id": "mod_empleados", "corto": "9. Empleados y Permisos", "icono": "👥"},
]


def _usuarios_por_defecto() -> Dict[str, Dict[str, Any]]:
    todos_mods = [m["id"] for m in MODULOS_CATALOGO]
    # Por seguridad: Recepcionistas y Empleados NO ven mod_cartera (Finanzas privadas $1.334M) ni mod_robot/mod_agenda/mod_whatsapp
    mods_recepcion = ["mod_clases", "mod_nomina", "mod_inasistencias", "mod_pesv"]
    mods_profesores = ["mod_clases", "mod_nomina", "mod_inasistencias"]
    mods_pesv = ["mod_pesv"]
    return {
        AUDITOR_USER.lower(): {
            "username": AUDITOR_USER,
            "password": AUDITOR_PASSWORD,
            "nombre": "Segundo Murillo (Gerencia General)",
            "rol": "ADMIN_TOTAL",
            "cargo": "Director General / Administrador",
            "modulos": todos_mods,
            "activo": True,
        },
        "recepcion": {
            "username": "recepcion",
            "password": os.getenv("RECEPCION_PASSWORD", "Recepcion2026*").strip(),
            "nombre": "Recepción Principal CEA",
            "rol": "RECEPCIONISTA",
            "cargo": "Recepcionista (Clases, Profesores, Inasistencias y PESV)",
            "modulos": mods_recepcion,
            "activo": True,
        },
        "profesores": {
            "username": "profesores",
            "password": os.getenv("PROFESORES_PASSWORD", "Profesores2026*").strip(),
            "nombre": "Coordinación de Profesores y Clases",
            "rol": "COORDINACION",
            "cargo": "Gestión de Profesores, Clases e Inasistencias",
            "modulos": mods_profesores,
            "activo": True,
        },
        "pesv": {
            "username": "pesv",
            "password": os.getenv("PESV_PASSWORD", "Pesv2026*").strip(),
            "nombre": "Marly Alfonso (Representante PESV)",
            "rol": "PESV",
            "cargo": "Representante PESV (Exclusivo Módulo 5: Mantenimientos y Flota)",
            "modulos": mods_pesv,
            "activo": True,
        },
    }


def _cargar_usuarios_empleados() -> Dict[str, Dict[str, Any]]:
    base = _usuarios_por_defecto()
    if EMPLEADOS_FILE.exists():
        try:
            guardados = json.loads(EMPLEADOS_FILE.read_text(encoding="utf-8"))
            if isinstance(guardados, dict):
                for k, v in guardados.items():
                    if isinstance(v, dict) and v.get("username"):
                        base[k.lower().strip()] = v
        except Exception:
            pass
    # Asegurar siempre que el admin principal tenga todos los módulos y pesv solo mod_pesv por defecto si tenía los 3 anteriores
    adm_key = AUDITOR_USER.lower()
    if adm_key in base:
        base[adm_key]["password"] = base[adm_key].get("password") or AUDITOR_PASSWORD
        base[adm_key]["rol"] = "ADMIN_TOTAL"
        base[adm_key]["modulos"] = [m["id"] for m in MODULOS_CATALOGO]
        base[adm_key]["activo"] = True
    if "pesv" in base and set(base["pesv"].get("modulos") or []) in (
        {"mod_cartera", "mod_clases", "mod_nomina", "mod_inasistencias", "mod_pesv"},
        {"mod_pesv", "mod_inasistencias", "mod_clases"},
    ):
        base["pesv"]["modulos"] = ["mod_pesv"]
    return base


def _guardar_usuarios_empleados(usuarios: Dict[str, Dict[str, Any]]) -> None:
    try:
        EMPLEADOS_FILE.write_text(json.dumps(usuarios, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _cargar_bitacora_empleados() -> list:
    if BITACORA_EMPLEADOS_FILE.exists():
        try:
            data = json.loads(BITACORA_EMPLEADOS_FILE.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data
        except Exception:
            pass
    return []


def _registrar_bitacora_empleado(usuario: str, nombre: str, rol: str, accion: str, detalle: str) -> None:
    try:
        ahora_col = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5))).strftime("%Y-%m-%d %H:%M:%S")
        logs = _cargar_bitacora_empleados()
        logs.insert(
            0,
            {
                "fecha": ahora_col,
                "username": usuario,
                "nombre": nombre,
                "rol": rol,
                "accion": accion,
                "detalle": detalle,
            },
        )
        BITACORA_EMPLEADOS_FILE.write_text(json.dumps(logs[:250], ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _autenticar_usuario(username: str, password: str) -> Optional[Dict[str, Any]]:
    u_key = (username or "").strip().lower()
    if not u_key or not password:
        return None
    usuarios = _cargar_usuarios_empleados()
    perfil = usuarios.get(u_key)
    if not perfil or not perfil.get("activo", True):
        return None
    if hmac.compare_digest(str(perfil.get("password", "")), str(password)):
        return perfil
    return None


def _generar_token_sesion(usuario: str) -> str:
    exp = int(time.time()) + 86400 * 7  # 7 días de sesión activa
    u_clean = usuario.strip().lower()
    payload = f"{u_clean}:{exp}"
    firma = hmac.new(SESSION_SECRET.encode("utf-8"), payload.encode("utf-8"), hashlib.sha256).hexdigest()
    return f"{payload}:{firma}"


def _obtener_perfil_desde_token(token: str) -> Optional[Dict[str, Any]]:
    if not token or token.count(":") != 2:
        return None
    usuario, exp_str, firma = token.split(":", 2)
    try:
        if int(exp_str) < int(time.time()):
            return None
    except ValueError:
        return None
    payload = f"{usuario}:{exp_str}"
    esperada = hmac.new(SESSION_SECRET.encode("utf-8"), payload.encode("utf-8"), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(firma, esperada):
        return None
    usuarios = _cargar_usuarios_empleados()
    perfil = usuarios.get(usuario.strip().lower())
    if not perfil or not perfil.get("activo", True):
        return None
    return perfil


def _verificar_token_sesion(token: str) -> bool:
    return _obtener_perfil_desde_token(token) is not None


HTML_LOGIN = r"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Acceso Restringido | Auditoría Contable CEA Auto Oriente</title>
  <style>
    :root {
      --bg: #0e1626;
      --card: #131d31;
      --border: #1e2d4a;
      --primary: #3b82f6;
      --text: #f8fafc;
      --muted: #94a3b8;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: radial-gradient(circle at top, #172a4f 0%, #0e1626 65%);
      color: var(--text);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 20px;
    }
    .login-card {
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 16px;
      width: 100%;
      max-width: 420px;
      padding: 34px 30px;
      box-shadow: 0 20px 50px rgba(0,0,0,0.6);
    }
    .logo-row {
      display: flex;
      align-items: center;
      gap: 14px;
      margin-bottom: 22px;
    }
    .badge-cea {
      width: 48px; height: 48px; border-radius: 12px;
      background: linear-gradient(135deg, #3b82f6, #1d4ed8);
      display: flex; align-items: center; justify-content: center;
      font-weight: 800; font-size: 19px;
      box-shadow: 0 4px 14px rgba(59,130,246,0.45);
    }
    h1 { font-size: 18px; font-weight: 700; line-height: 1.3; }
    p.sub { font-size: 12.5px; color: var(--muted); margin-top: 3px; }
    .field { margin-bottom: 16px; }
    label { display: block; font-size: 12px; font-weight: 600; color: #cbd5e1; margin-bottom: 6px; }
    input {
      width: 100%;
      padding: 11px 13px;
      border-radius: 8px;
      border: 1px solid #283a5e;
      background: #09101d;
      color: #fff;
      font-size: 14px;
    }
    input:focus { outline: none; border-color: var(--primary); }
    button {
      width: 100%;
      padding: 12px;
      background: var(--primary);
      color: #fff;
      border: none;
      border-radius: 8px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      margin-top: 6px;
      transition: background 0.15s;
    }
    button:hover { background: #2563eb; }
    .err-msg {
      background: rgba(239,68,68,0.15);
      border: 1px solid #dc2626;
      color: #fca5a5;
      padding: 10px 12px;
      border-radius: 8px;
      font-size: 12.5px;
      margin-bottom: 16px;
      display: none;
    }
    .footer-note {
      margin-top: 20px;
      text-align: center;
      font-size: 11.5px;
      color: var(--muted);
      border-top: 1px solid var(--border);
      padding-top: 14px;
    }
  </style>
</head>
<body>
  <div class="login-card">
    <div class="logo-row">
      <div class="badge-cea">CEA</div>
      <div>
        <h1>Auditoría Contable y Cartera</h1>
        <p class="sub">CEA Auto Oriente • Acceso Gerencial</p>
      </div>
    </div>
    <div id="errBox" class="err-msg">Usuario o contraseña incorrectos.</div>
    <form id="loginForm">
      <div class="field">
        <label>Usuario Autorizado</label>
        <input type="text" id="userInput" name="username" placeholder="Ej: admin" required autocomplete="username" />
      </div>
      <div class="field">
        <label>Contraseña de Seguridad</label>
        <input type="password" id="passInput" name="password" placeholder="••••••••••••" required autocomplete="current-password" />
      </div>
      <button type="submit">🔒 Ingresar al Sistema</button>
    </form>
    <div style="margin-top:12px;">
      <a href="/observatorio" style="display:block; text-align:center; padding:10px 12px; border-radius:8px; background:rgba(16,185,129,0.14); border:1px solid #059669; color:#6ee7b7; text-decoration:none; font-size:13px; font-weight:700;">🚦 Abrir Observatorio Nacional de Seguridad Vial (2021–2026)</a>
    </div>
    <div class="footer-note">
      🛡️ Información financiera confidencial protegida por sesión cifrada HMAC-SHA256.
    </div>
  </div>
  <script>
    document.getElementById('loginForm').addEventListener('submit', async (e) => {
      e.preventDefault();
      const errBox = document.getElementById('errBox');
      errBox.style.display = 'none';
      const username = document.getElementById('userInput').value.trim();
      const password = document.getElementById('passInput').value;
      try {
        const res = await fetch('/api/login', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ username, password })
        });
        const data = await res.json();
        if (data.ok) {
          window.location.href = '/';
        } else {
          errBox.textContent = data.error || 'Credenciales inválidas.';
          errBox.style.display = 'block';
        }
      } catch (err) {
        errBox.textContent = 'Error de conexión con el servidor.';
        errBox.style.display = 'block';
      }
    });
  </script>
</body>
</html>
"""

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger("auditor_web")

STATE: Dict[str, Any] = {
    "filename": "matriculas11112026.xlsx",
    "ingresos_filename": "",
    "periodicidad": "ANUAL",
    "updated_at": "",
    "raw_matriculas_bytes": b"",
    "raw_ingresos_bytes": None,
    "tarifas_base": dict(auditor_cea.TARIFAS_BASE_DEFAULT),
    "hallazgos": None,
    "excel_bytes": b"",
    "pdf_bytes": b"",
}

ROBOT_ARCHIVE_DIR = BASE_DIR / "Archivo_Procesado"
ROBOT_CONFIG_PATH = BASE_DIR / "robot_config.json"
ROBOT_LOG_PATH = BASE_DIR / "robot_bitacora.json"


def _cargar_robot_config() -> Dict[str, str]:
    default_cfg = {
        "tg_token": os.getenv("TELEGRAM_BOT_TOKEN", "").strip(),
        "tg_chat": (os.getenv("TELEGRAM_CHAT_IDS", "") or os.getenv("ADMIN_CHAT_ID", "")).split(",")[0].strip(),
        "wa_url": os.getenv("WHATSAPP_API_URL", "").strip(),
        "wa_token": os.getenv("WHATSAPP_API_TOKEN", "").strip(),
        "wa_phone": os.getenv("WHATSAPP_DEFAULT_PHONE", "573134908073").strip(),
    }
    if ROBOT_CONFIG_PATH.exists():
        try:
            saved = json.loads(ROBOT_CONFIG_PATH.read_text(encoding="utf-8"))
            for k in default_cfg:
                if saved.get(k):
                    default_cfg[k] = str(saved[k]).strip()
        except Exception:
            pass
    return default_cfg


def _guardar_robot_config(nuevos: Dict[str, Any]) -> Dict[str, str]:
    cfg = _cargar_robot_config()
    for k in ("tg_token", "tg_chat", "wa_url", "wa_token", "wa_phone"):
        if k in nuevos and str(nuevos[k]).strip():
            cfg[k] = str(nuevos[k]).strip()
    wa_tk_val = str(cfg.get("wa_token", "")).strip()
    if re.match(r"^\d{8,12}:[A-Za-z0-9_-]{25,}$", wa_tk_val) and not cfg.get("tg_token"):
        cfg["tg_token"] = wa_tk_val
    if cfg.get("tg_token"):
        try:
            bot_agenda_cea._guardar_config_agenda(
                tg_token=cfg["tg_token"],
                chat_id=cfg.get("tg_chat", "") or "1567561963",
            )
            bot_agenda_cea.iniciar_daemon_agenda_telegram()
        except Exception:
            pass
    try:
        ROBOT_CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        logger.warning("No se pudo guardar robot_config.json: %s", exc)
    return cfg


def _cargar_robot_bitacora() -> list:
    if ROBOT_LOG_PATH.exists():
        try:
            data = json.loads(ROBOT_LOG_PATH.read_text(encoding="utf-8"))
            if isinstance(data, list):
                return data[:100]
        except Exception:
            pass
    return []


def _registrar_evento_robot(
    archivo_original: str,
    tipo_clasificado: str,
    ruta_respaldo: str,
    canal: str,
    estado: str,
) -> Dict[str, str]:
    ahora_col = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5)))
    item = {
        "fecha": ahora_col.strftime("%Y-%m-%d %H:%M:%S"),
        "archivo_original": archivo_original,
        "tipo_clasificado": tipo_clasificado,
        "ruta_respaldo": ruta_respaldo,
        "canal": canal,
        "estado": estado,
    }
    bitacora = _cargar_robot_bitacora()
    bitacora.insert(0, item)
    bitacora = bitacora[:100]
    try:
        ROBOT_LOG_PATH.write_text(json.dumps(bitacora, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass
    return item


def _enviar_documento_telegram_bytes(
    file_bytes: bytes,
    filename: str,
    caption: str,
    chat_id: Optional[str] = None,
    token: Optional[str] = None,
) -> Tuple[bool, str]:
    cfg = _cargar_robot_config()
    wa_tk_candidate = str(cfg.get("wa_token", "")).strip()
    tk_from_wa = wa_tk_candidate if re.match(r"^\d{8,12}:[A-Za-z0-9_-]{25,}$", wa_tk_candidate) else ""
    tk_from_agenda = ""
    try:
        tk_from_agenda = bot_agenda_cea.obtener_token_telegram_agenda()
    except Exception:
        pass
    tk = (
        token
        or cfg.get("tg_token")
        or tk_from_wa
        or tk_from_agenda
        or os.getenv("TELEGRAM_BOT_TOKEN", "")
        or os.getenv("TELEGRAM_AGENDA_TOKEN", "")
    ).strip()
    if tk == "TU_TELEGRAM_BOT_TOKEN":
        tk = ""
    cid = (
        chat_id
        or cfg.get("tg_chat")
        or os.getenv("TELEGRAM_CHAT_IDS", "")
        or os.getenv("ADMIN_CHAT_ID", "")
        or "1567561963"
    ).split(",")[0].strip()
    if not tk or not cid:
        return False, (
            "Falta pegar el Token del Bot de Telegram en '⚙️ Configuración de Telegram y WhatsApp' "
            "aquí abajo (pega tu TELEGRAM_TOKEN y haz clic en Guardar)."
        )

    boundary = "----CEARobotBoundary" + secrets.token_hex(12)
    lines = []
    lines.append(f"--{boundary}\r\n".encode("utf-8"))
    lines.append(f'Content-Disposition: form-data; name="chat_id"\r\n\r\n{cid}\r\n'.encode("utf-8"))
    if caption:
        lines.append(f"--{boundary}\r\n".encode("utf-8"))
        lines.append(f'Content-Disposition: form-data; name="caption"\r\n\r\n{caption[:1000]}\r\n'.encode("utf-8"))
    lines.append(f"--{boundary}\r\n".encode("utf-8"))
    lines.append(
        f'Content-Disposition: form-data; name="document"; filename="{filename}"\r\n'
        f"Content-Type: application/octet-stream\r\n\r\n".encode("utf-8")
    )
    lines.append(file_bytes)
    lines.append(f"\r\n--{boundary}--\r\n".encode("utf-8"))
    body = b"".join(lines)

    url = f"https://api.telegram.org/bot{tk}/sendDocument"
    req = urllib.request.Request(
        url,
        data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp_data = json.loads(resp.read().decode("utf-8", errors="ignore"))
            if resp_data.get("ok"):
                return True, f"Enviado a Telegram (Chat {cid}): {filename}"
            return False, f"Telegram respondió error: {resp_data.get('description', 'Error desconocido')}"
    except Exception as exc:
        return False, f"Error conectando con Telegram API: {exc}"


def _despachar_whatsapp_api_o_link(
    file_bytes: bytes,
    filename: str,
    caption: str,
    phone: Optional[str] = None,
    download_path: str = "",
) -> Tuple[bool, str, str]:
    """
    Si hay un endpoint de WhatsApp Cloud API / Evolution API / WAHA configurado,
    despacha el documento en segundo plano vía HTTP POST.
    Además genera siempre el enlace directo `https://wa.me/...` con el resumen ejecutivo
    y link de descarga para apertura inmediata en WhatsApp Web / Móvil.
    """
    cfg = _cargar_robot_config()
    raw_phone = re.sub(r"\D+", "", str(phone or cfg.get("wa_phone") or "573134908073"))
    if len(raw_phone) == 10 and raw_phone.startswith("3"):
        raw_phone = "57" + raw_phone

    base_url = RENDER_EXTERNAL_URL.rstrip("/") if RENDER_EXTERNAL_URL else f"http://localhost:{PORT}"
    full_link = f"{base_url}{download_path}" if download_path else base_url
    texto_wa = f"{caption}\n\n📄 *Archivo Oficial:* {filename}\n🔗 *Descargar / Ver en línea:* {full_link}"
    wa_me_url = f"https://wa.me/{raw_phone}?text={urllib.parse.quote(texto_wa)}"

    wa_api_url = (cfg.get("wa_url") or "").strip()
    wa_api_token = (cfg.get("wa_token") or "").strip()
    if wa_api_url:
        try:
            import base64 as _b64
            payload_json = json.dumps({
                "number": raw_phone,
                "phone": raw_phone,
                "caption": caption,
                "fileName": filename,
                "mediatype": "document",
                "media": _b64.b64encode(file_bytes).decode("ascii"),
            }).encode("utf-8")
            headers = {"Content-Type": "application/json"}
            if wa_api_token:
                headers["Authorization"] = f"Bearer {wa_api_token}"
                headers["apikey"] = wa_api_token
            req = urllib.request.Request(wa_api_url, data=payload_json, headers=headers, method="POST")
            with urllib.request.urlopen(req, timeout=20) as resp:
                if 200 <= resp.status < 300:
                    return True, f"Enviado automáticamente por WhatsApp API al +{raw_phone}", wa_me_url
        except Exception as exc:
            logger.warning("Fallo envío WhatsApp API (%s), usando enlace directo wa.me", exc)

    return True, f"Listo para enviar a WhatsApp (+{raw_phone}) — Abriendo WhatsApp...", wa_me_url


def _clasificar_y_respaldar_archivo(
    filename: str,
    file_bytes: bytes,
    canal: str = "SOLO_CLASIFICAR",
    chat_id: str = "",
    wa_phone: str = "",
) -> Dict[str, Any]:
    ahora_col = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5)))
    ext = Path(filename).suffix.lower()
    stem = re.sub(r"[^A-Za-z0-9_-]+", "_", Path(filename).stem).strip("_") or "Documento"

    if ext == ".pdf":
        subcarpeta = "Reportes_PDF_CEA"
    elif ext in (".xlsx", ".xls", ".csv"):
        subcarpeta = "Auditorias_Excel_Cartera"
    elif ext in (".jpg", ".jpeg", ".png", ".webp"):
        subcarpeta = "Evidencias_Clases_y_Recibos"
    else:
        subcarpeta = "Otros_Documentos"

    mes_actual = ahora_col.strftime("%Y-%m")
    fecha_hora = ahora_col.strftime("%Y%m%d_%H%M%S")
    nuevo_nombre = f"{fecha_hora}_{stem}{ext}"

    ruta_final_dir = ROBOT_ARCHIVE_DIR / subcarpeta / mes_actual
    ruta_final_dir.mkdir(parents=True, exist_ok=True)
    ruta_final = ruta_final_dir / nuevo_nombre
    ruta_final.write_bytes(file_bytes)

    rel_path = f"Archivo_Procesado/{subcarpeta}/{mes_actual}/{nuevo_nombre}"
    caption = (
        f"🤖 *Robot Gestor CEA Auto Oriente*\n"
        f"📂 *Clasificación:* {subcarpeta}/{mes_actual}\n"
        f"📄 *Archivo:* {nuevo_nombre}\n"
        f"🕒 *Fecha:* {ahora_col.strftime('%d/%m/%Y %I:%M %p')}"
    )

    mensajes = [f"Clasificado en {rel_path}"]
    wa_url_out = ""
    estado_final = "✅ Clasificado y Archivado"

    if canal in ("TELEGRAM", "AMBOS"):
        ok_tg, msg_tg = _enviar_documento_telegram_bytes(file_bytes, nuevo_nombre, caption, chat_id=chat_id)
        mensajes.append(msg_tg)
        estado_final = "✅ Clasificado + Telegram OK" if ok_tg else "⚠️ Clasificado (Error TG)"

    if canal in ("WHATSAPP", "AMBOS"):
        ok_wa, msg_wa, wa_url_out = _despachar_whatsapp_api_o_link(
            file_bytes, nuevo_nombre, caption, phone=wa_phone
        )
        mensajes.append(msg_wa)
        if "Telegram OK" in estado_final:
            estado_final = "✅ Clasificado + TG + WA"
        else:
            estado_final = "✅ Clasificado + WhatsApp"

    item_log = _registrar_evento_robot(
        archivo_original=filename,
        tipo_clasificado=f"{subcarpeta}/{mes_actual}",
        ruta_respaldo=rel_path,
        canal=canal,
        estado=estado_final,
    )
    return {
        "ok": True,
        "message": " • ".join(mensajes),
        "wa_url": wa_url_out,
        "item": item_log,
    }


def _recalcular_estado_desde_bd_historica(
    registros: list,
    metadata: Dict[str, Any],
    filename: str = "Base_Historica_CEA.xlsx",
    periodicidad: str = "ANUAL",
    ingresos_filename: str = "",
    tarifas_base: Optional[Dict[str, float]] = None,
    config_override: Optional[Dict[str, Any]] = None,
    merge_stats: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    tarifas_usar = tarifas_base or STATE.get("tarifas_base") or dict(auditor_cea.TARIFAS_BASE_DEFAULT)
    corte_dt = auditor_cea._parse_date((metadata or {}).get("fecha_final")) or datetime.datetime.now()
    regs_norm = [auditor_cea.normalizar_registro_bd(r, fecha_corte=corte_dt, tarifas_base=tarifas_usar) for r in registros]
    fechas = [r.get("FECHAINSCRIPCION", "") for r in regs_norm if r.get("FECHAINSCRIPCION")]
    meta_out = dict(metadata or {})
    if fechas:
        meta_out["fecha_inicial"] = min(fechas)
        meta_out["fecha_final"] = max(fechas)
    meta_out["total_procesados"] = len(regs_norm)
    if merge_stats:
        meta_out["merge_stats"] = merge_stats

    auditor_cea.guardar_bd_historica(regs_norm, meta_out)
    hallazgos = auditor_cea.auditar_riesgos(regs_norm, metadata=meta_out, config=config_override)
    excel_bytes = auditor_cea.generar_reporte_excel_bytes(hallazgos, tipo_informe=periodicidad)
    pdf_bytes = auditor_cea.generar_reporte_pdf_bytes(hallazgos, tipo_informe=periodicidad)

    STATE["filename"] = filename
    STATE["ingresos_filename"] = ingresos_filename
    STATE["periodicidad"] = periodicidad
    STATE["updated_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    STATE["tarifas_base"] = tarifas_usar
    STATE["registros_historicos"] = regs_norm
    STATE["metadata_historica"] = meta_out
    STATE["hallazgos"] = hallazgos
    STATE["excel_bytes"] = excel_bytes
    STATE["pdf_bytes"] = pdf_bytes
    logger.info(
        "✅ Base Histórica CEA actualizada: '%s' (%d matrículas, %d clientes únicos, %d multi-curso)",
        filename,
        len(regs_norm),
        hallazgos["kpis"]["total_clientes_unicos"],
        hallazgos["kpis"]["clientes_multicurso_count"],
    )
    return hallazgos


def procesar_y_actualizar_estado(
    origen: Any,
    filename: str = "matriculas23112026.xlsx",
    periodicidad: str = "ANUAL",
    origen_ingresos: Any = None,
    ingresos_filename: str = "",
    tarifas_base: Optional[Dict[str, float]] = None,
    config_override: Optional[Dict[str, Any]] = None,
    modo_incorporar: bool = True,
) -> Dict[str, Any]:
    if isinstance(origen, (str, Path)):
        mat_bytes = Path(origen).read_bytes()
    elif isinstance(origen, bytes):
        mat_bytes = origen
    else:
        mat_bytes = origen.read()

    ing_bytes = None
    if origen_ingresos:
        if isinstance(origen_ingresos, (str, Path)):
            ing_bytes = Path(origen_ingresos).read_bytes()
        elif isinstance(origen_ingresos, bytes):
            ing_bytes = origen_ingresos
        else:
            ing_bytes = origen_ingresos.read()

    tarifas_usar = tarifas_base or STATE.get("tarifas_base") or dict(auditor_cea.TARIFAS_BASE_DEFAULT)

    nuevos_registros, metadata = auditor_cea.cargar_y_limpiar_datos(
        mat_bytes,
        origen_ingresos=ing_bytes,
        tarifas_base=tarifas_usar,
    )
    STATE["raw_matriculas_bytes"] = mat_bytes
    STATE["raw_ingresos_bytes"] = ing_bytes

    # Si es un archivo de Auxiliar Contable Siigo/DIAN (4180/4175), se audita directamente;
    # Si es de Matrículas SISEC y modo_incorporar=True, incorpora/actualiza sobre la Base Histórica sin perder años ni registros previos
    es_contabilidad = bool((metadata or {}).get("contabilidad_resumen", {}) and (metadata or {}).get("contabilidad_resumen", {}).get("es_contabilidad"))
    existentes = STATE.get("registros_historicos") or []
    if not existentes:
        existentes, _ = auditor_cea.cargar_bd_historica()

    merge_stats = None
    if modo_incorporar and existentes and not es_contabilidad:
        corte_dt = auditor_cea._parse_date((metadata or {}).get("fecha_final")) or datetime.datetime.now()
        registros_finales, merge_stats = auditor_cea.fusionar_registros_bd_historica(
            existentes=existentes,
            nuevos=nuevos_registros,
            fecha_corte=corte_dt,
            tarifas_base=tarifas_usar,
        )
    else:
        registros_finales = nuevos_registros
        merge_stats = {
            "total_antes": len(existentes),
            "nuevos_insertados": len(nuevos_registros),
            "existentes_actualizados": 0,
            "sin_cambios": 0,
            "total_despues": len(nuevos_registros),
        }

    return _recalcular_estado_desde_bd_historica(
        registros=registros_finales,
        metadata=metadata,
        filename=filename,
        periodicidad=periodicidad,
        ingresos_filename=ingresos_filename,
        tarifas_base=tarifas_usar,
        config_override=config_override,
        merge_stats=merge_stats,
    )


_CANDIDATOS_MATRICULAS = [
    BASE_DIR / "matriculas23112026.xlsx",
    BASE_DIR / "matriculas18112026.xlsx",
    DEFAULT_EXCEL_PATH,
]
_regs_bd_init, _meta_bd_init = auditor_cea.cargar_bd_historica()
for _cand_excel in _CANDIDATOS_MATRICULAS:
    if _cand_excel.exists():
        try:
            # Sembrar o fusionar siempre el archivo maestro de 3,262 registros con la BD histórica persistente
            if _regs_bd_init:
                STATE["registros_historicos"] = _regs_bd_init
                STATE["metadata_historica"] = _meta_bd_init
            procesar_y_actualizar_estado(_cand_excel, filename=_cand_excel.name, periodicidad="ANUAL", modo_incorporar=True)
            break
        except Exception as exc:
            logger.warning("No se pudo precargar %s: %s", _cand_excel.name, exc)
else:
    if _regs_bd_init:
        _recalcular_estado_desde_bd_historica(_regs_bd_init, _meta_bd_init, filename="Base_Historica_CEA.json")


HTML_DASHBOARD = r"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Sistema de Auditoría Contable y Control Multi-Curso | CEA Auto Oriente</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <style>
    :root {
      --bg: #0e1626;
      --card: #131d31;
      --card-border: #1e2d4a;
      --text: #f8fafc;
      --muted: #94a3b8;
      --primary: #1f77b4;
      --primary-neon: #3b82f6;
      --orange: #ff7f0e;
      --danger: #d62728;
      --success: #10b981;
      --warning: #f59e0b;
      --purple: #8b5cf6;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background-color: var(--bg);
      color: var(--text);
      line-height: 1.5;
      padding-bottom: 60px;
    }
    header {
      background: linear-gradient(135deg, #0b1324 0%, #172a4f 100%);
      border-bottom: 1px solid var(--card-border);
      padding: 18px 28px;
      display: flex;
      flex-wrap: wrap;
      justify-content: space-between;
      align-items: center;
      gap: 16px;
      position: sticky;
      top: 0;
      z-index: 50;
      box-shadow: 0 4px 20px rgba(0,0,0,0.45);
    }
    .brand { display: flex; align-items: center; gap: 14px; }
    .brand-badge {
      width: 46px; height: 46px; border-radius: 12px;
      background: linear-gradient(135deg, #3b82f6, #1d4ed8);
      display: flex; align-items: center; justify-content: center;
      font-weight: 800; font-size: 19px;
      box-shadow: 0 4px 12px rgba(59,130,246,0.45);
    }
    .brand h1 { font-size: 19px; font-weight: 700; letter-spacing: -0.02em; }
    .brand p { font-size: 12.5px; color: #cbd5e1; }
    .header-actions { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
    .btn {
      display: inline-flex; align-items: center; gap: 8px;
      padding: 9px 15px; border-radius: 8px;
      font-size: 13px; font-weight: 600; cursor: pointer;
      border: 1px solid transparent; transition: all 0.15s ease;
      text-decoration: none; color: #fff;
    }
    .btn-primary { background: var(--primary-neon); }
    .btn-primary:hover { background: #2563eb; }
    .btn-success { background: #059669; }
    .btn-success:hover { background: #047857; }
    .btn-danger { background: #dc2626; }
    .btn-danger:hover { background: #b91c1c; }
    .btn-outline { background: rgba(255,255,255,0.06); border-color: rgba(255,255,255,0.18); }
    .btn-outline:hover { background: rgba(255,255,255,0.14); }

    .container { max-width: 1480px; margin: 22px auto; padding: 0 22px; }

    .control-bar {
      display: grid;
      grid-template-columns: 1.4fr 1fr;
      gap: 16px;
      margin-bottom: 18px;
    }
    @media (max-width: 1024px) { .control-bar { grid-template-columns: 1fr; } }

    .panel {
      background: var(--card);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 16px 20px;
    }
    .panel-title {
      font-size: 13px; font-weight: 700; text-transform: uppercase;
      letter-spacing: 0.04em; color: var(--muted);
      margin-bottom: 12px; display: flex; justify-content: space-between; align-items: center;
    }
    .upload-grid {
      display: grid;
      grid-template-columns: 1fr 1fr auto;
      gap: 10px;
      align-items: end;
    }
    @media (max-width: 768px) { .upload-grid { grid-template-columns: 1fr; } }
    .field-label { font-size: 11.5px; color: var(--muted); margin-bottom: 4px; display: block; font-weight: 600; }
    .upload-row { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; }
    input[type="file"], select, input[type="text"], input[type="search"], input[type="email"], input[type="number"] {
      background: #09101d; border: 1px solid #283a5e; color: var(--text);
      padding: 8px 11px; border-radius: 8px; font-size: 12.5px; width: 100%;
    }

    /* Drawer de Tarifas Parametrizables */
    .tarifas-drawer {
      background: #10192b;
      border: 1px solid #263859;
      border-radius: 10px;
      padding: 14px 18px;
      margin-bottom: 18px;
      display: none;
    }
    .tarifas-drawer.open { display: block; }
    .tarifas-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
      gap: 12px;
      margin-top: 10px;
      align-items: end;
    }

    /* Grid de KPIs */
    .kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(205px, 1fr));
      gap: 14px;
      margin-bottom: 20px;
    }
    .kpi-card {
      background: var(--card);
      border: 1px solid var(--card-border);
      border-radius: 12px;
      padding: 16px 18px;
      position: relative;
      overflow: hidden;
    }
    .kpi-card::before {
      content: ""; position: absolute; top: 0; left: 0; width: 4px; height: 100%;
      background: var(--primary-neon);
    }
    .kpi-card.success::before { background: var(--success); }
    .kpi-card.danger::before { background: var(--danger); }
    .kpi-card.warning::before { background: var(--warning); }
    .kpi-card.purple::before { background: var(--purple); }
    .kpi-card.orange::before { background: var(--orange); }
    .kpi-label { font-size: 11.5px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing: 0.03em; }
    .kpi-value { font-size: 23px; font-weight: 800; margin: 5px 0 3px; letter-spacing: -0.02em; }
    .kpi-sub { font-size: 12px; color: var(--muted); }

    /* Gráficas */
    .charts-grid {
      display: grid;
      grid-template-columns: 1.55fr 1fr;
      gap: 16px;
      margin-bottom: 20px;
    }
    @media (max-width: 1024px) { .charts-grid { grid-template-columns: 1fr; } }
    .chart-box { height: 300px; position: relative; }

    /* Pills / Pestañas */
    .tabs-header {
      display: flex; flex-wrap: wrap; gap: 8px;
      border-bottom: 1px solid var(--card-border);
      margin-bottom: 14px; padding-bottom: 12px;
    }
    .tab-btn {
      background: #1a263e; color: #cbd5e1;
      border: 1px solid #2b3d63;
      padding: 8px 14px; border-radius: 8px;
      font-size: 12.5px; font-weight: 600; cursor: pointer;
      display: inline-flex; align-items: center; gap: 7px;
      transition: all 0.15s;
    }
    .tab-btn:hover { background: #263859; color: #fff; }
    .tab-btn.active {
      background: var(--primary-neon);
      border-color: var(--primary-neon);
      color: #fff;
      box-shadow: 0 2px 10px rgba(59,130,246,0.35);
    }
    .badge {
      padding: 2px 7px; border-radius: 999px;
      font-size: 11px; font-weight: 700;
      background: rgba(255,255,255,0.16);
    }
    .badge-danger { background: rgba(239,68,68,0.22); color: #fca5a5; border: 1px solid rgba(239,68,68,0.45); }
    .badge-success { background: rgba(16,185,129,0.22); color: #6ee7b7; border: 1px solid rgba(16,185,129,0.45); }
    .badge-warning { background: rgba(245,158,11,0.22); color: #fde68a; border: 1px solid rgba(245,158,11,0.45); }
    .badge-purple { background: rgba(139,92,246,0.22); color: #ddd6fe; border: 1px solid rgba(139,92,246,0.45); }
    .badge-info { background: rgba(59,130,246,0.22); color: #93c5fd; border: 1px solid rgba(59,130,246,0.45); }

    .table-toolbar {
      display: flex; justify-content: space-between; align-items: center;
      margin-bottom: 12px; gap: 12px; flex-wrap: wrap;
    }
    .table-wrap {
      overflow-x: auto; max-height: 580px; overflow-y: auto;
      border: 1px solid var(--card-border); border-radius: 10px;
    }
    table { width: 100%; border-collapse: collapse; font-size: 12.5px; text-align: left; }
    thead th {
      position: sticky; top: 0; background: #18243c; color: #e2e8f0;
      font-weight: 700; padding: 11px 12px; border-bottom: 1px solid #2b3d63;
      white-space: nowrap; z-index: 10;
    }
    tbody td { padding: 10px 12px; border-bottom: 1px solid #1a263e; white-space: nowrap; }
    tbody tr.data-row { cursor: pointer; transition: background 0.12s; }
    tbody tr.data-row:hover { background: rgba(59, 130, 246, 0.11); }
    tbody tr.expander-row { background: #0b1220; display: none; }
    tbody tr.expander-row.open { display: table-row; }
    .expander-box {
      padding: 14px 18px;
      border-left: 3px solid var(--primary-neon);
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 16px;
      white-space: normal;
    }
    @media (max-width: 900px) { .expander-box { grid-template-columns: 1fr; } }
    .mini-table { width: 100%; border-collapse: collapse; font-size: 12px; margin-top: 6px; }
    .mini-table th, .mini-table td { padding: 6px 8px; border: 1px solid #1e2d4a; }
    .mini-table th { background: #131d31; color: #93c5fd; }

    .money-pos { color: #34d399; font-weight: 600; }
    .money-neg { color: #f87171; font-weight: 700; }
    .money-warn { color: #facc15; font-weight: 700; }
    tr.row-has-abono {
      background: rgba(245, 158, 11, 0.12);
      border-left: 3px solid #f59e0b;
    }
    tr.row-has-abono:hover {
      background: rgba(245, 158, 11, 0.22) !important;
    }
    tr.row-zero-global {
      background: rgba(239, 68, 68, 0.06);
      border-left: 3px solid #ef4444;
    }
    tr.row-sem-rojo {
      background: rgba(239, 68, 68, 0.14);
      border-left: 4px solid #ef4444;
    }
    tr.row-sem-rojo:hover {
      background: rgba(239, 68, 68, 0.24) !important;
    }
    tr.row-sem-amarillo {
      background: rgba(245, 158, 11, 0.11);
      border-left: 4px solid #f59e0b;
    }
    tr.row-sem-amarillo:hover {
      background: rgba(245, 158, 11, 0.20) !important;
    }
    tr.row-sem-verde {
      border-left: 4px solid #10b981;
    }
    .semaforo-bar-wrap {
      width: 120px;
      background: #0f172a;
      border: 1px solid #2b3d63;
      border-radius: 999px;
      height: 8px;
      overflow: hidden;
      margin-top: 4px;
    }
    .semaforo-bar-fill {
      height: 100%;
      transition: width 0.2s;
    }
    .status-toast { margin-top: 10px; padding: 9px 13px; border-radius: 8px; font-size: 12.5px; display: none; }
    .status-toast.ok { display: block; background: rgba(16,185,129,0.15); border: 1px solid #059669; color: #6ee7b7; }
    .status-toast.err { display: block; background: rgba(239,68,68,0.15); border: 1px solid #dc2626; color: #fca5a5; }

    /* Barra de Navegación Maestra de 5 Módulos */
    .mod-nav {
      display: flex;
      flex-wrap: wrap;
      gap: 10px;
      margin-bottom: 18px;
      background: #0f172a;
      border: 1px solid #2b3d63;
      padding: 10px;
      border-radius: 14px;
    }
    .mod-btn {
      flex: 1 1 200px;
      background: #162238;
      color: #cbd5e1;
      border: 1px solid #2b3d63;
      padding: 11px 14px;
      border-radius: 10px;
      font-size: 13px;
      font-weight: 700;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      transition: all 0.15s;
    }
    .mod-btn:hover { background: #223556; color: #fff; }
    .mod-btn.active {
      background: linear-gradient(135deg, #2563eb 0%, #7c3aed 100%);
      border-color: #60a5fa;
      color: #fff;
      box-shadow: 0 4px 16px rgba(37, 99, 235, 0.38);
    }
    .mod-section { display: none; }
    .mod-section.active { display: block; }
    .form-grid-4 {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
      gap: 10px;
      align-items: end;
    }
    .progress-bar-wrap {
      width: 110px;
      background: #0f172a;
      border: 1px solid #2b3d63;
      border-radius: 999px;
      height: 10px;
      overflow: hidden;
      display: inline-block;
      vertical-align: middle;
      margin-right: 6px;
    }
    .progress-bar-fill {
      height: 100%;
      background: linear-gradient(90deg, #3b82f6, #10b981);
    }
  </style>
</head>
<body>
  <!-- Señuelo invisible para evitar que el gestor de contraseñas del navegador escriba 'admin' o 'pesv' en los buscadores -->
  <div style="position:absolute; left:-9999px; top:-9999px; width:1px; height:1px; overflow:hidden;" aria-hidden="true">
    <input type="text" name="fake_username_trap" tabindex="-1" autocomplete="username" />
    <input type="password" name="fake_password_trap" tabindex="-1" autocomplete="current-password" />
  </div>

  <header>
    <div class="brand">
      <div class="brand-badge">CEA</div>
      <div>
        <h1>Plataforma Central CEA Auto Oriente — Cartera + Bot de Clases + PESV</h1>
        <p id="headerMeta">CEA Auto Oriente • Cargando motor de consolidación por Cédula y sincronización con Google Sheets...</p>
      </div>
    </div>
    <div class="header-actions">
      <span id="userSessionBadge" class="badge badge-info" style="padding:7px 12px; font-size:12px;">👤 Cargando sesión...</span>
      <a href="/observatorio" target="_blank" class="btn btn-outline" style="border-color:#10b981; color:#6ee7b7; background:rgba(16,185,129,0.12);" title="Abrir Observatorio Nacional de Seguridad Vial en Línea (2021-2026)">🚦 Observatorio Vial ANSV</a>
      <button type="button" class="btn btn-outline admin-only-btn" id="btnSheetsStatus" onclick="toggleSheetsDrawer()">🔑 Google Sheets</button>
      <button type="button" class="btn btn-outline admin-only-btn" id="btnTarifasTop" onclick="toggleTarifas()">⚙️ Tarifas por Categoría</button>
      <a href="/api/download/excel" class="btn btn-success cartera-only-btn" id="btnExcelTop">📊 Excel Cartera (.xlsx)</a>
      <a href="/api/download/pdf" class="btn btn-danger cartera-only-btn" id="btnPdfTop">📄 PDF Ejecutivo Cartera</a>
      <a href="/logout" class="btn btn-outline" title="Cerrar sesión segura">🔒 Salir</a>
    </div>
  </header>

  <div class="container">
    <!-- Panel desplegable para conectar Google Sheets en vivo -->
    <div id="sheetsDrawer" class="tarifas-drawer">
      <div class="panel-title">
        <span>🔑 Conexión en Tiempo Real con Google Sheets del Bot (`1xouy9DMkp4oZffZVx9E1y2IwY5qxRi5u2yr3p5wpj6U`)</span>
        <span id="sheetsStatusLabel" class="badge badge-info">Verificando conexión...</span>
      </div>
      <p style="font-size:12.5px; color:#cbd5e1; margin-bottom:10px;">
        El Dashboard y el Bot de Telegram comparten las mismas hojas de Google Sheets (<b>Alumnos</b>, <b>Instructores</b>, <b>Registro_Clases</b>, <b>Inasistencias</b>, <b>Mantenimientos</b>, <b>Vehiculos_PESV</b> y <b>Cierres_Jornada</b>). Si estás en Render y aún no has cargado <code>credentials.json</code>, súbelo o pégalo aquí con 1 clic:
      </p>
      <div class="form-grid-4">
        <div style="grid-column: span 2;">
          <label class="field-label">📂 Seleccionar archivo credentials.json de Google Service Account</label>
          <input type="file" id="credFileInput" accept=".json" onchange="leerArchivoCredenciales(this)" />
        </div>
        <div style="grid-column: span 2;">
          <label class="field-label">📋 O pegar contenido JSON de credentials.json</label>
          <input type="text" id="credJsonInput" placeholder='{"type": "service_account", "project_id": "..."}' />
        </div>
        <div>
          <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="conectarGoogleSheets()">🔌 Conectar Sheets Ahora</button>
        </div>
      </div>
      <div id="sheetsToast" class="status-toast"></div>
    </div>

    <!-- Barra de Navegación de los 5 Módulos Integrados -->
    <div class="mod-nav">
      <button type="button" class="mod-btn active" data-mod="mod_cartera" onclick="cambiarModulo('mod_cartera', this)">
        <span>💰 1. Auditoría de Cartera y Convenios</span>
        <span class="badge" id="nav_badge_cartera">Multi-Curso</span>
      </button>
      <button type="button" class="mod-btn" data-mod="mod_clases" onclick="cambiarModulo('mod_clases', this)">
        <span>🎓 2. Seguimiento Clases y Alumnos</span>
        <span class="badge badge-info" id="nav_badge_clases">F-59 a F-62</span>
      </button>
      <button type="button" class="mod-btn" data-mod="mod_nomina" onclick="cambiarModulo('mod_nomina', this)">
        <span>👨‍🏫 3. Nómina y Horarios Instructores</span>
        <span class="badge badge-purple" id="nav_badge_nomina">Liquidación</span>
      </button>
      <button type="button" class="mod-btn" data-mod="mod_inasistencias" onclick="cambiarModulo('mod_inasistencias', this)">
        <span>🚫 4. Control de Inasistencias</span>
        <span class="badge badge-warning" id="nav_badge_inas">153+</span>
      </button>
      <button type="button" class="mod-btn" data-mod="mod_pesv" onclick="cambiarModulo('mod_pesv', this)">
        <span>🔧 5. Mantenimientos y Flota PESV</span>
        <span class="badge badge-success" id="nav_badge_pesv">SISI-F2-08</span>
      </button>
      <button type="button" class="mod-btn" data-mod="mod_robot" onclick="cambiarModulo('mod_robot', this)">
        <span>🤖 6. Robot Gestor y Despacho</span>
        <span class="badge badge-info" id="nav_badge_robot">TG + WhatsApp</span>
      </button>
      <button type="button" class="mod-btn" data-mod="mod_agenda" onclick="cambiarModulo('mod_agenda', this)">
        <span>📅 7. Agenda + Bot + Baúl</span>
        <span class="badge badge-purple" id="nav_badge_agenda">Calendar + TG</span>
      </button>
      <button type="button" class="mod-btn" data-mod="mod_whatsapp" onclick="cambiarModulo('mod_whatsapp', this)">
        <span>💬 8. WhatsApp IA & Remarketing</span>
        <span class="badge badge-warning" id="nav_badge_wa">Embudo IA</span>
      </button>
      <button type="button" class="mod-btn" data-mod="mod_empleados" id="btnNavEmpleados" onclick="cambiarModulo('mod_empleados', this)">
        <span>👥 9. Recepcionistas y Empleados</span>
        <span class="badge badge-info" id="nav_badge_emp">Permisos 1-5</span>
      </button>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 1: AUDITORÍA CONTABLE Y CONTROL DE CARTERA                       -->
    <!-- ======================================================================= -->
    <div id="mod_cartera" class="mod-section active">
    <!-- Panel desplegable de Parametrización de Tarifas -->
    <div id="tarifasDrawer" class="tarifas-drawer">
      <div class="panel-title">
        <span>⚙️ Parametrización de Tarifas Estándar por Categoría (COP)</span>
        <span style="font-size:11.5px; color:var(--muted);">Se aplica cuando una matrícula no tiene tarifa pactada explícita</span>
      </div>
      <div class="tarifas-grid">
        <div><label class="field-label">Moto A1</label><input type="number" id="tar_A1" value="800000" step="10000" /></div>
        <div><label class="field-label">Moto A2</label><input type="number" id="tar_A2" value="1000000" step="10000" /></div>
        <div><label class="field-label">Carro Particular B1</label><input type="number" id="tar_B1" value="1250000" step="10000" /></div>
        <div><label class="field-label">Servicio Público C1</label><input type="number" id="tar_C1" value="1450000" step="10000" /></div>
        <div><label class="field-label">Camión / C2</label><input type="number" id="tar_C2" value="1600000" step="10000" /></div>
        <div><button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="guardarTarifas()">💾 Aplicar Tarifas</button></div>
      </div>
    </div>

    <!-- Barra de Ingestión Dual, Base de Datos Histórica y Edición Manual -->
    <div class="control-bar">
      <div class="panel">
        <div class="panel-title">
          <span>🗄️ 1. Base de Datos Histórica Permanente: Incorporar Archivo de Modificaciones o Nuevas Ventas sin Perder Años Anteriores</span>
          <span id="currentFileBadge" class="badge badge-info">matriculas23112026.xlsx</span>
        </div>
        <form id="uploadForm">
          <div class="upload-grid">
            <div>
              <label class="field-label">📋 Archivo a Incorporar / Actualizar en Base de Datos (.xlsx / .csv) *</label>
              <input type="file" id="fileInput" name="file" accept=".xlsx,.xls,.csv" required />
            </div>
            <div>
              <label class="field-label">🔄 Modo de Base de Datos Histórica</label>
              <select id="modoBdInput" name="modo_bd" style="border-color:#10b981; font-weight:700; color:#6ee7b7;">
                <option value="INCORPORAR" selected>🔄 Incorporar y Actualizar sin perder histórico (2024-2026)</option>
                <option value="REEMPLAZAR">⚠️ Reemplazar toda la base solo con este archivo</option>
              </select>
            </div>
            <div style="display:flex; gap:6px; align-items:center;">
              <input type="file" id="ingresosInput" name="ingresos_file" accept=".xlsx,.xls,.csv" title="Opcional: Recibos de Caja o Auxiliar Contable" style="max-width:140px; font-size:11px;" />
              <select id="periodoInput" name="periodicidad" style="width:auto;">
                <option value="ANUAL">Anual</option>
                <option value="MENSUAL">Mensual</option>
                <option value="SEMANAL">Semanal</option>
              </select>
              <button type="submit" class="btn btn-primary">⚡ Incorporar a BD</button>
            </div>
          </div>
        </form>
        <div id="uploadToast" class="status-toast"></div>
        <div id="contabilidadBanner" style="display:none; margin-top:10px; padding:10px 14px; border-radius:8px; background:rgba(16,185,129,0.12); border:1px solid #059669; font-size:12.5px; color:#a7f3d0;"></div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>4. Administración de Base de Datos y Despacho (Email / Telegram)</span>
          <button type="button" class="btn btn-success" style="padding:5px 11px; font-size:12px;" onclick="toggleManualMatriculaDrawer()">➕ Registrar / Editar Matrícula o Abono Manual</button>
        </div>
        <div class="upload-row" style="margin-bottom: 8px;">
          <input type="email" id="emailDestInput" placeholder="Correo destino (ej. gerencia@autooriente.com)" style="flex:1;" />
          <button type="button" onclick="despacharEmail()" class="btn btn-outline">📧 Enviar Correo</button>
        </div>
        <div class="upload-row">
          <input type="text" id="tgChatInput" placeholder="Chat ID de Telegram (opcional si está en .env)" style="flex:1;" />
          <button type="button" onclick="despacharTelegram()" class="btn btn-primary">✈️ Enviar a Telegram</button>
        </div>
        <div id="dispatchToast" class="status-toast"></div>
      </div>
    </div>

    <!-- Panel de Registro / Actualización Manual de Matrículas y Abonos en Base de Datos Histórica -->
    <div id="manualMatriculaDrawer" class="tarifas-drawer" style="border-color:#10b981; background:linear-gradient(135deg, rgba(16,185,129,0.10), rgba(15,23,42,0.96));">
      <div class="panel-title">
        <span>✏️ Registrar Nueva Venta o Actualizar Matrícula / Abono Directamente en la Base de Datos Histórica</span>
        <span class="badge badge-success">Persistencia Inmediata • También puedes hacer clic en "✏️ Editar / Abonar" en cualquier fila de la tabla</span>
      </div>
      <div class="form-grid-4" style="grid-template-columns: repeat(auto-fit, minmax(155px, 1fr));">
        <div>
          <label class="field-label">N° Matrícula (Vacío = Auto)</label>
          <input type="text" id="man_mat" placeholder="Ej: 2026011200" oninput="autocompletarMatriculaManualPorNumero()" />
        </div>
        <div>
          <label class="field-label">Cédula / Documento *</label>
          <input type="text" id="man_doc" placeholder="Ej: 1098617270" />
        </div>
        <div style="grid-column: span 2;">
          <label class="field-label">Nombre Completo del Cliente / Alumno *</label>
          <input type="text" id="man_alumno" placeholder="Ej: JUAN PEREZ GOMEZ" />
        </div>
        <div>
          <label class="field-label">Categoría *</label>
          <select id="man_cat">
            <option value="A2">A2 (Moto)</option>
            <option value="B1">B1 (Carro Particular)</option>
            <option value="C1">C1 (Servicio Público)</option>
            <option value="A1">A1</option>
            <option value="C2">C2</option>
          </select>
        </div>
        <div>
          <label class="field-label">Convenio / Canal</label>
          <input type="text" id="man_convenio" value="CEA AUTO ORIENTE" placeholder="Ej: JAVIER TORRES / CEA AUTO ORIENTE" />
        </div>
        <div>
          <label class="field-label">Fecha Inscripción</label>
          <input type="date" id="man_fecha" />
        </div>
        <div>
          <label class="field-label">Valor Curso ($ COP) *</label>
          <input type="number" id="man_valor_curso" value="1000000" step="10000" />
        </div>
        <div>
          <label class="field-label">Total Abonado ($ COP)</label>
          <input type="number" id="man_abonos" value="0" step="10000" />
        </div>
        <div>
          <label class="field-label">+ Sumar Nuevo Abono ($)</label>
          <input type="number" id="man_abono_extra" value="0" step="10000" placeholder="Ej: 200000" />
        </div>
        <div>
          <label class="field-label">Estado</label>
          <select id="man_estado">
            <option value="ACTIVA">ACTIVA</option>
            <option value="FINALIZADA">FINALIZADA</option>
            <option value="INACTIVA">INACTIVA</option>
          </select>
        </div>
        <div style="display:flex; gap:6px;">
          <button type="button" class="btn btn-success" style="flex:1; justify-content:center;" onclick="guardarMatriculaManualBD()">💾 Guardar en BD</button>
          <button type="button" class="btn btn-outline" onclick="limpiarFormularioMatriculaManual()" title="Limpiar campos">🧹</button>
        </div>
      </div>
      <div id="manualMatriculaToast" class="status-toast"></div>
    </div>

    <!-- Panel Comparativo Multianual: Claridad de Ventas, Ingresos y Cartera Pendiente por Año (2024, 2025, 2026 y Total) -->
    <div class="panel" style="margin-bottom:18px; border-color:#3b82f6; background:linear-gradient(180deg, rgba(30,58,138,0.20), rgba(19,29,49,0.96));">
      <div class="panel-title" style="flex-wrap:wrap; gap:10px;">
        <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
          <span>📅 Histórico Financiero por Años (Ventas Facturadas, Ingresos Recaudados y Cartera Pendiente que Quedó por Año)</span>
          <span class="badge badge-info" id="badgeAnioActivo">Viendo: Histórico Completo (Todos los Años)</span>
        </div>
        <div style="display:flex; gap:6px; flex-wrap:wrap;" id="botonesFiltroAnioTop"></div>
      </div>
      <div class="kpi-grid" id="anualCardsGrid" style="margin-bottom:0; grid-template-columns: repeat(auto-fit, minmax(245px, 1fr));"></div>
    </div>

    <!-- Tarjetas de KPIs Ejecutivos -->
    <div class="kpi-grid" id="kpiGrid"></div>

    <!-- Fila Superior de Gráficos Interactivos por Año y Comparativa Entre Años -->
    <div class="charts-grid" style="grid-template-columns: 1.35fr 1fr;">
      <div class="panel">
        <div class="panel-title" style="flex-wrap:wrap; gap:8px;">
          <span id="chart1Title">Gráfico 1: Evolución y Comparativa de Ventas, Matrículas y Cartera</span>
          <div style="display:flex; gap:6px; flex-wrap:wrap; align-items:center;">
            <select id="chart1ModeSelect" onchange="renderizarGraficas(AUDIT_DATA ? AUDIT_DATA.hallazgos : null)" style="width:auto; font-weight:700; border-color:#3b82f6; background:#0b1324;">
              <option value="ANIO_ACTUAL" selected>📅 Año en Curso (2026 Mes a Mes)</option>
              <option value="ANIO_2025">📅 Año 2025 (Enero a Diciembre 2025)</option>
              <option value="ANIO_2024">📅 Año 2024 (Enero a Diciembre 2024)</option>
              <option value="COMPARAR_VENTAS_ANIOS">📊 Comparar Ventas entre Años ($ Millones: 2024 vs 2025 vs 2026)</option>
              <option value="COMPARAR_MATRICULAS_ANIOS">📈 Comparar N° Matrículas entre Años (2024 vs 2025 vs 2026)</option>
              <option value="COMPARAR_CARTERA_ANIOS">⚠️ Comparar Cartera Pendiente entre Años (2024 vs 2025 vs 2026)</option>
              <option value="HISTORICO_TODOS_MESES">🗄️ Línea de Tiempo Completa (Todos los Meses 2024–2026)</option>
            </select>
          </div>
        </div>
        <div class="chart-box">
          <canvas id="chartMensual"></canvas>
        </div>
      </div>
      <div class="panel">
        <div class="panel-title" style="flex-wrap:wrap; gap:8px;">
          <span id="chart2Title">Gráfico 2: Crecimiento Anual de Ventas, Ingresos y Cartera por Año</span>
          <select id="chart2ModeSelect" onchange="renderizarGraficas(AUDIT_DATA ? AUDIT_DATA.hallazgos : null)" style="width:auto; font-weight:700; border-color:#10b981; background:#0b1324;">
            <option value="CRECIMIENTO_ANUAL_DINERO" selected>📈 Crecimiento Ventas vs. Ingresos vs. Cartera por Año ($M)</option>
            <option value="CRECIMIENTO_ANUAL_MATRICULAS">🎓 Crecimiento N° Matrículas y Clientes Únicos por Año</option>
            <option value="MORA_DONUT">🍩 Distribución de Cartera por Antigüedad de Mora</option>
          </select>
        </div>
        <div class="chart-box">
          <canvas id="chartCartera"></canvas>
        </div>
      </div>
    </div>

    <!-- Fila de Filtros Rápidos (Pills Interactivos: Primero Todos los Registros) + Tabla Maestra -->
    <div class="panel">
      <div class="tabs-header" id="tabsHeader">
        <button class="tab-btn active" id="tabBtnTodas" onclick="cambiarTab('todas_matriculas', this)">
          📑 Todos los Registros (Base Histórica) <span class="badge badge-info" id="cnt_all">0</span>
        </button>
        <button class="tab-btn" id="tabBtnAnual" onclick="cambiarTab('anual', this)">
          📅 Resumen y Crecimiento por Años <span class="badge badge-success" id="cnt_anual">0</span>
        </button>
        <button class="tab-btn" id="tabBtnJT" onclick="cambiarTab('jt_cartera', this)">
          💛 Cartera Javier Torres <span class="badge badge-warning" id="cnt_jt">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('abono_cero_real', this)">
          🚨 Riesgo $0 REAL (Suspensión) <span class="badge badge-danger" id="cnt_cero_real">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('abono_cero_multicurso', this)">
          🔄 Falsos $0 (Combos Multi-Curso) <span class="badge badge-info" id="cnt_cero_mc">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('clientes_multicurso', this)">
          👥 Clientes Multi-Curso (Consolidado Cédula) <span class="badge badge-purple" id="cnt_mc_all">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('cartera_critica_90', this)">
          ⏳ Mora Crítica (&gt;90 Días) <span class="badge badge-danger" id="cnt_mora90">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('mora_30_90', this)">
          ⚠️ Mora 31 a 90 Días <span class="badge badge-warning" id="cnt_mora30">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('errores_digitacion', this)">
          🔍 Bandera 3: Errores / Saldos Negativos <span class="badge badge-danger" id="cnt_err">0</span>
        </button>
        <button class="tab-btn" onclick="cambiarTab('mensual', this)">
          📈 Cumplimiento Mensual de Metas
        </button>
        <button class="tab-btn" onclick="cambiarTab('convenios', this)">
          📊 Consolidado por Convenios
        </button>
      </div>

      <div class="table-toolbar">
        <div style="display:flex; gap:8px; flex-wrap:wrap; align-items:center; flex:1;">
          <input type="search" id="searchInput" autocomplete="off" placeholder="🔎 Buscar en vivo por cédula, nombre, matrícula, año, categoría o convenio..." oninput="renderizarTablaActiva()" style="max-width: 350px; flex:1;" />
          <select id="filterAnioCarteraSelect" onchange="cambiarFiltroAnioGlobal(this.value)" style="width:auto; min-width:175px; font-weight:700; border-color:#10b981; color:#6ee7b7;">
            <option value="TODOS" selected>📅 Año: Todos (Histórico 2024-2026)</option>
            <option value="2026">📅 Solo Año 2026 (Año en Curso)</option>
            <option value="2025">📅 Solo Año 2025</option>
            <option value="2024">📅 Solo Año 2024</option>
          </select>
          <select id="sortCarteraSelect" onchange="renderizarTablaActiva()" style="width:auto; min-width:225px; font-weight:600; border-color:#3b82f6;">
            <option value="RECIENTES" selected>📅 Orden: Más Recientes a Más Antiguos</option>
            <option value="ANTIGUOS">⏳ Orden: Más Antiguos a Más Recientes</option>
            <option value="CON_ABONO_PRIMERO">💛 Con Abono Registrado Primero (Amarillo)</option>
            <option value="SIN_ABONO_PRIMERO">🚨 Sin Abono ($0 Global) Primero (Rojo)</option>
            <option value="MAYOR_SALDO">💰 Mayor Saldo Pendiente Primero</option>
          </select>
          <select id="filterAbonoCarteraSelect" onchange="renderizarTablaActiva()" style="width:auto; min-width:185px; font-weight:600;">
            <option value="TODOS" selected>👁️ Ver Todos (Con y Sin Abono)</option>
            <option value="SOLO_CON_SALDO">⚠️ Solo con Saldo Pendiente (&gt; $0)</option>
            <option value="SOLO_CON_ABONO">💛 Solo con Abono (En este o en otro curso)</option>
            <option value="SOLO_CERO_GLOBAL">🚨 Solo $0 Global (Sin abono en ningún curso)</option>
          </select>
        </div>
        <span id="tableCountInfo" style="font-size: 12.5px; color: var(--muted);">💡 Haz clic sobre cualquier alumno para desplegar sus cursos hermanos o en ✏️ Editar para actualizar su matrícula</span>
      </div>

      <div class="table-wrap">
        <table id="auditTable">
          <thead id="auditThead"></thead>
          <tbody id="auditTbody"></tbody>
        </table>
      </div>
    </div>
    </div> <!-- Fin mod_cartera -->

    <!-- ======================================================================= -->
    <!-- MÓDULO 2: SEGUIMIENTO DE CLASES PRÁCTICAS Y ALUMNOS (F-59 a F-62)       -->
    <!-- ======================================================================= -->
    <div id="mod_clases" class="mod-section">
      <div class="kpi-grid" id="kpiClasesGrid"></div>

      <div class="control-bar">
        <datalist id="dlAlumnosGlobal"></datalist>
        <datalist id="dlCedulasGlobal"></datalist>
        <datalist id="dlInstructoresGlobal"></datalist>

        <div class="panel">
          <div class="panel-title">
            <span>➕ Registrar, Buscar o Actualizar Alumno en Google Sheets (Sincronizado con el Bot)</span>
            <span class="badge badge-info">Hoja: Alumnos</span>
          </div>
          <p style="font-size:11.5px; color:#93c5fd; margin-bottom:8px;">
            💡 Escribe la <b>Cédula</b> o el <b>Nombre</b> para autocompletar sus datos actuales, o haz clic en el botón <b>✏️ Editar</b> de cualquier alumno en la tabla inferior.
          </p>
          <div class="form-grid-4">
            <div>
              <label class="field-label">Cédula / Documento *</label>
              <input type="text" id="al_cedula" list="dlCedulasGlobal" autocomplete="off" placeholder="Ej: 1094950123 (Escribe para buscar)" oninput="autocompletarFormularioAlumno('cedula')" />
            </div>
            <div>
              <label class="field-label">Nombre Completo del Alumno *</label>
              <input type="text" id="al_nombre" list="dlAlumnosGlobal" autocomplete="off" placeholder="Ej: JUAN PEREZ GOMEZ (Escribe para buscar)" oninput="autocompletarFormularioAlumno('nombre')" />
            </div>
            <div>
              <label class="field-label">Categoría *</label>
              <select id="al_cat">
                <option value="A2">🏍️ A2 Moto (15 clases - F-60)</option>
                <option value="B1">🚗 B1 Carro Particular (20 clases - F-61)</option>
                <option value="C1">🚕 C1 Servicio Público (30 clases - F-62)</option>
                <option value="RECAT_C1">🔄 Recategorización C1 (10 clases - F-62)</option>
                <option value="A1">🏍️ A1 Moto (8 clases - F-59)</option>
              </select>
            </div>
            <div>
              <label class="field-label">Estado RUNT (Por Categoría)</label>
              <select id="al_runt">
                <option value="No Registrado RUNT">⏳ No Registrado RUNT (Pendiente)</option>
                <option value="Cargue RUNT">✅ Cargue RUNT (Ocultar de pendientes)</option>
              </select>
            </div>
            <div>
              <label class="field-label">Teléfono / Celular</label>
              <input type="text" id="al_tel" autocomplete="off" placeholder="Ej: 3101234567" />
            </div>
            <div style="display:flex; gap:6px; grid-column: span 2;">
              <button type="button" class="btn btn-primary" style="flex:1; justify-content:center;" onclick="registrarAlumnoWeb()">💾 Guardar / Actualizar Alumno</button>
              <button type="button" class="btn btn-success" style="flex:1; justify-content:center;" onclick="marcarCargueRuntDesdeFormulario()" title="Marca Cargue RUNT para este alumno y esta categoría y lo filtra del listado">✅ Marcar Cargue RUNT</button>
              <button type="button" class="btn btn-outline" title="Buscar en la tabla" onclick="buscarAlumnoDesdeFormulario()">🔎</button>
            </div>
          </div>
          <div id="alumnoToast" class="status-toast"></div>
        </div>

        <div class="panel">
          <div class="panel-title">
            <span>📝 Registrar Clase Práctica Manualmente</span>
            <span class="badge badge-success">Hoja: Registro_Clases</span>
          </div>
          <p style="font-size:11.5px; color:#6ee7b7; margin-bottom:8px;">
            💡 Selecciona o escribe la <b>Cédula</b> o <b>Nombre</b> del alumno y se cargará su categoría y su siguiente clase pendiente automáticamente.
          </p>
          <div class="form-grid-4">
            <div>
              <label class="field-label">Cédula Alumno *</label>
              <input type="text" id="cl_cedula" list="dlCedulasGlobal" autocomplete="off" placeholder="Cédula (Autocompleta)" oninput="autocompletarFormularioClase('cedula')" />
            </div>
            <div>
              <label class="field-label">Nombre Alumno *</label>
              <input type="text" id="cl_alumno" list="dlAlumnosGlobal" autocomplete="off" placeholder="Nombre Alumno (Autocompleta)" oninput="autocompletarFormularioClase('nombre')" />
            </div>
            <div>
              <label class="field-label">Categoría</label>
              <select id="cl_cat">
                <option value="A2">A2</option>
                <option value="B1">B1</option>
                <option value="C1">C1</option>
                <option value="RECAT_C1">RECAT_C1</option>
                <option value="A1">A1</option>
              </select>
            </div>
            <div>
              <label class="field-label">Instructor *</label>
              <input type="text" id="cl_instructor" list="dlInstructoresGlobal" autocomplete="off" placeholder="Nombre Instructor" />
            </div>
            <div>
              <label class="field-label">Clases Dictadas (Ej: 1, 2)</label>
              <input type="text" id="cl_nums" value="1, 2" />
            </div>
            <div>
              <button type="button" class="btn btn-success" style="width:100%; justify-content:center;" onclick="registrarClaseWeb()">✅ Registrar Clase</button>
            </div>
          </div>
          <div id="claseToast" class="status-toast"></div>
        </div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>🎓 Listado de Alumnos en Clases Prácticas + Cruce en Vivo con Cartera + Descarga PDF Oficial (F-59 a F-62)</span>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <select id="filtroCatClases" onchange="renderizarModuloClases()" style="width:auto;">
              <option value="TODAS">Todas las Categorías</option>
              <option value="A1">A1 (F-59)</option>
              <option value="A2">A2 (F-60)</option>
              <option value="B1">B1 (F-61)</option>
              <option value="C1">C1 (F-62)</option>
              <option value="RECAT_C1">RECAT_C1 (F-62)</option>
            </select>
            <button type="button" class="btn btn-outline" onclick="cargarBotOverview()">🔄 Sincronizar con Bot</button>
          </div>
        </div>

        <!-- Sub-Pestañas para Control Manual de Cargue RUNT por Categoría y Semáforo de Vencimiento 90 Días -->
        <div class="tabs-header" id="clasesTabsHeader">
          <button type="button" class="tab-btn active" onclick="cambiarSubTabClases('PENDIENTE_RUNT', this)">
            ⏳ Pendientes Cargue RUNT <span class="badge badge-warning" id="cnt_cl_runt_pend">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarSubTabClases('SEMAFORO_ROJO', this)" style="border-color:rgba(239,68,68,0.65);">
            🔴 Semáforo Crítico / Próximos a Vencer 90d <span class="badge badge-danger" id="cnt_cl_sem_rojo">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarSubTabClases('SEMAFORO_AMARILLO', this)" style="border-color:rgba(245,158,11,0.55);">
            🟡 Semáforo en Alerta (90d) <span class="badge badge-warning" id="cnt_cl_sem_amarillo">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarSubTabClases('CARGADOS_RUNT', this)">
            ✅ Ya Cargados en RUNT <span class="badge badge-success" id="cnt_cl_runt_ok">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarSubTabClases('EN_CURSO', this)">
            🟢 En Curso (&lt; 100% + No RUNT) <span class="badge badge-info" id="cnt_cl_activos">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarSubTabClases('FINALIZADOS', this)">
            🎓 Finalizaron Clases (100%) <span class="badge badge-success" id="cnt_cl_fin">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarSubTabClases('CON_SALDO', this)">
            ⚠️ Con Saldo Pendiente <span class="badge badge-warning" id="cnt_cl_saldo">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarSubTabClases('CERO_REAL', this)">
            🚨 Con Abono $0 (Bloquear) <span class="badge badge-danger" id="cnt_cl_cero">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarSubTabClases('TODOS', this)">
            📚 Todos (Histórico Completo) <span class="badge" id="cnt_cl_todos">0</span>
          </button>
        </div>
        <input type="hidden" id="filtroAlertaCartera" value="PENDIENTE_RUNT" />

        <div class="table-toolbar">
          <div style="display:flex; gap:8px; flex-wrap:wrap; align-items:center; flex:1;">
            <input type="search" id="searchClasesInput" autocomplete="off" placeholder="🔎 Buscar alumno por nombre, cédula, categoría o semáforo..." oninput="renderizarModuloClases()" style="max-width: 390px; flex:1;" />
            <select id="sortClasesSelect" onchange="renderizarModuloClases()" style="width:auto; min-width:255px; font-weight:600; border-color:#3b82f6;">
              <option value="RECIENTES" selected>📅 Orden: Más Recientes a Más Antiguos</option>
              <option value="SEMAFORO_URGENCIA">🚦 Semáforo 90 Días: Más Críticos / Próximos a Vencer Primero</option>
              <option value="MENOS_DIAS">⏰ Menos Días Restantes (De 90 Días) Primero</option>
              <option value="ANTIGUOS">⏳ Orden: Más Antiguos a Más Recientes</option>
              <option value="AVANCE_DESC">🎓 Mayor Avance de Clases Primero</option>
            </select>
          </div>
          <span id="clasesCountInfo" style="font-size:12.5px; color:var(--muted);">💡 Control Semáforo 90 Días desde 1ra Clase Práctica + Control Manual Cargue RUNT por categoría</span>
        </div>

        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Cédula</th>
                <th>Alumno</th>
                <th>Categoría</th>
                <th>🚦 Semáforo 90 Días (Desde 1ra Clase)</th>
                <th>Estado RUNT (Control Manual)</th>
                <th>Formato Oficial</th>
                <th>Avance Clases Prácticas</th>
                <th>Estado Curso</th>
                <th>Cruce Contable Cartera</th>
                <th>Saldo Real Cédula</th>
                <th>Acciones / PDF Oficial / Despacho</th>
              </tr>
            </thead>
            <tbody id="clasesTbody"></tbody>
          </table>
        </div>
      </div>

      <!-- =================================================================== -->
      <!-- PANEL INTEGRADO: HOJA DE CÁLCULO CONTROL DIGITAL RUNT               -->
      <!-- ID: 1XHbqfDa7dK5CWcgWKNgJZiXN_OSrjJCEJJu8F4z-u2M                    -->
      <!-- =================================================================== -->
      <div class="panel" style="margin-top:18px; border-color:#3b82f6; background:linear-gradient(180deg, rgba(30,58,138,0.18), rgba(15,23,42,0.96));">
        <div class="panel-title">
          <div style="display:flex; align-items:center; gap:10px; flex-wrap:wrap;">
            <span>🏛️ Hoja de Cálculo Integrada: Control Digital de Reporte de Clases en Plataforma RUNT</span>
            <span class="badge badge-info" id="runtSheetStatusBadge">ID: 1XHbqfDa7dK5CWcgWKNgJZiXN_OSrjJCEJJu8F4z-u2M</span>
          </div>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <button type="button" class="btn btn-primary" onclick="sincronizarHorasBotHaciaHojaRunt()" title="Toma las horas prácticas dictadas por los instructores en el Bot y actualiza la columna Practica de esta hoja">🔄 Pasar Horas Prácticas del Bot a Hoja RUNT</button>
            <button type="button" class="btn btn-outline" onclick="cargarPanoramaRuntDigital(true)">🔃 Recargar Hoja RUNT</button>
            <a id="runtSheetExternalLink" href="https://docs.google.com/spreadsheets/d/1XHbqfDa7dK5CWcgWKNgJZiXN_OSrjJCEJJu8F4z-u2M/edit" target="_blank" class="btn btn-outline">↗️ Abrir en Google Sheets</a>
          </div>
        </div>

        <!-- KPIs de la Hoja RUNT -->
        <div class="kpi-grid" id="kpiRuntSheetGrid" style="margin-bottom:14px;"></div>

        <!-- ================================================================= -->
        <!-- BARRA DE CARGUE MASIVO RUNT (TEORÍA Y TALLER + TOPE 8 CLASES/DÍA) -->
        <!-- ================================================================= -->
        <div class="panel" style="margin-bottom:14px; background:linear-gradient(135deg, rgba(16,185,129,0.12), rgba(15,23,42,0.95)); border-color:#10b981;">
          <div class="panel-title">
            <span>🚀 Centro de Cargue Masivo de Teoría y Taller (Guía Paso a Paso + Botón Deshacer Errores)</span>
            <span class="badge badge-success">Con Confirmación Segura y Selección por Casillas ☑️</span>
          </div>
          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(240px, 1fr)); gap:8px; margin-bottom:12px;">
            <div style="background:rgba(15,23,42,0.85); border:1px solid #334155; border-left:4px solid #a855f7; border-radius:8px; padding:8px 10px; font-size:11.5px; color:#e2e8f0;">
              <b>1️⃣ PASO 1: Elige el Filtro abajo y Marca ☑️</b><br/>
              Haz clic en la pestaña que necesitas (ej. <b>🏍️ Moto A2 sin 25h Teoría</b> o <b>🔧 Moto A2 sin 3h Taller</b>) y marca con chulito <b>☑️</b> en la tabla los alumnos que vas a reportar.
            </div>
            <div style="background:rgba(15,23,42,0.85); border:1px solid #334155; border-left:4px solid #fde047; border-radius:8px; padding:8px 10px; font-size:11.5px; color:#e2e8f0;">
              <b>2️⃣ PASO 2: Copia sus Cédulas para RUNT</b><br/>
              Haz clic en <b>📋 Copiar Cédulas</b> para llevarte las cédulas de los alumnos marcados y pegarlas en el cargue masivo de la plataforma RUNT.
            </div>
            <div style="background:rgba(15,23,42,0.85); border:1px solid #334155; border-left:4px solid #10b981; border-radius:8px; padding:8px 10px; font-size:11.5px; color:#e2e8f0;">
              <b>3️⃣ PASO 3: Registra las Horas o Deshaz</b><br/>
              Elige la fecha (<b>Ayer / Hoy / Mañana</b>) y pulsa <b>📘 Sumar Teoría</b> o <b>🔧 Sumar Taller</b>. El sistema te pedirá confirmación antes de guardar, y si te equivocas puedes pulsar <b>↩️ Deshacer / Corregir</b>.
            </div>
          </div>
          <div class="form-grid-4">
            <div>
              <label class="field-label">📅 Fecha del Reporte / Cargue (Ayer / Hoy / Mañana)</label>
              <div style="display:flex; gap:4px;">
                <input type="date" id="rn_batch_fecha" style="flex:1;" />
                <button type="button" class="btn btn-outline" style="padding:4px 7px; font-size:10.5px;" title="Poner fecha de Ayer" onclick="ajustarFechaBatchRunt(-1)">Ayer</button>
                <button type="button" class="btn btn-outline" style="padding:4px 7px; font-size:10.5px;" title="Poner fecha de Hoy" onclick="ajustarFechaBatchRunt(0)">Hoy</button>
                <button type="button" class="btn btn-outline" style="padding:4px 7px; font-size:10.5px;" title="Poner fecha de Mañana (Ej: Lunes adelantando Martes)" onclick="ajustarFechaBatchRunt(1)">Mañana</button>
              </div>
            </div>
            <div>
              <label class="field-label">🏍️/🚗 Categoría para Cargue Masivo</label>
              <select id="rn_batch_cat">
                <option value="A2" selected>🏍️ A2 Moto (Meta: 25 Teo / 3 Taller)</option>
                <option value="B1">🚗 B1 Carro Particular (Meta: 25 Teo / 5 Taller)</option>
                <option value="C1">🚕 C1 Público (Meta: 30 Teo / 5 Taller)</option>
                <option value="RECAT_C1">🔄 RAC. C1 (Meta: 8 Teo / 0 Taller)</option>
                <option value="A1">🏍️ A1 Moto (Meta: 25 Teo / 3 Taller)</option>
                <option value="TODAS">📚 Todas las Categorías</option>
              </select>
            </div>
            <div>
              <label class="field-label">⚖️ Regla de Cupo Diario RUNT</label>
              <select id="rn_batch_modo">
                <option value="MAX_8" selected>🛡️ Tope Reglamentario: Sumar máx. +8 clases en esta fecha</option>
                <option value="COMPLETAR_META">⚡ Excepción (Ponerse al día): Completar el 100% que lleva</option>
              </select>
            </div>
            <div>
              <label class="field-label">💰 Valor Sanción por Curso Vencido (90d)</label>
              <input type="number" id="rn_sancion_valor" value="50000" step="5000" min="0" title="Valor de la sanción económica cuando se vence el curso de 90 días y toca reiniciar el reporte en RUNT" />
            </div>
          </div>
          <div style="display:flex; gap:8px; flex-wrap:wrap; margin-top:10px; align-items:center;">
            <button type="button" class="btn btn-outline" style="border-color:#fde047; color:#fde047; font-weight:700;" onclick="copiarCedulasFiltroActualRunt()" title="Copia al portapapeles las Cédulas de los alumnos marcados ☑️ (o de toda la pestaña actual) para pegarlas en RUNT">📋 1. Copiar Cédulas (Seleccionados ☑️ o Filtro)</button>
            <button type="button" class="btn btn-primary" onclick="ejecutarCargueMasivoRunt('TEORIA')" title="Suma +8h de Teoría (o completa las 25h/30h) a los alumnos seleccionados ☑️ (pide confirmación antes de aplicar)">📘 2. Cargue Masivo TEORÍA</button>
            <button type="button" class="btn btn-primary" style="background:#0284c7; border-color:#38bdf8;" onclick="ejecutarCargueMasivoRunt('TALLER')" title="Suma las horas de Taller (3h en Moto / 5h en Carro) a los alumnos seleccionados ☑️ (pide confirmación antes de aplicar)">🔧 2. Cargue Masivo TALLER</button>
            <button type="button" class="btn btn-success" onclick="ejecutarCargueMasivoRunt('TEORIA_Y_TALLER')" title="Actualiza Teoría + Taller en lote para los alumnos seleccionados ☑️">✨ Cargue Masivo TEORÍA + TALLER</button>
            <button type="button" class="btn btn-danger" style="background:#dc2626; border-color:#f87171; font-weight:700;" onclick="deshacerCargueMasivoRunt('RESTAURAR_14_A2_ORIGINAL')" title="Corrige inmediatamente el cargue masivo hecho por error y devuelve los 14 alumnos de Moto A2 a sus horas originales (16h, 8h, 24h, 0h)">↩️ Deshacer / Corregir Cargue Masivo por Error (Restaurar 14 Alumnos A2)</button>
            <button type="button" class="btn btn-outline" style="border-color:#f97316; color:#fdba74;" onclick="deshacerCargueMasivoRunt('RESTAR_8_TEORIA')" title="Resta -8 horas de Teoría a los alumnos marcados con ☑️">➖ Restar -8h Teoría (a Seleccionados ☑️)</button>
          </div>
        </div>

        <!-- Formulario de Actualización / Registro Individual en Pestaña 'Practica' de la Hoja RUNT -->
        <div class="panel" style="margin-bottom:14px; background:#0f172a; border-color:#2b3d63;">
          <div class="panel-title">
            <span>📝 Reportar / Actualizar Conteo Individual de Horas en RUNT (Práctica Individual • Máx. 8 Clases/Día o Ponerse al Día)</span>
            <span class="badge badge-warning" id="runtEditRowBadge">Nuevo / Buscar por Cédula o Clic en ✏️ de la Tabla</span>
          </div>
          <input type="hidden" id="rn_fila" value="0" />
          <div class="form-grid-4">
            <div>
              <label class="field-label">Cédula Alumno *</label>
              <input type="text" id="rn_cedula" list="dlCedulasGlobal" autocomplete="off" placeholder="Ej: 1094278698" oninput="autocompletarFormularioRuntSheet('cedula')" />
            </div>
            <div>
              <label class="field-label">Nombre y Apellido *</label>
              <input type="text" id="rn_nombre" list="dlAlumnosGlobal" autocomplete="off" placeholder="Ej: ADRIAN MAURICIO RAMIREZ" oninput="autocompletarFormularioRuntSheet('nombre')" />
            </div>
            <div>
              <label class="field-label">Categoría *</label>
              <select id="rn_cat" onchange="ajustarMetasFormularioRunt()">
                <option value="A2">🏍️ A2 (25 Teo + 3 Tal + 15 Prác = 43h)</option>
                <option value="B1">🚗 B1 (25 Teo + 5 Tal + 20 Prác = 50h)</option>
                <option value="C1">🚕 C1 (30 Teo + 5 Tal + 30 Prác = 65h)</option>
                <option value="RECAT_C1">🔄 RAC. C1 (8 Teo + 0 Tal + 8 Prác = 16h)</option>
                <option value="A1">🏍️ A1 (25 Teo + 3 Tal + 8 Prác = 36h)</option>
              </select>
            </div>
            <div>
              <label class="field-label">Prioridad Operativa</label>
              <select id="rn_prioridad">
                <option value="EN PRACTICA">🚨 EN PRACTICA (Urgente a Certificar)</option>
                <option value="Termino">🔥 Termino (Terminó Práctica • Listo Certificar)</option>
                <option value="revisado">🔍 revisado</option>
                <option value="Novedad">⚠️ Novedad</option>
              </select>
            </div>
            <div>
              <label class="field-label">📅 Fecha Matrícula (Calcula Vence 90d)</label>
              <input type="date" id="rn_fecha_mat" />
            </div>
            <div>
              <label class="field-label">📘 Horas Teoría Cargadas</label>
              <input type="number" id="rn_teoria" value="25" min="0" max="40" />
            </div>
            <div>
              <label class="field-label">🔧 Horas Taller Cargadas</label>
              <input type="number" id="rn_taller" value="3" min="0" max="15" />
            </div>
            <div>
              <label class="field-label">🚗 Horas Práctica Cargadas (Individual)</label>
              <div style="display:flex; gap:4px;">
                <input type="number" id="rn_practica" value="15" min="0" max="40" style="flex:1;" />
                <button type="button" class="btn btn-outline" style="padding:4px 7px; font-size:10.5px;" title="Sumar +2 clases prácticas" onclick="sumarHorasInputFormulario('rn_practica', 2)">+2h</button>
                <button type="button" class="btn btn-outline" style="padding:4px 7px; font-size:10.5px; border-color:#38bdf8; color:#38bdf8;" title="Sumar tope máximo diario (+8 clases prácticas)" onclick="sumarHorasInputFormulario('rn_practica', 8)">+8h (Tope)</button>
              </div>
            </div>
            <div>
              <label class="field-label">📅 Fecha de Cargue (Ayer / Hoy / Mañana)</label>
              <input type="date" id="rn_fecha_cargue" />
            </div>
            <div>
              <label class="field-label">📌 Etiqueta Cargue Diario</label>
              <select id="rn_cargue_diario">
                <option value="">-- Sin etiqueta diaria --</option>
                <option value="TEORIA">📘 TEORIA</option>
                <option value="TALLER">🔧 TALLER</option>
                <option value="PRACTICA">🚗 PRACTICA</option>
                <option value="COMPLETO">✅ COMPLETO</option>
              </select>
            </div>
            <div>
              <label class="field-label">🏛️ Actividad / Estado RUNT</label>
              <select id="rn_estado_runt">
                <option value="PENDIENTE RUNT">⚠️ PENDIENTE RUNT</option>
                <option value="CARGADO RUNT">✅ CARGADO RUNT</option>
              </select>
            </div>
            <div style="display:flex; gap:6px;">
              <button type="button" class="btn btn-primary" style="flex:1; justify-content:center;" onclick="guardarRegistroHojaRunt(false)">💾 Guardar Conteo</button>
              <button type="button" class="btn btn-success" style="flex:1; justify-content:center;" onclick="guardarRegistroHojaRunt(true)" title="Completa el 100% de Teoría+Taller+Práctica y marca ✅ CARGADO RUNT">✅ 100% + CARGADO RUNT</button>
            </div>
          </div>
          <div id="runtSheetToast" class="status-toast"></div>
        </div>

        <!-- Sub-Pestañas Operativas por Prioridad de Certificación y Filtros Masivos -->
        <div class="tabs-header" id="runtSheetTabsHeader">
          <button type="button" class="tab-btn active" onclick="cambiarTabHojaRunt('PENDIENTES_PRACTICA', this)">
            📋 Todos Pendientes (Orden Prioridad) <span class="badge badge-warning" id="cnt_rn_pend">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('TERMINARON_SIN_CARGUE', this)" style="border-color:#f97316; color:#fed7aa;">
            🔥 1. Terminaron y Falta Cargue Manual RUNT <span class="badge badge-danger" id="cnt_rn_term_sin">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('URGENTES_EN_PRACTICA', this)" style="border-color:#38bdf8; color:#bae6fd;">
            🚨 2. Urgentes: En Práctica (A Certificar) <span class="badge badge-info" id="cnt_rn_urg_prac">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('MOTO_SIN_25_TEORIA', this)" style="border-color:#a855f7;">
            🏍️ Moto A2 sin 25h Teoría (Masivo) <span class="badge badge-purple" id="cnt_rn_moto_teo">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('MOTO_SIN_3_TALLER', this)" style="border-color:#a855f7;">
            🔧 Moto A2 sin 3h Taller (Masivo) <span class="badge badge-purple" id="cnt_rn_moto_tal">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('CARRO_SIN_TEO_TAL', this)">
            🚗 Carro B1/C1 sin Teoría o Taller <span class="badge badge-info" id="cnt_rn_carro_tt">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('DESFASE_BOT', this)">
            🔄 Con Horas Nuevas del Bot <span class="badge badge-info" id="cnt_rn_bot">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('POR_VENCER_90', this)" style="border-color:#ef4444;">
            💸 Vencidos / Por Vencer 90d (Reiniciar + Sanción) <span class="badge badge-danger" id="cnt_rn_vence">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('CARGADOS_PRACTICA', this)">
            ✅ Ya Cargados RUNT <span class="badge badge-success" id="cnt_rn_cargados">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('TODOS_PRACTICA', this)">
            📚 Histórico (191) <span class="badge" id="cnt_rn_todos">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('TAB_NO_RUNT', this)">
            🚫 Pestaña "No runt" <span class="badge badge-danger" id="cnt_rn_norunt">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('TAB_INICIO_PRACTICA', this)">
            🕒 Pestaña "Inicio Practica" <span class="badge badge-purple" id="cnt_rn_horarios">0</span>
          </button>
          <button type="button" class="tab-btn" onclick="cambiarTabHojaRunt('TAB_CODIGOS_RUNT', this)">
            🔢 Pestaña "Codigo de Runt" <span class="badge badge-info" id="cnt_rn_codigos">0</span>
          </button>
        </div>

        <div class="table-toolbar">
          <input type="search" id="searchRuntSheetInput" autocomplete="off" placeholder="🔎 Buscar en esta hoja por nombre, cédula, categoría, prioridad o estado..." oninput="renderizarTablaHojaRunt()" style="max-width:460px;" />
          <span id="runtSheetCountInfo" style="font-size:12.5px; color:var(--muted);">Haz clic en <b>✏️ Editar</b> o en <b>✅ CARGADO RUNT</b> para actualizar en tiempo real</span>
        </div>

        <div class="table-wrap">
          <table>
            <thead id="runtSheetThead"></thead>
            <tbody id="runtSheetTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 3: LIQUIDACIÓN / NÓMINA Y HORARIOS DE INSTRUCTORES               -->
    <!-- ======================================================================= -->
    <div id="mod_nomina" class="mod-section">
      <div class="kpi-grid" id="kpiNominaGrid"></div>

      <div class="panel" style="margin-bottom:18px;">
        <div class="panel-title">
          <span>👨‍🏫 Filtros de Liquidación de Instructores y Validación de Horario de Envío (Min 18-28 / Min 40-50)</span>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <select id="nom_periodo" onchange="cargarBotOverview()" style="width:auto;">
              <option value="mes">📅 Este Mes</option>
              <option value="dia">📅 Hoy</option>
              <option value="ayer">📅 Ayer</option>
              <option value="semana">📅 Esta Semana</option>
              <option value="semana_ant">📅 Semana Anterior</option>
              <option value="mes_ant">📅 Mes Anterior</option>
              <option value="anio">📅 Este Año</option>
              <option value="todo">📚 Todo el Histórico</option>
            </select>
            <select id="nom_instructor" onchange="cargarBotOverview()" style="width:auto;">
              <option value="TODOS">👥 Todos los Instructores</option>
            </select>
            <button type="button" class="btn btn-danger" onclick="descargarPdfNomina()">📄 Descargar PDF Oficial de Nómina / Liquidación</button>
          </div>
        </div>

        <div class="table-wrap" style="max-height:290px; margin-bottom:16px;">
          <table>
            <thead>
              <tr>
                <th>Instructor</th>
                <th>Sesiones Registradas</th>
                <th>Total Clases (Horas)</th>
                <th>✅ A Tiempo (En Horario)</th>
                <th>⚠️ Fuera de Horario</th>
                <th>Desglose por Categoría</th>
                <th>Acción PDF Individual</th>
              </tr>
            </thead>
            <tbody id="nominaResumenTbody"></tbody>
          </table>
        </div>

        <div class="panel-title" style="margin-top:10px;">
          <span>📋 Detalle de Clases Dictadas y Validación de Minuto de Envío</span>
          <input type="search" id="searchNominaInput" autocomplete="off" placeholder="🔎 Filtrar por instructor, alumno, categoría o validación..." oninput="renderizarModuloNomina()" style="max-width: 360px;" />
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Fecha</th>
                <th>Hora Envío</th>
                <th>Instructor</th>
                <th>Cédula Alumno</th>
                <th>Alumno</th>
                <th>Cat.</th>
                <th>Clases Marcadas</th>
                <th>N° Horas</th>
                <th>Validación Horario</th>
                <th>Detalle Minuto</th>
              </tr>
            </thead>
            <tbody id="nominaDetalleTbody"></tbody>
          </table>
        </div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>📸 Cierres de Jornada Diarios (Foto de Agenda / Programación Enviada al Bot)</span>
          <span class="badge badge-purple" id="cntCierresBadge">0 cierres</span>
        </div>
        <div class="table-wrap" style="max-height:260px;">
          <table>
            <thead>
              <tr>
                <th>Fecha</th>
                <th>Hora</th>
                <th>Instructor</th>
                <th>Estado Cierre</th>
                <th>Evidencia / Enlace Drive</th>
              </tr>
            </thead>
            <tbody id="cierresTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 4: CONTROL DE INASISTENCIAS Y RECAUDO DE MULTAS                  -->
    <!-- ======================================================================= -->
    <div id="mod_inasistencias" class="mod-section">
      <div class="kpi-grid" id="kpiInasGrid"></div>

      <div class="panel" style="margin-bottom:18px;">
        <div class="panel-title">
          <span>➕ Registrar o Actualizar Inasistencia de Alumno (Sincronizado con Google Sheets y Bot)</span>
          <span class="badge badge-warning">Hoja: Inasistencias</span>
        </div>
        <p style="font-size:11.5px; color:#fde68a; margin-bottom:8px;">
          💡 Al escribir el <b>Nombre del Alumno</b>, el sistema busca automáticamente sus inasistencias en la tabla de abajo para que puedas actualizar su medio de pago (💵 Efectivo / 🏦 Transf. / ⏳ Pendiente) con 1 clic, o registrar una nueva inasistencia.
        </p>
        <div class="form-grid-4">
          <div>
            <label class="field-label">Fecha Programada *</label>
            <input type="text" id="in_fecha" autocomplete="off" placeholder="DD/MM/YYYY" />
          </div>
          <div>
            <label class="field-label">Nombre Completo del Alumno *</label>
            <input type="text" id="in_alumno" list="dlAlumnosGlobal" autocomplete="off" placeholder="Escribe para buscar o registrar (Ej: Nicolás Vera)" oninput="autocompletarYFiltrarInasistencia(this.value)" />
          </div>
          <div>
            <label class="field-label">Categoría *</label>
            <select id="in_cat">
              <option value="B1">🚗 B1 Carro</option>
              <option value="C1">🚕 C1 Público</option>
              <option value="A2">🏍️ A2 Moto</option>
              <option value="A1">🏍️ A1 Moto</option>
            </select>
          </div>
          <div>
            <label class="field-label">Hora de Clase *</label>
            <input type="text" id="in_hora" autocomplete="off" placeholder="Ej: 6:00 a 8:00 AM" value="6:00 a 8:00 AM" />
          </div>
          <div>
            <label class="field-label">Instructor Asignado *</label>
            <input type="text" id="in_instructor" list="dlInstructoresGlobal" autocomplete="off" placeholder="Ej: ALDEMAR / HECTOR" />
          </div>
          <div>
            <label class="field-label">Estado / Medio de Pago *</label>
            <select id="in_estado">
              <option value="Pendiente Pago">⏳ Pendiente Pago</option>
              <option value="Cancelo Efectivo">💵 Cancelo Efectivo</option>
              <option value="Cancelo Trasferencia">🏦 Cancelo Trasferencia</option>
            </select>
          </div>
          <div style="display:flex; gap:6px; grid-column: span 2;">
            <button type="button" class="btn btn-primary" style="flex:1; justify-content:center;" onclick="registrarInasistenciaWeb()">💾 Registrar Nueva Inasistencia</button>
            <button type="button" class="btn btn-success" style="flex:1; justify-content:center;" onclick="actualizarPagoAlumnoSeleccionadoInas()" title="Actualiza el estado de pago de la inasistencia encontrada del alumno">🔄 Actualizar Estado de Pago de Este Alumno</button>
          </div>
        </div>
        <div id="inasToast" class="status-toast"></div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>🚫 Historial Oficial de Inasistencias y Control de Pagos</span>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <select id="inas_periodo" onchange="cargarBotOverview()" style="width:auto;">
              <option value="todo">📚 Todo el Histórico</option>
              <option value="mes">📅 Este Mes</option>
              <option value="semana">📅 Esta Semana</option>
              <option value="dia">📅 Hoy</option>
              <option value="ayer">📅 Ayer</option>
              <option value="mes_ant">📅 Mes Anterior</option>
              <option value="anio">📅 Este Año</option>
            </select>
            <select id="inas_estado_filtro" onchange="cargarBotOverview()" style="width:auto;">
              <option value="TODOS">Todos los Estados</option>
              <option value="PENDIENTE">⏳ Solo Pendiente Pago</option>
              <option value="PAGADO">✅ Solo Pagadas (Efectivo / Transf.)</option>
            </select>
            <button type="button" class="btn btn-danger" onclick="descargarPdfInasistenciasGeneral()">📄 Descargar PDF Consolidado Inasistencias</button>
          </div>
        </div>

        <div class="table-toolbar">
          <input type="search" id="searchInasInput" autocomplete="off" placeholder="🔎 Buscar inasistencia por alumno, instructor, fecha o estado..." oninput="renderizarModuloInasistencias()" style="max-width: 440px;" />
          <span id="inasCountInfo" style="font-size:12.5px; color:var(--muted);">Haz clic en 💵 Efectivo / 🏦 Transf. para actualizar el pago en vivo</span>
        </div>

        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>#</th>
                <th>Fecha</th>
                <th>Alumno</th>
                <th>Cat.</th>
                <th>Horario</th>
                <th>Instructor</th>
                <th>Estado de Pago</th>
                <th>Actualizar Pago (1 Clic)</th>
                <th>PDF Individual</th>
              </tr>
            </thead>
            <tbody id="inasTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 5: MANTENIMIENTOS Y FLOTA VEHICULAR PESV (FORMATO SISI-F2-08)    -->
    <!-- ======================================================================= -->
    <div id="mod_pesv" class="mod-section">
      <div class="kpi-grid" id="kpiPesvGrid"></div>

      <div class="control-bar">
        <div class="panel">
          <div class="panel-title">
            <span>🔧 Registrar Nuevo Mantenimiento Vehicular (Formato Oficial SISI-F2-08)</span>
            <span class="badge badge-success">Hoja: Mantenimientos</span>
          </div>
          <div class="form-grid-4">
            <div>
              <label class="field-label">Placa del Vehículo *</label>
              <select id="mt_placa"></select>
            </div>
            <div>
              <label class="field-label">Tipo Mantenimiento *</label>
              <select id="mt_tipo">
                <option value="PREVENTIVO">🟢 PREVENTIVO</option>
                <option value="CORRECTIVO">🟠 CORRECTIVO</option>
              </select>
            </div>
            <div>
              <label class="field-label">Kilometraje Actual *</label>
              <input type="text" id="mt_km" placeholder="Ej: 161500" />
            </div>
            <div>
              <label class="field-label">Taller / Proveedor *</label>
              <input type="text" id="mt_taller" placeholder="Ej: TALLER / MONTALLANTAS" />
            </div>
            <div>
              <label class="field-label">Mecánico / Técnico *</label>
              <input type="text" id="mt_mecanico" placeholder="Ej: HECTOR / JAIRO" />
            </div>
            <div style="grid-column: span 2;">
              <label class="field-label">Actividad Realizada / Repuestos *</label>
              <input type="text" id="mt_actividad" placeholder="Ej: CAMBIO DE ACEITE Y FILTROS / REVISION FRENOS" />
            </div>
            <div>
              <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="registrarMantenimientoWeb()">💾 Guardar Mantenimiento</button>
            </div>
          </div>
          <div id="mtToast" class="status-toast"></div>
        </div>

        <div class="panel">
          <div class="panel-title">
            <span>🚘 Gestión de Flota Activa PESV</span>
            <span class="badge badge-info">Hoja: Vehiculos_PESV</span>
          </div>
          <div class="form-grid-4" style="margin-bottom:10px;">
            <div>
              <label class="field-label">Placa *</label>
              <input type="text" id="vh_placa" placeholder="Ej: ABC123" />
            </div>
            <div>
              <label class="field-label">Tipo (CARRO / MOTO)</label>
              <select id="vh_tipo">
                <option value="CARRO">🚗 CARRO</option>
                <option value="MOTO">🏍️ MOTO</option>
              </select>
            </div>
            <div>
              <label class="field-label">Cat. (B1/C1/A2)</label>
              <input type="text" id="vh_cat" value="B1" />
            </div>
            <div>
              <label class="field-label">Marca / Línea</label>
              <input type="text" id="vh_marca" placeholder="RENAULT / AKT" />
            </div>
          </div>
          <button type="button" class="btn btn-success" style="width:100%; justify-content:center;" onclick="agregarVehiculoWeb()">➕ Agregar / Reactivar Vehículo en Flota</button>
          <div id="vhToast" class="status-toast"></div>
          <div id="flotaChipsWrap" style="margin-top:10px; display:flex; flex-wrap:wrap; gap:6px;"></div>
        </div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>📋 Bitácora Oficial de Mantenimientos Preventivos y Correctivos (SISI-F2-08 Firmado)</span>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <select id="pesv_periodo" onchange="cargarBotOverview()" style="width:auto;">
              <option value="todo">📚 Todo el Histórico (2026)</option>
              <option value="mes">📅 Este Mes</option>
              <option value="mes_ant">📅 Mes Anterior</option>
              <option value="semana">📅 Esta Semana</option>
              <option value="anio">📅 Este Año</option>
              <option value="mes_01">Enero 2026</option>
              <option value="mes_02">Febrero 2026</option>
              <option value="mes_03">Marzo 2026</option>
              <option value="mes_04">Abril 2026</option>
              <option value="mes_05">Mayo 2026</option>
              <option value="mes_06">Junio 2026</option>
              <option value="mes_07">Julio 2026</option>
              <option value="mes_08">Agosto 2026</option>
              <option value="mes_09">Septiembre 2026</option>
              <option value="mes_10">Octubre 2026</option>
              <option value="mes_11">Noviembre 2026</option>
              <option value="mes_12">Diciembre 2026</option>
            </select>
            <select id="pesv_placa_filtro" onchange="cargarBotOverview()" style="width:auto;">
              <option value="TODAS">🚘 Todas las Placas</option>
            </select>
            <button type="button" class="btn btn-danger" onclick="descargarPdfPesv()">📄 Descargar Formato Oficial SISI-F2-08 (Con Firmas)</button>
          </div>
        </div>

        <div class="table-toolbar">
          <input type="search" id="searchPesvInput" autocomplete="off" placeholder="🔎 Filtrar por placa, taller, mecánico, actividad o tipo..." oninput="renderizarModuloPesv()" style="max-width: 440px;" />
          <span id="pesvCountInfo" style="font-size:12.5px; color:var(--muted);">Firmado digitalmente por SEGUNDO MURILLO (Rep. Legal) y MARLY ALFONSO (Líder PESV)</span>
        </div>

        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>ID</th>
                <th>Fecha</th>
                <th>Placa</th>
                <th>Tipo Vehículo</th>
                <th>Tipo Mantenimiento</th>
                <th>Kilometraje</th>
                <th>Actividad Realizada / Intervención</th>
                <th>Taller / Proveedor</th>
                <th>Mecánico</th>
                <th>PDF Placa</th>
              </tr>
            </thead>
            <tbody id="pesvTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 6: ROBOT GESTOR DE ARCHIVOS Y DESPACHO (TELEGRAM + WHATSAPP)     -->
    <!-- ======================================================================= -->
    <div id="mod_robot" class="mod-section">
      <div class="kpi-grid" id="kpiRobotGrid">
        <div class="kpi-card">
          <div class="kpi-label">Estado del Robot Clasificador</div>
          <div class="kpi-value" style="color:#34d399;">🟢 Activo 24/7</div>
          <div class="kpi-sub">Clasificación por extensión y fecha (`YYYY-MM`)</div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">Conector Telegram Bot API</div>
          <div class="kpi-value" style="color:#c4b5fd;">✈️ Listo</div>
          <div class="kpi-sub">Envío directo de PDFs, Excels y Alertas</div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">Conector WhatsApp (API + Web)</div>
          <div class="kpi-value" style="color:#6ee7b7;">📲 Multicanal</div>
          <div class="kpi-sub">Cloud API / Evolution API + Enlace Directo</div>
        </div>
        <div class="kpi-card orange">
          <div class="kpi-label">Archivos Procesados y Despachados</div>
          <div class="kpi-value" id="robotTotalFiles" style="color:#fdba74;">0 Archivos</div>
          <div class="kpi-sub">Documentos PDF, Reportes Excel y Evidencias</div>
        </div>
      </div>

      <div class="control-bar">
        <div class="panel">
          <div class="panel-title">
            <span>🚀 1. Despachador Automático de Formatos Oficiales CEA (Telegram + WhatsApp)</span>
            <span class="badge badge-info">Envío Instantáneo</span>
          </div>
          <div class="form-grid-4">
            <div style="grid-column: span 2;">
              <label class="field-label">Seleccionar Reporte Oficial a Generar y Enviar *</label>
              <select id="rb_tipo_reporte">
                <option value="cartera_pdf">📄 Informe Ejecutivo de Auditoría de Cartera (.pdf)</option>
                <option value="cartera_excel">📊 Base Auditada Multi-Pestaña de Cartera (.xlsx)</option>
                <option value="alumno_f59">🎓 Ficha Oficial de Clases de Alumno (F-59 a F-62 .pdf)</option>
                <option value="nomina_pdf">👨‍🏫 Liquidación de Nómina de Instructores (.pdf)</option>
                <option value="inasistencias_pdf">🚫 Reporte Consolidado de Inasistencias (.pdf)</option>
                <option value="pesv_pdf">🔧 Formato Oficial Mantenimientos PESV SISI-F2-08 (.pdf)</option>
              </select>
            </div>
            <div style="grid-column: span 2;">
              <label class="field-label">Cédula / Nombre Alumno o Placa (Si aplica)</label>
              <input type="text" id="rb_criterio" list="listaAlumnosRobot" placeholder="Ej: 1098800332 o nombre del alumno / instructor" />
              <datalist id="listaAlumnosRobot"></datalist>
            </div>
            <div>
              <label class="field-label">✈️ Chat ID Telegram Destino</label>
              <input type="text" id="rb_tg_chat" value="1567561963" placeholder="Ej: 1567561963" />
            </div>
            <div>
              <label class="field-label">📲 Teléfono WhatsApp (Ej: 3155680575)</label>
              <input type="text" id="rb_wa_phone" value="3155680575" placeholder="Ej: 3155680575" />
            </div>
            <div>
              <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="despacharReporteRobot('TELEGRAM')">✈️ Enviar PDF/Excel a Telegram</button>
            </div>
            <div>
              <button type="button" class="btn btn-success" style="width:100%; justify-content:center;" onclick="despacharReporteRobot('WHATSAPP')">📲 Enviar a WhatsApp</button>
            </div>
          </div>
          <div id="robotDispatchToast" class="status-toast"></div>
        </div>

        <div class="panel">
          <div class="panel-title">
            <span>📂 2. Bandeja Web del Vigilante / Clasificador de Archivos</span>
            <a href="/api/bot/download_robot_script" class="btn btn-outline" style="padding:4px 10px; font-size:11.5px;">📥 Descargar robot_gestor_cea.py (PC)</a>
          </div>
          <p style="font-size:12px; color:#cbd5e1; margin-bottom:10px;">
            Sube cualquier archivo (`.pdf`, `.xlsx`, `.jpg`, `.png`) para clasificarlo automáticamente en <code>Archivo_Procesado/&lt;Tipo&gt;/&lt;YYYY-MM&gt;/</code> y despacharlo en vivo a Telegram / WhatsApp:
          </p>
          <form id="robotUploadForm">
            <div class="form-grid-4">
              <div style="grid-column: span 2;">
                <label class="field-label">Archivo a clasificar y archivar *</label>
                <input type="file" id="rb_file_input" accept=".pdf,.xlsx,.xls,.csv,.jpg,.jpeg,.png" required />
              </div>
              <div>
                <label class="field-label">Acción de Despacho</label>
                <select id="rb_auto_tg">
                  <option value="TELEGRAM">✈️ Clasificar + Enviar a Telegram</option>
                  <option value="WHATSAPP">📲 Clasificar + Enviar a WhatsApp</option>
                  <option value="AMBOS">🚀 Clasificar + Enviar a Telegram y WhatsApp</option>
                  <option value="SOLO_CLASIFICAR">📂 Solo clasificar y archivar</option>
                </select>
              </div>
              <div>
                <button type="submit" class="btn btn-primary" style="width:100%; justify-content:center;">🤖 Clasificar y Despachar</button>
              </div>
            </div>
          </form>
          <div id="robotUploadToast" class="status-toast"></div>
        </div>
      </div>

      <div class="panel" style="margin-bottom:18px; border-color:#3b82f6;">
        <div class="panel-title">
          <span>⚙️ Configuración de Telegram Bot + Pasarela WhatsApp API (Cloud API / Evolution API / WAHA)</span>
          <span class="badge badge-purple">Conecta aquí tu Token de Telegram para envío directo de PDFs</span>
        </div>
        <div class="form-grid-4">
          <div>
            <label class="field-label">🤖 Token del Bot de Telegram (`TELEGRAM_TOKEN`) *</label>
            <input type="password" id="cfg_tg_token" placeholder="Pega aquí tu token de BotFather..." />
          </div>
          <div>
            <label class="field-label">✈️ Chat ID Telegram Predeterminado</label>
            <input type="text" id="cfg_tg_chat" value="1567561963" placeholder="Ej: 1567561963" />
          </div>
          <div>
            <label class="field-label">📲 URL WhatsApp API (Opcional)</label>
            <input type="text" id="cfg_wa_url" placeholder="https://api.tudominio.com/message/sendMedia/..." />
          </div>
          <div>
            <label class="field-label">🔑 Token / API Key (WhatsApp o Telegram)</label>
            <div style="display:flex; gap:6px;">
              <input type="password" id="cfg_wa_token" placeholder="Token API" />
              <button type="button" class="btn btn-primary" style="white-space:nowrap;" onclick="guardarConfigRobot()">💾 Guardar</button>
            </div>
          </div>
        </div>
        <div id="robotCfgToast" class="status-toast"></div>
      </div>

      <div class="panel">
        <div class="panel-title">
          <span>📋 Bitácora en Tiempo Real de Archivos Clasificados y Despachados por el Robot</span>
          <button type="button" class="btn btn-outline" onclick="cargarBitacoraRobot()">🔄 Actualizar Bitácora</button>
        </div>
        <div class="table-wrap" style="max-height:360px;">
          <table>
            <thead>
              <tr>
                <th>Fecha / Hora</th>
                <th>Archivo</th>
                <th>Carpeta Clasificada (`REGLAS_CARPETAS`)</th>
                <th>Canal de Despacho</th>
                <th>Destinatario</th>
                <th>Resultado</th>
              </tr>
            </thead>
            <tbody id="robotLogTbody">
              <tr><td colspan="6" style="text-align:center; color:var(--muted);">Listo para clasificar y despachar documentos.</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 7: BOT DE AGENDA, GOOGLE CALENDAR Y BAÚL DIGITAL INTEGRADO        -->
    <!-- ======================================================================= -->
    <div id="mod_agenda" class="mod-section">
      <!-- Franja de Conexión en 1 Clic: Bot de Telegram + Google Calendar (token.json) -->
      <div class="panel" style="margin-bottom:16px; border-color:#3b82f6; background:linear-gradient(180deg, rgba(30,58,138,0.22), rgba(15,23,42,0.92));">
        <div class="panel-title">
          <span>🔌 Conexión en Tiempo Real: Bot de Agenda Telegram + Google Calendar (`token.json`) + Baúl Digital</span>
          <span class="badge badge-purple" id="gcalStatusBadge">Conectar Bot y Calendar</span>
        </div>
        <div class="form-grid-4">
          <div>
            <label class="field-label">🤖 Token Bot de Agenda Telegram (`TELEGRAM_TOKEN`)</label>
            <input type="password" id="ag_cfg_token" placeholder="Pega aquí tu TELEGRAM_TOKEN de BotFather..." />
          </div>
          <div>
            <label class="field-label">✈️ Tu Chat ID de Telegram (Notificaciones y Baúl)</label>
            <input type="text" id="ag_cfg_chat" value="1567561963" placeholder="Ej: 1567561963" />
          </div>
          <div>
            <label class="field-label">📅 Subir o Pegar `token.json` de Google Calendar</label>
            <div style="display:flex; gap:6px;">
              <input type="file" accept=".json" onchange="leerTokenCalendarArchivo(this)" style="max-width:135px; font-size:11px;" />
              <input type="text" id="ag_cfg_gcal" placeholder='Pega aquí {"token":..., "refresh_token":...}' />
            </div>
          </div>
          <div>
            <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="conectarAgendaTelegramYCalendar()">🔌 Conectar Calendar y Bot Telegram</button>
          </div>
        </div>
        <div id="agConfigToast" class="status-toast"></div>
      </div>

      <div class="kpi-grid" id="kpiAgendaGrid"></div>

      <div class="two-col-forms" style="margin-bottom:18px;">
        <div class="panel">
          <div class="panel-title">
            <span>➕ Agendar Nuevo Evento Puntual o Diario (Sincronizado con Google Calendar + Bot Telegram)</span>
            <span class="badge badge-info">Auto-detecta hora en título (ej: "4 pm")</span>
          </div>
          <div class="form-row-grid">
            <div style="grid-column: span 2;">
              <label class="field-label">📌 Título o Actividad de la Agenda *</label>
              <input type="text" id="ag_titulo" placeholder="Ej: Reunión a las 4 pm / Auditoría Cartera / Cierre Exámenes CRC La 27" />
            </div>
            <div>
              <label class="field-label">📅 Fecha (`YYYY-MM-DD`)</label>
              <input type="date" id="ag_fecha" />
            </div>
            <div>
              <label class="field-label">🕒 Hora Inicio (24h)</label>
              <input type="time" id="ag_hora" value="09:00" />
            </div>
            <div>
              <label class="field-label">⏳ Duración</label>
              <select id="ag_duracion">
                <option value="30">30 minutos</option>
                <option value="45">45 minutos</option>
                <option value="60" selected>1 hora (60 min)</option>
                <option value="120">2 horas (120 min)</option>
              </select>
            </div>
            <div>
              <label class="field-label">🔁 Frecuencia</label>
              <select id="ag_tipo">
                <option value="PUNTUAL">📅 Evento Puntual (Una vez)</option>
                <option value="DIARIO">🔁 Evento Diario Recurrente (Todos los días)</option>
              </select>
            </div>
            <div>
              <label class="field-label">👤 Responsable / Área</label>
              <input type="text" id="ag_responsable" value="Gerencia CEA Auto Oriente" />
            </div>
            <div>
              <label class="field-label">✈️ Chat ID Telegram (Recordatorio)</label>
              <input type="text" id="ag_chat_id" value="1567561963" placeholder="1567561963" />
            </div>
            <div style="grid-column: span 2;">
              <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="crearEventoAgendaWeb()">📅 Guardar en Google Calendar y Notificar por Telegram</button>
            </div>
          </div>
          <div id="agendaToast" class="status-toast"></div>
        </div>

        <div class="panel">
          <div class="panel-title">
            <span>☀️ Tu Agenda para Hoy (<span id="lblFechaHoyAgenda">Hoy</span>)</span>
            <button type="button" class="btn btn-outline" onclick="cargarAgendaWeb()">🔄 Sincronizar Calendar</button>
          </div>
          <div class="table-wrap" style="max-height:290px;">
            <table>
              <thead>
                <tr>
                  <th>Hora</th>
                  <th>Actividad</th>
                  <th>Frecuencia</th>
                  <th>Responsable</th>
                  <th>Recordatorio</th>
                </tr>
              </thead>
              <tbody id="agendaHoyTbody">
                <tr><td colspan="5" style="text-align:center; color:var(--muted);">Cargando agenda de hoy...</td></tr>
              </tbody>
            </table>
          </div>
        </div>
      </div>

      <div class="panel" style="margin-bottom:18px;">
        <div class="panel-title">
          <span>📅 Agenda de los Próximos 7 Días (Eventos Puntuales y Rutinas Diarias del CEA)</span>
          <a href="/api/bot/download_agenda_script" class="btn btn-outline">⬇️ Descargar `bot_agenda_cea.py`</a>
        </div>
        <div class="table-wrap" style="max-height:360px;">
          <table>
            <thead>
              <tr>
                <th>Fecha</th>
                <th>Hora</th>
                <th>Título / Actividad</th>
                <th>Duración</th>
                <th>Tipo</th>
                <th>Responsable</th>
                <th>Estado</th>
                <th>Acciones</th>
              </tr>
            </thead>
            <tbody id="agendaSemanaTbody"></tbody>
          </table>
        </div>
      </div>

      <!-- BAÚL DIGITAL SINCRONIZADO ENTRE BOT DE TELEGRAM Y DASHBOARD WEB -->
      <div class="panel" style="border-color:#8b5cf6;">
        <div class="panel-title">
          <span>🗄️ Baúl Digital del Bot y Dashboard (Notas, Accesos, Contabilidad, Enlaces y Apuntes Rápidos)</span>
          <span class="badge badge-purple">Sincronizado con `/baul` en Telegram</span>
        </div>
        <div class="form-grid-4" style="margin-bottom:14px;">
          <div>
            <label class="field-label">🏷️ Categoría del Baúl</label>
            <select id="baul_cat">
              <option value="NOTA / PENDIENTE">📌 Nota / Pendiente</option>
              <option value="ACCESO / LINK">🔗 Acceso / Enlace / Clave</option>
              <option value="CONTABILIDAD">💰 Contabilidad / Facturación</option>
              <option value="DOCUMENTO / RUNT">📄 Documento / RUNT / CRC</option>
            </select>
          </div>
          <div>
            <label class="field-label">📌 Título del Apunte o Recurso *</label>
            <input type="text" id="baul_titulo" placeholder="Ej: Clave Portal SISEC / Nota Reunión 4 PM" />
          </div>
          <div>
            <label class="field-label">📝 Contenido / Detalle / Enlace *</label>
            <input type="text" id="baul_contenido" placeholder="Escribe el detalle, link, cuenta o recordatorio..." />
          </div>
          <div>
            <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="guardarItemBaulWeb()">🗄️ Guardar en Baúl + Enviar a TG</button>
          </div>
        </div>
        <div id="baulToast" class="status-toast" style="margin-bottom:10px;"></div>
        <div class="table-wrap" style="max-height:320px;">
          <table>
            <thead>
              <tr>
                <th>Categoría</th>
                <th>Título</th>
                <th>Contenido / Detalle Guardado en el Baúl</th>
                <th>Fecha</th>
                <th>Origen</th>
                <th>Acciones</th>
              </tr>
            </thead>
            <tbody id="baulTbody">
              <tr><td colspan="6" style="text-align:center; color:var(--muted);">Cargando Baúl Digital...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 8: CRM WHATSAPP IA, ESTADÍSTICAS DE ABANDONO Y REMARKETING       -->
    <!-- ======================================================================= -->
    <div id="mod_whatsapp" class="mod-section">
      <!-- KPIs Estadísticos del Embudo de WhatsApp IA -->
      <div class="kpi-grid" id="kpiWhatsappGrid"></div>

      <!-- Dos paneles: 1. Plantilla y Campaña de Remarketing | 2. Registro e Importación Masiva de Chats -->
      <div class="two-col-forms" style="margin-bottom:18px;">
        <div class="panel" style="border-color:#059669;">
          <div class="panel-title">
            <span>📲 1. Plantilla Inteligente de Remarketing WhatsApp (Para Clientes en Estado `Responde la IA`)</span>
            <span class="badge badge-success">Cierre Comercial</span>
          </div>
          <div class="form-row-grid">
            <div style="grid-column: span 2;">
              <label class="field-label">🎯 Seleccionar Plantilla de Recuperación</label>
              <select id="wa_plantilla_sel" onchange="cambiarPlantillaRemarketing()">
                <option value="IA_RECUPERACION_50">🔥 Recuperación IA — Inicia hoy con el 50% (General)</option>
                <option value="MOTO_A2_RAPIDO">🏍️ Cierre Curso Moto A2 (Duración 11 días - Todo Incluido)</option>
                <option value="CARRO_B1_C1">🚗 Cierre Carro Particular B1 / Público C1 (20 días)</option>
                <option value="COMBO_MOTO_CARRO">⚡ Promoción Combo Moto A2 + Carro B1/C1</option>
              </select>
            </div>
            <div style="grid-column: span 2;">
              <label class="field-label">✍️ Mensaje de la Plantilla (Usa `{nombre}` y `{categoria}` para personalizar automáticamente)</label>
              <textarea id="wa_plantilla_txt" rows="5" style="width:100%; background:#09101d; color:#fff; border:1px solid #283a5e; border-radius:8px; padding:10px; font-size:12.5px; font-family:inherit;"></textarea>
            </div>
            <div>
              <button type="button" class="btn btn-outline" style="width:100%; justify-content:center;" onclick="guardarPlantillaRemarketingWeb()">💾 Guardar Cambios en Plantilla</button>
            </div>
            <div>
              <button type="button" class="btn btn-success" style="width:100%; justify-content:center;" onclick="lanzarSiguienteRemarketingPendiente()">🚀 Enviar a Siguiente Pendiente IA</button>
            </div>
          </div>
          <div id="waPlantillaToast" class="status-toast"></div>
        </div>

        <div class="panel">
          <div class="panel-title">
            <span>➕ 2. Capturar Prospecto de WhatsApp o Pegar Lista de Chats (`Responde la IA`)</span>
            <span class="badge badge-warning">Sincronizado con Cartera</span>
          </div>
          <div class="form-row-grid" style="margin-bottom:12px;">
            <div>
              <label class="field-label">📱 WhatsApp (+57...)</label>
              <input type="text" id="wa_tel" placeholder="Ej: 573112472122" />
            </div>
            <div>
              <label class="field-label">👤 Nombre Cliente / Perfil</label>
              <input type="text" id="wa_nombre" placeholder="Ej: Mary / YRTD / Rocío" />
            </div>
            <div>
              <label class="field-label">🏍️ Curso Consultado</label>
              <select id="wa_cat">
                <option value="A2">🏍️ Moto A2 (11 días)</option>
                <option value="B1">🚗 Carro Particular B1 (20 días)</option>
                <option value="C1">🚕 Servicio Público C1 (20 días)</option>
                <option value="COMBO A2+B1">⚡ Combo Moto A2 + Carro B1</option>
                <option value="COMBO A2+C1">🔥 Combo Moto A2 + Público C1</option>
                <option value="POR DEFINIR">❓ Preguntó Info General (Sin elegir)</option>
              </select>
            </div>
            <div>
              <label class="field-label">🤖 Estado en WhatsApp</label>
              <select id="wa_estado">
                <option value="RESPONDE_IA_PENDIENTE">🟢 Responde la IA (Quedó sin concluir)</option>
                <option value="TRANSFERENCIA_ASESOR">🔴 Transferencia de IA (Pide Asesor)</option>
                <option value="REMARKETING_ENVIADO">📲 Remarketing Enviado</option>
                <option value="MATRICULADO">✅ Matriculado (Cerrado)</option>
              </select>
            </div>
            <div>
              <label class="field-label">💬 Último Mensaje / Consulta</label>
              <input type="text" id="wa_msg" placeholder="Ej: ¡Hola! Quiero más información" />
            </div>
            <div style="display:flex; align-items:end;">
              <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="registrarLeadWhatsappWeb()">➕ Guardar Prospecto</button>
            </div>
          </div>

          <div style="border-top:1px solid var(--card-border); padding-top:10px;">
            <label class="field-label">📋 Importación Rápida en Lote: Pega aquí líneas o chats copiados de WhatsApp Web (`+57 311 2472122 ...`)</label>
            <div style="display:flex; gap:8px; margin-top:4px;">
              <textarea id="wa_lote_txt" rows="2" placeholder="Pega aquí varios teléfonos o texto copiado de WhatsApp Web (ej: +57 311 2472122 YRTD curso moto A2)..." style="flex:1; background:#09101d; color:#fff; border:1px solid #283a5e; border-radius:8px; padding:8px; font-size:12px;"></textarea>
              <button type="button" class="btn btn-outline" onclick="importarLoteWhatsappWeb()">📥 Importar Lote</button>
            </div>
          </div>
          <div id="waLeadToast" class="status-toast"></div>
        </div>
      </div>

      <!-- Tabla de Embudo de Prospectos WhatsApp IA con Filtros Rápidos -->
      <div class="panel">
        <div class="panel-title" style="flex-wrap:wrap; gap:10px;">
          <span>📊 Listado Estadístico de Consultas WhatsApp IA y Cola de Remarketing</span>
          <div style="display:flex; gap:8px; flex-wrap:wrap;">
            <select id="wa_filtro_estado" onchange="renderizarModuloWhatsapp()" style="width:auto;">
              <option value="SIN_CONCLUIR">🎯 Todos los Sin Concluir (Pendientes IA + Asesor + Remarketing)</option>
              <option value="RESPONDE_IA_PENDIENTE">🟢 Solo Responde la IA (Quedaron en visto sin concluir)</option>
              <option value="TRANSFERENCIA_ASESOR">🔴 Solo Transferencia de IA (Esperando Asesor)</option>
              <option value="REMARKETING_ENVIADO">📲 Remarketing Ya Enviado</option>
              <option value="MATRICULADO">✅ Matriculados (Cerrados)</option>
              <option value="TODOS">📚 Todos los Registros</option>
            </select>
            <button type="button" class="btn btn-outline" onclick="cargarWhatsappCrm()">🔄 Actualizar Cruce</button>
          </div>
        </div>

        <div class="search-bar">
          <input type="search" id="searchWhatsappInput" autocomplete="off" placeholder="🔎 Buscar por teléfono (+57...), nombre, categoría o mensaje..." oninput="renderizarModuloWhatsapp()" />
          <div id="waCountInfo" style="font-size:12.5px; color:var(--muted); align-self:center;"></div>
        </div>

        <div class="table-wrap" style="max-height:520px;">
          <table>
            <thead>
              <tr>
                <th>WhatsApp</th>
                <th>Prospecto / Cliente</th>
                <th>Interés</th>
                <th>Valor Est.</th>
                <th>Estado IA / Embudo</th>
                <th>Última Consulta / Respuesta IA</th>
                <th>Fecha / Hora</th>
                <th>Intentos RMK</th>
                <th>Acción / Remarketing con Plantilla</th>
              </tr>
            </thead>
            <tbody id="waLeadsTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- ======================================================================= -->
    <!-- MÓDULO 9: GESTIÓN DE RECEPCIONISTAS, EMPLEADOS Y PERMISOS (MÓDS 1 AL 5) -->
    <!-- ======================================================================= -->
    <div id="mod_empleados" class="mod-section">
      <div class="kpi-grid">
        <div class="kpi-card">
          <div class="kpi-label">👥 Usuarios / Empleados Activos</div>
          <div class="kpi-value" id="emp_kpi_total">4</div>
          <div class="kpi-sub">Con acceso personalizado por módulo</div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">👩‍💼 Perfil Recepcionistas (Módulos 1 al 5)</div>
          <div class="kpi-value" id="emp_kpi_recepcion">2</div>
          <div class="kpi-sub">Cartera, Alumnos, Profesores, Fallas y PESV</div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">👨‍🏫 Gestión Profesores / Clases (2 al 5)</div>
          <div class="kpi-value" id="emp_kpi_profesores">1</div>
          <div class="kpi-sub">Seguimiento Clases, Nómina, Inasistencias y Flota</div>
        </div>
        <div class="kpi-card warning">
          <div class="kpi-label">📋 Acciones Registradas en Bitácora</div>
          <div class="kpi-value" id="emp_kpi_logs">0</div>
          <div class="kpi-sub">Auditoría por empleado en tiempo real</div>
        </div>
      </div>

      <div class="control-bar" style="grid-template-columns: 1.25fr 1fr;">
        <!-- Formulario Crear / Editar Recepcionista o Empleado -->
        <div class="panel">
          <div class="panel-title">
            <span>➕ Crear o Actualizar Cuenta de Recepcionista / Empleado</span>
            <span class="badge badge-success">Control de Acceso por Módulos</span>
          </div>
          <div class="form-grid-4" style="margin-bottom:12px;">
            <div>
              <label class="field-label">👤 Usuario de Ingreso</label>
              <input type="text" id="emp_user" placeholder="Ej: recepcion / laura / camila" />
            </div>
            <div>
              <label class="field-label">🔑 Contraseña de Acceso</label>
              <input type="text" id="emp_pass" placeholder="Ej: Recepcion2026*" />
            </div>
            <div>
              <label class="field-label">📛 Nombre Completo Empleado(a)</label>
              <input type="text" id="emp_nombre" placeholder="Ej: Laura Gómez (Recepción)" />
            </div>
            <div>
              <label class="field-label">🏷️ Rol / Cargo Operativo</label>
              <select id="emp_rol" onchange="aplicarPlantillaRolEmpleado(this.value)">
                <option value="RECEPCIONISTA">👩‍💼 Recepcionista (Módulos 1, 2, 3, 4 y 5)</option>
                <option value="COORDINACION">👨‍🏫 Gestión Profesores y Alumnos (Módulos 2, 3, 4 y 5)</option>
                <option value="PESV">🔧 Auxiliar PESV e Inasistencias (Módulos 2, 3, 4 y 5)</option>
                <option value="SOLO_CLASES">🎓 Auxiliar Clases e Instructores (Módulos 2 y 3)</option>
                <option value="ADMIN_TOTAL">👑 Administrador General (Todos los Módulos 1 al 9)</option>
              </select>
            </div>
          </div>

          <div style="background:#09101d; border:1px solid #283a5e; border-radius:10px; padding:12px 14px; margin-bottom:12px;">
            <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px; margin-bottom:8px;">
              <label class="field-label" style="margin:0; color:#93c5fd;">✅ Selecciona qué módulos puede ver y gestionar este empleado:</label>
              <div style="display:flex; gap:6px; flex-wrap:wrap;">
                <button type="button" class="btn btn-outline" style="padding:4px 9px; font-size:11px;" onclick="marcarModulosEmpleado(['mod_cartera','mod_clases','mod_nomina','mod_inasistencias','mod_pesv'])">⚡ Marcar Módulos 1, 2, 3, 4 y 5</button>
                <button type="button" class="btn btn-outline" style="padding:4px 9px; font-size:11px;" onclick="marcarModulosEmpleado(['mod_clases','mod_nomina','mod_inasistencias','mod_pesv'])">👨‍🏫 Marcar Módulos 2, 3, 4 y 5 (Sin Cartera)</button>
                <button type="button" class="btn btn-outline" style="padding:4px 9px; font-size:11px;" onclick="marcarModulosEmpleado(['mod_cartera','mod_clases','mod_nomina','mod_inasistencias','mod_pesv','mod_robot','mod_agenda','mod_whatsapp','mod_empleados'])">👑 Todos</button>
              </div>
            </div>
            <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap:8px; font-size:12.5px;">
              <label style="display:flex; align-items:center; gap:7px; cursor:pointer;"><input type="checkbox" class="emp-mod-chk" value="mod_cartera" style="width:auto;" /> 💰 1. Cartera Financiera (Privado Gerencia)</label>
              <label style="display:flex; align-items:center; gap:7px; cursor:pointer;"><input type="checkbox" class="emp-mod-chk" value="mod_clases" checked style="width:auto;" /> 🎓 2. Seguimiento Clases y Alumnos</label>
              <label style="display:flex; align-items:center; gap:7px; cursor:pointer;"><input type="checkbox" class="emp-mod-chk" value="mod_nomina" checked style="width:auto;" /> 👨‍🏫 3. Nómina y Horarios Profesores</label>
              <label style="display:flex; align-items:center; gap:7px; cursor:pointer;"><input type="checkbox" class="emp-mod-chk" value="mod_inasistencias" checked style="width:auto;" /> 🚫 4. Control de Inasistencias</label>
              <label style="display:flex; align-items:center; gap:7px; cursor:pointer;"><input type="checkbox" class="emp-mod-chk" value="mod_pesv" checked style="width:auto;" /> 🔧 5. Mantenimientos y Flota PESV</label>
              <label style="display:flex; align-items:center; gap:7px; cursor:pointer;"><input type="checkbox" class="emp-mod-chk" value="mod_robot" style="width:auto;" /> 🤖 6. Robot Gestor y Despacho</label>
              <label style="display:flex; align-items:center; gap:7px; cursor:pointer;"><input type="checkbox" class="emp-mod-chk" value="mod_agenda" style="width:auto;" /> 📅 7. Agenda Google Calendar + Baúl</label>
              <label style="display:flex; align-items:center; gap:7px; cursor:pointer;"><input type="checkbox" class="emp-mod-chk" value="mod_whatsapp" style="width:auto;" /> 💬 8. WhatsApp IA & Remarketing</label>
              <label style="display:flex; align-items:center; gap:7px; cursor:pointer;"><input type="checkbox" class="emp-mod-chk" value="mod_empleados" style="width:auto;" /> 👥 9. Gestión de Empleados</label>
            </div>
          </div>

          <button type="button" class="btn btn-primary" style="width:100%; justify-content:center;" onclick="guardarEmpleadoWeb()">💾 Guardar Usuario y Permisos de Módulos</button>
          <div id="empToast" class="status-toast"></div>
        </div>

        <!-- Guía rápida de Cuentas Preconfiguradas -->
        <div class="panel">
          <div class="panel-title">
            <span>🔑 Accesos Directos Listos para tus Recepcionistas y Empleados</span>
            <span class="badge badge-info">Listos para compartir</span>
          </div>
          <p style="font-size:12.5px; color:#cbd5e1; margin-bottom:10px;">
            Comparte el mismo link del sistema (<code>https://bot-clases-autooriente.onrender.com</code>) con tus recepcionistas o coordinadores. Según el usuario con el que entren, el sistema les mostrará <b>únicamente los módulos autorizados</b>:
          </p>
          <div style="display:flex; flex-direction:column; gap:8px; font-size:12.5px;">
            <div style="background:#09101d; border:1px solid #283a5e; border-radius:8px; padding:10px 12px;">
              <b>👩‍💼 1. Recepción Principal (Revisa Módulos 1, 2, 3, 4 y 5):</b><br/>
              Usuario: <code>recepcion</code> &nbsp;|&nbsp; Contraseña: <code>Recepcion2026*</code>
            </div>
            <div style="background:#09101d; border:1px solid #283a5e; border-radius:8px; padding:10px 12px;">
              <b>👨‍🏫 2. Gestión de Profesores y Alumnos (Revisa Módulos 2, 3, 4 y 5):</b><br/>
              Usuario: <code>profesores</code> &nbsp;|&nbsp; Contraseña: <code>Profesores2026*</code>
            </div>
            <div style="background:#09101d; border:1px solid #283a5e; border-radius:8px; padding:10px 12px;">
              <b>🔧 3. Coordinación PESV e Inasistencias — Marly Alfonso (Módulos 1 al 5):</b><br/>
              Usuario: <code>pesv</code> &nbsp;|&nbsp; Contraseña: <code>Pesv2026*</code>
            </div>
          </div>
        </div>
      </div>

      <!-- Tabla de Usuarios y Permisos -->
      <div class="panel" style="margin-bottom:18px;">
        <div class="panel-title">
          <span>📋 Cuentas de Recepcionistas, Coordinadores y Empleados Configuradas</span>
          <button type="button" class="btn btn-outline" onclick="cargarEmpleadosWeb()">🔄 Actualizar Lista</button>
        </div>
        <div class="table-wrap" style="max-height:360px;">
          <table>
            <thead>
              <tr>
                <th>Usuario</th>
                <th>Contraseña</th>
                <th>Nombre Empleado / Recepcionista</th>
                <th>Rol / Cargo</th>
                <th>Módulos Habilitados</th>
                <th>Acciones</th>
              </tr>
            </thead>
            <tbody id="empleadosTbody"></tbody>
          </table>
        </div>
      </div>

      <!-- Bitácora de Auditoría de Acciones por Empleado -->
      <div class="panel">
        <div class="panel-title">
          <span>🕵️‍♂️ Bitácora de Gestión de Recepcionistas y Empleados (Quién hizo qué)</span>
          <span class="badge badge-purple" id="empLogsBadge">En vivo</span>
        </div>
        <div class="table-wrap" style="max-height:360px;">
          <table>
            <thead>
              <tr>
                <th>Fecha y Hora</th>
                <th>Usuario</th>
                <th>Nombre Empleado</th>
                <th>Rol</th>
                <th>Acción Realizada</th>
                <th>Detalle</th>
              </tr>
            </thead>
            <tbody id="bitacoraEmpleadosTbody"></tbody>
          </table>
        </div>
      </div>
    </div>

  </div>

  <script>
    let AUDIT_DATA = null;
    let BOT_DATA = null;
    let ACTIVE_TAB = 'todas_matriculas';
    let ACTIVE_MODULE = 'mod_cartera';
    let FILTRO_ANIO_GLOBAL = 'TODOS';
    let chartMensualInstance = null;
    let chartCarteraInstance = null;

    const fmtCOP = (val) => '$' + Math.round(Number(val || 0)).toLocaleString('es-CO');

    function obtenerAnioDeRegistro(r) {
      if (!r) return 'SIN_ANIO';
      if (r.ANIO && r.ANIO !== 'SIN_ANIO') return String(r.ANIO);
      const f = String(r.FECHAINSCRIPCION || r.FECHA_PRIMERA_INSC || r.mes || '');
      if (f.length >= 4 && /^\d{4}/.test(f)) return f.slice(0, 4);
      return 'SIN_ANIO';
    }

    function toggleManualMatriculaDrawer(forceOpen) {
      const dr = document.getElementById('manualMatriculaDrawer');
      if (!dr) return;
      if (typeof forceOpen === 'boolean') {
        dr.classList.toggle('open', forceOpen);
      } else {
        dr.classList.toggle('open');
      }
      const fIn = document.getElementById('man_fecha');
      if (fIn && !fIn.value) {
        fIn.value = new Date().toISOString().slice(0, 10);
      }
    }

    function limpiarFormularioMatriculaManual() {
      ['man_mat', 'man_doc', 'man_alumno'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
      });
      const catEl = document.getElementById('man_cat');
      if (catEl) catEl.value = 'A2';
      const convEl = document.getElementById('man_convenio');
      if (convEl) convEl.value = 'CEA AUTO ORIENTE';
      const fecEl = document.getElementById('man_fecha');
      if (fecEl) fecEl.value = new Date().toISOString().slice(0, 10);
      const vcEl = document.getElementById('man_valor_curso');
      if (vcEl) vcEl.value = '1000000';
      const abEl = document.getElementById('man_abonos');
      if (abEl) abEl.value = '0';
      const exEl = document.getElementById('man_abono_extra');
      if (exEl) exEl.value = '0';
      const estEl = document.getElementById('man_estado');
      if (estEl) estEl.value = 'ACTIVA';
    }

    function cargarMatriculaEnEditorManual(numMatricula) {
      if (!AUDIT_DATA || !AUDIT_DATA.hallazgos) return;
      const lista = AUDIT_DATA.hallazgos.todas_matriculas || [];
      const reg = lista.find(x => String(x.NUMEROMATRICULA || '') === String(numMatricula || ''));
      if (!reg) return;
      toggleManualMatriculaDrawer(true);
      document.getElementById('man_mat').value = reg.NUMEROMATRICULA || '';
      document.getElementById('man_doc').value = reg.NUMERODOCUMENTO || '';
      document.getElementById('man_alumno').value = reg.ALUMNO || '';
      document.getElementById('man_cat').value = reg.CATEGORIA || 'A2';
      document.getElementById('man_convenio').value = reg.CONVENIO_NORM || reg.CONVENIO || 'CEA AUTO ORIENTE';
      document.getElementById('man_fecha').value = reg.FECHAINSCRIPCION || new Date().toISOString().slice(0, 10);
      document.getElementById('man_valor_curso').value = Math.round(Number(reg.VALORCURSO || 0));
      document.getElementById('man_abonos').value = Math.round(Number(reg.ABONOS || 0));
      document.getElementById('man_abono_extra').value = '0';
      document.getElementById('man_estado').value = reg.ESTADO || 'ACTIVA';
      const dr = document.getElementById('manualMatriculaDrawer');
      if (dr) dr.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    function autocompletarMatriculaManualPorNumero() {
      if (!AUDIT_DATA || !AUDIT_DATA.hallazgos) return;
      const matVal = (document.getElementById('man_mat').value || '').trim();
      if (!matVal) return;
      const lista = AUDIT_DATA.hallazgos.todas_matriculas || [];
      const reg = lista.find(x => String(x.NUMEROMATRICULA || '').toLowerCase() === matVal.toLowerCase());
      if (reg) {
        document.getElementById('man_doc').value = reg.NUMERODOCUMENTO || '';
        document.getElementById('man_alumno').value = reg.ALUMNO || '';
        document.getElementById('man_cat').value = reg.CATEGORIA || 'A2';
        document.getElementById('man_convenio').value = reg.CONVENIO_NORM || reg.CONVENIO || 'CEA AUTO ORIENTE';
        document.getElementById('man_fecha').value = reg.FECHAINSCRIPCION || '';
        document.getElementById('man_valor_curso').value = Math.round(Number(reg.VALORCURSO || 0));
        document.getElementById('man_abonos').value = Math.round(Number(reg.ABONOS || 0));
        document.getElementById('man_estado').value = reg.ESTADO || 'ACTIVA';
      }
    }

    async function guardarMatriculaManualBD() {
      const payload = {
        action: 'cartera_guardar_matricula_manual',
        numero_matricula: (document.getElementById('man_mat').value || '').trim(),
        cedula: (document.getElementById('man_doc').value || '').trim(),
        alumno: (document.getElementById('man_alumno').value || '').trim(),
        categoria: document.getElementById('man_cat').value,
        convenio: (document.getElementById('man_convenio').value || 'CEA AUTO ORIENTE').trim(),
        fecha_inscripcion: document.getElementById('man_fecha').value,
        valor_curso: Number(document.getElementById('man_valor_curso').value || 0),
        abonos: Number(document.getElementById('man_abonos').value || 0),
        abono_adicional: Number(document.getElementById('man_abono_extra').value || 0),
        estado: document.getElementById('man_estado').value,
      };
      const toast = document.getElementById('manualMatriculaToast');
      if (!payload.cedula && !payload.alumno && !payload.numero_matricula) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Ingresa al menos la Cédula y el Nombre del Alumno (o el N° de Matrícula).';
        return;
      }
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Guardando y recalculando Base de Datos Histórica...';
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });
        const data = await res.json();
        if (data.ok && data.hallazgos) {
          AUDIT_DATA = data;
          actualizarVista();
          document.getElementById('man_abono_extra').value = '0';
          if (data.registro) {
            document.getElementById('man_mat').value = data.registro.NUMEROMATRICULA || '';
            document.getElementById('man_abonos').value = Math.round(Number(data.registro.ABONOS || 0));
          }
          toast.className = 'status-toast ok';
          toast.textContent = data.message || '✅ Base de Datos Histórica actualizada.';
        } else {
          toast.className = 'status-toast err';
          toast.textContent = '❌ ' + (data.message || data.error || 'No se pudo guardar.');
        }
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error de red: ' + err;
      }
    }

    function cambiarFiltroAnioGlobal(anioSel) {
      FILTRO_ANIO_GLOBAL = anioSel || 'TODOS';
      const selEl = document.getElementById('filterAnioCarteraSelect');
      if (selEl && selEl.value !== FILTRO_ANIO_GLOBAL) {
        selEl.value = FILTRO_ANIO_GLOBAL;
      }
      const c1Sel = document.getElementById('chart1ModeSelect');
      if (c1Sel) {
        if (FILTRO_ANIO_GLOBAL === '2026') c1Sel.value = 'ANIO_ACTUAL';
        else if (FILTRO_ANIO_GLOBAL === '2025') c1Sel.value = 'ANIO_2025';
        else if (FILTRO_ANIO_GLOBAL === '2024') c1Sel.value = 'ANIO_2024';
        else if (FILTRO_ANIO_GLOBAL === 'TODOS' && c1Sel.value.startsWith('ANIO_202')) c1Sel.value = 'COMPARAR_VENTAS_ANIOS';
      }
      actualizarVista();
    }

    function obtenerBusquedaLimpia(inputId) {
      const el = document.getElementById(inputId);
      if (!el) return '';
      let val = (el.value || '').trim().toLowerCase();
      const currUser = (window.__INITIAL_USER__ && window.__INITIAL_USER__.username) ? String(window.__INITIAL_USER__.username).toLowerCase() : '';
      if (['admin', 'recepcion', 'profesores', 'pesv', currUser].includes(val) && val !== '') {
        el.value = '';
        return '';
      }
      return val;
    }

    function toggleTarifas() {
      document.getElementById('tarifasDrawer').classList.toggle('open');
    }

    async function guardarTarifas() {
      const tarifas = {
        A1: Number(document.getElementById('tar_A1').value || 800000),
        A2: Number(document.getElementById('tar_A2').value || 1000000),
        B1: Number(document.getElementById('tar_B1').value || 1250000),
        C1: Number(document.getElementById('tar_C1').value || 1450000),
        C2: Number(document.getElementById('tar_C2').value || 1600000),
      };
      const res = await fetch('/api/tarifas', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ tarifas })
      });
      const data = await res.json();
      if (data.ok) {
        AUDIT_DATA = data;
        actualizarVista();
        toggleTarifas();
      }
    }

    async function cargarDatos() {
      if (window.__INITIAL_USER__ && window.__INITIAL_USER__.rol !== 'ADMIN_TOTAL') {
        const mods = window.__INITIAL_USER__.modulos || [];
        if (!mods.includes('mod_cartera')) {
          const hm = document.getElementById('headerMeta');
          if (hm) {
            hm.textContent = `CEA Auto Oriente • Panel Operativo (${window.__INITIAL_USER__.nombre} — ${window.__INITIAL_USER__.cargo || window.__INITIAL_USER__.rol})`;
          }
          return;
        }
      }
      const res = await fetch('/api/data');
      if (!res.ok) return;
      AUDIT_DATA = await res.json();
      if (AUDIT_DATA && AUDIT_DATA.ok) {
        actualizarVista();
      }
    }

    function actualizarVista() {
      if (!AUDIT_DATA || !AUDIT_DATA.hallazgos) return;
      const h = AUDIT_DATA.hallazgos;
      const k = h.kpis;
      const m = h.metadata || {};
      const listaAnual = h.anual || [];

      const badgeTxt = AUDIT_DATA.ingresos_filename
        ? `${AUDIT_DATA.filename} + ${AUDIT_DATA.ingresos_filename}`
        : (AUDIT_DATA.filename || 'matriculas23112026.xlsx');
      document.getElementById('currentFileBadge').textContent = `${badgeTxt} (${k.total_matriculas} reg.)`;

      const aniosDisponibles = listaAnual.map(a => a.anio).join(', ') || '2024-2026';
      document.getElementById('headerMeta').textContent =
        `Sede: ${m.sede || 'CEA'} • Base Histórica (${aniosDisponibles}): ${m.fecha_inicial || '2024-02-01'} al ${m.fecha_final || '2026-11-23'} • ${k.total_matriculas} Matrículas (${k.total_clientes_unicos} Clientes Únicos | ${k.clientes_multicurso_count} Multi-Curso)`;

      const contBanner = document.getElementById('contabilidadBanner');
      if (contBanner) {
        if (m.contabilidad_resumen && m.contabilidad_resumen.es_contabilidad) {
          const cr = m.contabilidad_resumen;
          contBanner.style.display = 'block';
          contBanner.innerHTML = `📑 <b>MODO AUXILIAR CONTABLE DIAN/SIIGO DETECTADO:</b> Facturación Servicios Cuenta <code>41800101</code> (${cr.total_facturas_fv} Facturas <b>FV-2</b>): <b>${fmtCOP(cr.total_credito_servicios_4180)}</b> • Devoluciones Cuenta <code>41750502</code> (${cr.total_notas_credito_nc} Notas Crédito <b>NC-2</b>): <b style="color:#fca5a5;">-${fmtCOP(cr.total_debito_devoluciones_4175)}</b> • <b>Ingreso Neto Contable: ${fmtCOP(cr.ingreso_neto_contable)}</b>`;
        } else if (m.merge_stats && (m.merge_stats.nuevos_insertados > 0 || m.merge_stats.existentes_actualizados > 0)) {
          const ms = m.merge_stats;
          contBanner.style.display = 'block';
          contBanner.innerHTML = `🗄️ <b>BASE DE DATOS HISTÓRICA ACTIVA Y SINCRONIZADA:</b> Total conservado: <b>${ms.total_despues} registros</b> (${aniosDisponibles}) • Nuevos incorporados: <b>+${ms.nuevos_insertados}</b> • Existentes actualizados: <b>${ms.existentes_actualizados}</b> • Sin pérdida de años anteriores.`;
        } else if (m.recibos_externos_cargados > 0) {
          contBanner.style.display = 'block';
          contBanner.innerHTML = `🔗 <b>CRUCE MULTIFORMATO ACTIVO:</b> Se cruzaron <b>${m.recibos_externos_cargados} comprobantes/recibos externos</b> (${m.recibos_huerfanos_cruzados} acreditados automáticamente por Cédula/NIT).`;
        } else {
          contBanner.style.display = 'block';
          contBanner.innerHTML = `🗄️ <b>BASE DE DATOS HISTÓRICA PERMANENTE ACTIVA:</b> Conservando <b>${k.total_matriculas} matrículas históricas</b> de los años <b>${aniosDisponibles}</b>. Cualquier archivo nuevo que subas incorporará modificaciones sin borrar años anteriores.`;
        }
      }

      // Renderizar Botones y Tarjetas Comparativas por Año (2024, 2025, 2026 + Histórico Total)
      const badgeAnioEl = document.getElementById('badgeAnioActivo');
      if (badgeAnioEl) {
        badgeAnioEl.textContent = FILTRO_ANIO_GLOBAL === 'TODOS'
          ? `Viendo: Histórico Completo (${aniosDisponibles})`
          : `Viendo: Solo Año ${FILTRO_ANIO_GLOBAL}`;
      }

      const btnsAnioTop = document.getElementById('botonesFiltroAnioTop');
      if (btnsAnioTop) {
        const btns = [
          `<button type="button" class="btn ${FILTRO_ANIO_GLOBAL === 'TODOS' ? 'btn-primary' : 'btn-outline'}" style="padding:5px 11px; font-size:12px;" onclick="cambiarFiltroAnioGlobal('TODOS')">🗄️ Todos los Años (${k.total_matriculas})</button>`
        ];
        listaAnual.forEach(a => {
          const isAct = FILTRO_ANIO_GLOBAL === String(a.anio);
          btns.push(
            `<button type="button" class="btn ${isAct ? 'btn-primary' : 'btn-outline'}" style="padding:5px 11px; font-size:12px;" onclick="cambiarFiltroAnioGlobal('${a.anio}')">📅 Año ${a.anio} (${a.matriculas})</button>`
          );
        });
        btnsAnioTop.innerHTML = btns.join('');
      }

      const anualGrid = document.getElementById('anualCardsGrid');
      if (anualGrid) {
        const cardsAnios = listaAnual.map(a => {
          const isSel = FILTRO_ANIO_GLOBAL === String(a.anio);
          const borderStyle = isSel ? 'border:2px solid #3b82f6; box-shadow:0 0 16px rgba(59,130,246,0.35);' : 'cursor:pointer;';
          const crecBadge = a.es_primer_anio
            ? `<span class="badge badge-info">Año Base Histórico</span>`
            : `<span class="badge ${a.crecimiento_ventas_pct >= 0 ? 'badge-success' : 'badge-danger'}">${a.crecimiento_ventas_pct >= 0 ? '▲ +' : '▼ '}${a.crecimiento_ventas_pct}% Ventas vs año ant.</span>`;
          return `
            <div class="kpi-card ${a.anio === '2026' ? 'success' : (a.anio === '2025' ? 'purple' : 'orange')}" style="${borderStyle} cursor:pointer;" onclick="cambiarFiltroAnioGlobal('${isSel ? 'TODOS' : a.anio}')" title="Haz clic para filtrar el dashboard y la tabla por el año ${a.anio}">
              <div style="display:flex; justify-content:space-between; align-items:center; gap:6px; margin-bottom:4px;">
                <span class="kpi-label" style="font-size:13px; color:#fff;">📅 AÑO ${a.anio} (${a.matriculas} Mat.)</span>
                ${crecBadge}
              </div>
              <div style="font-size:12.5px; color:#cbd5e1; line-height:1.65; margin-top:4px;">
                <div>💰 <b>Ventas Facturadas:</b> <span style="color:#93c5fd; font-weight:700;">${fmtCOP(a.facturado)}</span></div>
                <div>✅ <b>Ingresos Recaudados:</b> <span style="color:#34d399; font-weight:700;">${fmtCOP(a.recaudado)}</span> <span style="font-size:11px; color:var(--muted);">(${a.tasa_recaudo_pct}%)</span></div>
                <div>⚠️ <b>Cartera que Quedó Pendiente:</b> <span style="color:#f87171; font-weight:800;">${fmtCOP(a.pendiente)}</span> <span style="font-size:11px; color:#fca5a5;">(${a.con_saldo} mat.)</span></div>
              </div>
            </div>
          `;
        });
        const isAllSel = FILTRO_ANIO_GLOBAL === 'TODOS';
        cardsAnios.push(`
          <div class="kpi-card" style="${isAllSel ? 'border:2px solid #3b82f6; box-shadow:0 0 16px rgba(59,130,246,0.35);' : ''} cursor:pointer;" onclick="cambiarFiltroAnioGlobal('TODOS')" title="Ver consolidado histórico de todos los años">
            <div style="display:flex; justify-content:space-between; align-items:center; gap:6px; margin-bottom:4px;">
              <span class="kpi-label" style="font-size:13px; color:#fff;">🗄️ TOTAL HISTÓRICO (${aniosDisponibles})</span>
              <span class="badge badge-info">${k.total_matriculas} Matrículas</span>
            </div>
            <div style="font-size:12.5px; color:#cbd5e1; line-height:1.65; margin-top:4px;">
              <div>💰 <b>Ventas Totales:</b> <span style="color:#93c5fd; font-weight:700;">${fmtCOP(k.total_facturado)}</span></div>
              <div>✅ <b>Ingresos Totales:</b> <span style="color:#34d399; font-weight:700;">${fmtCOP(k.total_abonos)}</span> <span style="font-size:11px; color:var(--muted);">(${k.tasa_recaudo_pct}%)</span></div>
              <div>⚠️ <b>Cartera Total Acumulada:</b> <span style="color:#f87171; font-weight:800;">${fmtCOP(k.total_pendiente)}</span> <span style="font-size:11px; color:#fca5a5;">(${k.alumnos_con_saldo} mat.)</span></div>
            </div>
          </div>
        `);
        anualGrid.innerHTML = cardsAnios.join('');
      }

      const matchAnio = (r) => FILTRO_ANIO_GLOBAL === 'TODOS' || obtenerAnioDeRegistro(r) === FILTRO_ANIO_GLOBAL;
      const todasFiltradasAnio = (h.todas_matriculas || []).filter(matchAnio);
      const jtFiltradasAnio = (h.jt_cartera || []).filter(matchAnio);
      const ceroRealFiltradasAnio = (h.abono_cero_real || []).filter(matchAnio);
      const ceroMcFiltradasAnio = (h.abono_cero_multicurso || []).filter(matchAnio);
      const mcFiltradasAnio = (h.clientes_multicurso || []).filter(c => FILTRO_ANIO_GLOBAL === 'TODOS' || (c.CURSOS_DETALLE || []).some(cd => String(cd.FECHAINSCRIPCION || '').startsWith(FILTRO_ANIO_GLOBAL)));
      const mora90FiltradasAnio = (h.cartera_critica_90 || []).filter(matchAnio);
      const mora30FiltradasAnio = [...(h.mora_media_60 || []), ...(h.mora_temprana_30 || [])].filter(matchAnio);
      const errFiltradasAnio = (h.errores_digitacion || []).filter(matchAnio);

      // Contadores de pestañas sincronizados con el filtro de año activo
      const cntAllEl = document.getElementById('cnt_all');
      if (cntAllEl) cntAllEl.textContent = todasFiltradasAnio.length;
      const cntAnualEl = document.getElementById('cnt_anual');
      if (cntAnualEl) cntAnualEl.textContent = listaAnual.length;
      document.getElementById('cnt_jt').textContent = jtFiltradasAnio.length;
      document.getElementById('cnt_cero_real').textContent = ceroRealFiltradasAnio.length;
      document.getElementById('cnt_cero_mc').textContent = ceroMcFiltradasAnio.length;
      document.getElementById('cnt_mc_all').textContent = mcFiltradasAnio.length;
      document.getElementById('cnt_mora90').textContent = mora90FiltradasAnio.length;
      document.getElementById('cnt_mora30').textContent = mora30FiltradasAnio.length;
      document.getElementById('cnt_err').textContent = errFiltradasAnio.length;

      // Calcular KPIs dinámicos según el año seleccionado (o Total Histórico si TODOS)
      const factKpi = todasFiltradasAnio.reduce((acc, r) => acc + Number(r.VALORCURSO || 0), 0);
      const recKpi = todasFiltradasAnio.reduce((acc, r) => acc + Number(r.ABONOS || 0), 0);
      const pendKpi = todasFiltradasAnio.reduce((acc, r) => acc + Math.max(0, Number(r.VALORPENDIENTE || 0)), 0);
      const conSaldoKpi = todasFiltradasAnio.filter(r => Number(r.VALORPENDIENTE || 0) > 0).length;
      const pazSalvoKpi = todasFiltradasAnio.length - conSaldoKpi;
      const tasaRecKpi = factKpi > 0 ? ((recKpi / factKpi) * 100).toFixed(2) : '0.00';
      const tasaMoraKpi = factKpi > 0 ? ((pendKpi / factKpi) * 100).toFixed(2) : '0.00';
      const cedulasUnicasKpi = new Set(todasFiltradasAnio.map(r => r.NUMERODOCUMENTO || r.NUMEROMATRICULA)).size;
      const montoCeroRealKpi = ceroRealFiltradasAnio.reduce((acc, r) => acc + Math.max(0, Number(r.VALORPENDIENTE || 0)), 0);
      const clientesCeroRealKpi = new Set(ceroRealFiltradasAnio.map(r => r.NUMERODOCUMENTO)).size;
      const jtTotalMatAnio = (h.jt_todos || []).filter(matchAnio).length;
      const jtPendKpi = jtFiltradasAnio.reduce((acc, r) => acc + Math.max(0, Number(r.VALORPENDIENTE || 0)), 0);
      const jtCeroRealKpi = jtFiltradasAnio.filter(r => r.TIPO_RIESGO_CERO === 'RIESGO_CERO_REAL').length;
      const tagPeriodoKpi = FILTRO_ANIO_GLOBAL === 'TODOS' ? 'Histórico' : `Año ${FILTRO_ANIO_GLOBAL}`;

      const kpisHtml = `
        <div class="kpi-card">
          <div class="kpi-label">Total Facturado (${tagPeriodoKpi})</div>
          <div class="kpi-value">${fmtCOP(factKpi)}</div>
          <div class="kpi-sub"><b>${todasFiltradasAnio.length} matrículas</b> (${cedulasUnicasKpi} cédulas únicas)</div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">Ingresos / Recaudo (${tagPeriodoKpi})</div>
          <div class="kpi-value" style="color:#34d399;">${fmtCOP(recKpi)}</div>
          <div class="kpi-sub">Tasa Recaudo: <b>${tasaRecKpi}%</b> (${pazSalvoKpi} al día)</div>
        </div>
        <div class="kpi-card danger">
          <div class="kpi-label">Cartera Pendiente (${tagPeriodoKpi})</div>
          <div class="kpi-value" style="color:#f87171;">${fmtCOP(pendKpi)}</div>
          <div class="kpi-sub"><b>${conSaldoKpi} matrículas</b> con saldo pendiente (${tasaMoraKpi}%)</div>
        </div>
        <div class="kpi-card danger">
          <div class="kpi-label">🚨 Riesgo $0 REAL (${tagPeriodoKpi})</div>
          <div class="kpi-value" style="color:#fca5a5;">${ceroRealFiltradasAnio.length} Mat. (${clientesCeroRealKpi} Clientes)</div>
          <div class="kpi-sub">$0 en todos sus cursos: <b>${fmtCOP(montoCeroRealKpi)}</b></div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">🔄 Clientes Multi-Curso (${tagPeriodoKpi})</div>
          <div class="kpi-value" style="color:#c4b5fd;">${mcFiltradasAnio.length} Clientes</div>
          <div class="kpi-sub"><b>${ceroMcFiltradasAnio.length} falsos $0</b> (abonaron en su 1ra matrícula)</div>
        </div>
        <div class="kpi-card orange">
          <div class="kpi-label">💛 Cartera Javier Torres (${tagPeriodoKpi})</div>
          <div class="kpi-value" style="color:#fdba74;">${fmtCOP(jtPendKpi)}</div>
          <div class="kpi-sub"><b>${jtFiltradasAnio.length} mat.</b> con saldo de ${jtTotalMatAnio} (${jtCeroRealKpi} en $0 real)</div>
        </div>
      `;
      document.getElementById('kpiGrid').innerHTML = kpisHtml;

      renderizarGraficas(h);
      renderizarTablaActiva();
    }

    function renderizarGraficas(h) {
      if (!h) return;
      const mode1 = (document.getElementById('chart1ModeSelect') || {}).value || 'ANIO_ACTUAL';
      const mode2 = (document.getElementById('chart2ModeSelect') || {}).value || 'CRECIMIENTO_ANUAL_DINERO';
      const mesesValidos = (h.mensual || []).filter(x => x.mes !== 'SIN_FECHA');
      const listaAnual = h.anual || [];

      if (chartMensualInstance) chartMensualInstance.destroy();
      const ctx1 = document.getElementById('chartMensual').getContext('2d');
      const NOMBRES_MESES_CORTOS = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
      const COLORES_ANIO = {
        '2024': { bg: 'rgba(245, 158, 11, 0.78)', border: '#f59e0b' },
        '2025': { bg: 'rgba(139, 92, 246, 0.82)', border: '#8b5cf6' },
        '2026': { bg: 'rgba(16, 185, 129, 0.85)', border: '#10b981' },
      };

      if (mode1.startsWith('COMPARAR_')) {
        // Comparativa Mes a Mes (Enero a Diciembre) entre Años (2024 vs 2025 vs 2026)
        const aniosUnicos = [...new Set(mesesValidos.map(m => (m.mes || '').slice(0, 4)).filter(Boolean))].sort();
        const esVentas = mode1 === 'COMPARAR_VENTAS_ANIOS';
        const esCartera = mode1 === 'COMPARAR_CARTERA_ANIOS';
        const datasetsComp = aniosUnicos.map(anio => {
          const col = COLORES_ANIO[anio] || { bg: 'rgba(59, 130, 246, 0.8)', border: '#3b82f6' };
          const valores12 = [];
          for (let mIdx = 1; mIdx <= 12; mIdx++) {
            const mm = String(mIdx).padStart(2, '0');
            const found = mesesValidos.find(x => x.mes === `${anio}-${mm}`);
            if (!found) {
              valores12.push(0);
            } else if (esVentas) {
              valores12.push(Math.round((found.facturado || 0) / 1000000 * 10) / 10);
            } else if (esCartera) {
              valores12.push(Math.round(Math.max(0, found.pendiente || 0) / 1000000 * 10) / 10);
            } else {
              valores12.push(found.total_matriculas || 0);
            }
          }
          return {
            label: esVentas ? `Ventas ${anio} ($ Millones)` : (esCartera ? `Cartera Pendiente ${anio} ($M)` : `Matrículas ${anio}`),
            data: valores12,
            backgroundColor: col.bg,
            borderColor: col.border,
            borderWidth: 2,
            borderRadius: 4,
            tension: 0.25,
          };
        });

        const titleEl = document.getElementById('chart1Title');
        if (titleEl) {
          titleEl.textContent = esVentas
            ? 'Gráfico 1: Comparativa de Ventas Facturadas Entre Años (Enero a Diciembre: 2024 vs 2025 vs 2026)'
            : (esCartera
                ? 'Gráfico 1: Comparativa de Cartera Pendiente Entre Años (2024 vs 2025 vs 2026)'
                : 'Gráfico 1: Comparativa de N° Matrículas Mensuales Entre Años (2024 vs 2025 vs 2026)');
        }

        chartMensualInstance = new Chart(ctx1, {
          type: 'bar',
          data: { labels: NOMBRES_MESES_CORTOS, datasets: datasetsComp },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#cbd5e1', font: { size: 11.5 } } } },
            scales: {
              x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.04)' } },
              y: {
                position: 'left',
                beginAtZero: true,
                ticks: {
                  color: '#93c5fd',
                  callback: v => (esVentas || esCartera) ? ('$' + v + 'M') : v
                },
                grid: { color: 'rgba(255,255,255,0.06)' },
                title: {
                  display: true,
                  text: (esVentas || esCartera) ? 'Millones COP ($M)' : 'N° Matrículas',
                  color: '#93c5fd'
                }
              }
            }
          }
        });
      } else {
        // Vista de un año específico (2026, 2025, 2024) o línea de tiempo completa
        let anioTarget = '';
        if (mode1 === 'ANIO_ACTUAL') anioTarget = '2026';
        else if (mode1 === 'ANIO_2025') anioTarget = '2025';
        else if (mode1 === 'ANIO_2024') anioTarget = '2024';

        let mesesGraf = anioTarget
          ? mesesValidos.filter(x => String(x.mes || '').startsWith(anioTarget))
          : mesesValidos;
        if (mesesGraf.length === 0) mesesGraf = mesesValidos;

        const titleEl = document.getElementById('chart1Title');
        if (titleEl) {
          titleEl.textContent = anioTarget
            ? `Gráfico 1: Ventas, Matrículas y Cartera Pendiente por Mes — Año ${anioTarget}`
            : 'Gráfico 1: Evolución Histórica Mensual (2024 a 2026)';
        }

        const labels = mesesGraf.map(x => x.nombre_mes);
        const dataMat = mesesGraf.map(x => x.total_matriculas);
        const dataJT = mesesGraf.map(x => x.matriculas_jt);
        const dataCarteraMillones = mesesGraf.map(x => Math.round(Math.max(0, x.pendiente || 0) / 1000000 * 10) / 10);

        chartMensualInstance = new Chart(ctx1, {
          type: 'bar',
          data: {
            labels,
            datasets: [
              {
                label: 'Matrículas Totales Mes (Meta: 100)',
                data: dataMat,
                backgroundColor: 'rgba(59, 130, 246, 0.82)',
                borderColor: '#3b82f6',
                borderWidth: 1,
                borderRadius: 5,
                yAxisID: 'y',
              },
              {
                label: 'Matrículas Javier Torres (Meta: 30)',
                data: dataJT,
                backgroundColor: 'rgba(139, 92, 246, 0.85)',
                borderColor: '#8b5cf6',
                borderWidth: 1,
                borderRadius: 5,
                yAxisID: 'y',
              },
              {
                label: 'Cartera Pendiente Real Mes ($ Millones COP)',
                data: dataCarteraMillones,
                type: 'line',
                borderColor: '#ff7f0e',
                backgroundColor: 'rgba(255, 127, 14, 0.2)',
                borderWidth: 2.5,
                pointRadius: 4,
                tension: 0.3,
                yAxisID: 'y1',
              }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#cbd5e1', font: { size: 11.5 } } } },
            scales: {
              x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.04)' } },
              y: { position: 'left', beginAtZero: true, ticks: { color: '#93c5fd' }, grid: { color: 'rgba(255,255,255,0.06)' }, title: { display: true, text: 'N° Matrículas', color: '#93c5fd' } },
              y1: { position: 'right', beginAtZero: true, ticks: { color: '#fdba74', callback: v => '$' + v + 'M' }, grid: { drawOnChartArea: false }, title: { display: true, text: 'Cartera Pendiente ($M)', color: '#fdba74' } }
            }
          }
        });
      }

      // GRÁFICO 2: Crecimiento Anual por Años (Ventas vs Ingresos vs Cartera) o Antigüedad de Mora
      if (chartCarteraInstance) chartCarteraInstance.destroy();
      const ctx2 = document.getElementById('chartCartera').getContext('2d');
      const title2El = document.getElementById('chart2Title');

      if (mode2 === 'CRECIMIENTO_ANUAL_DINERO' && listaAnual.length > 0) {
        if (title2El) title2El.textContent = 'Gráfico 2: Crecimiento Anual de Ventas, Ingresos y Cartera por Año';
        const labelsAnios = listaAnual.map(a => `Año ${a.anio}`);
        const factMill = listaAnual.map(a => Math.round((a.facturado || 0) / 1000000 * 10) / 10);
        const recMill = listaAnual.map(a => Math.round((a.recaudado || 0) / 1000000 * 10) / 10);
        const pendMill = listaAnual.map(a => Math.round(Math.max(0, a.pendiente || 0) / 1000000 * 10) / 10);

        chartCarteraInstance = new Chart(ctx2, {
          type: 'bar',
          data: {
            labels: labelsAnios,
            datasets: [
              {
                label: 'Ventas Facturadas ($M)',
                data: factMill,
                backgroundColor: 'rgba(59, 130, 246, 0.85)',
                borderColor: '#3b82f6',
                borderWidth: 1,
                borderRadius: 6,
              },
              {
                label: 'Ingresos Recaudados ($M)',
                data: recMill,
                backgroundColor: 'rgba(16, 185, 129, 0.85)',
                borderColor: '#10b981',
                borderWidth: 1,
                borderRadius: 6,
              },
              {
                label: 'Cartera Pendiente ($M)',
                data: pendMill,
                backgroundColor: 'rgba(239, 68, 68, 0.88)',
                borderColor: '#ef4444',
                borderWidth: 1,
                borderRadius: 6,
              }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { color: '#cbd5e1', font: { size: 11 } } } },
            scales: {
              x: { ticks: { color: '#e2e8f0', font: { weight: 'bold' } }, grid: { color: 'rgba(255,255,255,0.04)' } },
              y: {
                beginAtZero: true,
                ticks: { color: '#93c5fd', callback: v => '$' + v + 'M' },
                grid: { color: 'rgba(255,255,255,0.06)' },
                title: { display: true, text: 'Millones COP ($M)', color: '#93c5fd' }
              }
            }
          }
        });
      } else if (mode2 === 'CRECIMIENTO_ANUAL_MATRICULAS' && listaAnual.length > 0) {
        if (title2El) title2El.textContent = 'Gráfico 2: Crecimiento de Matrículas y Clientes por Año';
        chartCarteraInstance = new Chart(ctx2, {
          type: 'bar',
          data: {
            labels: listaAnual.map(a => `Año ${a.anio}`),
            datasets: [
              {
                label: 'Total Matrículas del Año',
                data: listaAnual.map(a => a.matriculas),
                backgroundColor: 'rgba(139, 92, 246, 0.85)',
                borderColor: '#8b5cf6',
                borderWidth: 1,
                borderRadius: 6,
              },
              {
                label: 'Clientes Únicos (Cédulas)',
                data: listaAnual.map(a => a.clientes_unicos),
                backgroundColor: 'rgba(59, 130, 246, 0.82)',
                borderColor: '#3b82f6',
                borderWidth: 1,
                borderRadius: 6,
              },
              {
                label: 'Matrículas con Saldo Pendiente',
                data: listaAnual.map(a => a.con_saldo),
                backgroundColor: 'rgba(245, 158, 11, 0.88)',
                borderColor: '#f59e0b',
                borderWidth: 1,
                borderRadius: 6,
              }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { color: '#cbd5e1', font: { size: 11 } } } },
            scales: {
              x: { ticks: { color: '#e2e8f0', font: { weight: 'bold' } }, grid: { color: 'rgba(255,255,255,0.04)' } },
              y: { beginAtZero: true, ticks: { color: '#c4b5fd' }, grid: { color: 'rgba(255,255,255,0.06)' } }
            }
          }
        });
      } else {
        if (title2El) title2El.textContent = 'Gráfico 2: Distribución de la Cartera por Antigüedad de Mora';
        const k = h.kpis;
        chartCarteraInstance = new Chart(ctx2, {
          type: 'doughnut',
          data: {
            labels: [
              `Crítica > 90 días (${k.mora_90_count})`,
              `Mora Media 61-90 días (${k.mora_60_count})`,
              `Mora Temprana 31-60 días (${k.mora_30_count})`,
              `Reciente <= 30 días (${k.mora_aldia_count})`
            ],
            datasets: [{
              data: [k.mora_90_monto, k.mora_60_monto, k.mora_30_monto, k.mora_aldia_monto],
              backgroundColor: ['#ef4444', '#ff7f0e', '#eab308', '#3b82f6'],
              borderWidth: 2,
              borderColor: '#131d31'
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '62%',
            plugins: {
              legend: { position: 'bottom', labels: { color: '#cbd5e1', font: { size: 11 } } }
            }
          }
        });
      }
    }

    function cambiarTab(tabKey, btnEl) {
      ACTIVE_TAB = tabKey;
      document.querySelectorAll('#tabsHeader .tab-btn').forEach(b => b.classList.remove('active'));
      if (btnEl) btnEl.classList.add('active');
      obtenerBusquedaLimpia('searchInput');
      renderizarTablaActiva();
    }

    function toggleExpander(rowId) {
      const el = document.getElementById(rowId);
      if (el) el.classList.toggle('open');
    }

    function renderizarTablaActiva() {
      if (!AUDIT_DATA || !AUDIT_DATA.hallazgos) return;
      const h = AUDIT_DATA.hallazgos;
      const q = obtenerBusquedaLimpia('searchInput');

      const thead = document.getElementById('auditThead');
      const tbody = document.getElementById('auditTbody');

      // Pestaña de Resumen y Crecimiento por Años (2024, 2025, 2026)
      if (ACTIVE_TAB === 'anual') {
        thead.innerHTML = `<tr>
          <th>Año</th><th>Matrículas Totales</th><th>Crecimiento Matrículas</th><th>Clientes Únicos</th>
          <th>Total Facturado (Ventas)</th><th>Crecimiento Ventas</th><th>Ingresos / Recaudo Efectivo</th>
          <th>% Recaudo</th><th>Cartera Pendiente que Quedó</th><th>Matrículas con Saldo</th><th>Riesgo $0 Real</th><th>Acción</th>
        </tr>`;
        const rowsAnual = (h.anual || []).filter(a => !q || String(a.anio).includes(q));
        document.getElementById('tableCountInfo').textContent = `Mostrando ${rowsAnual.length} años históricos en la Base de Datos`;
        tbody.innerHTML = rowsAnual.map(a => {
          const crecV = a.es_primer_anio
            ? '<span class="badge badge-info">Año Base</span>'
            : `<span class="badge ${a.crecimiento_ventas_pct >= 0 ? 'badge-success' : 'badge-danger'}">${a.crecimiento_ventas_pct >= 0 ? '▲ +' : '▼ '}${a.crecimiento_ventas_pct}%</span>`;
          const crecM = a.es_primer_anio
            ? '<span class="badge badge-info">Año Base</span>'
            : `<span class="badge ${a.crecimiento_matriculas_pct >= 0 ? 'badge-success' : 'badge-danger'}">${a.crecimiento_matriculas_pct >= 0 ? '▲ +' : '▼ '}${a.crecimiento_matriculas_pct}%</span>`;
          return `<tr>
            <td><b style="font-size:14px; color:#93c5fd;">📅 Año ${a.anio}</b></td>
            <td><b>${a.matriculas}</b></td>
            <td>${crecM}</td>
            <td>${a.clientes_unicos}</td>
            <td><b>${fmtCOP(a.facturado)}</b></td>
            <td>${crecV}</td>
            <td class="money-pos"><b>${fmtCOP(a.recaudado)}</b></td>
            <td><span class="badge badge-success">${a.tasa_recaudo_pct}%</span></td>
            <td class="${a.pendiente > 0 ? 'money-neg' : 'money-pos'}"><b>${fmtCOP(a.pendiente)}</b></td>
            <td><b>${a.con_saldo}</b> mat. (${a.tasa_morosidad_pct}%)</td>
            <td><span class="badge badge-danger">${a.abono_cero_real}</span></td>
            <td><button type="button" class="btn btn-outline" style="padding:4px 9px; font-size:11.5px;" onclick="cambiarFiltroAnioGlobal('${a.anio}'); cambiarTab('todas_matriculas', document.getElementById('tabBtnTodas'));">🔎 Ver Registros ${a.anio}</button></td>
          </tr>`;
        }).join('');
        return;
      }

      // Pestaña de Cumplimiento Mensual
      if (ACTIVE_TAB === 'mensual') {
        thead.innerHTML = `<tr>
          <th>Mes</th><th>Matrículas CEA</th><th>Meta CEA</th><th>Estado CEA</th>
          <th>Matrículas Javier Torres</th><th>Meta JT</th><th>Estado JT</th>
          <th>Total Facturado</th><th>Total Recaudado</th><th>Cartera Pendiente</th><th>Riesgo $0 Real</th>
        </tr>`;
        const rows = h.mensual.filter(m => {
          if (FILTRO_ANIO_GLOBAL !== 'TODOS' && !String(m.mes || '').startsWith(FILTRO_ANIO_GLOBAL)) return false;
          return !q || m.nombre_mes.toLowerCase().includes(q);
        });
        document.getElementById('tableCountInfo').textContent = `Mostrando ${rows.length} periodos mensuales`;
        tbody.innerHTML = rows.map(m => `<tr>
          <td><b>${m.nombre_mes}</b></td>
          <td><b>${m.total_matriculas}</b></td>
          <td>${m.meta_empresa}</td>
          <td><span class="badge ${m.estado_meta_empresa === 'CUMPLIDA' ? 'badge-success' : 'badge-warning'}">${m.cumplimiento_empresa_pct}% (${m.estado_meta_empresa})</span></td>
          <td><b>${m.matriculas_jt}</b></td>
          <td>${m.meta_jt}</td>
          <td><span class="badge ${m.estado_meta_jt === 'CUMPLIDA' ? 'badge-success' : 'badge-warning'}">${m.cumplimiento_jt_pct}% (${m.estado_meta_jt})</span></td>
          <td>${fmtCOP(m.facturado)}</td>
          <td class="money-pos">${fmtCOP(m.recaudado)}</td>
          <td class="${m.pendiente > 0 ? 'money-neg' : ''}">${fmtCOP(m.pendiente)}</td>
          <td><span class="badge badge-danger">${m.abono_cero_real_count}</span></td>
        </tr>`).join('');
        return;
      }

      // Pestaña de Consolidado por Convenios
      if (ACTIVE_TAB === 'convenios') {
        thead.innerHTML = `<tr>
          <th>Convenio / Canal</th><th>Matrículas</th><th>Clientes Únicos</th><th>Total Facturado</th>
          <th>Recaudado en Caja</th><th>% Recaudo</th><th>Cartera Pendiente</th><th>% Morosidad</th>
          <th>Deudores Activos</th><th>Riesgo $0 Real</th>
        </tr>`;
        const rows = h.convenios.filter(c => !q || c.convenio.toLowerCase().includes(q));
        document.getElementById('tableCountInfo').textContent = `Mostrando ${rows.length} convenios activos`;
        tbody.innerHTML = rows.map(c => `<tr>
          <td><b>${c.convenio}</b></td>
          <td><b>${c.matriculas}</b></td>
          <td>${c.clientes_unicos}</td>
          <td>${fmtCOP(c.facturado)}</td>
          <td class="money-pos">${fmtCOP(c.recaudado)}</td>
          <td><span class="badge badge-success">${c.tasa_recaudo}%</span></td>
          <td class="${c.pendiente > 0 ? 'money-neg' : ''}">${fmtCOP(c.pendiente)}</td>
          <td><span class="badge ${c.tasa_morosidad > 10 ? 'badge-danger' : 'badge-warning'}">${c.tasa_morosidad}%</span></td>
          <td>${c.con_saldo}</td>
          <td><span class="badge badge-danger">${c.abono_cero_real}</span></td>
        </tr>`).join('');
        return;
      }

      const sortMode = (document.getElementById('sortCarteraSelect') || {}).value || 'RECIENTES';
      const filterAbono = (document.getElementById('filterAbonoCarteraSelect') || {}).value || 'TODOS';

      // Pestaña de Clientes Multi-Curso Consolidados por Cédula
      if (ACTIVE_TAB === 'clientes_multicurso') {
        let rows = h.clientes_multicurso.filter(c => {
          if (FILTRO_ANIO_GLOBAL !== 'TODOS') {
            const tieneCursoEnAnio = (c.CURSOS_DETALLE || []).some(cd => String(cd.FECHAINSCRIPCION || '').startsWith(FILTRO_ANIO_GLOBAL));
            if (!tieneCursoEnAnio) return false;
          }
          const totAb = Number(c.TOTAL_ABONADO || 0);
          const salCl = Number(c.SALDO_REAL_CLIENTE || 0);
          if (filterAbono === 'SOLO_CON_SALDO' && salCl <= 0) return false;
          if (filterAbono === 'SOLO_CON_ABONO' && totAb <= 0) return false;
          if (filterAbono === 'SOLO_CERO_GLOBAL' && totAb > 0) return false;
          if (!q) return true;
          return `${c.NUMERODOCUMENTO} ${c.ALUMNO} ${c.CONVENIO_NORM} ${c.CATEGORIAS} ${c.MATRICULAS}`.toLowerCase().includes(q);
        });
        rows = [...rows].sort((a, b) => {
          const fa = String(a.FECHA_PRIMERA_INSC || '');
          const fb = String(b.FECHA_PRIMERA_INSC || '');
          if (sortMode === 'ANTIGUOS') return fa.localeCompare(fb);
          if (sortMode === 'CON_ABONO_PRIMERO') return Number(b.TOTAL_ABONADO || 0) - Number(a.TOTAL_ABONADO || 0) || fb.localeCompare(fa);
          if (sortMode === 'SIN_ABONO_PRIMERO') return Number(a.TOTAL_ABONADO || 0) - Number(b.TOTAL_ABONADO || 0) || fb.localeCompare(fa);
          if (sortMode === 'MAYOR_SALDO') return Number(b.SALDO_REAL_CLIENTE || 0) - Number(a.SALDO_REAL_CLIENTE || 0);
          return fb.localeCompare(fa);
        });
        const sumaSaldo = rows.reduce((a, b) => a + Number(b.SALDO_REAL_CLIENTE || 0), 0);
        document.getElementById('tableCountInfo').textContent =
          `Mostrando ${rows.length} Clientes Multi-Curso (2+ matrículas) | Saldo Real Consolidado: ${fmtCOP(sumaSaldo)} (Haz clic en una fila para ver el desglose)`;

        thead.innerHTML = `<tr>
          <th>Cédula / Doc</th><th>Alumno (Cliente)</th><th>Convenio</th><th>N° Cursos</th>
          <th>Combo Categorías</th><th>Matrículas Asociadas</th><th>Total Contratado (Cliente)</th>
          <th>Total Abonado (Caja Global)</th><th>Saldo Real Cliente</th><th>Antigüedad Máx.</th>
        </tr>`;

        tbody.innerHTML = rows.map((c, idx) => {
          const expId = `exp_mc_${idx}`;
          const cursosHtml = (c.CURSOS_DETALLE || []).map(cd => `
            <tr>
              <td><b>${cd.NUMEROMATRICULA}</b></td>
              <td><span class="badge">${cd.CATEGORIA}</span></td>
              <td>${cd.FECHAINSCRIPCION}</td>
              <td>${fmtCOP(cd.VALORCURSO)}</td>
              <td class="money-pos">${fmtCOP(cd.ABONOS_ORIGINAL)}</td>
              <td class="${cd.VALORPENDIENTE_ORIGINAL > 0 ? 'money-neg' : ''}">${fmtCOP(cd.VALORPENDIENTE_ORIGINAL)}</td>
              <td>${cd.ESTADO}</td>
            </tr>
          `).join('');

          const recibosHtml = (c.RECIBOS_DETALLE || []).map(rc => `
            <tr>
              <td><b>${rc.RECIBO}</b></td>
              <td>${rc.FECHA}</td>
              <td>${rc.NUMEROMATRICULA || '<span class="badge badge-warning">HUÉRFANO (Por Cédula)</span>'}</td>
              <td class="money-pos">${fmtCOP(rc.VALOR)}</td>
              <td>${rc.TIPOPAGO}</td>
            </tr>
          `).join('') || '<tr><td colspan="5" style="color:var(--muted);">Sin recibos registrados</td></tr>';

          return `
            <tr class="data-row" onclick="toggleExpander('${expId}')">
              <td><code>${c.NUMERODOCUMENTO}</code></td>
              <td><b>🔍 ${c.ALUMNO}</b></td>
              <td>${c.CONVENIO_NORM}</td>
              <td><span class="badge badge-purple">${c.CANTIDAD_CURSOS} cursos</span></td>
              <td><b>${c.CATEGORIAS}</b></td>
              <td>${c.MATRICULAS}</td>
              <td>${fmtCOP(c.TOTAL_CONTRATADO)}</td>
              <td class="money-pos">${fmtCOP(c.TOTAL_ABONADO)}</td>
              <td class="${c.SALDO_REAL_CLIENTE > 0 ? 'money-neg' : 'money-pos'}">${fmtCOP(c.SALDO_REAL_CLIENTE)}</td>
              <td><span class="badge ${c.DIAS_ANTIGUEDAD_MAX > 90 ? 'badge-danger' : 'badge-warning'}">${c.DIAS_ANTIGUEDAD_MAX} días</span></td>
            </tr>
            <tr id="${expId}" class="expander-row">
              <td colspan="10">
                <div class="expander-box">
                  <div>
                    <b style="color:#93c5fd;">📚 Matrículas Simultáneas de ${c.ALUMNO} (Cédula: ${c.NUMERODOCUMENTO})</b>
                    <table class="mini-table">
                      <thead><tr><th>Matrícula</th><th>Cat.</th><th>Fecha Insc.</th><th>Valor Curso</th><th>Abono Reg.</th><th>Saldo Reg.</th><th>Estado</th></tr></thead>
                      <tbody>${cursosHtml}</tbody>
                    </table>
                  </div>
                  <div>
                    <b style="color:#6ee7b7;">🧾 Historial de Recibos de Caja y Abonos Aplicados a la Cédula</b>
                    <table class="mini-table">
                      <thead><tr><th>Recibo</th><th>Fecha</th><th>Matrícula</th><th>Valor Pagado</th><th>Medio / Concepto</th></tr></thead>
                      <tbody>${recibosHtml}</tbody>
                    </table>
                  </div>
                </div>
              </td>
            </tr>
          `;
        }).join('');
        return;
      }

      // Listados por Matrícula con Contexto Multi-Curso y Expander
      let sourceList = [];
      let showMotivo = false;
      if (ACTIVE_TAB === 'jt_cartera') sourceList = h.jt_cartera;
      else if (ACTIVE_TAB === 'todas_matriculas') sourceList = h.todas_matriculas || [];
      else if (ACTIVE_TAB === 'abono_cero_real') sourceList = h.abono_cero_real;
      else if (ACTIVE_TAB === 'abono_cero_multicurso') sourceList = h.abono_cero_multicurso;
      else if (ACTIVE_TAB === 'cartera_critica_90') sourceList = h.cartera_critica_90;
      else if (ACTIVE_TAB === 'mora_30_90') sourceList = [...h.mora_media_60, ...h.mora_temprana_30];
      else if (ACTIVE_TAB === 'errores_digitacion') {
        sourceList = h.errores_digitacion;
        showMotivo = true;
      }

      let filtered = sourceList.filter(r => {
        if (FILTRO_ANIO_GLOBAL !== 'TODOS' && obtenerAnioDeRegistro(r) !== FILTRO_ANIO_GLOBAL) return false;
        const abMat = Number(r.ABONOS || 0);
        const abGlobal = Number(r.TOTAL_ABONADO_CLIENTE || abMat || 0);
        const vpMat = Number(r.VALORPENDIENTE || 0);
        const tieneAbonoEnAlgunCurso = abMat > 0 || abGlobal > 0;
        if (filterAbono === 'SOLO_CON_SALDO' && vpMat <= 0) return false;
        if (filterAbono === 'SOLO_CON_ABONO' && !tieneAbonoEnAlgunCurso) return false;
        if (filterAbono === 'SOLO_CERO_GLOBAL' && tieneAbonoEnAlgunCurso) return false;
        if (!q) return true;
        const txt = `${r.NUMEROMATRICULA} ${r.NUMERODOCUMENTO} ${r.ALUMNO} ${r.CONVENIO_NORM} ${r.CATEGORIA} ${r.FECHAINSCRIPCION || ''} ${r.ETIQUETA_MULTICURSO || ''} ${r.MOTIVO_ERROR || ''}`.toLowerCase();
        return txt.includes(q);
      });

      const numMatOf = (x) => Number(String(x.NUMEROMATRICULA || '').replace(/\D+/g, '')) || 0;
      filtered = [...filtered].sort((a, b) => {
        const fa = String(a.FECHAINSCRIPCION || '');
        const fb = String(b.FECHAINSCRIPCION || '');
        const abA = Math.max(Number(a.ABONOS || 0), Number(a.TOTAL_ABONADO_CLIENTE || 0));
        const abB = Math.max(Number(b.ABONOS || 0), Number(b.TOTAL_ABONADO_CLIENTE || 0));
        const hasAbA = abA > 0 ? 1 : 0;
        const hasAbB = abB > 0 ? 1 : 0;
        if (sortMode === 'ANTIGUOS') {
          return fa.localeCompare(fb) || (numMatOf(a) - numMatOf(b));
        }
        if (sortMode === 'CON_ABONO_PRIMERO') {
          if (hasAbB !== hasAbA) return hasAbB - hasAbA;
          return fb.localeCompare(fa) || (numMatOf(b) - numMatOf(a));
        }
        if (sortMode === 'SIN_ABONO_PRIMERO') {
          if (hasAbA !== hasAbB) return hasAbA - hasAbB;
          return fb.localeCompare(fa) || (numMatOf(b) - numMatOf(a));
        }
        if (sortMode === 'MAYOR_SALDO') {
          return (Number(b.VALORPENDIENTE || 0) - Number(a.VALORPENDIENTE || 0)) || fb.localeCompare(fa);
        }
        // Por defecto: RECIENTES (de los más recientes a los más antiguos)
        return fb.localeCompare(fa) || (numMatOf(b) - numMatOf(a));
      });

      const totalSaldo = filtered.reduce((acc, r) => acc + Math.max(0, Number(r.VALORPENDIENTE || 0)), 0);
      const cntConAbono = filtered.filter(r => Number(r.ABONOS || 0) > 0 || Number(r.TOTAL_ABONADO_CLIENTE || 0) > 0).length;
      const cntCeroGlobal = filtered.length - cntConAbono;
      const LIMIT_RENDER = 600;
      const displayRows = filtered.slice(0, LIMIT_RENDER);
      const avisoLimit = filtered.length > LIMIT_RENDER ? ` (Mostrando primeros ${LIMIT_RENDER} — usa el buscador o filtro de año para acotar)` : '';
      document.getElementById('tableCountInfo').innerHTML =
        `Total en vista: <b>${filtered.length}</b> registros${avisoLimit} (<span style="color:#fde047; font-weight:700;">💛 ${cntConAbono} con abono</span> • <span style="color:#fca5a5; font-weight:700;">🚨 ${cntCeroGlobal} en $0 global</span>) | Cartera Pendiente Real: <b>${fmtCOP(totalSaldo)}</b>`;

      thead.innerHTML = `<tr>
        <th>Matrícula</th><th>Cédula / Doc</th><th>Alumno</th><th>Convenio</th>
        <th>Fecha Insc.</th><th>Cat.</th><th>Valor Curso</th><th>Abono Mat.</th>
        <th>Saldo Mat.</th><th>Contexto Cliente (Multi-Curso / Caja Global)</th>
        <th>Antigüedad</th><th>Estado / Editar</th>
        ${showMotivo ? '<th>Hallazgo / Error Detectado</th>' : ''}
      </tr>`;

      tbody.innerHTML = displayRows.map((r, idx) => {
        const expId = `exp_row_${idx}`;
        const abMat = Number(r.ABONOS || 0);
        const abGlobal = Number(r.TOTAL_ABONADO_CLIENTE || abMat || 0);
        const tieneAbonoOtroCurso = (abMat === 0 && abGlobal > 0);
        const tieneAlgunAbono = (abMat > 0 || abGlobal > 0);
        const rowClass = tieneAlgunAbono ? 'data-row row-has-abono' : 'data-row row-zero-global';

        let badgeCtx = `<span class="badge">${r.ETIQUETA_MULTICURSO || 'Curso Único'}</span>`;
        if (r.TIPO_RIESGO_CERO === 'RIESGO_CERO_REAL') {
          badgeCtx = `<span class="badge badge-danger">🚨 $0 Global (Sin abono en ningún curso)</span>`;
        } else if (r.TIPO_RIESGO_CERO === 'MULTICURSO_CON_ABONO' || tieneAbonoOtroCurso) {
          badgeCtx = `<span class="badge badge-warning" style="background:rgba(245,158,11,0.28); color:#fde047; border:1px solid #f59e0b; font-weight:800;">💛 SÍ TIENE ABONO • Combo ${r.CATEGORIAS_CLIENTE} (Abonó ${fmtCOP(abGlobal)} en otra cat.)</span>`;
        } else if (abMat > 0 && r.ES_MULTICURSO) {
          badgeCtx = `<span class="badge badge-warning" style="background:rgba(245,158,11,0.24); color:#fde047; border:1px solid #f59e0b;">💛 SÍ TIENE ABONO (${r.CATEGORIAS_CLIENTE}) • Pagó ${fmtCOP(abGlobal)}</span>`;
        } else if (abMat > 0) {
          badgeCtx = `<span class="badge badge-warning" style="background:rgba(245,158,11,0.24); color:#fde047; border:1px solid #f59e0b;">💛 SÍ TIENE ABONO (Abonó ${fmtCOP(abMat)})</span>`;
        } else if (r.ES_MULTICURSO) {
          badgeCtx = `<span class="badge badge-purple">👥 Multi-Curso (${r.CATEGORIAS_CLIENTE}) • Global: ${fmtCOP(r.SALDO_REAL_CLIENTE)}</span>`;
        }

        const abonoCellHtml = abMat > 0
          ? `<span class="money-warn" title="Abonó en esta categoría">💛 ${fmtCOP(abMat)}</span>`
          : (tieneAbonoOtroCurso
              ? `<span class="money-warn" title="En esta matrícula figura $0, pero abonó ${fmtCOP(abGlobal)} en su otra categoría (${r.CATEGORIAS_CLIENTE})">💛 $0 (Abonó ${fmtCOP(abGlobal)})</span>`
              : `<span class="money-neg">$0</span>`);

        const cursosHtml = (r.CURSOS_HERMANOS || []).map(cd => `
          <tr>
            <td><b>${cd.NUMEROMATRICULA}</b></td>
            <td><span class="badge">${cd.CATEGORIA}</span></td>
            <td>${cd.FECHAINSCRIPCION}</td>
            <td>${fmtCOP(cd.VALORCURSO)}</td>
            <td class="money-pos">${fmtCOP(cd.ABONOS_ORIGINAL)}</td>
            <td class="${cd.VALORPENDIENTE_ORIGINAL > 0 ? 'money-neg' : ''}">${fmtCOP(cd.VALORPENDIENTE_ORIGINAL)}</td>
            <td>${cd.ESTADO}</td>
          </tr>
        `).join('');

        const recibosHtml = (r.RECIBOS_CLIENTE || []).map(rc => `
          <tr>
            <td><b>${rc.RECIBO}</b></td>
            <td>${rc.FECHA}</td>
            <td>${rc.NUMEROMATRICULA || '<span class="badge badge-warning">HUÉRFANO (Cédula)</span>'}</td>
            <td class="money-pos">${fmtCOP(rc.VALOR)}</td>
            <td>${rc.TIPOPAGO}</td>
          </tr>
        `).join('') || '<tr><td colspan="5" style="color:#fca5a5;">🚨 Sin recibos ni abonos registrados en caja ($0 Global)</td></tr>';

        return `
          <tr class="${rowClass}" onclick="toggleExpander('${expId}')">
            <td><b>${r.NUMEROMATRICULA}</b></td>
            <td><code>${r.NUMERODOCUMENTO}</code></td>
            <td><b>🔍 ${r.ALUMNO}</b></td>
            <td>${r.CONVENIO_NORM}</td>
            <td>${r.FECHAINSCRIPCION}</td>
            <td><span class="badge">${r.CATEGORIA}</span></td>
            <td>${fmtCOP(r.VALORCURSO)}</td>
            <td>${abonoCellHtml}</td>
            <td class="${Number(r.VALORPENDIENTE || 0) > 0 ? 'money-neg' : 'money-pos'}">${fmtCOP(Math.max(0, Number(r.VALORPENDIENTE || 0)))}</td>
            <td>${badgeCtx}</td>
            <td><span class="badge ${r.DIAS_ANTIGUEDAD > 90 ? 'badge-danger' : (r.DIAS_ANTIGUEDAD > 30 ? 'badge-warning' : '')}">${r.DIAS_ANTIGUEDAD} días</span></td>
            <td>
              <div style="display:flex; align-items:center; gap:6px;">
                <span>${r.ESTADO}</span>
                <button type="button" class="btn btn-outline" style="padding:2px 7px; font-size:11px; border-color:#10b981; color:#6ee7b7;" onclick="event.stopPropagation(); cargarMatriculaEnEditorManual('${String(r.NUMEROMATRICULA || '').replace(/'/g, '')}')" title="Editar o registrar abono para esta matrícula">✏️ Editar</button>
              </div>
            </td>
            ${showMotivo ? `<td style="color:#fca5a5;font-weight:600;">${r.MOTIVO_ERROR}</td>` : ''}
          </tr>
          <tr id="${expId}" class="expander-row">
            <td colspan="${showMotivo ? 13 : 12}">
              <div class="expander-box">
                <div>
                  <b style="color:#93c5fd;">📊 Consolidado por Cédula (${r.NUMERODOCUMENTO}) — ${r.ALUMNO}</b>
                  <p style="font-size:12px; color:#cbd5e1; margin:4px 0;">
                    Total Contratado Cliente: <b>${fmtCOP(r.TOTAL_CONTRATADO_CLIENTE)}</b> |
                    Total Pagado Cliente: <b style="color:#34d399;">${fmtCOP(r.TOTAL_ABONADO_CLIENTE)}</b> |
                    Saldo Real Cliente: <b style="color:#f87171;">${fmtCOP(r.SALDO_REAL_CLIENTE)}</b>
                  </p>
                  <table class="mini-table">
                    <thead><tr><th>Matrícula</th><th>Cat.</th><th>Fecha Insc.</th><th>Valor Curso</th><th>Abono Reg.</th><th>Saldo Reg.</th><th>Estado</th></tr></thead>
                    <tbody>${cursosHtml}</tbody>
                  </table>
                </div>
                <div>
                  <b style="color:#6ee7b7;">🧾 Recibos de Caja Asociados a la Cédula (${r.NUMERODOCUMENTO})</b>
                  <table class="mini-table">
                    <thead><tr><th>N° Recibo</th><th>Fecha</th><th>Matrícula</th><th>Valor</th><th>Medio de Pago</th></tr></thead>
                    <tbody>${recibosHtml}</tbody>
                  </table>
                </div>
              </div>
            </td>
          </tr>
        `;
      }).join('');
    }

    document.getElementById('uploadForm').addEventListener('submit', async (e) => {
      e.preventDefault();
      const fileEl = document.getElementById('fileInput');
      const ingEl = document.getElementById('ingresosInput');
      if (!fileEl.files.length) return;
      const fd = new FormData();
      fd.append('file', fileEl.files[0]);
      if (ingEl.files.length) {
        fd.append('ingresos_file', ingEl.files[0]);
      }
      fd.append('periodicidad', document.getElementById('periodoInput').value);
      const modoBdEl = document.getElementById('modoBdInput');
      fd.append('modo_bd', modoBdEl ? modoBdEl.value : 'INCORPORAR');

      const toast = document.getElementById('uploadToast');
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Incorporando archivo en la Base de Datos Histórica y recalculando...';

      try {
        const res = await fetch('/api/upload', { method: 'POST', body: fd });
        const data = await res.json();
        if (data.ok) {
          AUDIT_DATA = data;
          actualizarVista();
          await cargarBotOverview();
          toast.className = 'status-toast ok';
          const k = data.hallazgos.kpis;
          const ms = data.merge_stats || {};
          if (ms.total_despues) {
            toast.textContent = `✅ Base Histórica actualizada: ${ms.total_despues} matrículas conservadas (+${ms.nuevos_insertados || 0} nuevas, ${ms.existentes_actualizados || 0} modificadas).`;
          } else {
            toast.textContent = `✅ Auditado: ${k.total_matriculas} matrículas (${k.total_clientes_unicos} clientes únicos, ${k.clientes_multicurso_count} multi-curso).`;
          }
        } else {
          toast.className = 'status-toast err';
          toast.textContent = `❌ Error: ${data.error}`;
        }
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = `❌ Error de red: ${err}`;
      }
    });

    async function despacharEmail() {
      const email = document.getElementById('emailDestInput').value.trim();
      const toast = document.getElementById('dispatchToast');
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Enviando informe por correo SMTP...';
      try {
        const res = await fetch('/api/send/email', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ email })
        });
        const data = await res.json();
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (data.ok ? '✅ ' : '⚠️ ') + data.message;
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error: ' + err;
      }
    }

    async function despacharTelegram() {
      const chat_id = document.getElementById('tgChatInput').value.trim();
      const toast = document.getElementById('dispatchToast');
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Enviando resumen + Excel + PDF a Telegram...';
      try {
        const res = await fetch('/api/send/telegram', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ chat_id })
        });
        const data = await res.json();
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (data.ok ? '✅ ' : '⚠️ ') + data.message;
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error: ' + err;
      }
    }

    // =========================================================================
    // NAVEGACIÓN ENTRE LOS 5 MÓDULOS Y SINCRONIZACIÓN CON EL BOT / SHEETS
    // =========================================================================
    function cambiarModulo(modId, btnEl) {
      ACTIVE_MODULE = modId;
      document.querySelectorAll('.mod-section').forEach(s => s.classList.remove('active'));
      document.querySelectorAll('.mod-btn').forEach(b => b.classList.remove('active'));
      const target = document.getElementById(modId);
      if (target) target.classList.add('active');
      if (btnEl) btnEl.classList.add('active');
      if (CURRENT_USER && CURRENT_USER.rol !== 'ADMIN_TOTAL') {
        const modsOk = CURRENT_USER.modulos || [];
        if (!modsOk.includes(modId)) {
          alert('⛔ Tu perfil de usuario no tiene permiso para acceder a este módulo.');
          return;
        }
      }
      if (modId === 'mod_empleados') {
        cargarEmpleadosWeb();
      } else if (modId === 'mod_agenda') {
        cargarAgendaWeb();
      } else if (modId === 'mod_whatsapp') {
        cargarWhatsappCrm();
      } else if (modId === 'mod_robot') {
        cargarBitacoraRobot();
      } else if (modId !== 'mod_cartera' && !BOT_DATA) {
        cargarBotOverview();
      }
    }

    function toggleSheetsDrawer() {
      document.getElementById('sheetsDrawer').classList.toggle('open');
    }

    function leerArchivoCredenciales(inputEl) {
      if (!inputEl.files || !inputEl.files.length) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        document.getElementById('credJsonInput').value = ev.target.result || '';
      };
      reader.readAsText(inputEl.files[0]);
    }

    async function conectarGoogleSheets() {
      const rawJson = document.getElementById('credJsonInput').value.trim();
      const toast = document.getElementById('sheetsToast');
      if (!rawJson) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Selecciona el archivo credentials.json o pega su contenido JSON.';
        return;
      }
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Conectando con Google Sheets del Bot...';
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'conectar_sheets', credentials_json: rawJson })
        });
        const data = await res.json();
        const isOk = Array.isArray(data) ? Boolean(data[0]) : Boolean(data && data.ok);
        const msg = Array.isArray(data) ? String(data[1] || '') : String((data && (data.message || data.error)) || '');
        toast.className = isOk ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (isOk ? '✅ ' : '❌ ') + msg;
        if (isOk) {
          await cargarBotOverview();
        }
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error: ' + err;
      }
    }

    async function cargarBotOverview() {
      const nomPeriodo = (document.getElementById('nom_periodo') || {}).value || 'mes';
      const nomInst = (document.getElementById('nom_instructor') || {}).value || 'TODOS';
      const inasPeriodo = (document.getElementById('inas_periodo') || {}).value || 'todo';
      const inasEstado = (document.getElementById('inas_estado_filtro') || {}).value || 'TODOS';
      const pesvPeriodo = (document.getElementById('pesv_periodo') || {}).value || 'todo';
      const pesvPlaca = (document.getElementById('pesv_placa_filtro') || {}).value || 'TODAS';

      const qs = new URLSearchParams({
        periodo_nomina: nomPeriodo,
        instructor_nomina: nomInst,
        periodo_inasistencias: inasPeriodo,
        estado_inasistencias: inasEstado,
        periodo_mantenimientos: pesvPeriodo,
        placa_mantenimientos: pesvPlaca,
      });

      try {
        const res = await fetch('/api/bot/overview?' + qs.toString());
        const data = await res.json();
        if (!data.ok) return;
        if (BOT_DATA && BOT_DATA.clases && (BOT_DATA.clases.alumnos || []).length > 0 && data.clases && (data.clases.alumnos || []).length === 0) {
          // Si una consulta transitoria devolvió 0 alumnos pero ya teníamos alumnos cargados en memoria, conservar la lista previa
          data.clases = BOT_DATA.clases;
          data.alumnos = BOT_DATA.alumnos;
        }
        BOT_DATA = data;

        // Estado de Google Sheets en Header
        const btnSheets = document.getElementById('btnSheetsStatus');
        const lblSheets = document.getElementById('sheetsStatusLabel');
        if (data.sheets_conectado) {
          btnSheets.innerHTML = '🟢 Sheets Conectado';
          lblSheets.className = 'badge badge-success';
          lblSheets.textContent = '🟢 Sincronizado en Vivo con Google Sheets';
        } else {
          btnSheets.innerHTML = '🟡 Conectar Google Sheets';
          lblSheets.className = 'badge badge-warning';
          lblSheets.textContent = '🟡 Modo Base Local / Semilla Histórica (Clic para conectar Google Sheets en vivo)';
        }

        // Actualizar badges de la barra de módulos
        document.getElementById('nav_badge_clases').textContent = `${data.clases.kpis.en_curso} Activos • ${data.clases.kpis.finalizados} Fin.`;
        document.getElementById('nav_badge_nomina').textContent = `${data.nomina.total_clases} Clases`;
        document.getElementById('nav_badge_inas').textContent = `${data.inasistencias.kpis.total} (${data.inasistencias.kpis.pendientes} Pend.)`;
        document.getElementById('nav_badge_pesv').textContent = `${data.mantenimientos.kpis.total} Mant.`;

        // Poblar datalist de alumnos en el Módulo 6 (Robot Gestor) y datalists globales de autocompletado
        const dlAl = document.getElementById('listaAlumnosRobot');
        const listaAl = (data.clases && data.clases.alumnos) ? data.clases.alumnos : [];
        if (dlAl && listaAl.length) {
          dlAl.innerHTML = listaAl.map(a =>
            `<option value="${a.cedula || a.nombre}|${a.categoria}">${a.nombre} (${a.categoria} - ${a.clases_completadas}/${a.total_requerido} - ${a.estado_curso})</option>`
          ).join('');
        }

        // Poblar datalists globales (Alumnos + Matrículas Excel + Inasistencias + Instructores)
        const mapNom = new Map();
        const mapCed = new Map();
        listaAl.forEach(a => {
          if (a.nombre) mapNom.set(a.nombre.toUpperCase(), a);
          if (a.cedula) mapCed.set(String(a.cedula), a);
        });
        if (AUDIT_DATA && AUDIT_DATA.hallazgos && AUDIT_DATA.hallazgos.todas_matriculas) {
          AUDIT_DATA.hallazgos.todas_matriculas.forEach(m => {
            const nomM = String(m.ALUMNO || '').trim();
            const docM = String(m.NUMERODOCUMENTO || '').trim();
            const catM = String(m.CATEGORIA || 'B1').trim();
            if (nomM && !mapNom.has(nomM.toUpperCase())) {
              mapNom.set(nomM.toUpperCase(), { cedula: docM, nombre: nomM, categoria: catM });
            }
            if (docM && !mapCed.has(docM)) {
              mapCed.set(docM, { cedula: docM, nombre: nomM, categoria: catM });
            }
          });
        }
        if (data.inasistencias && data.inasistencias.registros) {
          data.inasistencias.registros.forEach(r => {
            const nomI = String(r.alumno || '').trim();
            if (nomI && !mapNom.has(nomI.toUpperCase())) {
              mapNom.set(nomI.toUpperCase(), { cedula: r.cedula || '', nombre: nomI, categoria: r.categoria || 'B1' });
            }
          });
        }
        const dlNomGlobal = document.getElementById('dlAlumnosGlobal');
        if (dlNomGlobal) {
          dlNomGlobal.innerHTML = Array.from(mapNom.values()).slice(0, 600).map(a =>
            `<option value="${a.nombre}">${a.cedula ? 'CC ' + a.cedula + ' • ' : ''}${a.categoria || ''}</option>`
          ).join('');
        }
        const dlCedGlobal = document.getElementById('dlCedulasGlobal');
        if (dlCedGlobal) {
          dlCedGlobal.innerHTML = Array.from(mapCed.values()).slice(0, 600).map(a =>
            `<option value="${a.cedula}">${a.nombre} (${a.categoria || ''})</option>`
          ).join('');
        }
        const dlInstGlobal = document.getElementById('dlInstructoresGlobal');
        if (dlInstGlobal && data.nomina && data.nomina.instructores_disponibles) {
          dlInstGlobal.innerHTML = data.nomina.instructores_disponibles.map(i => `<option value="${i}"></option>`).join('');
        }
        const inFechaEl = document.getElementById('in_fecha');
        if (inFechaEl && !inFechaEl.value) {
          const hoy = new Date();
          inFechaEl.value = `${String(hoy.getDate()).padStart(2, '0')}/${String(hoy.getMonth() + 1).padStart(2, '0')}/${hoy.getFullYear()}`;
        }

        renderizarModuloClases();
        renderizarModuloNomina();
        renderizarModuloInasistencias();
        renderizarModuloPesv();
        cargarBitacoraRobot();
        cargarPanoramaRuntDigital(false);
      } catch (err) {
        console.error('Error cargando panorama del bot:', err);
      }
    }

    let RUNT_SHEET_DATA = null;
    let ACTIVE_RUNT_TAB = 'PENDIENTES_PRACTICA';
    let RUNT_SELECTED_FILAS = new Set();

    function toggleSeleccionFilaRunt(filaIdx, checked) {
      const f = Number(filaIdx);
      if (checked) RUNT_SELECTED_FILAS.add(f);
      else RUNT_SELECTED_FILAS.delete(f);
    }

    function toggleSeleccionarTodosRunt(checked) {
      const rows = obtenerFilasFiltradasPracticaRunt();
      rows.forEach(r => {
        if (checked) RUNT_SELECTED_FILAS.add(Number(r.fila));
        else RUNT_SELECTED_FILAS.delete(Number(r.fila));
      });
      renderizarTablaHojaRunt();
    }

    function ajustarFechaBatchRunt(offsetDias = 0) {
      const d = new Date();
      d.setDate(d.getDate() + Number(offsetDias || 0));
      const iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
      const elBatch = document.getElementById('rn_batch_fecha');
      const elForm = document.getElementById('rn_fecha_cargue');
      if (elBatch) elBatch.value = iso;
      if (elForm) elForm.value = iso;
    }

    function sumarHorasInputFormulario(inputId, delta) {
      const el = document.getElementById(inputId);
      if (!el) return;
      const cat = (document.getElementById('rn_cat') || {}).value || 'A2';
      const maxPrac = { 'A1': 8, 'A2': 15, 'B1': 20, 'C1': 30, 'RECAT_C1': 8 }[cat] || 20;
      const actual = Number(el.value || 0);
      el.value = Math.min(maxPrac, actual + Number(delta || 0));
      const cg = document.getElementById('rn_cargue_diario');
      if (cg) cg.value = 'PRACTICA';
    }

    async function cargarPanoramaRuntDigital(forzar = false) {
      try {
        const elBatchF = document.getElementById('rn_batch_fecha');
        if (elBatchF && !elBatchF.value) ajustarFechaBatchRunt(0);
        const res = await fetch('/api/bot/runt_control?forzar=' + (forzar ? '1' : '0'));
        const data = await res.json();
        if (!data || !data.ok) return;
        RUNT_SHEET_DATA = data;

        const stBadge = document.getElementById('runtSheetStatusBadge');
        const extLink = document.getElementById('runtSheetExternalLink');
        if (extLink && data.sheet_url) extLink.href = data.sheet_url;
        if (stBadge) {
          if (data.conectado_en_vivo) {
            stBadge.className = 'badge badge-success';
            stBadge.textContent = `🟢 Conectado en Vivo • ID: ${data.sheet_id}`;
          } else {
            stBadge.className = 'badge badge-warning';
            stBadge.textContent = `🟡 Base Activa RUNT (${data.kpis.total_practica} reg.) • ID: ${data.sheet_id}`;
          }
        }

        const k = data.kpis || {};
        const setCnt = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
        setCnt('cnt_rn_pend', k.pendientes_runt || 0);
        setCnt('cnt_rn_term_sin', k.terminaron_sin_cargue || 0);
        setCnt('cnt_rn_urg_prac', k.urgentes_en_practica || 0);
        setCnt('cnt_rn_moto_teo', k.moto_sin_25_teoria || 0);
        setCnt('cnt_rn_moto_tal', k.moto_sin_3_taller || 0);
        setCnt('cnt_rn_carro_tt', k.carro_sin_teoria_taller || 0);
        setCnt('cnt_rn_bot', k.con_clases_bot_por_subir || 0);
        setCnt('cnt_rn_vence', k.por_vencer_o_vencidos || 0);
        setCnt('cnt_rn_cargados', k.cargados_runt || 0);
        setCnt('cnt_rn_todos', k.total_practica || 0);
        setCnt('cnt_rn_norunt', k.no_runt_pendientes || 0);
        setCnt('cnt_rn_horarios', k.total_inicio_practica || 0);
        setCnt('cnt_rn_codigos', k.total_codigos_runt || 0);

        const kpiEl = document.getElementById('kpiRuntSheetGrid');
        if (kpiEl) {
          kpiEl.innerHTML = `
            <div class="kpi-card danger" style="cursor:pointer;" onclick="cambiarTabHojaRunt('TERMINARON_SIN_CARGUE')">
              <div class="kpi-label">🔥 1. Terminaron y Falta Cargue Manual RUNT</div>
              <div class="kpi-value" style="color:#fdba74;">${k.terminaron_sin_cargue || 0} Listos</div>
              <div class="kpi-sub">Prioridad #1 de certificación + <b>🚨 ${k.urgentes_en_practica || 0} en práctica</b></div>
            </div>
            <div class="kpi-card purple" style="cursor:pointer;" onclick="cambiarTabHojaRunt('MOTO_SIN_25_TEORIA')">
              <div class="kpi-label">🏍️ Moto A2 Pendientes Cargue Masivo</div>
              <div class="kpi-value" style="color:#d8b4fe;">${k.moto_sin_25_teoria || 0} Sin 25h Teoría</div>
              <div class="kpi-sub"><b>🔧 ${k.moto_sin_3_taller || 0}</b> sin 3h Taller • <b>🚗 ${k.carro_sin_teoria_taller || 0}</b> Carros sin Teo/Tal</div>
            </div>
            <div class="kpi-card warning" style="cursor:pointer;" onclick="cambiarTabHojaRunt('POR_VENCER_90')">
              <div class="kpi-label">💸 Vencidos / Por Vencer 90 Días (Reiniciar + Sanción)</div>
              <div class="kpi-value" style="color:#fca5a5;">${k.por_vencer_o_vencidos || 0} Alumnos</div>
              <div class="kpi-sub"><b>${k.vencidos_sancion || 0} vencidos</b> para reiniciar reporte RUNT con sanción</div>
            </div>
            <div class="kpi-card success" style="cursor:pointer;" onclick="cambiarTabHojaRunt('CARGADOS_PRACTICA')">
              <div class="kpi-label">✅ Ya Cargados 100% en Plataforma RUNT</div>
              <div class="kpi-value" style="color:#34d399;">${k.cargados_runt || 0} Certificados</div>
              <div class="kpi-sub"><b>⚠️ ${k.pendientes_runt || 0} pendientes</b> en total en hoja Practica</div>
            </div>
          `;
        }

        renderizarTablaHojaRunt();
      } catch (err) {
        console.error('Error cargando hoja RUNT digital:', err);
      }
    }

    function cambiarTabHojaRunt(tabVal, btnEl) {
      ACTIVE_RUNT_TAB = tabVal;
      document.querySelectorAll('#runtSheetTabsHeader .tab-btn').forEach(b => b.classList.remove('active'));
      if (btnEl) {
        btnEl.classList.add('active');
      } else {
        const idxMap = {
          'PENDIENTES_PRACTICA': 0,
          'TERMINARON_SIN_CARGUE': 1,
          'URGENTES_EN_PRACTICA': 2,
          'MOTO_SIN_25_TEORIA': 3,
          'MOTO_SIN_3_TALLER': 4,
          'CARRO_SIN_TEO_TAL': 5,
          'DESFASE_BOT': 6,
          'POR_VENCER_90': 7,
          'CARGADOS_PRACTICA': 8,
          'TODOS_PRACTICA': 9,
          'TAB_NO_RUNT': 10,
          'TAB_INICIO_PRACTICA': 11,
          'TAB_CODIGOS_RUNT': 12
        };
        const btns = document.querySelectorAll('#runtSheetTabsHeader .tab-btn');
        if (btns[idxMap[tabVal] || 0]) btns[idxMap[tabVal] || 0].classList.add('active');
      }
      renderizarTablaHojaRunt();
    }

    function ajustarMetasFormularioRunt() {
      const cat = (document.getElementById('rn_cat') || {}).value || 'A2';
      const metas = {
        'A1': { t: 25, tl: 3, p: 8 },
        'A2': { t: 25, tl: 3, p: 15 },
        'B1': { t: 25, tl: 5, p: 20 },
        'C1': { t: 30, tl: 5, p: 30 },
        'RECAT_C1': { t: 8, tl: 0, p: 8 }
      }[cat] || { t: 25, tl: 3, p: 15 };
      const filaAct = Number((document.getElementById('rn_fila') || {}).value || 0);
      if (filaAct === 0) {
        document.getElementById('rn_teoria').value = metas.t;
        document.getElementById('rn_taller').value = metas.tl;
        document.getElementById('rn_practica').value = metas.p;
      }
    }

    function cargarFilaRuntEnFormulario(filaIdx) {
      if (!RUNT_SHEET_DATA || !RUNT_SHEET_DATA.practica) return;
      const item = RUNT_SHEET_DATA.practica.find(x => Number(x.fila) === Number(filaIdx));
      if (!item) return;
      document.getElementById('rn_fila').value = item.fila;
      document.getElementById('rn_cedula').value = item.cedula || '';
      document.getElementById('rn_nombre').value = item.nombre || '';
      document.getElementById('rn_cat').value = item.categoria || 'A2';
      document.getElementById('rn_prioridad').value = item.prioridad || 'EN PRACTICA';
      document.getElementById('rn_fecha_mat').value = item.fecha_matricula || '';
      document.getElementById('rn_teoria').value = item.teoria || 0;
      document.getElementById('rn_taller').value = item.taller || 0;
      document.getElementById('rn_practica').value = Math.max(item.practica || 0, item.practica_bot || 0);
      document.getElementById('rn_cargue_diario').value = item.cargue_diario || '';
      document.getElementById('rn_estado_runt').value = item.estado_runt || 'PENDIENTE RUNT';
      const badge = document.getElementById('runtEditRowBadge');
      if (badge) {
        badge.className = 'badge badge-info';
        badge.textContent = `✏️ Editando Fila #${item.fila}: ${item.nombre} (${item.categoria})`;
      }
      document.getElementById('rn_cedula').scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    function autocompletarFormularioRuntSheet(origen) {
      if (!RUNT_SHEET_DATA || !RUNT_SHEET_DATA.practica) return;
      const cedVal = (document.getElementById('rn_cedula').value || '').trim();
      const nomVal = (document.getElementById('rn_nombre').value || '').trim().toUpperCase();
      const catVal = (document.getElementById('rn_cat').value || 'A2').trim();
      let match = null;
      if (origen === 'cedula' && cedVal.length >= 5) {
        match = RUNT_SHEET_DATA.practica.find(x => String(x.cedula) === cedVal && x.categoria === catVal)
          || RUNT_SHEET_DATA.practica.find(x => String(x.cedula) === cedVal);
      } else if (origen === 'nombre' && nomVal.length >= 4) {
        match = RUNT_SHEET_DATA.practica.find(x => String(x.nombre || '').toUpperCase() === nomVal && x.categoria === catVal)
          || RUNT_SHEET_DATA.practica.find(x => String(x.nombre || '').toUpperCase() === nomVal);
      }
      if (match) {
        document.getElementById('rn_fila').value = match.fila;
        if (origen === 'cedula' && match.nombre) document.getElementById('rn_nombre').value = match.nombre;
        if (origen === 'nombre' && match.cedula) document.getElementById('rn_cedula').value = match.cedula;
        document.getElementById('rn_cat').value = match.categoria || 'A2';
        document.getElementById('rn_prioridad').value = match.prioridad || 'EN PRACTICA';
        if (match.fecha_matricula) document.getElementById('rn_fecha_mat').value = match.fecha_matricula;
        document.getElementById('rn_teoria').value = match.teoria || 0;
        document.getElementById('rn_taller').value = match.taller || 0;
        document.getElementById('rn_practica').value = Math.max(match.practica || 0, match.practica_bot || 0);
        document.getElementById('rn_estado_runt').value = match.estado_runt || 'PENDIENTE RUNT';
        const badge = document.getElementById('runtEditRowBadge');
        if (badge) {
          badge.className = 'badge badge-info';
          badge.textContent = `✏️ Fila #${match.fila} encontrada (${match.nombre} - ${match.categoria})`;
        }
      } else {
        document.getElementById('rn_fila').value = '0';
      }
    }

    async function guardarRegistroHojaRunt(completarCienPorCiento = false) {
      const fila = Number(document.getElementById('rn_fila').value || 0);
      const cedula = document.getElementById('rn_cedula').value.trim();
      const nombre = document.getElementById('rn_nombre').value.trim();
      const categoria = document.getElementById('rn_cat').value;
      let prioridad = document.getElementById('rn_prioridad').value;
      const fecha_matricula = document.getElementById('rn_fecha_mat').value;
      let teoria = Number(document.getElementById('rn_teoria').value || 0);
      let taller = Number(document.getElementById('rn_taller').value || 0);
      let practica = Number(document.getElementById('rn_practica').value || 0);
      let cargue_diario = document.getElementById('rn_cargue_diario').value;
      let estado_runt = document.getElementById('rn_estado_runt').value;
      const fecha_cargue = (document.getElementById('rn_fecha_cargue') || {}).value || (document.getElementById('rn_batch_fecha') || {}).value || '';
      const toast = document.getElementById('runtSheetToast');

      if (!cedula && !nombre) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Ingresa la Cédula o el Nombre del alumno.';
        return;
      }

      if (completarCienPorCiento) {
        const metas = {
          'A1': { t: 25, tl: 3, p: 8 },
          'A2': { t: 25, tl: 3, p: 15 },
          'B1': { t: 25, tl: 5, p: 20 },
          'C1': { t: 30, tl: 5, p: 30 },
          'RECAT_C1': { t: 8, tl: 0, p: 8 }
        }[categoria] || { t: 25, tl: 3, p: 15 };
        teoria = metas.t;
        taller = metas.tl;
        practica = metas.p;
        prioridad = 'Termino';
        estado_runt = 'CARGADO RUNT';
        if (!cargue_diario) cargue_diario = 'PRACTICA';
      }

      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Guardando conteo de horas en Hoja RUNT y sincronizando con Módulo 2...';

      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'runt_guardar_practica',
            fila, cedula, nombre, categoria, prioridad, fecha_matricula,
            teoria, taller, practica, cargue_diario, estado_runt, fecha_cargue
          })
        });
        const data = await res.json();
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = data.message || 'Actualizado.';
        if (data.ok) {
          await cargarPanoramaRuntDigital(false);
          await cargarBotOverview();
        }
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error guardando en Hoja RUNT: ' + err;
      }
    }

    async function ejecutarCargueMasivoRunt(tipo) {
      const categoria = (document.getElementById('rn_batch_cat') || {}).value || 'A2';
      const modo_cupo = (document.getElementById('rn_batch_modo') || {}).value || 'MAX_8';
      const fecha_cargue = (document.getElementById('rn_batch_fecha') || {}).value || '';
      const toast = document.getElementById('runtSheetToast');

      const todasFilas = (RUNT_SHEET_DATA && RUNT_SHEET_DATA.practica) ? RUNT_SHEET_DATA.practica : [];
      const seleccionadasArr = Array.from(RUNT_SELECTED_FILAS);
      let candidatos = [];
      if (seleccionadasArr.length > 0) {
        candidatos = todasFilas.filter(r => RUNT_SELECTED_FILAS.has(Number(r.fila)) && r.estado_runt !== 'CARGADO RUNT');
      } else {
        candidatos = todasFilas.filter(r => {
          if (r.estado_runt === 'CARGADO RUNT') return false;
          if (categoria !== 'TODAS' && r.categoria !== categoria) return false;
          if (tipo === 'TEORIA') return r.teoria < r.meta_teoria;
          if (tipo === 'TALLER') return r.taller < r.meta_taller;
          return (r.teoria < r.meta_teoria) || (r.taller < r.meta_taller);
        });
      }

      if (!candidatos.length) {
        alert('⚠️ No hay alumnos seleccionados o pendientes que necesiten horas de ' + tipo + ' en la categoría ' + categoria + '.');
        return;
      }

      const previewLineas = candidatos.slice(0, 8).map(r => {
        const nTeo = (tipo === 'TEORIA' || tipo === 'TEORIA_Y_TALLER')
          ? (modo_cupo === 'COMPLETAR_META' ? r.meta_teoria : Math.min(r.meta_teoria, r.teoria + 8))
          : r.teoria;
        const nTal = (tipo === 'TALLER' || tipo === 'TEORIA_Y_TALLER')
          ? (modo_cupo === 'COMPLETAR_META' ? r.meta_taller : Math.min(r.meta_taller, r.taller + 8))
          : r.taller;
        return `• ${r.nombre} (${r.cedula}): Teoría ${r.teoria}→${nTeo}/${r.meta_teoria} | Taller ${r.taller}→${nTal}/${r.meta_taller}`;
      }).join('\n');

      const okConfirm = confirm(
        `🚀 CONFIRMAR CARGUE MASIVO DE ${tipo}\n\n` +
        `• Fecha de Cargue: ${fecha_cargue || 'Hoy'}\n` +
        `• Regla: ${modo_cupo === 'MAX_8' ? 'Sumar máx. +8 clases hoy (Tope RUNT)' : 'Completar 100% meta'}\n` +
        `• Total alumnos a actualizar: ${candidatos.length} ${seleccionadasArr.length > 0 ? '(Marcados con chulito ☑️)' : '(Categoría ' + categoria + ')'}\n\n` +
        `Vista previa:\n${previewLineas}${candidatos.length > 8 ? '\n...y ' + (candidatos.length - 8) + ' alumno(s) más.' : ''}\n\n` +
        `¿Confirmas aplicar este cargue masivo?`
      );
      if (!okConfirm) return;

      if (toast) {
        toast.className = 'status-toast ok';
        toast.textContent = `⏳ Ejecutando Cargue Masivo de ${tipo} para ${candidatos.length} alumno(s)...`;
      }
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'runt_cargue_masivo',
            tipo,
            categoria,
            modo_cupo,
            fecha_cargue,
            filas: candidatos.map(r => r.fila)
          })
        });
        const data = await res.json();
        if (toast) {
          toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
          toast.textContent = data.message || 'Cargue masivo ejecutado.';
        }
        RUNT_SELECTED_FILAS.clear();
        await cargarPanoramaRuntDigital(false);
      } catch (err) {
        if (toast) {
          toast.className = 'status-toast err';
          toast.textContent = '❌ Error en cargue masivo: ' + err;
        }
      }
    }

    async function deshacerCargueMasivoRunt(modo = 'RESTAURAR_14_A2_ORIGINAL') {
      const seleccionadasArr = Array.from(RUNT_SELECTED_FILAS);
      let msgConfirm = '';
      if (modo === 'RESTAR_8_TEORIA') {
        msgConfirm = seleccionadasArr.length > 0
          ? `¿Deseas restar -8 horas de Teoría a los ${seleccionadasArr.length} alumno(s) marcados con ☑️?`
          : `¿Deseas restar -8 horas de Teoría a los 14 alumnos de Moto A2 que recibieron el cargue masivo?`;
      } else {
        msgConfirm = `↩️ CORREGIR CARGUE MASIVO POR ERROR\n\n` +
          `Esto devolverá inmediatamente en Google Sheets y en el sistema a los 14 alumnos de Moto A2 (Diana Marcela Vivienscas, Angel Cárdenas, Cristo Humberto Llain, Jonathan González, Jesús Sebastián Suárez, Rosalía Plata, etc.) a sus horas originales exactas (16h, 8h, 24h, 0h) tal como estaban antes del cargue masivo.\n\n` +
          `¿Confirmas restaurar sus horas originales ahora?`;
      }
      if (!confirm(msgConfirm)) return;

      const toast = document.getElementById('runtSheetToast');
      if (toast) {
        toast.className = 'status-toast ok';
        toast.textContent = '⏳ Restaurando horas originales de los alumnos en Google Sheets y en el sistema...';
      }
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'runt_deshacer_masivo',
            modo,
            filas: seleccionadasArr
          })
        });
        const data = await res.json();
        if (toast) {
          toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
          toast.textContent = data.message || 'Horas restauradas.';
        }
        RUNT_SELECTED_FILAS.clear();
        await cargarPanoramaRuntDigital(false);
      } catch (err) {
        if (toast) {
          toast.className = 'status-toast err';
          toast.textContent = '❌ Error restaurando horas: ' + err;
        }
      }
    }

    async function sumarLotePracticaRunt(filaIdx, horasSumar = 8, completarTodo = false) {
      const fecha_cargue = (document.getElementById('rn_batch_fecha') || {}).value || (document.getElementById('rn_fecha_cargue') || {}).value || '';
      const toast = document.getElementById('runtSheetToast');
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'runt_sumar_practica_individual',
            fila: filaIdx,
            horas_sumar: horasSumar,
            completar_todo: completarTodo,
            fecha_cargue
          })
        });
        const data = await res.json();
        if (toast) {
          toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
          toast.textContent = data.message || '';
        }
        await cargarPanoramaRuntDigital(false);
        if (completarTodo) await cargarBotOverview();
      } catch (err) {
        if (toast) {
          toast.className = 'status-toast err';
          toast.textContent = '❌ Error sumando práctica: ' + err;
        }
      }
    }

    async function reiniciarCursoVencidoRunt(filaIdx, nombreAlumno, catAlumno) {
      const valorSancion = Number((document.getElementById('rn_sancion_valor') || {}).value || 50000);
      const nuevaFecha = (document.getElementById('rn_batch_fecha') || {}).value || '';
      const confirmar = confirm(
        `⚠️ REINICIO DE CURSO VENCIDO (90 DÍAS) + SANCIÓN ECONÓMICA\n\n` +
        `Alumno: ${nombreAlumno} (${catAlumno})\n` +
        `• Se reiniciará su vigencia de 90 días a partir de ${nuevaFecha || 'Hoy'}.\n` +
        `• Se reiniciará el conteo de clases reportadas en RUNT a 0h para volver a subirlas.\n` +
        `• Se generará automáticamente la Sanción Económica por $${valorSancion.toLocaleString('es-CO')} en Control de Pagos.\n\n` +
        `¿Confirmas el reinicio del reporte RUNT?`
      );
      if (!confirmar) return;
      const toast = document.getElementById('runtSheetToast');
      if (toast) {
        toast.className = 'status-toast ok';
        toast.textContent = `⏳ Reiniciando vigencia RUNT de ${nombreAlumno} y registrando sanción económica...`;
      }
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'runt_reiniciar_vencido',
            fila: filaIdx,
            valor_sancion: valorSancion,
            nueva_fecha: nuevaFecha,
            conservar_teoria: false
          })
        });
        const data = await res.json();
        if (toast) {
          toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
          toast.textContent = data.message || '';
        }
        await cargarPanoramaRuntDigital(false);
        await cargarBotOverview();
      } catch (err) {
        if (toast) {
          toast.className = 'status-toast err';
          toast.textContent = '❌ Error reiniciando curso: ' + err;
        }
      }
    }

    function obtenerFilasFiltradasPracticaRunt() {
      if (!RUNT_SHEET_DATA || !RUNT_SHEET_DATA.practica) return [];
      const q = obtenerBusquedaLimpia('searchRuntSheetInput');
      let listP = RUNT_SHEET_DATA.practica || [];
      listP = listP.filter(r => {
        const esCarg = r.estado_runt === 'CARGADO RUNT';
        if (ACTIVE_RUNT_TAB === 'PENDIENTES_PRACTICA' && esCarg) return false;
        if (ACTIVE_RUNT_TAB === 'TERMINARON_SIN_CARGUE' && !r.es_terminado_sin_cargue) return false;
        if (ACTIVE_RUNT_TAB === 'URGENTES_EN_PRACTICA' && !r.es_urgente_en_practica) return false;
        if (ACTIVE_RUNT_TAB === 'MOTO_SIN_25_TEORIA' && (esCarg || !['A1', 'A2'].includes(r.categoria) || r.teoria >= r.meta_teoria)) return false;
        if (ACTIVE_RUNT_TAB === 'MOTO_SIN_3_TALLER' && (esCarg || !['A1', 'A2'].includes(r.categoria) || r.taller >= r.meta_taller)) return false;
        if (ACTIVE_RUNT_TAB === 'CARRO_SIN_TEO_TAL' && (esCarg || !['B1', 'C1', 'RECAT_C1'].includes(r.categoria) || (r.teoria >= r.meta_teoria && r.taller >= r.meta_taller))) return false;
        if (ACTIVE_RUNT_TAB === 'CARGADOS_PRACTICA' && !esCarg) return false;
        if (ACTIVE_RUNT_TAB === 'DESFASE_BOT' && (esCarg || (r.pendiente_subir_bot || 0) <= 0)) return false;
        if (ACTIVE_RUNT_TAB === 'POR_VENCER_90' && (esCarg || !['POR_VENCER', 'VENCIDO'].includes(r.semaforo_90))) return false;
        if (!q) return true;
        return `${r.nombre} ${r.cedula} ${r.categoria} ${r.prioridad} ${r.estado_runt} ${r.cargue_diario} ${r.vigente_texto} ${r.etiqueta_prioridad_op || ''}`.toLowerCase().includes(q);
      });
      // Ordenar por Prioridad de Certificación (1. Terminó práctica sin cargue -> 2. En Práctica -> 3. Por vencer -> 4. Teoría/Taller)
      listP = [...listP].sort((a, b) => {
        const pa = Number(a.orden_prioridad || 4);
        const pb = Number(b.orden_prioridad || 4);
        if (pa !== pb) return pa - pb;
        const da = a.dias_restantes_90 !== null && a.dias_restantes_90 !== undefined ? Number(a.dias_restantes_90) : 999;
        const db = b.dias_restantes_90 !== null && b.dias_restantes_90 !== undefined ? Number(b.dias_restantes_90) : 999;
        return da - db;
      });
      return listP;
    }

    async function copiarCedulasFiltroActualRunt() {
      let rows = obtenerFilasFiltradasPracticaRunt();
      if (RUNT_SELECTED_FILAS.size > 0) {
        const selRows = rows.filter(r => RUNT_SELECTED_FILAS.has(Number(r.fila)));
        if (selRows.length > 0) rows = selRows;
      }
      if (!rows.length) {
        alert('⚠️ No hay alumnos seleccionados o en el filtro actual para copiar.');
        return;
      }
      const lineas = rows.map(r => `${r.cedula}\t${r.nombre}\t${r.categoria}\tTeo:${r.teoria}/${r.meta_teoria}\tTal:${r.taller}/${r.meta_taller}\tPrac:${r.practica}/${r.meta_practica}`).join('\n');
      const soloCedulas = rows.map(r => r.cedula).filter(Boolean).join('\n');
      try {
        await navigator.clipboard.writeText(soloCedulas);
        const toast = document.getElementById('runtSheetToast');
        if (toast) {
          toast.className = 'status-toast ok';
          toast.textContent = `📋 Copiadas ${rows.length} cédulas al portapapeles listas para pegar en el cargue masivo de RUNT.`;
        }
      } catch (_) {
        prompt('Copia las cédulas para RUNT:', soloCedulas || lineas);
      }
    }

    async function marcarCargadoHojaRuntRapido(filaIdx, marcarCargado = true) {
      if (!RUNT_SHEET_DATA || !RUNT_SHEET_DATA.practica) return;
      const item = RUNT_SHEET_DATA.practica.find(x => Number(x.fila) === Number(filaIdx));
      if (!item) return;
      const payload = {
        action: 'runt_guardar_practica',
        fila: item.fila,
        cedula: item.cedula,
        nombre: item.nombre,
        categoria: item.categoria,
        prioridad: marcarCargado ? 'Termino' : 'EN PRACTICA',
        fecha_matricula: item.fecha_matricula,
        teoria: marcarCargado ? item.meta_teoria : item.teoria,
        taller: marcarCargado ? item.meta_taller : item.taller,
        practica: marcarCargado ? item.meta_practica : Math.max(item.practica, item.practica_bot),
        cargue_diario: marcarCargado ? 'PRACTICA' : item.cargue_diario,
        estado_runt: marcarCargado ? 'CARGADO RUNT' : 'PENDIENTE RUNT'
      };
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();
      const toast = document.getElementById('runtSheetToast');
      if (toast) {
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = data.message || '';
      }
      await cargarPanoramaRuntDigital(false);
      await cargarBotOverview();
    }

    async function sincronizarHorasBotHaciaHojaRunt() {
      const toast = document.getElementById('runtSheetToast');
      if (toast) {
        toast.className = 'status-toast ok';
        toast.textContent = '⏳ Cruzando clases prácticas del Bot de Telegram con la hoja RUNT...';
      }
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'runt_sincronizar_bot' })
      });
      const data = await res.json();
      if (toast) {
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = data.message || '';
      }
      await cargarPanoramaRuntDigital(false);
    }

    async function cambiarEstadoPestanaNoRunt(filaIdx, nuevoEstado) {
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'runt_cambiar_no_runt', fila: filaIdx, estado_registro: nuevoEstado })
      });
      const data = await res.json();
      const toast = document.getElementById('runtSheetToast');
      if (toast) {
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = data.message || '';
      }
      await cargarPanoramaRuntDigital(false);
    }

    function renderizarTablaHojaRunt() {
      if (!RUNT_SHEET_DATA) return;
      const q = obtenerBusquedaLimpia('searchRuntSheetInput');
      const thead = document.getElementById('runtSheetThead');
      const tbody = document.getElementById('runtSheetTbody');
      const infoEl = document.getElementById('runtSheetCountInfo');
      if (!thead || !tbody) return;

      if (ACTIVE_RUNT_TAB === 'TAB_NO_RUNT') {
        const rows = (RUNT_SHEET_DATA.no_runt || []).filter(r => {
          if (!q) return true;
          return `${r.nombre} ${r.cedula} ${r.categoria} ${r.estado_registro} ${r.novedad}`.toLowerCase().includes(q);
        });
        infoEl.textContent = `Pestaña "No runt" — Mostrando ${rows.length} alumnos pendientes o verificados de inscripción RUNT`;
        thead.innerHTML = `<tr>
          <th>Fila</th><th>Fecha Reporte</th><th>Nombre y Apellido</th><th>Cédula</th><th>Tipo Doc.</th>
          <th>Cat.</th><th>Fecha Matrícula</th><th>Fecha Registro RUNT</th><th>Estado Registro</th><th>Acción 1 Clic</th>
        </tr>`;
        tbody.innerHTML = rows.map(r => {
          const esReg = r.estado_registro.includes('REGISTRADO') && !r.estado_registro.includes('SIN');
          return `<tr class="${esReg ? 'row-sem-verde' : 'row-sem-rojo'}">
            <td><code>#${r.fila}</code></td>
            <td>${r.timestamp || '-'}</td>
            <td><b>${r.nombre || 'SIN NOMBRE'}</b></td>
            <td><code>${r.cedula}</code></td>
            <td>${r.tipo_documento}</td>
            <td><span class="badge badge-info">${r.categoria}</span></td>
            <td>${r.fecha_matricula || '-'}</td>
            <td>${r.fecha_registro_runt || '-'}</td>
            <td><span class="badge ${esReg ? 'badge-success' : 'badge-danger'}">${r.estado_registro}</span></td>
            <td>
              ${esReg
                ? `<button type="button" class="btn btn-outline" style="padding:4px 8px; font-size:11px;" onclick="cambiarEstadoPestanaNoRunt(${r.fila}, 'SIN REGISTRO')">↩️ Marcar Sin Registro</button>`
                : `<button type="button" class="btn btn-success" style="padding:4px 8px; font-size:11px;" onclick="cambiarEstadoPestanaNoRunt(${r.fila}, 'REGISTRADO')">✅ Marcar REGISTRADO</button>`}
            </td>
          </tr>`;
        }).join('') || '<tr><td colspan="10" style="text-align:center; padding:18px; color:var(--muted);">Sin registros en pestaña No runt.</td></tr>';
        return;
      }

      if (ACTIVE_RUNT_TAB === 'TAB_INICIO_PRACTICA') {
        const rows = (RUNT_SHEET_DATA.inicio_practica || []).filter(r => {
          if (!q) return true;
          return `${r.nombre} ${r.cedula} ${r.categoria} ${r.novedades} ${r.horarios_practica}`.toLowerCase().includes(q);
        });
        infoEl.textContent = `Pestaña "Inicio Practica" — Mostrando ${rows.length} respuestas de disponibilidad de horarios y novedades`;
        thead.innerHTML = `<tr>
          <th>Fecha Formulario</th><th>Alumno</th><th>Cédula</th><th>Categoría</th><th>Novedades (Saldo / Viaje / Embarazo)</th><th>Mis Horarios de Práctica</th>
        </tr>`;
        tbody.innerHTML = rows.map(r => `<tr>
          <td>${r.fecha || '-'}</td>
          <td><b>${r.nombre}</b></td>
          <td><code>${r.cedula}</code></td>
          <td><span class="badge badge-info">${r.categoria}</span></td>
          <td>${r.novedades ? `<span class="badge badge-warning">${r.novedades}</span>` : '-'}</td>
          <td style="white-space:normal; max-width:420px; font-size:11.5px; color:#cbd5e1;">${r.horarios_practica || '-'}</td>
        </tr>`).join('') || '<tr><td colspan="6" style="text-align:center; padding:18px; color:var(--muted);">Sin registros.</td></tr>';
        return;
      }

      if (ACTIVE_RUNT_TAB === 'TAB_CODIGOS_RUNT') {
        const rows = (RUNT_SHEET_DATA.codigos_runt || []).filter(r => {
          if (!q) return true;
          return `${r.nombre} ${r.cedula} ${r.codigo_runt} ${r.responsable}`.toLowerCase().includes(q);
        });
        infoEl.textContent = `Pestaña "Codigo de Runt" — Mostrando ${rows.length} códigos asignados`;
        thead.innerHTML = `<tr>
          <th>N°</th><th>Alumno</th><th>Cédula</th><th>Fecha</th><th>Código RUNT</th><th>Responsable</th>
        </tr>`;
        tbody.innerHTML = rows.map(r => `<tr>
          <td><b>#${r.numeracion || r.fila}</b></td>
          <td><b>${r.nombre || '-'}</b></td>
          <td><code>${r.cedula || '-'}</code></td>
          <td>${r.fecha || '-'}</td>
          <td><code style="color:#fde047; font-weight:700;">${r.codigo_runt || '-'}</code></td>
          <td>${r.responsable ? `<span class="badge badge-purple">${r.responsable}</span>` : '-'}</td>
        </tr>`).join('') || '<tr><td colspan="6" style="text-align:center; padding:18px; color:var(--muted);">Sin códigos.</td></tr>';
        return;
      }

      // Pestaña principal 'Practica' (ordenada por Prioridad Operativa de Certificación y con filtros masivos Teoría/Taller)
      const listP = obtenerFilasFiltradasPracticaRunt();

      const titulosTabRunt = {
        'PENDIENTES_PRACTICA': '📋 Todos los Pendientes RUNT (Ordenados por Prioridad: 1° Terminaron Práctica → 2° En Práctica → 3° Por Vencer → 4° Teoría/Taller)',
        'TERMINARON_SIN_CARGUE': '🔥 PRIORIDAD #1: Alumnos que YA TERMINARON Práctica y Falta el Cargue Manual a RUNT',
        'URGENTES_EN_PRACTICA': '🚨 PRIORIDAD #2: Alumnos Urgentes EN PRÁCTICA (Próximos a Certificar)',
        'MOTO_SIN_25_TEORIA': '🏍️ FILTRO MASIVO MOTO (A2/A1): Alumnos que Aún NO Tienen las 25 Clases Teóricas',
        'MOTO_SIN_3_TALLER': '🔧 FILTRO MASIVO MOTO (A2/A1): Alumnos que Aún NO Tienen las 3 Clases de Taller',
        'CARRO_SIN_TEO_TAL': '🚗 FILTRO MASIVO CARRO (B1/C1): Alumnos sin Completar Teoría o Taller',
        'DESFASE_BOT': '🔄 Alumnos con Horas Prácticas Nuevas en el Bot Pendientes de Subir a RUNT',
        'POR_VENCER_90': '💸 Alumnos Vencidos o Por Vencer (90 Días) — Usa "🔄 Reiniciar + Sanción" si venció el plazo',
        'CARGADOS_PRACTICA': '✅ Alumnos con Cargue 100% Realizado en RUNT',
        'TODOS_PRACTICA': '📚 Histórico Completo Pestaña Practica'
      };
      infoEl.innerHTML = `${titulosTabRunt[ACTIVE_RUNT_TAB] || 'Pestaña "Practica"'} — Mostrando <b>${listP.length}</b> alumnos`;

      const todosMarcados = listP.length > 0 && listP.every(r => RUNT_SELECTED_FILAS.has(Number(r.fila)));
      thead.innerHTML = `<tr>
        <th style="width:36px; text-align:center;">
          <input type="checkbox" ${todosMarcados ? 'checked' : ''} title="Seleccionar / Desmarcar todos los alumnos de este filtro" onchange="toggleSeleccionarTodosRunt(this.checked)" style="width:16px; height:16px; cursor:pointer;" />
        </th>
        <th>Alumno</th>
        <th>Cédula</th>
        <th>Cat.</th>
        <th>Prioridad Operativa</th>
        <th>Fecha Mat.</th>
        <th>Vence 90 Días</th>
        <th>Vigencia</th>
        <th>📘 Teoría (Masivo)</th>
        <th>🔧 Taller (Masivo)</th>
        <th>🚗 Práctica (Individual • Máx 8/día)</th>
        <th>Σ Clases Cargadas</th>
        <th>Actividad RUNT</th>
        <th>Fecha Cargue (Ayer/Hoy/Mañana)</th>
        <th>Acciones Rápidas (Individual / Reinicio Sanción)</th>
      </tr>`;

      tbody.innerHTML = listP.map(r => {
        const esCarg = r.estado_runt === 'CARGADO RUNT';
        let trClass = esCarg ? 'row-sem-verde' : '';
        if (!esCarg && (r.semaforo_90 === 'VENCIDO' || r.es_terminado_sin_cargue)) trClass = 'row-sem-rojo';
        else if (!esCarg && (r.semaforo_90 === 'POR_VENCER' || r.es_urgente_en_practica || r.pendiente_subir_bot > 0)) trClass = 'row-sem-amarillo';

        let badgeVig = `<span class="badge badge-success">${r.vigente_texto || '✅ VIGENTE'}</span>`;
        if (r.semaforo_90 === 'VENCIDO') {
          badgeVig = `<span class="badge badge-danger">❌ VENCIDO ${r.dias_restantes_90 !== null ? '(' + r.dias_restantes_90 + 'd)' : ''}</span>`;
        } else if (r.semaforo_90 === 'POR_VENCER') {
          badgeVig = `<span class="badge badge-warning">⚠️ POR VENCER (${r.dias_restantes_90}d)</span>`;
        } else if (r.dias_restantes_90 !== null) {
          badgeVig = `<span class="badge badge-success">✅ VIGENTE (${r.dias_restantes_90}d)</span>`;
        }

        let badgePrioridad = `<span class="badge">${r.prioridad || '-'}</span>`;
        if (r.es_terminado_sin_cargue) {
          badgePrioridad = `<span class="badge badge-danger" style="background:rgba(249,115,22,0.28); color:#fed7aa; border-color:#f97316; font-weight:800;">🔥 1. TERMINÓ (FALTA RUNT)</span>`;
        } else if (r.es_urgente_en_practica) {
          badgePrioridad = `<span class="badge badge-info" style="font-weight:800;">🚨 2. EN PRÁCTICA (URGENTE)</span>`;
        } else if (esCarg) {
          badgePrioridad = `<span class="badge badge-success">✅ CERTIFICADO</span>`;
        }

        const teoFalta = r.teoria < r.meta_teoria;
        const talFalta = r.taller < r.meta_taller;
        const pracFalta = r.practica < r.meta_practica;

        const pracHtml = r.pendiente_subir_bot > 0
          ? `<b style="color:#fde047;">${r.practica}/${r.meta_practica}</b> <span class="badge badge-info" title="El Bot de Clases ya tiene ${r.practica_bot}h dictadas (+${r.pendiente_subir_bot}h por subir a RUNT)">🤖 Bot: ${r.practica_bot}h (+${r.pendiente_subir_bot})</span>`
          : `<b style="color:${r.practica >= r.meta_practica ? '#34d399' : '#fde047'};">${r.practica}/${r.meta_practica}</b>` + (pracFalta ? ` <span style="font-size:10.5px; color:#94a3b8;">(Faltan ${r.practica_faltante}h)</span>` : '');

        const safeNom = (r.nombre || '').replace(/'/g, "\\'");
        const btnPrac8 = (!esCarg && pracFalta)
          ? `<button type="button" class="btn btn-outline" style="padding:3px 7px; font-size:10.5px; border-color:#38bdf8; color:#38bdf8;" title="Reportar lote individual de hasta +8 clases prácticas en la fecha seleccionada arriba" onclick="sumarLotePracticaRunt(${r.fila}, 8, false)">+8h Prác.</button>`
          : '';
        const btnPrac2 = (!esCarg && pracFalta)
          ? `<button type="button" class="btn btn-outline" style="padding:3px 7px; font-size:10.5px;" title="Reportar +2 clases prácticas en la fecha seleccionada arriba" onclick="sumarLotePracticaRunt(${r.fila}, 2, false)">+2h</button>`
          : '';
        const btnReinicioSancion = (!esCarg && (r.semaforo_90 === 'VENCIDO' || r.semaforo_90 === 'POR_VENCER' || ACTIVE_RUNT_TAB === 'POR_VENCER_90'))
          ? `<button type="button" class="btn btn-danger" style="padding:3px 7px; font-size:10.5px;" title="El curso venció: Reinicia las horas RUNT a 0, crea nuevo plazo de 90 días y registra la sanción económica al alumno" onclick="reiniciarCursoVencidoRunt(${r.fila}, '${safeNom}', '${r.categoria}')">💸 Reiniciar + Sanción</button>`
          : '';
        const isChecked = RUNT_SELECTED_FILAS.has(Number(r.fila));

        return `<tr class="${trClass}">
          <td style="text-align:center;">
            <input type="checkbox" ${isChecked ? 'checked' : ''} onchange="toggleSeleccionFilaRunt(${r.fila}, this.checked)" style="width:16px; height:16px; cursor:pointer;" />
          </td>
          <td><b>${r.nombre}</b></td>
          <td><code>${r.cedula}</code></td>
          <td><span class="badge badge-info">${r.categoria_raw || r.categoria}</span></td>
          <td>${badgePrioridad}</td>
          <td>${r.fecha_matricula || '-'}</td>
          <td style="color:#fde047; font-weight:600;">${r.vence_90_dias || '-'}</td>
          <td>${badgeVig}</td>
          <td style="text-align:center;">
            <b style="color:${!teoFalta ? '#34d399' : '#f87171'};">${r.teoria}/${r.meta_teoria}</b>
            ${teoFalta ? `<div style="font-size:10px; color:#fca5a5;">Faltan ${r.teoria_faltante}h</div>` : ''}
          </td>
          <td style="text-align:center;">
            <b style="color:${!talFalta ? '#34d399' : '#f87171'};">${r.taller}/${r.meta_taller}</b>
            ${talFalta ? `<div style="font-size:10px; color:#fca5a5;">Faltan ${r.taller_faltante}h</div>` : ''}
          </td>
          <td>${pracHtml}</td>
          <td>
            <b>${r.clases_cargadas}/${r.meta_total}</b> (${r.porcentaje_runt}%)
          </td>
          <td>
            <span class="badge ${esCarg ? 'badge-success' : 'badge-danger'}">${esCarg ? '✅ CARGADO RUNT' : '⚠️ PENDIENTE RUNT'}</span>
          </td>
          <td>
            ${r.fecha_cargue ? `<span style="font-size:11px; color:#fde047; font-weight:600;">📅 ${r.fecha_cargue}</span>` : '<span style="font-size:11px; color:var(--muted);">-</span>'}
            ${r.cargue_diario ? `<div><span class="badge badge-info" style="font-size:10px;">${r.cargue_diario}</span></div>` : ''}
          </td>
          <td>
            <div style="display:flex; gap:4px; flex-wrap:wrap; align-items:center;">
              ${btnPrac2}
              ${btnPrac8}
              <button type="button" class="btn btn-outline" style="padding:3px 7px; font-size:10.5px;" onclick="cargarFilaRuntEnFormulario(${r.fila})">✏️ Editar</button>
              ${esCarg
                ? `<button type="button" class="btn btn-outline" style="padding:3px 7px; font-size:10.5px;" onclick="marcarCargadoHojaRuntRapido(${r.fila}, false)">↩️ Pendiente</button>`
                : `<button type="button" class="btn btn-success" style="padding:3px 7px; font-size:10.5px;" title="Ponerse al día (100% Teoría+Taller+Práctica) y marcar ✅ CARGADO RUNT" onclick="marcarCargadoHojaRuntRapido(${r.fila}, true)">✅ CARGADO RUNT</button>`}
              ${btnReinicioSancion}
            </div>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="14" style="text-align:center; padding:18px; color:var(--muted);">No hay alumnos que coincidan con este filtro en la hoja RUNT.</td></tr>';
    }

    function cambiarSubTabClases(filtroVal, btnEl) {
      const sel = document.getElementById('filtroAlertaCartera');
      if (sel) sel.value = filtroVal;
      document.querySelectorAll('#clasesTabsHeader .tab-btn').forEach(b => b.classList.remove('active'));
      const mapIdx = {
        'PENDIENTE_RUNT': 0,
        'SEMAFORO_ROJO': 1,
        'SEMAFORO_AMARILLO': 2,
        'CARGADOS_RUNT': 3,
        'EN_CURSO': 4,
        'FINALIZADOS': 5,
        'CON_SALDO': 6,
        'CERO_REAL': 7,
        'TODOS': 8
      };
      if (btnEl) {
        btnEl.classList.add('active');
      } else {
        const btns = document.querySelectorAll('#clasesTabsHeader .tab-btn');
        if (btns[mapIdx[filtroVal] || 0]) btns[mapIdx[filtroVal] || 0].classList.add('active');
      }
      if (filtroVal === 'SEMAFORO_ROJO' || filtroVal === 'SEMAFORO_AMARILLO') {
        const sortSel = document.getElementById('sortClasesSelect');
        if (sortSel) sortSel.value = 'SEMAFORO_URGENCIA';
      }
      renderizarModuloClases();
    }

    function renderizarModuloClases() {
      if (!BOT_DATA || !BOT_DATA.clases) return;
      const kc = BOT_DATA.clases.kpis;
      const todosAl = BOT_DATA.clases.alumnos || [];

      const esCargadoRunt = (a) => String(a.estado_runt || '').toLowerCase().includes('cargue') || String(a.estado_runt || '').toLowerCase().includes('cargado');

      // Contadores para las sub-pestañas (con control manual RUNT por categoría y Semáforo de 90 Días)
      const cntRuntPend = todosAl.filter(a => !esCargadoRunt(a)).length;
      const cntRuntOk = todosAl.filter(a => esCargadoRunt(a)).length;
      const cntActivos = todosAl.filter(a => a.estado_curso !== 'FINALIZADO' && !esCargadoRunt(a)).length;
      const cntFin = todosAl.filter(a => a.estado_curso === 'FINALIZADO' && !esCargadoRunt(a)).length;
      const cntSaldo = todosAl.filter(a => !esCargadoRunt(a) && (a.saldo_cartera || 0) > 0).length;
      const cntCero = todosAl.filter(a => !esCargadoRunt(a) && a.alerta_cartera === 'RIESGO_CERO_REAL').length;
      const cntSemRojo = todosAl.filter(a => !esCargadoRunt(a) && a.estado_curso !== 'FINALIZADO' && a.semaforo_color === 'ROJO').length;
      const cntSemAmarillo = todosAl.filter(a => !esCargadoRunt(a) && a.estado_curso !== 'FINALIZADO' && a.semaforo_color === 'AMARILLO').length;
      const cntSemVerde = todosAl.filter(a => !esCargadoRunt(a) && (a.semaforo_color === 'VERDE' || a.estado_curso === 'FINALIZADO')).length;

      const elRuntPend = document.getElementById('cnt_cl_runt_pend');
      const elSemRojo = document.getElementById('cnt_cl_sem_rojo');
      const elSemAmarillo = document.getElementById('cnt_cl_sem_amarillo');
      const elRuntOk = document.getElementById('cnt_cl_runt_ok');
      const elAct = document.getElementById('cnt_cl_activos');
      const elFin = document.getElementById('cnt_cl_fin');
      const elSal = document.getElementById('cnt_cl_saldo');
      const elCer = document.getElementById('cnt_cl_cero');
      const elTod = document.getElementById('cnt_cl_todos');
      if (elRuntPend) elRuntPend.textContent = cntRuntPend;
      if (elSemRojo) elSemRojo.textContent = cntSemRojo;
      if (elSemAmarillo) elSemAmarillo.textContent = cntSemAmarillo;
      if (elRuntOk) elRuntOk.textContent = cntRuntOk;
      if (elAct) elAct.textContent = cntActivos;
      if (elFin) elFin.textContent = cntFin;
      if (elSal) elSal.textContent = cntSaldo;
      if (elCer) elCer.textContent = cntCero;
      if (elTod) elTod.textContent = todosAl.length;

      document.getElementById('kpiClasesGrid').innerHTML = `
        <div class="kpi-card danger" style="cursor:pointer;" onclick="cambiarSubTabClases('SEMAFORO_ROJO')" title="Haz clic para ver los alumnos en riesgo crítico de perder el curso por vencimiento de los 90 días">
          <div class="kpi-label">🔴 Semáforo Crítico / Próximos a Vencer (90 Días)</div>
          <div class="kpi-value" style="color:#fca5a5;">${cntSemRojo} En Riesgo</div>
          <div class="kpi-sub"><b>🟡 ${cntSemAmarillo} en alerta preventiva</b> • <b>🟢 ${cntSemVerde} a tiempo</b> (clic para filtrar)</div>
        </div>
        <div class="kpi-card warning" style="cursor:pointer;" onclick="cambiarSubTabClases('PENDIENTE_RUNT')">
          <div class="kpi-label">⏳ Pendientes Cargue RUNT (No Registrado)</div>
          <div class="kpi-value" style="color:#fde68a;">${cntRuntPend} Matrículas</div>
          <div class="kpi-sub"><b>${cntActivos} en curso</b> y <b>${cntFin} listos (100%)</b> por subir a RUNT</div>
        </div>
        <div class="kpi-card success" style="cursor:pointer;" onclick="cambiarSubTabClases('CARGADOS_RUNT')">
          <div class="kpi-label">✅ Ya Cargados en RUNT (Ocultos del Filtro)</div>
          <div class="kpi-value" style="color:#34d399;">${cntRuntOk} Cargados</div>
          <div class="kpi-sub">Control independiente por categoría (<b>A2</b> vs <b>B1/C1</b>)</div>
        </div>
        <div class="kpi-card" style="cursor:pointer;" onclick="cambiarSubTabClases('CON_SALDO')">
          <div class="kpi-label">⚠️ Pendientes RUNT con Saldo en Cartera</div>
          <div class="kpi-value" style="color:#60a5fa;">${cntSaldo} Alumnos</div>
          <div class="kpi-sub"><b>${kc.total_sesiones_registradas} sesiones</b> registradas en total</div>
        </div>
      `;

      const q = obtenerBusquedaLimpia('searchClasesInput');
      const catF = document.getElementById('filtroCatClases').value;
      const alertaF = document.getElementById('filtroAlertaCartera').value || 'PENDIENTE_RUNT';

      // Sincronizar pestaña visual activa con el select
      const mapIdx = {
        'PENDIENTE_RUNT': 0,
        'SEMAFORO_ROJO': 1,
        'SEMAFORO_AMARILLO': 2,
        'CARGADOS_RUNT': 3,
        'EN_CURSO': 4,
        'FINALIZADOS': 5,
        'CON_SALDO': 6,
        'CERO_REAL': 7,
        'TODOS': 8
      };
      const tabBtns = document.querySelectorAll('#clasesTabsHeader .tab-btn');
      tabBtns.forEach((b, idx) => {
        if (idx === mapIdx[alertaF]) b.classList.add('active');
        else b.classList.remove('active');
      });

      let rows = todosAl.filter(a => {
        if (catF !== 'TODAS' && a.categoria !== catF) return false;
        const yaRunt = esCargadoRunt(a);
        if (alertaF === 'PENDIENTE_RUNT' && yaRunt) return false;
        if (alertaF === 'SEMAFORO_ROJO' && (yaRunt || a.estado_curso === 'FINALIZADO' || a.semaforo_color !== 'ROJO')) return false;
        if (alertaF === 'SEMAFORO_AMARILLO' && (yaRunt || a.estado_curso === 'FINALIZADO' || a.semaforo_color !== 'AMARILLO')) return false;
        if (alertaF === 'CARGADOS_RUNT' && !yaRunt) return false;
        if (alertaF === 'EN_CURSO' && (a.estado_curso === 'FINALIZADO' || yaRunt)) return false;
        if (alertaF === 'FINALIZADOS' && (a.estado_curso !== 'FINALIZADO' || yaRunt)) return false;
        if (alertaF === 'CON_SALDO' && (yaRunt || (a.saldo_cartera || 0) <= 0)) return false;
        if (alertaF === 'CERO_REAL' && (yaRunt || a.alerta_cartera !== 'RIESGO_CERO_REAL')) return false;
        if (!q) return true;
        return `${a.cedula} ${a.nombre} ${a.categoria} ${a.formato_codigo} ${a.estado_runt || ''} ${a.semaforo_etiqueta || ''} ${a.semaforo_detalle || ''}`.toLowerCase().includes(q);
      });

      const sortClases = (document.getElementById('sortClasesSelect') || {}).value || 'RECIENTES';
      if (sortClases === 'ANTIGUOS') {
        rows = [...rows].reverse();
      } else if (sortClases === 'AVANCE_DESC') {
        rows = [...rows].sort((a, b) => Number(b.porcentaje_avance || 0) - Number(a.porcentaje_avance || 0));
      } else if (sortClases === 'SEMAFORO_URGENCIA') {
        rows = [...rows].sort((a, b) => Number(b.urgencia_score || 0) - Number(a.urgencia_score || 0));
      } else if (sortClases === 'MENOS_DIAS') {
        rows = [...rows].sort((a, b) => {
          const da = a.estado_curso === 'FINALIZADO' ? 9999 : Number(a.dias_restantes_90 ?? 90);
          const db = b.estado_curso === 'FINALIZADO' ? 9999 : Number(b.dias_restantes_90 ?? 90);
          return da - db;
        });
      }

      const etiquetasVista = {
        'PENDIENTE_RUNT': '⏳ Alumnos Matriculados Pendientes de Cargue RUNT (No Registrado RUNT)',
        'SEMAFORO_ROJO': '🔴 SEMÁFORO CRÍTICO (90 DÍAS): Alumnos Próximos a Vencer o con Alto % de Clases Faltantes',
        'SEMAFORO_AMARILLO': '🟡 SEMÁFORO EN ALERTA PREVENTIVA (90 DÍAS): Alumnos que Requieren Acelerar Programación',
        'CARGADOS_RUNT': '✅ Alumnos con Cargue RUNT Realizado',
        'EN_CURSO': '🟢 Alumnos Activos en Curso (< 100% y Pendientes RUNT)',
        'FINALIZADOS': '🎓 Alumnos con 100% Clases Terminadas (Pendientes de Cargue RUNT)',
        'CON_SALDO': '⚠️ Alumnos Pendientes RUNT con Saldo en Cartera',
        'CERO_REAL': '🚨 Alumnos con Abono $0 Global',
        'TODOS': '📚 Histórico Completo (Pendientes RUNT + Cargados RUNT)'
      };
      document.getElementById('clasesCountInfo').innerHTML =
        `${etiquetasVista[alertaF] || ''} — Mostrando <b>${rows.length}</b> (<span style="color:#fca5a5;">🔴 ${cntSemRojo} críticos</span> • <span style="color:#fde047;">🟡 ${cntSemAmarillo} en alerta</span> • <span style="color:#6ee7b7;">🟢 ${cntSemVerde} a tiempo</span>)`;

      document.getElementById('clasesTbody').innerHTML = rows.map(a => {
        let badgeCartera = `<span class="badge badge-success">${a.alerta_texto || '✅ PAZ Y SALVO'}</span>`;
        if (a.alerta_cartera === 'RIESGO_CERO_REAL') {
          badgeCartera = `<span class="badge badge-danger">${a.alerta_texto}</span>`;
        } else if (a.alerta_cartera === 'MORA_CRITICA_90') {
          badgeCartera = `<span class="badge badge-danger">${a.alerta_texto}</span>`;
        } else if (a.alerta_cartera === 'CON_SALDO') {
          badgeCartera = `<span class="badge badge-warning">${a.alerta_texto}</span>`;
        } else if (a.alerta_cartera === 'NO_CRUZADO') {
          badgeCartera = `<span class="badge">Curso Práctico Activo</span>`;
        }

        const yaRunt = esCargadoRunt(a);
        const safeNombre = (a.nombre || '').replace(/'/g, "\\'");
        const safeCedula = (a.cedula || a.nombre || '').replace(/'/g, "\\'");
        const badgeRunt = yaRunt
          ? `<span class="badge badge-success" title="${a.fecha_runt || ''}">✅ Cargue RUNT</span>`
          : `<span class="badge badge-warning">⏳ No Registrado RUNT</span>`;
        const btnRuntAccion = yaRunt
          ? `<button type="button" class="btn btn-outline" style="padding:4px 8px; font-size:11px;" title="Devolver a No Registrado RUNT para que vuelva a aparecer en pendientes" onclick="cambiarEstadoRuntWeb('${safeCedula}', '${safeNombre}', '${a.categoria}', 'No Registrado RUNT')">↩️ Quitar RUNT</button>`
          : `<button type="button" class="btn btn-success" style="padding:4px 8px; font-size:11px;" title="Confirmar Cargue RUNT de esta categoría (${a.categoria}) y ocultarlo del filtro de pendientes" onclick="cambiarEstadoRuntWeb('${safeCedula}', '${safeNombre}', '${a.categoria}', 'Cargue RUNT')">✅ Cargue RUNT</button>`;

        // Renderizado visual del Semáforo de 90 Días (Desde 1ra Clase Práctica)
        const semColor = a.semaforo_color || 'AZUL';
        let semBadgeClass = 'badge-info';
        let semBarColor = '#3b82f6';
        let trSemClass = '';
        if (semColor === 'ROJO') {
          semBadgeClass = 'badge-danger';
          semBarColor = '#ef4444';
          trSemClass = 'row-sem-rojo';
        } else if (semColor === 'AMARILLO') {
          semBadgeClass = 'badge-warning';
          semBarColor = '#f59e0b';
          trSemClass = 'row-sem-amarillo';
        } else if (semColor === 'VERDE') {
          semBadgeClass = 'badge-success';
          semBarColor = '#10b981';
          trSemClass = 'row-sem-verde';
        }
        const diasTrans = Math.min(90, Math.max(0, Number(a.dias_transcurridos_90 || 0)));
        const pctTiempo90 = Math.min(100, Math.round((diasTrans / 90) * 100));
        const infoFechas90 = a.fecha_primera_clase
          ? `<div style="font-size:10.5px; color:#94a3b8; margin-top:2px;">1ra clase: <b>${a.fecha_primera_clase}</b> → Vence: <b style="color:#e2e8f0;">${a.fecha_vencimiento_90 || '-'}</b> (${a.dias_transcurridos_90}d/90d)</div>`
          : `<div style="font-size:10.5px; color:#94a3b8; margin-top:2px;">Sin 1ra clase fechada • Faltan ${a.clases_faltantes || 0} clases (${a.porcentaje_faltante || 0}%)</div>`;
        const semaforoCellHtml = `
          <div style="min-width:195px;">
            <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
              <span class="badge ${semBadgeClass}" style="font-weight:800;">${a.semaforo_etiqueta || '⚪ En Curso'}</span>
              <span style="font-size:11px; color:${semColor === 'ROJO' ? '#fca5a5' : (semColor === 'AMARILLO' ? '#fde047' : '#cbd5e1')}; font-weight:700;">${a.semaforo_detalle || ''}</span>
            </div>
            <div class="semaforo-bar-wrap" title="Tiempo consumido del plazo de 90 días: ${diasTrans}/90 días (${pctTiempo90}%) vs Avance de clases: ${a.porcentaje_avance}%">
              <div class="semaforo-bar-fill" style="width:${pctTiempo90}%; background:${semBarColor};"></div>
            </div>
            ${infoFechas90}
          </div>
        `;

        const qParam = encodeURIComponent(a.cedula || a.nombre);
        const catParam = encodeURIComponent(a.categoria || 'A2');

        return `<tr class="${trSemClass}">
          <td><code>${a.cedula || 'S/D'}</code></td>
          <td><b>${a.nombre}</b></td>
          <td><span class="badge badge-info">${a.categoria}</span></td>
          <td>${semaforoCellHtml}</td>
          <td>
            <div style="display:flex; gap:6px; align-items:center; flex-wrap:wrap;">
              ${badgeRunt}
              ${btnRuntAccion}
            </div>
          </td>
          <td><b>${a.formato_codigo}</b></td>
          <td>
            <div class="progress-bar-wrap"><div class="progress-bar-fill" style="width:${a.porcentaje_avance}%;"></div></div>
            <b>${a.clases_completadas}/${a.total_requerido}</b> (${a.porcentaje_avance}%)
            <div style="font-size:10.5px; color:${(a.porcentaje_faltante || 0) >= 80 ? '#fca5a5' : '#94a3b8'};">Faltan <b>${a.clases_faltantes ?? (a.total_requerido - a.clases_completadas)}</b> cl. (${a.porcentaje_faltante ?? (100 - a.porcentaje_avance)}%)</div>
          </td>
          <td><span class="badge ${a.estado_curso === 'FINALIZADO' ? 'badge-success' : 'badge-purple'}">${a.estado_curso}</span></td>
          <td>${badgeCartera}</td>
          <td class="${(a.saldo_cartera || 0) > 0 ? 'money-neg' : 'money-pos'}">${fmtCOP(a.saldo_cartera || 0)}</td>
          <td>
            <div style="display:flex; gap:5px; align-items:center; flex-wrap:wrap;">
              <button type="button" class="btn btn-outline" style="padding:5px 8px; font-size:11px;" title="Cargar datos del alumno arriba para editar, cambiar RUNT o registrar clase" onclick="cargarAlumnoEnFormularios('${safeCedula}', '${safeNombre}', '${a.categoria}', ${a.clases_completadas || 0}, ${a.total_requerido || 15}, '${a.estado_runt || 'No Registrado RUNT'}')">✏️ Editar / +Clase</button>
              <a href="/api/bot/pdf/alumno?q=${qParam}&cat=${catParam}" class="btn btn-primary" style="padding:5px 9px; font-size:11px;" title="Descargar Ficha Oficial PDF">📄 PDF ${a.formato_codigo}</a>
              <button type="button" class="btn btn-outline" style="padding:5px 7px; font-size:11px;" title="Enviar PDF F-59/F-62 al Telegram del CEA" onclick="enviarPdfAlumnoTelegram('${safeCedula}', '${safeNombre}', '${a.categoria}')">✈️</button>
              <button type="button" class="btn btn-success" style="padding:5px 7px; font-size:11px;" title="Enviar recordatorio de vencimiento 90 días y avance por WhatsApp al alumno" onclick="enviarInfoAlumnoWhatsApp('${safeCedula}', '${safeNombre}', '${a.categoria}', '${a.clases_completadas}/${a.total_requerido} (${a.porcentaje_avance}%)', ${a.saldo_cartera || 0}, '${a.dias_restantes_90 ?? 90}', '${a.fecha_vencimiento_90 || ''}', '${a.clases_faltantes ?? 0}')">📲</button>
            </div>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="11" style="text-align:center; color:var(--muted); padding:20px;">No hay alumnos que coincidan con el filtro seleccionado.</td></tr>';
    }

    function renderizarModuloNomina() {
      if (!BOT_DATA || !BOT_DATA.nomina) return;
      const nm = BOT_DATA.nomina;

      // Poblar selector de instructores conservando selección
      const selInst = document.getElementById('nom_instructor');
      const currentVal = selInst.value || 'TODOS';
      const listaInst = nm.instructores_disponibles || [];
      selInst.innerHTML = '<option value="TODOS">👥 Todos los Instructores</option>' +
        listaInst.map(i => `<option value="${i}" ${i === currentVal ? 'selected' : ''}>👨‍🏫 ${i}</option>`).join('');

      document.getElementById('kpiNominaGrid').innerHTML = `
        <div class="kpi-card">
          <div class="kpi-label">Periodo Consultado</div>
          <div class="kpi-value" style="font-size:18px;">${nm.titulo_periodo}</div>
          <div class="kpi-sub">Filtro: <b>${nm.instructor_filtro}</b></div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">Total Clases Prácticas (Horas)</div>
          <div class="kpi-value" style="color:#c4b5fd;">${nm.total_clases} Clases</div>
          <div class="kpi-sub"><b>${nm.total_sesiones} sesiones</b> registradas</div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">✅ Envíos A Tiempo (Min 18-28 / 40-50)</div>
          <div class="kpi-value" style="color:#34d399;">${nm.a_tiempo}</div>
          <div class="kpi-sub">Cumplen ventana oficial del CEA</div>
        </div>
        <div class="kpi-card danger">
          <div class="kpi-label">⚠️ Envíos Fuera de Horario</div>
          <div class="kpi-value" style="color:#fca5a5;">${nm.fuera_horario}</div>
          <div class="kpi-sub">Antes o después del minuto permitido</div>
        </div>
      `;

      const perCode = document.getElementById('nom_periodo').value || 'mes';
      document.getElementById('nominaResumenTbody').innerHTML = (nm.resumen_instructores || []).map(r => {
        const catsTxt = Object.entries(r.categorias || {}).map(([k, v]) => `${k}: <b>${v}</b>`).join(' • ');
        const instEnc = encodeURIComponent(r.instructor);
        return `<tr>
          <td><b>👨‍🏫 ${r.instructor}</b></td>
          <td>${r.sesiones} sesiones</td>
          <td><span class="badge badge-purple" style="font-size:12.5px;">${r.clases_dictadas} clases</span></td>
          <td><span class="badge badge-success">${r.a_tiempo} a tiempo</span></td>
          <td><span class="badge ${r.fuera_horario > 0 ? 'badge-danger' : ''}">${r.fuera_horario} fuera de hora</span></td>
          <td>${catsTxt || 'S/D'}</td>
          <td>
            <a href="/api/bot/pdf/nomina?instructor=${instEnc}&periodo=${perCode}" class="btn btn-outline" style="padding:5px 10px; font-size:11.5px;">📄 PDF Liquidación</a>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="7" style="text-align:center; color:var(--muted);">Sin registros en este periodo</td></tr>';

      const q = obtenerBusquedaLimpia('searchNominaInput');
      const detRows = (nm.detalle_clases || []).filter(d => {
        if (!q) return true;
        return `${d.fecha} ${d.instructor} ${d.alumno} ${d.cedula} ${d.categoria} ${d.estado_validacion}`.toLowerCase().includes(q);
      });

      document.getElementById('nominaDetalleTbody').innerHTML = detRows.map(d => {
        const okHora = d.estado_validacion.includes('A TIEMPO');
        return `<tr>
          <td>${d.fecha}</td>
          <td><code>${d.hora}</code></td>
          <td><b>${d.instructor}</b></td>
          <td><code>${d.cedula}</code></td>
          <td>${d.alumno}</td>
          <td><span class="badge">${d.categoria}</span></td>
          <td><b>${d.clases}</b></td>
          <td>${d.num_clases}</td>
          <td><span class="badge ${okHora ? 'badge-success' : 'badge-warning'}">${d.estado_validacion}</span></td>
          <td style="color:var(--muted); font-size:11.5px;">${d.detalle_validacion}</td>
        </tr>`;
      }).join('') || '<tr><td colspan="10" style="text-align:center; color:var(--muted);">No hay sesiones registradas para este filtro.</td></tr>';

      const cierres = BOT_DATA.cierres_jornada || [];
      document.getElementById('cntCierresBadge').textContent = `${cierres.length} cierres`;
      document.getElementById('cierresTbody').innerHTML = cierres.map(c => `<tr>
        <td>${c.fecha}</td>
        <td><code>${c.hora}</code></td>
        <td><b>${c.instructor}</b></td>
        <td><span class="badge badge-success">${c.estado || 'CERRADA'}</span></td>
        <td>${c.drive_link ? `<a href="${c.drive_link}" target="_blank" style="color:#60a5fa;">🔗 Ver Foto Agenda</a>` : 'Registrada en Bot'}</td>
      </tr>`).join('') || '<tr><td colspan="5" style="text-align:center; color:var(--muted);">Sin cierres de jornada recientes.</td></tr>';
    }

    function renderizarModuloInasistencias() {
      if (!BOT_DATA || !BOT_DATA.inasistencias) return;
      const ki = BOT_DATA.inasistencias.kpis;
      document.getElementById('kpiInasGrid').innerHTML = `
        <div class="kpi-card">
          <div class="kpi-label">Total Inasistencias (${BOT_DATA.inasistencias.titulo_periodo})</div>
          <div class="kpi-value">${ki.total}</div>
          <div class="kpi-sub">Registradas en Bot y Google Sheets</div>
        </div>
        <div class="kpi-card danger">
          <div class="kpi-label">⏳ Pendientes de Pago</div>
          <div class="kpi-value" style="color:#fca5a5;">${ki.pendientes}</div>
          <div class="kpi-sub">Alumnos por cancelar multa de inasistencia</div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">💵 Cancelaron en Efectivo</div>
          <div class="kpi-value" style="color:#34d399;">${ki.efectivo}</div>
          <div class="kpi-sub">Pagadas en caja física</div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">🏦 Cancelaron por Transferencia</div>
          <div class="kpi-value" style="color:#c4b5fd;">${ki.transferencia}</div>
          <div class="kpi-sub">Pagadas por banco / Nequi / Daviplata</div>
        </div>
      `;

      const q = obtenerBusquedaLimpia('searchInasInput');
      const rows = (BOT_DATA.inasistencias.registros || []).filter(r => {
        if (!q) return true;
        return `${r.fecha} ${r.alumno} ${r.categoria} ${r.instructor} ${r.estado_pago}`.toLowerCase().includes(q);
      });

      document.getElementById('inasCountInfo').textContent = `Mostrando ${rows.length} inasistencias`;

      document.getElementById('inasTbody').innerHTML = rows.map(r => {
        const estUp = (r.estado_pago || '').toUpperCase();
        let badgeEst = `<span class="badge badge-danger">⏳ ${r.estado_pago}</span>`;
        if (estUp.includes('EFECTIVO')) badgeEst = `<span class="badge badge-success">💵 ${r.estado_pago}</span>`;
        else if (estUp.includes('TRASFERENCIA') || estUp.includes('TRANSFERENCIA')) badgeEst = `<span class="badge badge-info">🏦 ${r.estado_pago}</span>`;

        const alEnc = encodeURIComponent(r.alumno);
        return `<tr>
          <td><code>#${r.fila}</code></td>
          <td><b>${r.fecha}</b></td>
          <td><b>${r.alumno}</b></td>
          <td><span class="badge">${r.categoria}</span></td>
          <td>${r.hora}</td>
          <td>${r.instructor}</td>
          <td>${badgeEst}</td>
          <td>
            <div style="display:flex; gap:4px;">
              <button type="button" class="btn btn-success" style="padding:4px 8px; font-size:11px;" onclick="cambiarEstadoInasistencia(${r.fila}, 'Cancelo Efectivo')">💵 Efectivo</button>
              <button type="button" class="btn btn-primary" style="padding:4px 8px; font-size:11px;" onclick="cambiarEstadoInasistencia(${r.fila}, 'Cancelo Trasferencia')">🏦 Transf.</button>
              <button type="button" class="btn btn-outline" style="padding:4px 8px; font-size:11px;" onclick="cambiarEstadoInasistencia(${r.fila}, 'Pendiente Pago')">⏳ Pend.</button>
            </div>
          </td>
          <td>
            <a href="/api/bot/pdf/inasistencias?alumno=${alEnc}" class="btn btn-outline" style="padding:4px 9px; font-size:11.5px;">📄 PDF Alumno</a>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="9" style="text-align:center; color:var(--muted);">No hay inasistencias para este filtro.</td></tr>';
    }

    function renderizarModuloPesv() {
      if (!BOT_DATA || !BOT_DATA.mantenimientos) return;
      const km = BOT_DATA.mantenimientos.kpis;
      const flota = BOT_DATA.mantenimientos.flota_activa || [];

      document.getElementById('kpiPesvGrid').innerHTML = `
        <div class="kpi-card">
          <div class="kpi-label">Total Intervenciones (${BOT_DATA.mantenimientos.titulo_periodo})</div>
          <div class="kpi-value">${km.total}</div>
          <div class="kpi-sub">Formato Oficial <b>SISI-F2-08</b></div>
        </div>
        <div class="kpi-card success">
          <div class="kpi-label">🟢 Mantenimientos Preventivos</div>
          <div class="kpi-value" style="color:#34d399;">${km.preventivos}</div>
          <div class="kpi-sub">Cambio aceite, filtros, alineación, revisión</div>
        </div>
        <div class="kpi-card orange">
          <div class="kpi-label">🟠 Mantenimientos Correctivos</div>
          <div class="kpi-value" style="color:#fdba74;">${km.correctivos}</div>
          <div class="kpi-sub">Reparación de frenos, embrague, llantas, etc.</div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">🚘 Vehículos Activos en Flota PESV</div>
          <div class="kpi-value" style="color:#c4b5fd;">${flota.length} Vehículos</div>
          <div class="kpi-sub">${km.vehiculos_intervenidos} intervenidos en el periodo</div>
        </div>
      `;

      // Poblar selectores de placas
      const selMtPlaca = document.getElementById('mt_placa');
      selMtPlaca.innerHTML = flota.map(v => `<option value="${v.placa}">${v.icono} ${v.placa} (${v.categoria} - ${v.marca})</option>`).join('');

      const selFiltroPlaca = document.getElementById('pesv_placa_filtro');
      const curPlaca = selFiltroPlaca.value || 'TODAS';
      selFiltroPlaca.innerHTML = '<option value="TODAS">🚘 Todas las Placas</option>' +
        flota.map(v => `<option value="${v.placa}" ${v.placa === curPlaca ? 'selected' : ''}>${v.icono} ${v.placa}</option>`).join('');

      // Chips de flota activa con opción de retiro
      document.getElementById('flotaChipsWrap').innerHTML = flota.map(v => `
        <span class="badge badge-info" style="padding:5px 10px; font-size:12px; display:inline-flex; align-items:center; gap:6px;">
          ${v.icono} <b>${v.placa}</b> (${v.categoria})
          <button type="button" onclick="retirarVehiculoWeb('${v.placa}')" title="Retirar vehículo de flota activa" style="background:none; border:none; color:#fca5a5; cursor:pointer; font-weight:800;">×</button>
        </span>
      `).join('');

      const q = obtenerBusquedaLimpia('searchPesvInput');
      const rows = (BOT_DATA.mantenimientos.registros || []).filter(r => {
        if (!q) return true;
        return `${r.id} ${r.fecha} ${r.placa} ${r.tipo_mantenimiento} ${r.actividad} ${r.taller} ${r.mecanico}`.toLowerCase().includes(q);
      });

      document.getElementById('pesvCountInfo').textContent = `Mostrando ${rows.length} mantenimientos registrados`;
      const perCode = document.getElementById('pesv_periodo').value || 'todo';

      document.getElementById('pesvTbody').innerHTML = rows.map(r => {
        const isPrev = (r.tipo_mantenimiento || '').toUpperCase().includes('PREVENTIVO');
        return `<tr>
          <td><code>${r.id}</code></td>
          <td><b>${r.fecha}</b></td>
          <td><span class="badge badge-info"><b>${r.placa}</b></span></td>
          <td>${r.tipo_vehiculo}</td>
          <td><span class="badge ${isPrev ? 'badge-success' : 'badge-warning'}">${r.tipo_mantenimiento}</span></td>
          <td><code>${r.kilometraje} km</code></td>
          <td style="white-space:normal; max-width:340px;"><b>${r.actividad}</b></td>
          <td>${r.taller}</td>
          <td>${r.mecanico}</td>
          <td>
            <a href="/api/bot/pdf/mantenimientos?periodo=${perCode}&placa=${encodeURIComponent(r.placa)}" class="btn btn-outline" style="padding:4px 9px; font-size:11.5px;">📄 SISI-F2-08</a>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="10" style="text-align:center; color:var(--muted);">No hay mantenimientos para este filtro.</td></tr>';
    }

    function buscarCandidatoAlumno(qVal, campo) {
      const q = String(qVal || '').trim().toLowerCase();
      if (!q || q.length < 3) return null;
      const listaBot = (BOT_DATA && BOT_DATA.clases && BOT_DATA.clases.alumnos) ? BOT_DATA.clases.alumnos : [];
      let hit = listaBot.find(a => campo === 'cedula'
        ? String(a.cedula || '').toLowerCase() === q
        : String(a.nombre || '').toLowerCase() === q
      );
      if (!hit) {
        hit = listaBot.find(a => campo === 'cedula'
          ? String(a.cedula || '').toLowerCase().startsWith(q)
          : String(a.nombre || '').toLowerCase().includes(q)
        );
      }
      if (hit) return hit;
      if (AUDIT_DATA && AUDIT_DATA.hallazgos && AUDIT_DATA.hallazgos.todas_matriculas) {
        const m = AUDIT_DATA.hallazgos.todas_matriculas.find(r => campo === 'cedula'
          ? String(r.NUMERODOCUMENTO || '').toLowerCase() === q
          : String(r.ALUMNO || '').toLowerCase() === q
        );
        if (m) {
          return {
            cedula: String(m.NUMERODOCUMENTO || ''),
            nombre: String(m.ALUMNO || ''),
            categoria: String(m.CATEGORIA || 'A2'),
            clases_completadas: 0,
            total_requerido: 15
          };
        }
      }
      return null;
    }

    function autocompletarFormularioAlumno(origen) {
      const cedEl = document.getElementById('al_cedula');
      const nomEl = document.getElementById('al_nombre');
      const catEl = document.getElementById('al_cat');
      const qVal = origen === 'cedula' ? cedEl.value : nomEl.value;
      const cand = buscarCandidatoAlumno(qVal, origen);
      if (cand) {
        if (origen === 'cedula' && cand.nombre && !nomEl.value) nomEl.value = cand.nombre;
        if (origen === 'nombre' && cand.cedula) cedEl.value = cand.cedula;
        if (cand.categoria && ['A1', 'A2', 'B1', 'C1', 'RECAT_C1'].includes(cand.categoria)) {
          catEl.value = cand.categoria;
        }
      }
      const searchEl = document.getElementById('searchClasesInput');
      if (searchEl && qVal.trim().length >= 3) {
        searchEl.value = qVal.trim();
        renderizarModuloClases();
      } else if (searchEl && qVal.trim().length === 0) {
        searchEl.value = '';
        renderizarModuloClases();
      }
    }

    function buscarAlumnoDesdeFormulario() {
      const q = (document.getElementById('al_cedula').value || document.getElementById('al_nombre').value || '').trim();
      const searchEl = document.getElementById('searchClasesInput');
      if (searchEl) {
        searchEl.value = q;
        cambiarSubTabClases('TODOS');
      }
    }

    function autocompletarFormularioClase(origen) {
      const cedEl = document.getElementById('cl_cedula');
      const nomEl = document.getElementById('cl_alumno');
      const catEl = document.getElementById('cl_cat');
      const numsEl = document.getElementById('cl_nums');
      const qVal = origen === 'cedula' ? cedEl.value : nomEl.value;
      const cand = buscarCandidatoAlumno(qVal, origen);
      if (cand) {
        if (origen === 'cedula' && cand.nombre) nomEl.value = cand.nombre;
        if (origen === 'nombre' && cand.cedula) cedEl.value = cand.cedula;
        if (cand.categoria && ['A1', 'A2', 'B1', 'C1', 'RECAT_C1'].includes(cand.categoria)) {
          catEl.value = cand.categoria;
        }
        const sig = Math.min((cand.total_requerido || 15), (cand.clases_completadas || 0) + 1);
        const sig2 = Math.min((cand.total_requerido || 15), sig + 1);
        numsEl.value = sig < sig2 ? `${sig}, ${sig2}` : `${sig}`;
      }
    }

    function cargarAlumnoEnFormularios(cedula, nombre, categoria, completadas, requeridas, estadoRunt) {
      document.getElementById('al_cedula').value = cedula || '';
      document.getElementById('al_nombre').value = nombre || '';
      if (categoria) document.getElementById('al_cat').value = categoria;
      const runtEl = document.getElementById('al_runt');
      if (runtEl && estadoRunt) {
        runtEl.value = String(estadoRunt).toLowerCase().includes('cargue') ? 'Cargue RUNT' : 'No Registrado RUNT';
      }

      document.getElementById('cl_cedula').value = cedula || '';
      document.getElementById('cl_alumno').value = nombre || '';
      if (categoria) document.getElementById('cl_cat').value = categoria;
      const sig = Math.min((requeridas || 15), (completadas || 0) + 1);
      const sig2 = Math.min((requeridas || 15), sig + 1);
      document.getElementById('cl_nums').value = sig < sig2 ? `${sig}, ${sig2}` : `${sig}`;

      const toast = document.getElementById('alumnoToast');
      if (toast) {
        toast.className = 'status-toast ok';
        toast.textContent = `✏️ Alumno cargado: ${nombre} (${categoria} - CC ${cedula} - ${estadoRunt || 'No Registrado RUNT'}). Puedes cambiar su Estado RUNT o registrarle las clases ${document.getElementById('cl_nums').value}.`;
      }
      window.scrollTo({ top: 120, behavior: 'smooth' });
    }

    async function cambiarEstadoRuntWeb(cedula, nombre, categoria, nuevoEstadoRunt) {
      const toast = document.getElementById('alumnoToast');
      if (toast) {
        toast.className = 'status-toast ok';
        toast.textContent = `⏳ Actualizando estado RUNT de ${nombre} (${categoria}) a "${nuevoEstadoRunt}"...`;
      }
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'estado_runt_alumno',
            cedula,
            nombre,
            categoria,
            estado_runt: nuevoEstadoRunt
          })
        });
        const rawTxt = await res.text();
        let data = {};
        try { data = JSON.parse(rawTxt); } catch (_) { data = { ok: false, message: rawTxt.slice(0, 180) }; }
        if (toast) {
          toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
          toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || 'Estado RUNT actualizado.');
        }
        if (data.ok) await cargarBotOverview();
      } catch (err) {
        if (toast) {
          toast.className = 'status-toast err';
          toast.textContent = '❌ Error actualizando RUNT: ' + err;
        }
      }
    }

    async function marcarCargueRuntDesdeFormulario() {
      const cedula = document.getElementById('al_cedula').value.trim();
      const nombre = document.getElementById('al_nombre').value.trim();
      const categoria = document.getElementById('al_cat').value;
      const toast = document.getElementById('alumnoToast');
      if (!cedula && !nombre) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Escribe o selecciona primero la Cédula o Nombre del alumno.';
        return;
      }
      const runtEl = document.getElementById('al_runt');
      if (runtEl) runtEl.value = 'Cargue RUNT';
      await cambiarEstadoRuntWeb(cedula, nombre, categoria, 'Cargue RUNT');
    }

    function autocompletarYFiltrarInasistencia(val) {
      const q = String(val || '').trim();
      const searchInas = document.getElementById('searchInasInput');
      if (searchInas) {
        searchInas.value = q;
        renderizarModuloInasistencias();
      }
      if (q.length < 3) return;
      const qLow = q.toLowerCase();
      const regsInas = (BOT_DATA && BOT_DATA.inasistencias && BOT_DATA.inasistencias.registros) ? BOT_DATA.inasistencias.registros : [];
      const hitInas = regsInas.find(r => String(r.alumno || '').toLowerCase() === qLow) || regsInas.find(r => String(r.alumno || '').toLowerCase().includes(qLow));
      if (hitInas) {
        if (hitInas.categoria && ['A1', 'A2', 'B1', 'C1'].includes(hitInas.categoria)) {
          document.getElementById('in_cat').value = hitInas.categoria;
        }
        if (hitInas.instructor && !document.getElementById('in_instructor').value) {
          document.getElementById('in_instructor').value = hitInas.instructor;
        }
        if (hitInas.hora && (!document.getElementById('in_hora').value || document.getElementById('in_hora').value === '6:00 a 8:00 AM')) {
          document.getElementById('in_hora').value = hitInas.hora;
        }
        const toast = document.getElementById('inasToast');
        if (toast) {
          toast.className = 'status-toast ok';
          toast.textContent = `🔎 Encontrada inasistencia #${hitInas.fila} de ${hitInas.alumno} (${hitInas.fecha} - Instructor: ${hitInas.instructor || 'S/D'} - Estado actual: ${hitInas.estado_pago}). Elige el Estado de Pago y haz clic en '🔄 Actualizar Estado de Pago de Este Alumno' o usa los botones de la tabla inferior.`;
        }
        return;
      }
      const candAl = buscarCandidatoAlumno(q, 'nombre');
      if (candAl && candAl.categoria && ['A1', 'A2', 'B1', 'C1'].includes(candAl.categoria)) {
        document.getElementById('in_cat').value = candAl.categoria;
      }
    }

    async function actualizarPagoAlumnoSeleccionadoInas() {
      const nom = (document.getElementById('in_alumno').value || '').trim().toLowerCase();
      const nuevoEstado = document.getElementById('in_estado').value;
      const toast = document.getElementById('inasToast');
      if (!nom) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Escribe el nombre del alumno para buscar su inasistencia y actualizar su estado de pago.';
        return;
      }
      const regsInas = (BOT_DATA && BOT_DATA.inasistencias && BOT_DATA.inasistencias.registros) ? BOT_DATA.inasistencias.registros : [];
      const coincidencias = regsInas.filter(r => String(r.alumno || '').toLowerCase().includes(nom));
      if (!coincidencias.length) {
        toast.className = 'status-toast err';
        toast.textContent = `⚠️ No se encontró una inasistencia previa para "${document.getElementById('in_alumno').value}". Si deseas crear una nueva, completa el Instructor y haz clic en '💾 Registrar Nueva Inasistencia'.`;
        return;
      }
      const target = coincidencias.find(r => String(r.estado_pago || '').toLowerCase().includes('pendiente')) || coincidencias[0];
      toast.className = 'status-toast ok';
      toast.textContent = `⏳ Actualizando inasistencia #${target.fila} de ${target.alumno} a "${nuevoEstado}" en Google Sheets...`;
      await cambiarEstadoInasistencia(target.fila, nuevoEstado);
      toast.className = 'status-toast ok';
      toast.textContent = `✅ Inasistencia #${target.fila} de ${target.alumno} actualizada a "${nuevoEstado}" en Google Sheets.`;
    }

    async function registrarAlumnoWeb() {
      const cedula = document.getElementById('al_cedula').value.trim();
      const nombre = document.getElementById('al_nombre').value.trim();
      const categoria = document.getElementById('al_cat').value;
      const estado_runt = (document.getElementById('al_runt') || {}).value || 'No Registrado RUNT';
      const telefono = document.getElementById('al_tel').value.trim();
      const toast = document.getElementById('alumnoToast');
      if (!cedula || !nombre) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Ingresa Cédula y Nombre del alumno.';
        return;
      }
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Guardando y sincronizando alumno en Google Sheets...';
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'registrar_alumno', cedula, nombre, categoria, telefono, estado_runt })
        });
        const rawTxt = await res.text();
        let data = {};
        try { data = JSON.parse(rawTxt); } catch (_) { data = { ok: false, message: 'Respuesta inesperada del servidor: ' + rawTxt.slice(0, 120) }; }
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || '');
        if (data.ok) await cargarBotOverview();
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error de red: ' + err;
      }
    }

    async function registrarClaseWeb() {
      const cedula = document.getElementById('cl_cedula').value.trim();
      const alumno = document.getElementById('cl_alumno').value.trim();
      const categoria = document.getElementById('cl_cat').value;
      const instructor = document.getElementById('cl_instructor').value.trim();
      const clases_str = document.getElementById('cl_nums').value.trim();
      const toast = document.getElementById('claseToast');
      if (!cedula || !alumno || !instructor) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Completa Cédula, Nombre e Instructor.';
        return;
      }
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Registrando clase práctica en Google Sheets...';
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'registrar_clase', cedula, alumno, categoria, instructor, clases_str })
        });
        const rawTxt = await res.text();
        let data = {};
        try { data = JSON.parse(rawTxt); } catch (_) { data = { ok: false, message: 'Respuesta inesperada del servidor: ' + rawTxt.slice(0, 120) }; }
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || '');
        if (data.ok) await cargarBotOverview();
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error de red: ' + err;
      }
    }

    async function registrarInasistenciaWeb() {
      const fecha_clase = document.getElementById('in_fecha').value.trim();
      const alumno = document.getElementById('in_alumno').value.trim();
      const categoria = document.getElementById('in_cat').value;
      const hora_programada = document.getElementById('in_hora').value.trim() || 'No especificada';
      const instructor = document.getElementById('in_instructor').value.trim() || 'POR ASIGNAR';
      const estado_pago = document.getElementById('in_estado').value;
      const toast = document.getElementById('inasToast');
      if (!alumno) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Ingresa el nombre del alumno.';
        return;
      }
      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Registrando inasistencia en Google Sheets...';
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'registrar_inasistencia',
            fecha_clase, alumno, categoria, hora_programada, instructor, estado_pago
          })
        });
        const data = await res.json();
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (data.ok ? '✅ ' : '❌ ') + data.message;
        if (data.ok) await cargarBotOverview();
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error de red: ' + err;
      }
    }

    async function cambiarEstadoInasistencia(filaSheet, nuevoEstado) {
      const toast = document.getElementById('inasToast');
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ action: 'estado_inasistencia', fila: filaSheet, estado_pago: nuevoEstado })
        });
        const data = await res.json();
        if (toast) {
          toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
          toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || '');
        }
        if (data.ok) await cargarBotOverview();
      } catch (err) {
        if (toast) {
          toast.className = 'status-toast err';
          toast.textContent = '❌ Error actualizando estado: ' + err;
        }
      }
    }

    async function registrarMantenimientoWeb() {
      const placa = document.getElementById('mt_placa').value;
      const tipo_mantenimiento = document.getElementById('mt_tipo').value;
      const kilometraje = document.getElementById('mt_km').value.trim();
      const taller = document.getElementById('mt_taller').value.trim();
      const mecanico = document.getElementById('mt_mecanico').value.trim();
      const actividad = document.getElementById('mt_actividad').value.trim();
      const toast = document.getElementById('mtToast');
      if (!placa || !actividad) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Selecciona la placa y describe la actividad realizada.';
        return;
      }
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'registrar_mantenimiento',
          placa, tipo_mantenimiento, kilometraje, taller, mecanico, actividad
        })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + data.message;
      if (data.ok) cargarBotOverview();
    }

    async function agregarVehiculoWeb() {
      const placa = document.getElementById('vh_placa').value.trim().toUpperCase();
      const tipo = document.getElementById('vh_tipo').value;
      const categoria = document.getElementById('vh_cat').value.trim().toUpperCase();
      const marca = document.getElementById('vh_marca').value.trim().toUpperCase();
      const toast = document.getElementById('vhToast');
      if (!placa) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Ingresa la placa del vehículo.';
        return;
      }
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'vehiculo_agregar', placa, tipo, categoria, marca })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + data.message;
      if (data.ok) cargarBotOverview();
    }

    async function retirarVehiculoWeb(placa) {
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'vehiculo_retirar', placa })
      });
      const data = await res.json();
      if (data.ok) cargarBotOverview();
    }

    function descargarPdfNomina() {
      const per = document.getElementById('nom_periodo').value || 'mes';
      const inst = document.getElementById('nom_instructor').value || 'TODOS';
      window.location.href = `/api/bot/pdf/nomina?instructor=${encodeURIComponent(inst)}&periodo=${encodeURIComponent(per)}`;
    }

    function descargarPdfInasistenciasGeneral() {
      const per = document.getElementById('inas_periodo').value || 'todo';
      const est = document.getElementById('inas_estado_filtro').value || 'TODOS';
      window.location.href = `/api/bot/pdf/inasistencias?periodo=${encodeURIComponent(per)}&estado=${encodeURIComponent(est)}`;
    }

    function descargarPdfPesv() {
      const per = document.getElementById('pesv_periodo').value || 'todo';
      const placa = document.getElementById('pesv_placa_filtro').value || 'TODAS';
      window.location.href = `/api/bot/pdf/mantenimientos?periodo=${encodeURIComponent(per)}&placa=${encodeURIComponent(placa)}`;
    }

    // =========================================================================
    // MÓDULO 6: ROBOT GESTOR DE ARCHIVOS Y DESPACHO MULTICANAL (TELEGRAM + WA)
    // =========================================================================
    async function enviarPdfAlumnoTelegram(cedulaOrNombre, nombre, cat) {
      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'robot_despachar_reporte',
            reporte: 'FICHA_ALUMNO',
            alumno_q: cedulaOrNombre,
            categoria: cat,
            canal: 'TELEGRAM'
          })
        });
        const data = await res.json();
        alert((data.ok ? '✅ ' : '❌ ') + (data.message || ''));
        cargarBitacoraRobot();
      } catch (err) {
        alert('❌ Error enviando a Telegram: ' + err);
      }
    }

    function enviarInfoAlumnoWhatsApp(cedulaOrNombre, nombre, cat, avance, saldo, diasRest90, fechaVenc90, clasesFalt) {
      const tel = (document.getElementById('rb_wa_phone') || {}).value || '573134908073';
      const cleanTel = tel.replace(/\D+/g, '') || '573134908073';
      const pdfUrl = `${window.location.origin}/api/bot/pdf/alumno?q=${encodeURIComponent(cedulaOrNombre)}&cat=${encodeURIComponent(cat)}`;
      const lineaPlazo = fechaVenc90
        ? `🚦 *Vigencia 90 Días (Desde 1ra Clase):* Quedan *${diasRest90} días* (Vence: ${fechaVenc90}) • Faltan *${clasesFalt} clases* por programar.`
        : `🚦 *Clases Pendientes:* Faltan *${clasesFalt || 0} clases* dentro del plazo reglamentario de 90 días.`;
      const msg = [
        `🎓 *CEA AUTO ORIENTE — CONTROL DE VIGENCIA 90 DÍAS Y CLASES PRÁCTICAS*`,
        `👤 *Alumno:* ${nombre}`,
        `🪪 *Cédula:* ${cedulaOrNombre}`,
        `🏍️ *Categoría:* ${cat}`,
        `📊 *Avance Clases:* ${avance}`,
        lineaPlazo,
        `💰 *Saldo Cartera:* ${fmtCOP(saldo || 0)}`,
        `📄 *Descargar Ficha Oficial PDF:* ${pdfUrl}`
      ].join('\n');
      window.open(`https://wa.me/${cleanTel}?text=${encodeURIComponent(msg)}`, '_blank');
    }

    async function despacharReporteRobot(canalRaw) {
      const canal = String(canalRaw || 'TELEGRAM').trim().toUpperCase();
      const reporte = ((document.getElementById('rb_tipo_reporte') || {}).value || 'cartera_pdf').trim().toUpperCase();
      const criterioRaw = ((document.getElementById('rb_criterio') || document.getElementById('rb_alumno_select') || {}).value || '').trim();
      const [alumno_q, categoria] = criterioRaw.includes('|') ? criterioRaw.split('|') : [criterioRaw, 'A2'];
      const chat_id = ((document.getElementById('rb_tg_chat') || document.getElementById('rb_chat_id') || {}).value || '1567561963').trim();
      const wa_phone = ((document.getElementById('rb_wa_phone') || {}).value || '3155680575').trim();
      const toast = document.getElementById('robotDispatchToast') || document.getElementById('rbDespachoToast');

      let waWin = null;
      if (canal === 'WHATSAPP' || canal === 'AMBOS') {
        waWin = window.open('about:blank', '_blank');
      }

      if (toast) {
        toast.className = 'status-toast ok';
        toast.textContent = '⏳ Generando documento oficial, clasificando en Archivo_Procesado y despachando...';
      }

      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'robot_despachar_reporte',
            reporte,
            alumno_q: (alumno_q || '').trim(),
            categoria: (categoria || 'A2').trim(),
            canal,
            chat_id,
            wa_phone
          })
        });
        const data = await res.json();
        if (toast) {
          toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
          let htmlMsg = `<span>${data.ok ? '✅ ' : '❌ '}${data.message || ''}</span>`;
          if (data.download_url) {
            htmlMsg += ` &nbsp;<a href="${data.download_url}" target="_blank" class="btn btn-outline" style="padding:3px 9px; font-size:11px; margin-left:6px;">📥 Descargar Archivo</a>`;
          }
          if (data.wa_url) {
            htmlMsg += ` &nbsp;<a href="${data.wa_url}" target="_blank" class="btn btn-success" style="padding:3px 9px; font-size:11px; margin-left:6px;">📲 Abrir WhatsApp</a>`;
          }
          toast.innerHTML = htmlMsg;
        }
        if (data.wa_url && (canal === 'WHATSAPP' || canal === 'AMBOS')) {
          if (waWin && !waWin.closed) {
            waWin.location.href = data.wa_url;
          } else {
            window.open(data.wa_url, '_blank');
          }
        } else if (waWin && !waWin.closed) {
          waWin.close();
        }
        cargarBitacoraRobot();
      } catch (err) {
        if (waWin && !waWin.closed) waWin.close();
        if (toast) {
          toast.className = 'status-toast err';
          toast.textContent = '❌ Error: ' + err;
        }
      }
    }

    async function guardarConfigRobot() {
      const tg_token = ((document.getElementById('cfg_tg_token') || {}).value || '').trim();
      const tg_chat = ((document.getElementById('cfg_tg_chat') || {}).value || '1567561963').trim();
      const wa_url = ((document.getElementById('cfg_wa_url') || {}).value || '').trim();
      const wa_token = ((document.getElementById('cfg_wa_token') || {}).value || '').trim();
      const wa_phone = ((document.getElementById('cfg_wa_phone') || document.getElementById('rb_wa_phone') || {}).value || '3155680575').trim();
      const toast = document.getElementById('robotCfgToast') || document.getElementById('cfgRobotToast');

      if (toast) {
        toast.className = 'status-toast ok';
        toast.textContent = '⏳ Guardando configuración del Robot y conectando con Telegram...';
      }

      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'robot_guardar_config',
          tg_token, tg_chat, wa_url, wa_token, wa_phone
        })
      });
      const data = await res.json();
      if (toast) {
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || '');
      }
    }

    async function cargarBitacoraRobot() {
      try {
        const res = await fetch('/api/bot/robot_log');
        const data = await res.json();
        if (!data.ok) return;

        const cfg = data.config || {};
        if (cfg.tg_chat && document.getElementById('cfg_tg_chat') && !document.getElementById('cfg_tg_chat').value) {
          document.getElementById('cfg_tg_chat').value = cfg.tg_chat;
        }
        if (cfg.wa_phone && document.getElementById('rb_wa_phone') && !document.getElementById('rb_wa_phone').value) {
          document.getElementById('rb_wa_phone').value = cfg.wa_phone;
        }
        if (cfg.wa_url && document.getElementById('cfg_wa_url') && !document.getElementById('cfg_wa_url').value) {
          document.getElementById('cfg_wa_url').value = cfg.wa_url;
        }

        const logs = data.bitacora || [];
        const cntEl = document.getElementById('cntRobotLogs');
        if (cntEl) cntEl.textContent = `${logs.length} procesados`;

        const tbody = document.getElementById('robotLogTbody');
        if (tbody) {
          tbody.innerHTML = logs.map(l => {
            const isErr = String(l.estado || '').includes('Error');
            return `<tr>
              <td><code>${l.fecha}</code></td>
              <td><b>${l.archivo_original}</b></td>
              <td><span class="badge badge-info">${l.tipo_clasificado}</span></td>
              <td><code style="font-size:11px; color:#93c5fd;">${l.ruta_respaldo}</code></td>
              <td><span class="badge badge-purple">${l.canal}</span></td>
              <td><span class="badge ${isErr ? 'badge-warning' : 'badge-success'}">${l.estado}</span></td>
            </tr>`;
          }).join('') || '<tr><td colspan="6" style="text-align:center; color:var(--muted); padding:18px;">Aún no se han clasificado archivos en esta sesión. Usa el clasificador o despacha cualquier reporte arriba.</td></tr>';
        }
      } catch (err) {
        console.error('Error cargando bitácora del robot:', err);
      }
    }

    const rbForm = document.getElementById('robotUploadForm');
    if (rbForm) {
      rbForm.addEventListener('submit', async (ev) => {
        ev.preventDefault();
        const fileIn = document.getElementById('rb_file_input') || document.getElementById('rbFileInput');
        const rawCanal = ((document.getElementById('rb_auto_tg') || document.getElementById('rbUploadAccion') || {}).value || 'TELEGRAM').toUpperCase();
        const canal = rawCanal === 'SI' ? 'TELEGRAM' : (rawCanal === 'NO' ? 'SOLO_CLASIFICAR' : rawCanal);
        const chat_id = ((document.getElementById('rb_tg_chat') || document.getElementById('rb_chat_id') || {}).value || '1567561963').trim();
        const wa_phone = ((document.getElementById('rb_wa_phone') || {}).value || '3155680575').trim();
        const toast = document.getElementById('robotUploadToast') || document.getElementById('rbUploadToast');

        if (!fileIn || !fileIn.files || !fileIn.files.length) {
          if (toast) {
            toast.className = 'status-toast err';
            toast.textContent = '⚠️ Selecciona un archivo (.pdf, .xlsx, .jpg, .png) primero.';
          }
          return;
        }

        let waWin = null;
        if (canal === 'WHATSAPP' || canal === 'AMBOS') {
          waWin = window.open('about:blank', '_blank');
        }

        if (toast) {
          toast.className = 'status-toast ok';
          toast.textContent = '⏳ Clasificando archivo en Archivo_Procesado y ejecutando despacho...';
        }

        const fd = new FormData();
        fd.append('file', fileIn.files[0]);
        fd.append('canal', canal);
        fd.append('chat_id', chat_id);
        fd.append('wa_phone', wa_phone);

        try {
          const res = await fetch('/api/bot/robot_upload', { method: 'POST', body: fd });
          const data = await res.json();
          if (toast) {
            toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
            let htmlMsg = `<span>${data.ok ? '✅ ' : '❌ '}${data.message || ''}</span>`;
            if (data.wa_url) {
              htmlMsg += ` &nbsp;<a href="${data.wa_url}" target="_blank" class="btn btn-success" style="padding:3px 9px; font-size:11px; margin-left:6px;">📲 Abrir WhatsApp</a>`;
            }
            toast.innerHTML = htmlMsg;
          }
          if (data.wa_url && (canal === 'WHATSAPP' || canal === 'AMBOS')) {
            if (waWin && !waWin.closed) {
              waWin.location.href = data.wa_url;
            } else {
              window.open(data.wa_url, '_blank');
            }
          } else if (waWin && !waWin.closed) {
            waWin.close();
          }
          cargarBitacoraRobot();
        } catch (err) {
          if (waWin && !waWin.closed) waWin.close();
          if (toast) {
            toast.className = 'status-toast err';
            toast.textContent = '❌ Error: ' + err;
          }
        }
      });
    }

    // =========================================================================
    // MÓDULO 7: BOT DE AGENDA, GOOGLE CALENDAR Y BAÚL DIGITAL INTEGRADO
    // =========================================================================
    let AGENDA_DATA = null;

    function leerTokenCalendarArchivo(inputEl) {
      if (!inputEl.files || !inputEl.files.length) return;
      const reader = new FileReader();
      reader.onload = (ev) => {
        const el = document.getElementById('ag_cfg_gcal');
        if (el) el.value = ev.target.result || '';
      };
      reader.readAsText(inputEl.files[0]);
    }

    async function conectarAgendaTelegramYCalendar() {
      const tg_token = (document.getElementById('ag_cfg_token') || {}).value.trim() || '';
      const chat_id = (document.getElementById('ag_cfg_chat') || {}).value.trim() || '1567561963';
      const calendar_token_json = (document.getElementById('ag_cfg_gcal') || {}).value.trim() || '';
      const toast = document.getElementById('agConfigToast');

      toast.className = 'status-toast ok';
      toast.textContent = '⏳ Conectando Bot de Agenda Telegram + Google Calendar y sincronizando eventos...';

      try {
        const res = await fetch('/api/bot/action', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            action: 'agenda_conectar_config',
            tg_token,
            chat_id,
            calendar_token_json
          })
        });
        const data = await res.json();
        toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
        toast.textContent = (data.ok ? '✅ ' : '⚠️ ') + (data.message || '');
        if (document.getElementById('ag_chat_id') && chat_id) {
          document.getElementById('ag_chat_id').value = chat_id;
        }
        cargarAgendaWeb();
      } catch (err) {
        toast.className = 'status-toast err';
        toast.textContent = '❌ Error: ' + err;
      }
    }

    async function cargarAgendaWeb() {
      try {
        const res = await fetch('/api/bot/agenda');
        const data = await res.json();
        if (!data.ok) return;
        AGENDA_DATA = data;
        const ka = data.kpis;
        const bAg = document.getElementById('nav_badge_agenda');
        if (bAg) bAg.textContent = `${ka.total_hoy} Hoy • ${ka.total_baul || 0} Baúl`;
        const lblHoy = document.getElementById('lblFechaHoyAgenda');
        if (lblHoy) lblHoy.textContent = data.fecha_hoy;
        const gcalBadge = document.getElementById('gcalStatusBadge');
        if (gcalBadge) {
          const calStr = data.calendar_conectado ? '🟢 Calendar Conectado' : '🟡 Calendar Local';
          const tgStr = data.bot_token_configurado ? '🟢 Bot TG Activo' : '🔴 Falta Token TG';
          gcalBadge.className = (data.calendar_conectado && data.bot_token_configurado) ? 'badge badge-success' : 'badge badge-purple';
          gcalBadge.textContent = `${calStr} • ${tgStr}`;
        }
        if (data.chat_id_actual) {
          if (document.getElementById('ag_cfg_chat') && !document.getElementById('ag_cfg_chat').value) {
            document.getElementById('ag_cfg_chat').value = data.chat_id_actual;
          }
          if (document.getElementById('ag_chat_id') && !document.getElementById('ag_chat_id').value) {
            document.getElementById('ag_chat_id').value = data.chat_id_actual;
          }
        }

        const kpiEl = document.getElementById('kpiAgendaGrid');
        if (kpiEl) {
          kpiEl.innerHTML = `
            <div class="kpi-card">
              <div class="kpi-label">☀️ Actividades para Hoy (${data.fecha_hoy})</div>
              <div class="kpi-value" style="color:#60a5fa;">${ka.total_hoy} Eventos</div>
              <div class="kpi-sub">Con recordatorio automático a Telegram (10m antes)</div>
            </div>
            <div class="kpi-card purple">
              <div class="kpi-label">📅 Eventos Próximos 7 Días</div>
              <div class="kpi-value" style="color:#c4b5fd;">${ka.total_semana} Eventos</div>
              <div class="kpi-sub">Sincronizados con Google Calendar y Bot</div>
            </div>
            <div class="kpi-card success">
              <div class="kpi-label">🔁 Rutinas Diarias + Puntuales</div>
              <div class="kpi-value" style="color:#34d399;">${ka.recurrentes_diarios} Diarias • ${ka.puntuales} Punt.</div>
              <div class="kpi-sub">Auditoría, Remarketing y Reuniones</div>
            </div>
            <div class="kpi-card warning">
              <div class="kpi-label">🗄️ Elementos en Baúl Digital</div>
              <div class="kpi-value" style="color:#fde68a;">${ka.total_baul || 0} Guardados</div>
              <div class="kpi-sub">Sincronizados con <code>/baul</code> en Telegram</div>
            </div>
          `;
        }

        const tbHoy = document.getElementById('agendaHoyTbody');
        if (tbHoy) {
          tbHoy.innerHTML = (data.eventos_hoy || []).map(e => {
            const safeTit = (e.titulo || '').replace(/'/g, "\\'");
            return `<tr>
              <td><code>${e.hora}</code></td>
              <td><b>${e.titulo}</b></td>
              <td><span class="badge ${e.tipo === 'DIARIO' ? 'badge-purple' : 'badge-info'}">${e.tipo}</span></td>
              <td>${e.responsable || 'CEA'}</td>
              <td>
                <button type="button" class="btn btn-primary" style="padding:4px 8px; font-size:11px;" onclick="recordarEventoTelegramWeb('${safeTit}', '${e.fecha}', '${e.hora}', '${e.responsable || 'CEA'}')">🔔 Telegram</button>
              </td>
            </tr>`;
          }).join('') || '<tr><td colspan="5" style="text-align:center; color:var(--muted);">🌴 No tienes actividades agendadas para hoy.</td></tr>';
        }

        const tbSem = document.getElementById('agendaSemanaTbody');
        if (tbSem) {
          tbSem.innerHTML = (data.eventos_semana || []).map(e => {
            const safeTit = (e.titulo || '').replace(/'/g, "\\'");
            return `<tr>
              <td><b>${e.fecha}</b></td>
              <td><code>${e.hora}</code></td>
              <td><b>${e.titulo}</b></td>
              <td>${e.duracion_min || 60} min</td>
              <td><span class="badge ${e.tipo === 'DIARIO' ? 'badge-purple' : 'badge-info'}">${e.tipo}</span></td>
              <td>${e.responsable || 'CEA'}</td>
              <td><span class="badge badge-success">${e.estado || 'PROGRAMADO'}</span></td>
              <td>
                <div style="display:flex; gap:5px;">
                  <button type="button" class="btn btn-outline" style="padding:4px 8px; font-size:11px;" onclick="recordarEventoTelegramWeb('${safeTit}', '${e.fecha}', '${e.hora}', '${e.responsable || 'CEA'}')">🔔 Recordar TG</button>
                  <button type="button" class="btn btn-danger" style="padding:4px 8px; font-size:11px;" onclick="eliminarEventoAgendaWeb('${e.id}')">🗑️</button>
                </div>
              </td>
            </tr>`;
          }).join('') || '<tr><td colspan="8" style="text-align:center; color:var(--muted);">No hay eventos programados en los próximos 7 días.</td></tr>';
        }

        const tbBaul = document.getElementById('baulTbody');
        if (tbBaul) {
          tbBaul.innerHTML = (data.baul_items || []).map(b => `<tr>
            <td><span class="badge badge-purple">${b.categoria || 'NOTA'}</span></td>
            <td><b>${b.titulo || ''}</b></td>
            <td style="white-space:normal; max-width:420px;">${b.contenido || ''}</td>
            <td><code>${b.fecha || ''}</code></td>
            <td><span class="badge badge-info">${b.origen || 'Web'}</span></td>
            <td>
              <div style="display:flex; gap:5px;">
                <button type="button" class="btn btn-primary" style="padding:4px 8px; font-size:11px;" onclick="enviarItemBaulTelegramWeb('${b.id}')">✈️ Enviar a TG</button>
                <button type="button" class="btn btn-danger" style="padding:4px 8px; font-size:11px;" onclick="eliminarItemBaulWeb('${b.id}')">🗑️</button>
              </div>
            </td>
          </tr>`).join('') || '<tr><td colspan="6" style="text-align:center; color:var(--muted);">Tu Baúl Digital está vacío. Agrega una nota arriba o escribe <code>/baul</code> en el Bot de Telegram.</td></tr>';
        }
      } catch (err) {
        console.error('Error cargando agenda:', err);
      }
    }

    async function crearEventoAgendaWeb() {
      const titulo = document.getElementById('ag_titulo').value.trim();
      const fecha = document.getElementById('ag_fecha').value;
      const hora = document.getElementById('ag_hora').value;
      const duracion_min = document.getElementById('ag_duracion').value;
      const tipo = document.getElementById('ag_tipo').value;
      const responsable = document.getElementById('ag_responsable').value.trim();
      const chat_id = document.getElementById('ag_chat_id').value.trim();
      const toast = document.getElementById('agendaToast');

      if (!titulo) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Escribe el título o actividad del evento.';
        return;
      }
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'agenda_crear_evento',
          titulo, fecha, hora, duracion_min,
          es_diario: tipo === 'DIARIO',
          responsable, chat_id
        })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || '');
      if (data.ok) {
        document.getElementById('ag_titulo').value = '';
        cargarAgendaWeb();
      }
    }

    async function recordarEventoTelegramWeb(titulo, fecha, hora, responsable) {
      const chat_id = (document.getElementById('ag_chat_id') || {}).value || (document.getElementById('ag_cfg_chat') || {}).value || '1567561963';
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'agenda_recordar_telegram', titulo, fecha, hora, responsable, chat_id })
      });
      const data = await res.json();
      alert((data.ok ? '✅ ' : '⚠️ ') + (data.message || ''));
    }

    async function eliminarEventoAgendaWeb(id) {
      await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'agenda_eliminar_evento', id })
      });
      cargarAgendaWeb();
    }

    async function guardarItemBaulWeb() {
      const categoria = document.getElementById('baul_cat').value;
      const titulo = document.getElementById('baul_titulo').value.trim();
      const contenido = document.getElementById('baul_contenido').value.trim();
      const chat_id = (document.getElementById('ag_cfg_chat') || {}).value || '1567561963';
      const toast = document.getElementById('baulToast');
      if (!titulo && !contenido) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Escribe el título y el contenido para guardar en el Baúl.';
        return;
      }
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'baul_guardar_item', categoria, titulo, contenido, chat_id })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || '');
      if (data.ok) {
        document.getElementById('baul_titulo').value = '';
        document.getElementById('baul_contenido').value = '';
        cargarAgendaWeb();
      }
    }

    async function enviarItemBaulTelegramWeb(id) {
      const chat_id = (document.getElementById('ag_cfg_chat') || {}).value || '1567561963';
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'baul_enviar_telegram', id, chat_id })
      });
      const data = await res.json();
      alert((data.ok ? '✅ ' : '⚠️ ') + (data.message || ''));
    }

    async function eliminarItemBaulWeb(id) {
      await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'baul_eliminar_item', id })
      });
      cargarAgendaWeb();
    }

    // =========================================================================
    // MÓDULO 8: CRM WHATSAPP IA, ESTADÍSTICAS DE ABANDONO Y REMARKETING
    // =========================================================================
    let WA_CRM_DATA = null;

    async function cargarWhatsappCrm() {
      try {
        const res = await fetch('/api/bot/whatsapp_crm');
        const data = await res.json();
        if (!data.ok) return;
        WA_CRM_DATA = data;

        const kw = data.kpis;
        const bWa = document.getElementById('nav_badge_wa');
        if (bWa) bWa.textContent = `${kw.sin_concluir_total} Pend. (${kw.tasa_abandono_ia_pct}%)`;

        const selPlant = document.getElementById('wa_plantilla_sel');
        const txtPlant = document.getElementById('wa_plantilla_txt');
        if (selPlant && txtPlant && data.plantillas && !txtPlant.value) {
          const cod = selPlant.value || 'IA_RECUPERACION_50';
          if (data.plantillas[cod]) txtPlant.value = data.plantillas[cod].texto;
        }

        renderizarModuloWhatsapp();
      } catch (err) {
        console.error('Error cargando CRM WhatsApp IA:', err);
      }
    }

    function cambiarPlantillaRemarketing() {
      if (!WA_CRM_DATA || !WA_CRM_DATA.plantillas) return;
      const cod = document.getElementById('wa_plantilla_sel').value;
      const p = WA_CRM_DATA.plantillas[cod];
      if (p) document.getElementById('wa_plantilla_txt').value = p.texto;
    }

    async function guardarPlantillaRemarketingWeb() {
      const codigo = document.getElementById('wa_plantilla_sel').value;
      const texto = document.getElementById('wa_plantilla_txt').value.trim();
      const toast = document.getElementById('waPlantillaToast');
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'wa_guardar_plantilla', codigo, texto })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || '');
      cargarWhatsappCrm();
    }

    function renderizarModuloWhatsapp() {
      if (!WA_CRM_DATA) return;
      const kw = WA_CRM_DATA.kpis;
      const kpiEl = document.getElementById('kpiWhatsappGrid');
      if (kpiEl) {
        kpiEl.innerHTML = `
          <div class="kpi-card">
            <div class="kpi-label">💬 Total Consultas WhatsApp IA</div>
            <div class="kpi-value">${kw.total_consultas} Prospectos</div>
            <div class="kpi-sub">Anuncios FB/IG (<code>fb.me</code>) y Chat Directo</div>
          </div>
          <div class="kpi-card warning">
            <div class="kpi-label">🟢 Atendidos por IA — Sin Concluir</div>
            <div class="kpi-value" style="color:#fde68a;">${kw.pendientes_ia} (${kw.tasa_abandono_ia_pct}%)</div>
            <div class="kpi-sub">Preguntaron, la IA respondió y quedaron ahí</div>
          </div>
          <div class="kpi-card danger">
            <div class="kpi-label">🔴 Piden Asesor + 📲 En Remarketing</div>
            <div class="kpi-value" style="color:#fca5a5;">${kw.transferidos_asesor} Asesor • ${kw.en_remarketing} RMK</div>
            <div class="kpi-sub">Potencial por cerrar: <b>${fmtCOP(kw.ingreso_potencial_pendiente)}</b></div>
          </div>
          <div class="kpi-card success">
            <div class="kpi-label">✅ Matriculados (Conversión Real)</div>
            <div class="kpi-value" style="color:#34d399;">${kw.matriculados} (${kw.tasa_conversion_pct}%)</div>
            <div class="kpi-sub">Cerrado en caja: <b>${fmtCOP(kw.ingreso_cerrado_real)}</b></div>
          </div>
        `;
      }

      const filtroEst = (document.getElementById('wa_filtro_estado') || {}).value || 'SIN_CONCLUIR';
      const q = obtenerBusquedaLimpia('searchWhatsappInput');

      const rows = (WA_CRM_DATA.leads || []).filter(l => {
        if (filtroEst === 'SIN_CONCLUIR' && (l.estado_ia === 'MATRICULADO' || l.estado_ia === 'DESCARTADO')) return false;
        if (filtroEst !== 'TODOS' && filtroEst !== 'SIN_CONCLUIR' && l.estado_ia !== filtroEst) return false;
        if (!q) return true;
        return `${l.telefono} ${l.nombre} ${l.categoria_interes} ${l.ultimo_mensaje_cliente} ${l.notas}`.toLowerCase().includes(q);
      });

      const infoEl = document.getElementById('waCountInfo');
      if (infoEl) infoEl.textContent = `Mostrando ${rows.length} prospectos en este filtro`;

      const tbody = document.getElementById('waLeadsTbody');
      if (!tbody) return;
      tbody.innerHTML = rows.map(l => {
        let badgeEst = '<span class="badge badge-warning">🟢 Responde la IA (Sin concluir)</span>';
        if (l.estado_ia === 'TRANSFERENCIA_ASESOR') {
          badgeEst = '<span class="badge badge-danger">🔴 Transferencia de IA (Asesor)</span>';
        } else if (l.estado_ia === 'REMARKETING_ENVIADO') {
          badgeEst = '<span class="badge badge-purple">📲 Remarketing Enviado</span>';
        } else if (l.estado_ia === 'MATRICULADO') {
          badgeEst = `<span class="badge badge-success">✅ MATRICULADO (${l.matricula_cruzada || 'Cerrado'})</span>`;
        } else if (l.estado_ia === 'DESCARTADO') {
          badgeEst = '<span class="badge">❄️ Descartado</span>';
        }

        const safeNombre = (l.nombre || '').replace(/'/g, "\\'");
        const safeCat = (l.categoria_interes || 'A2').replace(/'/g, "\\'");

        return `<tr>
          <td><code>+${l.telefono}</code></td>
          <td><b>${l.nombre}</b><br><span style="font-size:11px; color:var(--muted);">${l.origen || ''}</span></td>
          <td><span class="badge badge-info">${l.categoria_interes}</span></td>
          <td class="money-pos">${fmtCOP(l.valor_estimado || 0)}</td>
          <td>${badgeEst}</td>
          <td style="white-space:normal; max-width:290px; font-size:11.5px;">
            <b>Cliente:</b> "${l.ultimo_mensaje_cliente || ''}"<br>
            <span style="color:#93c5fd;"><b>IA:</b> ${l.ultima_respuesta_ia || ''}</span>
          </td>
          <td><code>${l.fecha_ultimo_contacto || l.fecha_consulta || ''}</code></td>
          <td style="text-align:center;"><span class="badge">${l.intentos_remarketing || 0} envíos</span></td>
          <td>
            <div style="display:flex; gap:5px; flex-wrap:wrap;">
              <button type="button" class="btn btn-success" style="padding:5px 9px; font-size:11px;" onclick="enviarRemarketingIndividual('${l.id}', '${l.telefono}', '${safeNombre}', '${safeCat}')">📲 Enviar Plantilla WA</button>
              <button type="button" class="btn btn-primary" style="padding:5px 7px; font-size:11px;" title="Marcar como Matriculado" onclick="cambiarEstadoLeadWeb('${l.id}', 'MATRICULADO')">✅ Cerró</button>
              <button type="button" class="btn btn-outline" style="padding:5px 7px; font-size:11px;" title="Marcar en Transferencia Asesor" onclick="cambiarEstadoLeadWeb('${l.id}', 'TRANSFERENCIA_ASESOR')">🔴</button>
            </div>
          </td>
        </tr>`;
      }).join('') || '<tr><td colspan="9" style="text-align:center; color:var(--muted); padding:20px;">No hay prospectos en este filtro.</td></tr>';
    }

    async function enviarRemarketingIndividual(leadId, telefono, nombre, categoria) {
      let plantilla = (document.getElementById('wa_plantilla_txt') || {}).value || '';
      if (!plantilla && WA_CRM_DATA && WA_CRM_DATA.plantillas) {
        plantilla = WA_CRM_DATA.plantillas.IA_RECUPERACION_50.texto;
      }
      const cleanNombre = (nombre || 'futuro conductor').replace(/\(.*\)/g, '').trim();
      const msg = plantilla
        .replace(/\{nombre\}/gi, cleanNombre)
        .replace(/\{categoria\}/gi, categoria || 'Moto A2 / Carro B1');

      const cleanTel = (telefono || '').replace(/\D+/g, '');
      window.open(`https://wa.me/${cleanTel}?text=${encodeURIComponent(msg)}`, '_blank');

      await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'wa_cambiar_estado',
          id: leadId,
          estado_ia: 'REMARKETING_ENVIADO',
          incrementar_remarketing: true
        })
      });
      cargarWhatsappCrm();
    }

    function lanzarSiguienteRemarketingPendiente() {
      if (!WA_CRM_DATA || !WA_CRM_DATA.leads) return;
      const pendiente = WA_CRM_DATA.leads.find(l => l.estado_ia === 'RESPONDE_IA_PENDIENTE' || l.estado_ia === 'TRANSFERENCIA_ASESOR');
      if (!pendiente) {
        alert('🎉 ¡Excelente! Todos los clientes pendientes de la IA ya recibieron su mensaje de remarketing.');
        return;
      }
      enviarRemarketingIndividual(pendiente.id, pendiente.telefono, pendiente.nombre, pendiente.categoria_interes);
    }

    async function registrarLeadWhatsappWeb() {
      const telefono = document.getElementById('wa_tel').value.trim();
      const nombre = document.getElementById('wa_nombre').value.trim();
      const categoria_interes = document.getElementById('wa_cat').value;
      const estado_ia = document.getElementById('wa_estado').value;
      const ultimo_mensaje_cliente = document.getElementById('wa_msg').value.trim();
      const toast = document.getElementById('waLeadToast');

      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'wa_registrar_lead',
          telefono, nombre, categoria_interes, estado_ia, ultimo_mensaje_cliente
        })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || '');
      if (data.ok) {
        document.getElementById('wa_tel').value = '';
        document.getElementById('wa_nombre').value = '';
        document.getElementById('wa_msg').value = '';
        cargarWhatsappCrm();
      }
    }

    async function importarLoteWhatsappWeb() {
      const texto = document.getElementById('wa_lote_txt').value.trim();
      const categoria_default = document.getElementById('wa_cat').value || 'A2';
      const toast = document.getElementById('waLeadToast');
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'wa_importar_lote', texto, categoria_default })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || '');
      if (data.ok) {
        document.getElementById('wa_lote_txt').value = '';
        cargarWhatsappCrm();
      }
    }

    async function cambiarEstadoLeadWeb(leadId, nuevoEstado) {
      await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'wa_cambiar_estado', id: leadId, estado_ia: nuevoEstado })
      });
      cargarWhatsappCrm();
    }

    // Inicializar fecha por defecto de inasistencia y agenda hoy
    (function initDefaults() {
      const hoy = new Date();
      const dd = String(hoy.getDate()).padStart(2, '0');
      const mm = String(hoy.getMonth() + 1).padStart(2, '0');
      const yyyy = hoy.getFullYear();
      const inFecha = document.getElementById('in_fecha');
      if (inFecha) inFecha.value = `${dd}/${mm}/${yyyy}`;
      const agFecha = document.getElementById('ag_fecha');
      if (agFecha) agFecha.value = `${yyyy}-${mm}-${dd}`;
    })();

    let CURRENT_USER = null;
    let EMPLEADOS_DATA = null;

    const MAPA_NOMBRES_MODULOS = {
      mod_cartera: '💰 1. Cartera',
      mod_clases: '🎓 2. Clases y Alumnos',
      mod_nomina: '👨‍🏫 3. Profesores / Nómina',
      mod_inasistencias: '🚫 4. Inasistencias',
      mod_pesv: '🔧 5. Mantenimientos PESV',
      mod_robot: '🤖 6. Robot',
      mod_agenda: '📅 7. Agenda + Baúl',
      mod_whatsapp: '💬 8. WhatsApp IA',
      mod_empleados: '👥 9. Empleados'
    };

    async function cargarSesionUsuario() {
      try {
        const res = await fetch('/api/me');
        const data = await res.json();
        if (!data.ok || !data.user) return;
        CURRENT_USER = data.user;
        const badge = document.getElementById('userSessionBadge');
        if (badge) {
          badge.innerHTML = `👤 <b>${CURRENT_USER.nombre}</b> (<code>${CURRENT_USER.username}</code> • ${CURRENT_USER.rol})`;
        }

        const modulosPermitidos = CURRENT_USER.modulos || ['mod_cartera','mod_clases','mod_nomina','mod_inasistencias','mod_pesv'];
        const esAdminTotal = CURRENT_USER.rol === 'ADMIN_TOTAL';

        // Mostrar u ocultar botones de navegación según permisos del empleado
        let primerModuloPermitido = null;
        let primerBtnPermitido = null;
        document.querySelectorAll('.mod-btn[data-mod]').forEach(btn => {
          const mId = btn.getAttribute('data-mod');
          if (esAdminTotal || modulosPermitidos.includes(mId)) {
            btn.style.display = 'flex';
            if (!primerModuloPermitido) {
              primerModuloPermitido = mId;
              primerBtnPermitido = btn;
            }
          } else {
            btn.style.display = 'none';
          }
        });

        // Ocultar botones administrativos sensibles si no es ADMIN_TOTAL
        document.querySelectorAll('.admin-only-btn').forEach(el => {
          el.style.display = esAdminTotal ? 'inline-flex' : 'none';
        });
        document.querySelectorAll('.cartera-only-btn').forEach(el => {
          el.style.display = (esAdminTotal || modulosPermitidos.includes('mod_cartera')) ? 'inline-flex' : 'none';
        });

        // Si el módulo activo actual no está permitido para este empleado, saltar al primero permitido
        if (!esAdminTotal && !modulosPermitidos.includes(ACTIVE_MODULE) && primerModuloPermitido) {
          cambiarModulo(primerModuloPermitido, primerBtnPermitido);
        }
      } catch (err) {
        console.error('Error verificando sesión:', err);
      }
    }

    function marcarModulosEmpleado(listaMods) {
      document.querySelectorAll('.emp-mod-chk').forEach(chk => {
        chk.checked = listaMods.includes(chk.value);
      });
    }

    function aplicarPlantillaRolEmpleado(rol) {
      if (rol === 'RECEPCIONISTA') {
        marcarModulosEmpleado(['mod_clases','mod_nomina','mod_inasistencias','mod_pesv']);
      } else if (rol === 'COORDINACION') {
        marcarModulosEmpleado(['mod_clases','mod_nomina','mod_inasistencias']);
      } else if (rol === 'PESV') {
        marcarModulosEmpleado(['mod_pesv']);
      } else if (rol === 'SOLO_CLASES') {
        marcarModulosEmpleado(['mod_clases','mod_nomina']);
      } else if (rol === 'ADMIN_TOTAL') {
        marcarModulosEmpleado(['mod_cartera','mod_clases','mod_nomina','mod_inasistencias','mod_pesv','mod_robot','mod_agenda','mod_whatsapp','mod_empleados']);
      }
    }

    async function cargarEmpleadosWeb() {
      try {
        const res = await fetch('/api/empleados');
        const data = await res.json();
        if (!data.ok) return;
        EMPLEADOS_DATA = data;
        const lista = data.empleados || [];
        const logs = data.bitacora || [];

        const kTot = document.getElementById('emp_kpi_total');
        if (kTot) kTot.textContent = lista.length;
        const kRec = document.getElementById('emp_kpi_recepcion');
        if (kRec) kRec.textContent = lista.filter(u => u.rol === 'RECEPCIONISTA' || u.rol === 'PESV').length;
        const kProf = document.getElementById('emp_kpi_profesores');
        if (kProf) kProf.textContent = lista.filter(u => (u.modulos || []).includes('mod_nomina')).length;
        const kLogs = document.getElementById('emp_kpi_logs');
        if (kLogs) kLogs.textContent = logs.length;
        const navEmp = document.getElementById('nav_badge_emp');
        if (navEmp) navEmp.textContent = `${lista.length} Usuarios`;

        const tbody = document.getElementById('empleadosTbody');
        if (tbody) {
          tbody.innerHTML = lista.map(u => {
            const modsBadges = (u.modulos || []).map(m => `<span class="badge badge-info" style="margin:2px;">${MAPA_NOMBRES_MODULOS[m] || m}</span>`).join(' ');
            const safeUser = (u.username || '').replace(/'/g, "\'");
            const safePass = (u.password || '').replace(/'/g, "\'");
            const safeNom = (u.nombre || '').replace(/'/g, "\'");
            const modsJson = JSON.stringify(u.modulos || []).replace(/"/g, '&quot;');
            return `<tr>
              <td><b><code>${u.username}</code></b></td>
              <td><code>${u.password}</code></td>
              <td><b>${u.nombre}</b></td>
              <td><span class="badge badge-purple">${u.rol}</span></td>
              <td style="white-space:normal; max-width:420px;">${modsBadges}</td>
              <td>
                <div style="display:flex; gap:6px; flex-wrap:wrap;">
                  <button type="button" class="btn btn-outline" style="padding:4px 9px; font-size:11px;" onclick="cargarEmpleadoEnFormulario('${safeUser}', '${safePass}', '${safeNom}', '${u.rol}', ${modsJson})">✏️ Editar</button>
                  <button type="button" class="btn btn-success" style="padding:4px 9px; font-size:11px;" onclick="copiarCredencialesEmpleado('${safeUser}', '${safePass}', '${safeNom}')">📲 Copiar Acceso</button>
                  ${u.rol !== 'ADMIN_TOTAL' ? `<button type="button" class="btn btn-danger" style="padding:4px 9px; font-size:11px;" onclick="eliminarEmpleadoWeb('${safeUser}')">🗑️</button>` : ''}
                </div>
              </td>
            </tr>`;
          }).join('');
        }

        const logTbody = document.getElementById('bitacoraEmpleadosTbody');
        if (logTbody) {
          logTbody.innerHTML = logs.map(l => `<tr>
            <td><code>${l.fecha}</code></td>
            <td><b><code>${l.username}</code></b></td>
            <td>${l.nombre}</td>
            <td><span class="badge">${l.rol}</span></td>
            <td><span class="badge badge-info">${l.accion}</span></td>
            <td style="white-space:normal;">${l.detalle}</td>
          </tr>`).join('') || '<tr><td colspan="6" style="text-align:center; color:var(--muted); padding:16px;">Aún no hay acciones registradas en la bitácora.</td></tr>';
        }
      } catch (err) {
        console.error('Error cargando empleados:', err);
      }
    }

    function cargarEmpleadoEnFormulario(username, password, nombre, rol, modulos) {
      document.getElementById('emp_user').value = username;
      document.getElementById('emp_pass').value = password;
      document.getElementById('emp_nombre').value = nombre;
      document.getElementById('emp_rol').value = rol || 'RECEPCIONISTA';
      marcarModulosEmpleado(modulos || []);
    }

    function copiarCredencialesEmpleado(username, password, nombre) {
      const txt = `🔐 *ACCESO PLATAFORMA ADMINISTRATIVA — CEA AUTO ORIENTE*\n👤 *Empleado(a):* ${nombre}\n🌐 *Enlace:* ${window.location.origin}\n🆔 *Usuario:* ${username}\n🔑 *Contraseña:* ${password}`;
      navigator.clipboard.writeText(txt).then(() => {
        alert('✅ Credenciales copiadas al portapapeles listas para pegar en WhatsApp o Telegram.');
      }).catch(() => {
        prompt('Copia las credenciales:', txt);
      });
    }

    async function guardarEmpleadoWeb() {
      const username = document.getElementById('emp_user').value.trim();
      const password = document.getElementById('emp_pass').value.trim();
      const nombre = document.getElementById('emp_nombre').value.trim();
      const rol = document.getElementById('emp_rol').value;
      const modulos = [];
      document.querySelectorAll('.emp-mod-chk:checked').forEach(chk => modulos.push(chk.value));
      const toast = document.getElementById('empToast');

      if (!username || !password || !nombre) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Completa el Usuario, Contraseña y Nombre del empleado(a).';
        return;
      }
      if (!modulos.length) {
        toast.className = 'status-toast err';
        toast.textContent = '⚠️ Selecciona al menos un módulo permitido.';
        return;
      }

      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          action: 'empleado_guardar',
          username, password, nombre, rol, modulos
        })
      });
      const data = await res.json();
      toast.className = data.ok ? 'status-toast ok' : 'status-toast err';
      toast.textContent = (data.ok ? '✅ ' : '❌ ') + (data.message || '');
      if (data.ok) {
        cargarEmpleadosWeb();
      }
    }

    async function eliminarEmpleadoWeb(username) {
      if (!confirm(`¿Seguro que deseas eliminar la cuenta '${username}'?`)) return;
      const res = await fetch('/api/bot/action', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ action: 'empleado_eliminar', username })
      });
      const data = await res.json();
      if (data.ok) cargarEmpleadosWeb();
      else alert(data.message || 'No se pudo eliminar');
    }

    cargarSesionUsuario();
    cargarDatos();
    cargarBotOverview();
    cargarAgendaWeb();
    cargarWhatsappCrm();
    cargarEmpleadosWeb();

    // Guardián anti-autocompletado tardío del gestor de contraseñas del navegador
    [350, 1000, 2200].forEach(delayMs => {
      setTimeout(() => {
        ['searchInput', 'searchClasesInput', 'searchNominaInput', 'searchInasInput', 'searchPesvInput', 'searchWhatsappInput'].forEach(id => {
          const el = document.getElementById(id);
          if (el && el.value) {
            const prev = el.value;
            obtenerBusquedaLimpia(id);
            if (el.value !== prev) {
              if (id === 'searchInput') renderizarTablaActiva();
              else if (id === 'searchClasesInput') renderizarModuloClases();
              else if (id === 'searchNominaInput') renderizarModuloNomina();
              else if (id === 'searchInasInput') renderizarModuloInasistencias();
              else if (id === 'searchPesvInput') renderizarModuloPesv();
              else if (id === 'searchWhatsappInput') renderizarModuloWhatsapp();
            }
          }
        });
      }, delayMs);
    });
  </script>
</body>
</html>
"""


class AuditorWebHandler(BaseHTTPRequestHandler):
    def _send_json(self, data: Dict[str, Any], status: int = 200, extra_headers: Optional[Dict[str, str]] = None) -> None:
        if isinstance(data, dict):
            if "error" in data and "message" not in data:
                data["message"] = str(data["error"])
            elif "mensaje" in data and "message" not in data:
                data["message"] = str(data["mensaje"])
            elif "message" in data and "error" not in data and not data.get("ok", True):
                data["error"] = str(data["message"])
        raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        if extra_headers:
            for k, v in extra_headers.items():
                self.send_header(k, v)
        self.end_headers()
        self.wfile.write(raw)

    def _get_current_user(self) -> Optional[Dict[str, Any]]:
        cookie_hdr = self.headers.get("Cookie", "")
        if not cookie_hdr:
            return None
        c = SimpleCookie()
        try:
            c.load(cookie_hdr)
        except Exception:
            return None
        if COOKIE_NAME not in c:
            return None
        return _obtener_perfil_desde_token(c[COOKIE_NAME].value)

    def _is_authenticated(self) -> bool:
        return self._get_current_user() is not None

    def do_GET(self) -> None:
        if self.path == "/health":
            self._send_json({"ok": True, "service": "Auditor CEA Auto Oriente Multi-Curso"})
            return

        if self.path == "/logout":
            self.send_response(302)
            self.send_header("Location", "/login")
            self.send_header("Set-Cookie", f"{COOKIE_NAME}=; Path=/; HttpOnly; Max-Age=0; SameSite=Lax")
            self.end_headers()
            return

        if self.path == "/login":
            raw = HTML_LOGIN.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return

        # =====================================================================
        # SERVICIO EN LÍNEA 24/7: OBSERVATORIO NACIONAL DE SEGURIDAD VIAL (ANSV)
        # =====================================================================
        if self.path.startswith("/observatorio"):
            raw = observatorio_vial_cea.HTML_OBSERVATORIO_STANDALONE.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return

        if self.path.startswith("/api/observatorio/logo"):
            rutas_logo = [
                BASE_DIR / "fallecidos" / "fallecidos" / "LOGO PEQUEÑO.png",
                BASE_DIR / "fallecidos" / "LOGO PEQUEÑO.png",
                BASE_DIR / "LOGO PEQUEÑO.png",
                BASE_DIR / "logo_autooriente.png",
                BASE_DIR / "SUBIR_A_GITHUB" / "logo_autooriente.png",
            ]
            logo_bytes = b""
            for rl in rutas_logo:
                if rl.exists():
                    logo_bytes = rl.read_bytes()
                    break
            if logo_bytes:
                self.send_response(200)
                self.send_header("Content-Type", "image/png")
                self.send_header("Content-Length", str(len(logo_bytes)))
                self.end_headers()
                self.wfile.write(logo_bytes)
            else:
                self.send_response(404)
                self.end_headers()
            return

        if self.path.startswith("/api/observatorio/data"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            anio = (qs.get("anio", ["Todos"])[0] or "Todos").strip()
            depto = (qs.get("departamento", ["Todos"])[0] or "Todos").strip()
            muni = (qs.get("municipio", ["Todos"])[0] or "Todos").strip()
            actor = (qs.get("actor_vial", ["Todos"])[0] or "Todos").strip()
            sexo = (qs.get("sexo", ["Todos"])[0] or "Todos").strip()
            res_obs = observatorio_vial_cea.observatorio_service.consultar(
                anio=anio,
                departamento=depto,
                municipio=muni,
                actor_vial=actor,
                sexo=sexo,
            )
            self._send_json(res_obs)
            return

        if not self._is_authenticated():
            if self.path.startswith("/api/"):
                self._send_json({"ok": False, "error": "Sesión expirada o no autorizada."}, 401)
            else:
                raw = HTML_LOGIN.encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(raw)))
                self.end_headers()
                self.wfile.write(raw)
            return

        if self.path in ("/", "/index.html"):
            u_curr = self._get_current_user() or {
                "username": "recepcion",
                "nombre": "Recepción",
                "rol": "RECEPCIONISTA",
                "modulos": ["mod_clases", "mod_nomina", "mod_inasistencias", "mod_pesv"],
            }
            es_admin_total = u_curr.get("rol") == "ADMIN_TOTAL"
            mods_permitidos = u_curr.get("modulos") or ["mod_clases", "mod_nomina", "mod_inasistencias", "mod_pesv"]
            todos_ids = [m["id"] for m in MODULOS_CATALOGO]
            mods_ocultos = [] if es_admin_total else [m for m in todos_ids if m not in mods_permitidos]

            reglas_css = []
            for m_hide in mods_ocultos:
                reglas_css.append(f'#{m_hide}, .mod-btn[data-mod="{m_hide}"] {{ display: none !important; }}')
            if not es_admin_total:
                reglas_css.append('.admin-only-btn, #sheetsDrawer, #tarifasDrawer { display: none !important; }')
            if not es_admin_total and "mod_cartera" not in mods_permitidos:
                reglas_css.append('.cartera-only-btn { display: none !important; }')

            u_safe = {k: v for k, v in u_curr.items() if k != "password"}
            primer_mod = "mod_cartera" if (es_admin_total or "mod_cartera" in mods_permitidos) else (mods_permitidos[0] if mods_permitidos else "mod_clases")

            html_custom = HTML_DASHBOARD
            # Si el usuario no tiene permiso para mod_cartera, activar desde el servidor su primer módulo permitido
            if primer_mod != "mod_cartera":
                html_custom = html_custom.replace(
                    '<button type="button" class="mod-btn active" data-mod="mod_cartera"',
                    '<button type="button" class="mod-btn" data-mod="mod_cartera"',
                    1,
                )
                html_custom = html_custom.replace(
                    f'<button type="button" class="mod-btn" data-mod="{primer_mod}"',
                    f'<button type="button" class="mod-btn active" data-mod="{primer_mod}"',
                    1,
                )
                html_custom = html_custom.replace(
                    '<div id="mod_cartera" class="mod-section active">',
                    '<div id="mod_cartera" class="mod-section">',
                    1,
                )
                html_custom = html_custom.replace(
                    f'<div id="{primer_mod}" class="mod-section">',
                    f'<div id="{primer_mod}" class="mod-section active">',
                    1,
                )
                html_custom = html_custom.replace(
                    "let ACTIVE_MODULE = 'mod_cartera';",
                    f"let ACTIVE_MODULE = '{primer_mod}';",
                    1,
                )

            inject_head = (
                f"<style id='server-role-lock'>{' '.join(reglas_css)}</style>"
                f"<script>window.__INITIAL_USER__ = {json.dumps(u_safe, ensure_ascii=False)};</script>"
            )
            html_custom = html_custom.replace("</head>", inject_head + "</head>", 1)

            raw = html_custom.encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return

        if self.path == "/api/me":
            u_curr = self._get_current_user()
            if not u_curr:
                self._send_json({"ok": False, "error": "No autenticado"}, 401)
                return
            u_safe = dict(u_curr)
            u_safe.pop("password", None)
            self._send_json({"ok": True, "user": u_safe})
            return

        if self.path == "/api/empleados":
            u_curr = self._get_current_user()
            if not u_curr:
                self._send_json({"ok": False, "error": "No autenticado"}, 401)
                return
            usuarios_map = _cargar_usuarios_empleados()
            lista_emp = list(usuarios_map.values())
            if u_curr.get("rol") != "ADMIN_TOTAL":
                # Si no es admin total, ocultar contraseñas de otros usuarios
                lista_emp = [
                    {**x, "password": "***"} for x in lista_emp if x.get("username") == u_curr.get("username")
                ]
            self._send_json({
                "ok": True,
                "empleados": lista_emp,
                "bitacora": _cargar_bitacora_empleados(),
            })
            return

        if self.path == "/api/data":
            u_curr = self._get_current_user() or {}
            if u_curr.get("rol") != "ADMIN_TOTAL" and "mod_cartera" not in (u_curr.get("modulos") or []):
                self._send_json({"ok": False, "restricted": True, "error": "Módulo financiero restringido a Gerencia."}, 403)
                return
            self._send_json({
                "ok": True,
                "filename": STATE["filename"],
                "ingresos_filename": STATE["ingresos_filename"],
                "periodicidad": STATE["periodicidad"],
                "updated_at": STATE["updated_at"],
                "tarifas_base": STATE["tarifas_base"],
                "hallazgos": STATE["hallazgos"],
            })
            return

        if self.path.startswith("/api/download/excel"):
            u_curr = self._get_current_user() or {}
            if u_curr.get("rol") != "ADMIN_TOTAL" and "mod_cartera" not in (u_curr.get("modulos") or []):
                self._send_json({"ok": False, "error": "Descarga financiera restringida a Gerencia."}, 403)
                return
            excel_bytes = STATE.get("excel_bytes") or b""
            fecha_str = datetime.date.today().strftime("%Y%m%d")
            fname = f"Auditoria_CEA_{STATE['periodicidad']}_{fecha_str}.xlsx"
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(excel_bytes)))
            self.end_headers()
            self.wfile.write(excel_bytes)
            return

        if self.path.startswith("/api/download/pdf"):
            u_curr = self._get_current_user() or {}
            if u_curr.get("rol") != "ADMIN_TOTAL" and "mod_cartera" not in (u_curr.get("modulos") or []):
                self._send_json({"ok": False, "error": "Descarga financiera restringida a Gerencia."}, 403)
                return
            pdf_bytes = STATE.get("pdf_bytes") or b""
            fecha_str = datetime.date.today().strftime("%Y%m%d")
            fname = f"Resumen_Ejecutivo_CEA_{STATE['periodicidad']}_{fecha_str}.pdf"
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(pdf_bytes)))
            self.end_headers()
            self.wfile.write(pdf_bytes)
            return

        if self.path.startswith("/api/bot/overview"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            p_nomina = (qs.get("periodo_nomina", ["mes"])[0] or "mes").strip()
            inst_nomina = (qs.get("instructor_nomina", ["TODOS"])[0] or "TODOS").strip()
            p_inas = (qs.get("periodo_inasistencias", ["todo"])[0] or "todo").strip()
            est_inas = (qs.get("estado_inasistencias", ["TODOS"])[0] or "TODOS").strip()
            p_mant = (qs.get("periodo_mantenimientos", ["todo"])[0] or "todo").strip()
            placa_mant = (qs.get("placa_mantenimientos", ["TODAS"])[0] or "TODAS").strip()

            panorama = clases_core.sheets_service.obtener_panorama_web(
                periodo_nomina=p_nomina,
                instructor_nomina=inst_nomina,
                periodo_inasistencias=p_inas,
                estado_inasistencias=est_inas,
                periodo_mantenimientos=p_mant,
                placa_mantenimientos=placa_mant,
                hallazgos_cartera=STATE.get("hallazgos") if ((self._get_current_user() or {}).get("rol") == "ADMIN_TOTAL" or "mod_cartera" in ((self._get_current_user() or {}).get("modulos") or [])) else None,
            )
            self._send_json(panorama)
            return

        if self.path.startswith("/api/bot/pdf/alumno"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            q_val = (qs.get("q", [""])[0] or "").strip()
            cat_val = (qs.get("cat", ["A2"])[0] or "A2").strip().upper()
            res_al = clases_core.sheets_service.obtener_historial_alumno(q_val)
            alumno = res_al.get("alumno") or {"cedula": q_val, "nombre": q_val or "ALUMNO CEA", "categoria": cat_val, "telefono": ""}
            if cat_val and cat_val in clases_core.FORMATOS_OFICIALES:
                alumno["categoria"] = cat_val
            registros = [
                r for r in (res_al.get("registros") or [])
                if not cat_val or str(r.get("categoria", "")).strip().upper() == cat_val
            ] or (res_al.get("registros") or [])
            pdf_buf = clases_core.generar_pdf_seguimiento(alumno, registros)
            pdf_bytes = pdf_buf.getvalue()
            safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", str(alumno.get("nombre", "Alumno")))
            fname = f"Seguimiento_Clases_{alumno.get('categoria', 'CEA')}_{safe_name}.pdf"
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(pdf_bytes)))
            self.end_headers()
            self.wfile.write(pdf_bytes)
            return

        if self.path.startswith("/api/bot/pdf/nomina"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            inst = (qs.get("instructor", ["TODOS"])[0] or "TODOS").strip()
            periodo = (qs.get("periodo", ["mes"])[0] or "mes").strip()
            reporte = clases_core.sheets_service.obtener_clases_por_instructor_periodo(inst, periodo)
            pdf_buf = clases_core.generar_pdf_liquidacion_instructor(reporte)
            pdf_bytes = pdf_buf.getvalue()
            safe_inst = re.sub(r"[^A-Za-z0-9_-]+", "_", inst)
            fname = f"Nomina_Liquidacion_{safe_inst}_{periodo}.pdf"
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(pdf_bytes)))
            self.end_headers()
            self.wfile.write(pdf_bytes)
            return

        if self.path.startswith("/api/bot/pdf/inasistencias"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            alumno_q = (qs.get("alumno", [""])[0] or "").strip()
            if alumno_q:
                datos_al = clases_core.sheets_service.buscar_inasistencias_alumno(alumno_q)
                pdf_buf = clases_core.generar_pdf_inasistencias_individual(datos_al)
                safe_al = re.sub(r"[^A-Za-z0-9_-]+", "_", alumno_q)
                fname = f"Inasistencias_Alumno_{safe_al}.pdf"
            else:
                periodo = (qs.get("periodo", ["todo"])[0] or "todo").strip()
                estado = (qs.get("estado", ["TODOS"])[0] or "TODOS").strip()
                datos_per = clases_core.sheets_service.obtener_inasistencias_periodo(periodo, estado)
                pdf_buf = clases_core.generar_pdf_inasistencias_periodo(datos_per)
                fname = f"Reporte_Inasistencias_{periodo}_{estado}.pdf"
            pdf_bytes = pdf_buf.getvalue()
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(pdf_bytes)))
            self.end_headers()
            self.wfile.write(pdf_bytes)
            return

        if self.path.startswith("/api/bot/pdf/mantenimientos"):
            parsed = urllib.parse.urlparse(self.path)
            qs = urllib.parse.parse_qs(parsed.query)
            periodo = (qs.get("periodo", ["todo"])[0] or "todo").strip()
            placa = (qs.get("placa", ["TODAS"])[0] or "TODAS").strip()
            datos_mt = clases_core.sheets_service.obtener_mantenimientos_periodo(periodo, placa)
            pdf_buf = clases_core.generar_pdf_mantenimientos_pesv(datos_mt)
            pdf_bytes = pdf_buf.getvalue()
            fname = f"PESV_SISI_F2_08_{placa}_{periodo}.pdf"
            self.send_response(200)
            self.send_header("Content-Type", "application/pdf")
            self.send_header("Content-Disposition", f'attachment; filename="{fname}"')
            self.send_header("Content-Length", str(len(pdf_bytes)))
            self.end_headers()
            self.wfile.write(pdf_bytes)
            return

        if self.path.startswith("/api/bot/robot_log"):
            cfg = _cargar_robot_config()
            cfg_safe = {
                "tg_chat": cfg.get("tg_chat", ""),
                "wa_url": cfg.get("wa_url", ""),
                "wa_phone": cfg.get("wa_phone", ""),
            }
            self._send_json({
                "ok": True,
                "config": cfg_safe,
                "bitacora": _cargar_robot_bitacora(),
            })
            return

        if self.path.startswith("/api/bot/download_robot_script"):
            script_path = BASE_DIR / "robot_gestor_cea.py"
            if script_path.exists():
                raw_script = script_path.read_bytes()
            else:
                raw_script = b"# robot_gestor_cea.py\n"
            self.send_response(200)
            self.send_header("Content-Type", "text/x-python; charset=utf-8")
            self.send_header("Content-Disposition", 'attachment; filename="robot_gestor_cea.py"')
            self.send_header("Content-Length", str(len(raw_script)))
            self.end_headers()
            self.wfile.write(raw_script)
            return

        if self.path.startswith("/api/bot/download_agenda_script"):
            script_path = BASE_DIR / "bot_agenda_cea.py"
            if script_path.exists():
                raw_script = script_path.read_bytes()
            else:
                raw_script = b"# bot_agenda_cea.py\n"
            self.send_response(200)
            self.send_header("Content-Type", "text/x-python; charset=utf-8")
            self.send_header("Content-Disposition", 'attachment; filename="bot_agenda_cea.py"')
            self.send_header("Content-Length", str(len(raw_script)))
            self.end_headers()
            self.wfile.write(raw_script)
            return

        if self.path.startswith("/api/bot/agenda"):
            try:
                res_ag = bot_agenda_cea.obtener_panorama_agenda_web()
                self._send_json(res_ag)
            except Exception as exc:
                self._send_json({"ok": False, "error": str(exc)}, 500)
            return

        if self.path.startswith("/api/bot/whatsapp_crm"):
            try:
                res_wa = crm_whatsapp_cea.obtener_panorama_whatsapp_crm(STATE.get("hallazgos"))
                self._send_json(res_wa)
            except Exception as exc:
                self._send_json({"ok": False, "error": str(exc)}, 500)
            return

        if self.path.startswith("/api/bot/runt_control"):
            try:
                parsed = urllib.parse.urlparse(self.path)
                qs = urllib.parse.parse_qs(parsed.query)
                forzar = (qs.get("forzar", ["0"])[0] == "1")
                res_rc = runt_control_cea.runt_service.obtener_panorama_runt_web(forzar=forzar)
                self._send_json(res_rc)
            except Exception as exc:
                self._send_json({"ok": False, "error": str(exc)}, 500)
            return

        self.send_response(404)
        self.end_headers()

    def do_POST(self) -> None:
        try:
            self._handle_post()
        except Exception as exc:
            logger.exception("Error inesperado en do_POST (%s)", self.path)
            try:
                self._send_json({"ok": False, "message": f"Error procesando solicitud: {exc}", "error": str(exc)}, 500)
            except Exception:
                pass

    def _handle_post(self) -> None:
        if self.path == "/api/login":
            length = int(self.headers.get("Content-Length", "0"))
            body_raw = self.rfile.read(length) if length > 0 else b"{}"
            try:
                payload = json.loads(body_raw.decode("utf-8"))
            except Exception:
                payload = {}
            user_in = str(payload.get("username", "")).strip()
            pass_in = str(payload.get("password", ""))
            perfil_login = _autenticar_usuario(user_in, pass_in)
            if perfil_login:
                token = _generar_token_sesion(perfil_login["username"])
                cookie_val = f"{COOKIE_NAME}={token}; Path=/; HttpOnly; Max-Age=604800; SameSite=Lax"
                _registrar_bitacora_empleado(
                    usuario=perfil_login["username"],
                    nombre=perfil_login.get("nombre", user_in),
                    rol=perfil_login.get("rol", "RECEPCIONISTA"),
                    accion="INICIO_SESION",
                    detalle=f"Ingresó al sistema web ({', '.join(perfil_login.get('modulos', []))})",
                )
                self._send_json({"ok": True, "rol": perfil_login.get("rol"), "nombre": perfil_login.get("nombre")}, 200, extra_headers={"Set-Cookie": cookie_val})
            else:
                time.sleep(0.6)
                self._send_json({"ok": False, "error": "Usuario o contraseña incorrectos."}, 401)
            return

        if self.path == "/api/observatorio/sync":
            ok_s, msg_s = observatorio_vial_cea.observatorio_service.sincronizar_con_ansv_ahora()
            self._send_json({"ok": ok_s, "message": msg_s})
            return

        if self.path == "/api/observatorio/upload_csv":
            try:
                content_type = self.headers.get("Content-Type", "")
                length = int(self.headers.get("Content-Length", "0"))
                raw_body = self.rfile.read(length) if length > 0 else b""
                parts = _parse_multipart_binary(content_type, raw_body)
                if "file" not in parts:
                    self._send_json({"ok": False, "message": "No se recibió el archivo CSV."}, 400)
                    return
                fn_csv, csv_bytes = parts["file"]
                ok_u, msg_u = observatorio_vial_cea.observatorio_service.guardar_csv_manual(
                    filename=fn_csv or "datos_ansv.csv",
                    file_bytes=csv_bytes,
                )
                self._send_json({"ok": ok_u, "message": msg_u})
            except Exception as exc:
                self._send_json({"ok": False, "message": f"Error al guardar CSV: {exc}"}, 500)
            return

        if not self._is_authenticated():
            self._send_json({"ok": False, "error": "No autorizado. Inicie sesión primero."}, 401)
            return

        if self.path == "/api/upload":
            try:
                content_type = self.headers.get("Content-Type", "")
                length = int(self.headers.get("Content-Length", "0"))
                raw_body = self.rfile.read(length) if length > 0 else b""
                parts = _parse_multipart_binary(content_type, raw_body)

                file_bytes = None
                filename = "matriculas.xlsx"
                ingresos_bytes = None
                ingresos_filename = ""
                periodicidad = "ANUAL"

                if "file" in parts:
                    fn, payload = parts["file"]
                    file_bytes = payload
                    if fn:
                        filename = fn

                if "ingresos_file" in parts:
                    fn_ing, payload_ing = parts["ingresos_file"]
                    # Si el usuario seleccionó por error el mismo archivo de matrículas en ambos campos,
                    # ignorar el segundo para no duplicar ni alterar el cruce.
                    if payload_ing and fn_ing and payload_ing != file_bytes:
                        ingresos_bytes = payload_ing
                        ingresos_filename = fn_ing

                if "periodicidad" in parts:
                    _, val_p = parts["periodicidad"]
                    if val_p:
                        periodicidad = val_p.decode("utf-8", errors="ignore").strip() or "ANUAL"

                modo_bd = "INCORPORAR"
                if "modo_bd" in parts:
                    _, val_m = parts["modo_bd"]
                    if val_m:
                        modo_bd = val_m.decode("utf-8", errors="ignore").strip().upper() or "INCORPORAR"
                modo_incorporar = modo_bd != "REEMPLAZAR"

                if not file_bytes:
                    self._send_json({"ok": False, "error": "No se recibió el archivo de matrículas."}, 400)
                    return

                hallazgos = procesar_y_actualizar_estado(
                    file_bytes,
                    filename=filename,
                    periodicidad=periodicidad,
                    origen_ingresos=ingresos_bytes,
                    ingresos_filename=ingresos_filename,
                    modo_incorporar=modo_incorporar,
                )
                m_stats = (hallazgos.get("metadata") or {}).get("merge_stats") or {}
                self._send_json({
                    "ok": True,
                    "filename": filename,
                    "ingresos_filename": ingresos_filename,
                    "periodicidad": periodicidad,
                    "modo_bd": modo_bd,
                    "merge_stats": m_stats,
                    "updated_at": STATE["updated_at"],
                    "hallazgos": hallazgos,
                })
            except Exception as exc:
                logger.exception("Error procesando archivos subidos")
                self._send_json({"ok": False, "error": str(exc)}, 500)
            return

        if self.path == "/api/tarifas":
            try:
                length = int(self.headers.get("Content-Length", "0"))
                body = json.loads(self.rfile.read(length).decode("utf-8")) if length > 0 else {}
                nuevas_tarifas = body.get("tarifas", {})
                if STATE["raw_matriculas_bytes"]:
                    hallazgos = procesar_y_actualizar_estado(
                        STATE["raw_matriculas_bytes"],
                        filename=STATE["filename"],
                        periodicidad=STATE["periodicidad"],
                        origen_ingresos=STATE["raw_ingresos_bytes"],
                        ingresos_filename=STATE["ingresos_filename"],
                        tarifas_base=nuevas_tarifas,
                    )
                    self._send_json({
                        "ok": True,
                        "filename": STATE["filename"],
                        "ingresos_filename": STATE["ingresos_filename"],
                        "periodicidad": STATE["periodicidad"],
                        "hallazgos": hallazgos,
                    })
                else:
                    self._send_json({"ok": False, "error": "No hay archivo cargado."}, 400)
            except Exception as exc:
                self._send_json({"ok": False, "error": str(exc)}, 500)
            return

        if self.path in ("/api/send/email", "/api/send/telegram"):
            length = int(self.headers.get("Content-Length", "0"))
            body_raw = self.rfile.read(length) if length > 0 else b"{}"
            try:
                payload = json.loads(body_raw.decode("utf-8"))
            except Exception:
                payload = {}

            if not STATE["hallazgos"]:
                self._send_json({"ok": False, "message": "Primero carga un archivo de matrículas."}, 400)
                return

            if self.path == "/api/send/email":
                email_extra = [payload.get("email", "").strip()] if payload.get("email") else None
                ok, msg_out = auditor_cea.enviar_correo_auditoria(
                    STATE["hallazgos"],
                    STATE["excel_bytes"],
                    STATE["pdf_bytes"],
                    periodicidad=STATE["periodicidad"],
                    destinatarios_extra=email_extra,
                )
                self._send_json({"ok": ok, "message": msg_out})
                return

            if self.path == "/api/send/telegram":
                chat_id = payload.get("chat_id", "").strip() or None
                ok, msg_out = auditor_cea.enviar_telegram_auditoria_sync(
                    STATE["hallazgos"],
                    STATE["excel_bytes"],
                    STATE["pdf_bytes"],
                    periodicidad=STATE["periodicidad"],
                    chat_id=chat_id,
                )
                self._send_json({"ok": ok, "message": msg_out})
                return

        if self.path == "/api/bot/action":
            length = int(self.headers.get("Content-Length", "0"))
            body_raw = self.rfile.read(length) if length > 0 else b"{}"
            try:
                payload = json.loads(body_raw.decode("utf-8"))
            except Exception:
                payload = {}

            action = str(payload.get("action", "")).strip()
            u_sesion = self._get_current_user() or {"username": AUDITOR_USER, "nombre": "Administrador", "rol": "ADMIN_TOTAL"}
            u_tag = f"{u_sesion.get('nombre', 'Usuario')} ({u_sesion.get('username', '')})"

            if action == "cartera_guardar_matricula_manual":
                existentes = STATE.get("registros_historicos") or []
                if not existentes:
                    existentes, _ = auditor_cea.cargar_bd_historica()
                meta_actual = dict(STATE.get("metadata_historica") or (STATE.get("hallazgos") or {}).get("metadata") or {})
                corte_dt = auditor_cea._parse_date(meta_actual.get("fecha_final")) or datetime.datetime.now()
                tarifas_usar = STATE.get("tarifas_base") or dict(auditor_cea.TARIFAS_BASE_DEFAULT)

                nuevos_regs, reg_guardado, msg_acc = auditor_cea.guardar_o_actualizar_matricula_manual(
                    datos=payload,
                    existentes=existentes,
                    fecha_corte=corte_dt,
                    tarifas_base=tarifas_usar,
                )
                hallazgos = _recalcular_estado_desde_bd_historica(
                    registros=nuevos_regs,
                    metadata=meta_actual,
                    filename=STATE.get("filename") or "Base_Historica_CEA.xlsx",
                    periodicidad=STATE.get("periodicidad") or "ANUAL",
                    ingresos_filename=STATE.get("ingresos_filename") or "",
                    tarifas_base=tarifas_usar,
                )
                _registrar_bitacora_empleado(
                    u_sesion["username"], u_sesion["nombre"], u_sesion["rol"],
                    "ACTUALIZAR_BD_CARTERA",
                    msg_acc,
                )
                self._send_json({
                    "ok": True,
                    "message": msg_acc,
                    "registro": reg_guardado,
                    "filename": STATE["filename"],
                    "ingresos_filename": STATE["ingresos_filename"],
                    "periodicidad": STATE["periodicidad"],
                    "updated_at": STATE["updated_at"],
                    "hallazgos": hallazgos,
                })
                return

            if action == "empleado_guardar":
                if u_sesion.get("rol") != "ADMIN_TOTAL":
                    self._send_json({"ok": False, "message": "Solo el Administrador General puede crear o modificar permisos de empleados."}, 403)
                    return
                u_new = str(payload.get("username", "")).strip()
                p_new = str(payload.get("password", "")).strip()
                nom_new = str(payload.get("nombre", "")).strip()
                rol_new = str(payload.get("rol", "RECEPCIONISTA")).strip().upper()
                mods_new = payload.get("modulos") or ["mod_cartera", "mod_clases", "mod_nomina", "mod_inasistencias", "mod_pesv"]
                if not u_new or not p_new or not nom_new:
                    self._send_json({"ok": False, "message": "Faltan datos del empleado (usuario, contraseña o nombre)."}, 400)
                    return
                usuarios_map = _cargar_usuarios_empleados()
                usuarios_map[u_new.lower()] = {
                    "username": u_new,
                    "password": p_new,
                    "nombre": nom_new,
                    "rol": rol_new,
                    "cargo": f"{rol_new} ({len(mods_new)} módulos)",
                    "modulos": list(mods_new),
                    "activo": True,
                }
                _guardar_usuarios_empleados(usuarios_map)
                _registrar_bitacora_empleado(
                    u_sesion["username"], u_sesion["nombre"], u_sesion["rol"],
                    "PERMISOS_EMPLEADO",
                    f"Configuró usuario '{u_new}' ({nom_new} - {rol_new}) con acceso a: {', '.join(mods_new)}"
                )
                self._send_json({
                    "ok": True,
                    "message": f"Usuario '{u_new}' ({nom_new}) guardado con acceso a {len(mods_new)} módulo(s).",
                })
                return

            if action == "empleado_eliminar":
                if u_sesion.get("rol") != "ADMIN_TOTAL":
                    self._send_json({"ok": False, "message": "Solo el Administrador General puede eliminar cuentas."}, 403)
                    return
                u_del = str(payload.get("username", "")).strip().lower()
                if u_del == AUDITOR_USER.lower():
                    self._send_json({"ok": False, "message": "No puedes eliminar la cuenta principal de Gerencia."}, 400)
                    return
                usuarios_map = _cargar_usuarios_empleados()
                if u_del in usuarios_map:
                    eliminado = usuarios_map.pop(u_del)
                    _guardar_usuarios_empleados(usuarios_map)
                    _registrar_bitacora_empleado(
                        u_sesion["username"], u_sesion["nombre"], u_sesion["rol"],
                        "ELIMINAR_EMPLEADO",
                        f"Eliminó cuenta '{u_del}' ({eliminado.get('nombre', '')})"
                    )
                self._send_json({"ok": True, "message": f"Cuenta '{u_del}' eliminada."})
                return

            if action == "conectar_sheets":
                raw_json = str(payload.get("credentials_json", "")).strip()
                try:
                    parsed_cred = json.loads(raw_json)
                    if "refresh_token" in parsed_cred and "client_id" in parsed_cred:
                        bot_agenda_cea.TOKEN_CALENDAR_PATH.write_text(raw_json, encoding="utf-8")
                        self._send_json({
                            "ok": True,
                            "message": "Token OAuth de Google Calendar guardado y conectado exitosamente.",
                        })
                        return
                except Exception:
                    pass
                res = clases_core.sheets_service.configurar_credenciales_en_vivo(raw_json)
                if isinstance(res, tuple):
                    res = {"ok": bool(res[0]), "message": str(res[1])}
                self._send_json(res)
                return

            if action == "registrar_alumno":
                cedula = str(payload.get("cedula", "")).strip()
                nombre = str(payload.get("nombre", "")).strip().upper()
                categoria = str(payload.get("categoria", "A2")).strip().upper()
                telefono = str(payload.get("telefono", "")).strip()
                estado_runt = str(payload.get("estado_runt", "")).strip() or None
                ok = clases_core.sheets_service.asegurar_alumno(cedula, nombre, categoria, telefono, estado_runt=estado_runt)
                _registrar_bitacora_empleado(u_sesion["username"], u_sesion["nombre"], u_sesion["rol"], "MATRICULAR_ALUMNO", f"Registró alumno {nombre} ({categoria} - CC {cedula} - {estado_runt or 'No Registrado RUNT'})")
                self._send_json({
                    "ok": bool(ok),
                    "message": f"Alumno {nombre} ({categoria} - C.C. {cedula}) guardado ({estado_runt or 'No Registrado RUNT'}).",
                })
                return

            if action == "estado_runt_alumno":
                cedula = str(payload.get("cedula", "")).strip()
                nombre = str(payload.get("nombre", "")).strip().upper()
                categoria = str(payload.get("categoria", "A2")).strip().upper()
                estado_runt = str(payload.get("estado_runt", "Cargue RUNT")).strip()
                res_r = clases_core.sheets_service.actualizar_estado_runt_alumno(
                    documento_o_nombre=cedula or nombre,
                    categoria=categoria,
                    estado_runt=estado_runt,
                    nombre_alumno=nombre,
                )
                est_final = res_r.get("Estado_RUNT", estado_runt)
                # Sincronizar también hacia la hoja de Control Digital RUNT (1XHbqfDa7dK5CWcgWKNgJZiXN_OSrjJCEJJu8F4z-u2M)
                try:
                    es_carg = est_final == "Cargue RUNT"
                    metas_c = {"A1": (25, 3, 8), "A2": (25, 3, 15), "B1": (25, 5, 20), "C1": (30, 5, 30), "RECAT_C1": (8, 0, 8)}.get(categoria, (25, 3, 15))
                    runt_control_cea.runt_service.guardar_o_actualizar_practica_runt({
                        "cedula": cedula,
                        "nombre": nombre or res_r.get("Nombre_Completo", ""),
                        "categoria": categoria,
                        "prioridad": "Termino" if es_carg else "EN PRACTICA",
                        "teoria": metas_c[0],
                        "taller": metas_c[1],
                        "practica": metas_c[2] if es_carg else 0,
                        "estado_runt": "CARGADO RUNT" if es_carg else "PENDIENTE RUNT",
                        "cargue_diario": "PRACTICA" if es_carg else "",
                    })
                except Exception:
                    pass
                _registrar_bitacora_empleado(
                    u_sesion["username"], u_sesion["nombre"], u_sesion["rol"],
                    "CONTROL_RUNT",
                    f"Marcó '{est_final}' para {nombre or cedula} (Categoría {categoria})"
                )
                self._send_json({
                    "ok": True,
                    "message": (
                        f"✅ {nombre or cedula} ({categoria}) marcado como 'Cargue RUNT' y ocultado del listado pendiente."
                        if est_final == "Cargue RUNT"
                        else f"⏳ {nombre or cedula} ({categoria}) devuelto a 'No Registrado RUNT' (visible en pendientes)."
                    ),
                    "alumno": res_r,
                })
                return

            if action == "registrar_clase":
                cedula = str(payload.get("cedula", "")).strip()
                alumno = str(payload.get("alumno", "")).strip().upper()
                categoria = str(payload.get("categoria", "A2")).strip().upper()
                instructor = str(payload.get("instructor", "")).strip().upper()
                clases_str = str(payload.get("clases_str", "1, 2")).strip()
                nums = [int(x) for x in re.findall(r"\d+", clases_str)] or [1]
                clases_core.sheets_service.asegurar_alumno(cedula, alumno, categoria, "")
                ahora_col = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=-5)))
                ts_str = ahora_col.strftime("%Y-%m-%d %H:%M:%S")
                f_str = ahora_col.strftime("%Y-%m-%d")
                h_str = ahora_col.strftime("%H:%M:%S")
                clases_fmt = clases_core.formatear_clases(nums)
                tema_auto = clases_core.obtener_temas_para_clases(categoria, nums)
                fila_idx = clases_core.sheets_service.registrar_clase(
                    timestamp=ts_str,
                    fecha_reporte=f_str,
                    hora_reporte=h_str,
                    instructor=instructor,
                    id_alumno=cedula,
                    nombre_alumno=alumno,
                    categoria=categoria,
                    clases_reportadas=clases_fmt,
                    tema_registrado=tema_auto,
                    url_foto_drive="Registrado en Web",
                    validacion_horario="EN TIEMPO (WEB)",
                    observacion_instructor=f"Registrado por {u_tag} desde Dashboard Web",
                )
                self._send_json({
                    "ok": bool(fila_idx),
                    "message": f"Clase(s) {clases_fmt} registrada(s) para {alumno} ({categoria}) con {instructor}.",
                })
                return

            if action == "registrar_inasistencia":
                res_in = clases_core.sheets_service.registrar_inasistencia(
                    documento_alumno=str(payload.get("cedula", "")).strip(),
                    nombre_alumno=str(payload.get("alumno", "")).strip().upper(),
                    categoria=str(payload.get("categoria", "B1")).strip().upper(),
                    clases_faltadas="No especificada",
                    nombre_instructor=str(payload.get("instructor", "")).strip().upper(),
                    no_horas=2,
                    valor_hora=25000,
                    estado=str(payload.get("estado_pago", "Pendiente Pago")).strip(),
                    fecha=str(payload.get("fecha_clase", "")).strip(),
                    hora=str(payload.get("hora_programada", "")).strip(),
                )
                self._send_json({
                    "ok": bool(res_in),
                    "message": f"Inasistencia #{res_in.get('Consecutivo', '')} registrada y sincronizada con el Bot.",
                })
                return

            if action == "estado_inasistencia":
                fila = str(payload.get("fila", "")).strip()
                nuevo_estado = str(payload.get("estado_pago", "Cancelo Efectivo")).strip()
                res_up = clases_core.sheets_service.actualizar_estado_inasistencia(fila, nuevo_estado)
                _registrar_bitacora_empleado(u_sesion["username"], u_sesion["nombre"], u_sesion["rol"], "COBRO_INASISTENCIA", f"Actualizó inasistencia #{fila} a {nuevo_estado}")
                self._send_json({
                    "ok": res_up is not None,
                    "message": f"Inasistencia #{fila} actualizada a {nuevo_estado}.",
                })
                return

            if action == "registrar_mantenimiento":
                res_mt = clases_core.sheets_service.registrar_mantenimiento(
                    placa=str(payload.get("placa", "")).strip().upper(),
                    tipo_mantenimiento=str(payload.get("tipo_mantenimiento", "PREVENTIVO")).strip(),
                    establecimiento=str(payload.get("taller", "TALLER AUTORIZADO")).strip().upper(),
                    nit_cc="0",
                    actividad_realizada=str(payload.get("actividad", "")).strip().upper(),
                    nombre_tecnico=str(payload.get("mecanico", "")).strip().upper(),
                    rol="MECANICO",
                    observaciones=f"Registrado desde Web por {u_tag}",
                    kilometraje_actual=str(payload.get("kilometraje", "0")).strip(),
                )
                self._send_json({
                    "ok": bool(res_mt),
                    "message": f"Mantenimiento registrado en SISI-F2-08 para placa {res_mt.get('Placa', '')}.",
                })
                return

            if action == "vehiculo_agregar":
                res_vh = clases_core.sheets_service.agregar_vehiculo_pesv(
                    placa=str(payload.get("placa", "")).strip().upper(),
                    tipo=str(payload.get("tipo", "CARRO")).strip().upper(),
                    categoria=str(payload.get("categoria", "B1")).strip().upper(),
                    marca=str(payload.get("marca", "")).strip().upper(),
                )
                self._send_json({
                    "ok": bool(res_vh.get("ok")),
                    "message": f"Vehículo {res_vh.get('placa', '')} activo en flota PESV.",
                })
                return

            if action == "vehiculo_retirar":
                res_ret = clases_core.sheets_service.retirar_vehiculo_pesv(
                    placa=str(payload.get("placa", "")).strip().upper(),
                )
                self._send_json({
                    "ok": bool(res_ret.get("ok")),
                    "message": f"Vehículo {res_ret.get('placa', '')} retirado de flota activa.",
                })
                return

            if action == "robot_guardar_config":
                cfg_saved = _guardar_robot_config(payload)
                self._send_json({
                    "ok": True,
                    "message": f"Configuración guardada (Telegram Chat: {cfg_saved.get('tg_chat') or 'Predeterminado'}, WhatsApp: +{cfg_saved.get('wa_phone')}).",
                })
                return

            if action == "robot_despachar_reporte":
                reporte_tipo = str(payload.get("reporte", "CARTERA_PDF")).strip().upper()
                canal = str(payload.get("canal", "TELEGRAM")).strip().upper()
                chat_id = str(payload.get("chat_id", "")).strip()
                wa_phone = str(payload.get("wa_phone", "")).strip()
                alumno_q = str(payload.get("alumno_q", "")).strip()
                cat_val = str(payload.get("categoria", "A2")).strip().upper()
                fecha_str = datetime.date.today().strftime("%Y%m%d")

                file_bytes = b""
                fname = f"Reporte_CEA_{fecha_str}.pdf"
                download_path = "/api/download/pdf"
                caption_extra = "Reporte Oficial CEA Auto Oriente"

                if reporte_tipo == "CARTERA_PDF":
                    file_bytes = STATE.get("pdf_bytes") or b""
                    fname = f"Resumen_Ejecutivo_Cartera_CEA_{fecha_str}.pdf"
                    download_path = "/api/download/pdf"
                    caption_extra = "📊 Dictamen Ejecutivo de Cartera y Auditoría Multi-Curso"
                elif reporte_tipo == "CARTERA_EXCEL":
                    file_bytes = STATE.get("excel_bytes") or b""
                    fname = f"Auditoria_Cartera_MultiPestana_CEA_{fecha_str}.xlsx"
                    download_path = "/api/download/excel"
                    caption_extra = "📗 Libro Excel Multi-Pestaña de Auditoría y Cartera CEA"
                elif reporte_tipo in ("FICHA_ALUMNO", "ALUMNO_F59"):
                    q_usar = alumno_q or "SEGUNDO MURILLO"
                    res_al = clases_core.sheets_service.obtener_historial_alumno(q_usar)
                    alumno = res_al.get("alumno") or {"cedula": q_usar, "nombre": q_usar, "categoria": cat_val, "telefono": ""}
                    if cat_val and cat_val in clases_core.FORMATOS_OFICIALES:
                        alumno["categoria"] = cat_val
                    registros = [
                        r for r in (res_al.get("registros") or [])
                        if not cat_val or str(r.get("categoria", "")).strip().upper() == cat_val
                    ] or (res_al.get("registros") or [])
                    pdf_buf = clases_core.generar_pdf_seguimiento(alumno, registros)
                    file_bytes = pdf_buf.getvalue()
                    safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", str(alumno.get("nombre", "Alumno")))
                    fname = f"Ficha_Seguimiento_{alumno.get('categoria', 'A2')}_{safe_name}.pdf"
                    download_path = f"/api/bot/pdf/alumno?q={urllib.parse.quote(q_usar)}&cat={urllib.parse.quote(cat_val)}"
                    caption_extra = f"🎓 Ficha Oficial de Clases Prácticas — {alumno.get('nombre')} ({alumno.get('categoria')})"
                elif reporte_tipo == "NOMINA_PDF":
                    rep_nom = clases_core.sheets_service.obtener_clases_por_instructor_periodo("TODOS", "mes")
                    file_bytes = clases_core.generar_pdf_liquidacion_instructor(rep_nom).getvalue()
                    fname = f"Nomina_Liquidacion_Instructores_{fecha_str}.pdf"
                    download_path = "/api/bot/pdf/nomina?instructor=TODOS&periodo=mes"
                    caption_extra = "👨‍🏫 Liquidación Mensual de Instructores y Auditoría de Horarios"
                elif reporte_tipo == "INASISTENCIAS_PDF":
                    rep_in = clases_core.sheets_service.obtener_inasistencias_periodo("todo", "TODOS")
                    file_bytes = clases_core.generar_pdf_inasistencias_periodo(rep_in).getvalue()
                    fname = f"Reporte_Inasistencias_CEA_{fecha_str}.pdf"
                    download_path = "/api/bot/pdf/inasistencias?periodo=todo&estado=TODOS"
                    caption_extra = "🚫 Reporte Consolidado de Inasistencias y Recaudo"
                elif reporte_tipo == "PESV_PDF":
                    rep_mt = clases_core.sheets_service.obtener_mantenimientos_periodo("todo", "TODAS")
                    file_bytes = clases_core.generar_pdf_mantenimientos_pesv(rep_mt).getvalue()
                    fname = f"PESV_SISI_F2_08_Flota_{fecha_str}.pdf"
                    download_path = "/api/bot/pdf/mantenimientos?periodo=todo&placa=TODAS"
                    caption_extra = "🔧 Historial Oficial PESV SISI-F2-08 — Mantenimiento de Flota CEA"

                if not file_bytes:
                    self._send_json({"ok": False, "message": "No se pudo generar el archivo solicitado."}, 400)
                    return

                res_arch = _clasificar_y_respaldar_archivo(
                    filename=fname,
                    file_bytes=file_bytes,
                    canal=canal,
                    chat_id=chat_id,
                    wa_phone=wa_phone,
                )
                res_arch["download_url"] = download_path
                if canal in ("WHATSAPP", "AMBOS") and not res_arch.get("wa_url"):
                    _, _, wa_link = _despachar_whatsapp_api_o_link(
                        file_bytes, fname, caption_extra, phone=wa_phone, download_path=download_path
                    )
                    res_arch["wa_url"] = wa_link
                self._send_json(res_arch)
                return

            if action == "agenda_conectar_config":
                res_cfg = bot_agenda_cea.configurar_credenciales_agenda_en_vivo(
                    tg_token=str(payload.get("tg_token", "")).strip(),
                    chat_id=str(payload.get("chat_id", "")).strip(),
                    calendar_token_json=str(payload.get("calendar_token_json", "")).strip(),
                )
                self._send_json(res_cfg)
                return

            if action == "agenda_crear_evento":
                ok_ev, msg_ev = bot_agenda_cea.crear_evento_agenda(
                    titulo=str(payload.get("titulo") or payload.get("summary") or "").strip(),
                    fecha=str(payload.get("fecha", "")).strip(),
                    hora=str(payload.get("hora", "09:00")).strip(),
                    duracion_min=int(payload.get("duracion_min") or 60),
                    es_diario=bool(payload.get("es_diario") or payload.get("recurrente_diario")),
                    responsable=str(payload.get("responsable") or payload.get("descripcion") or "Gerencia CEA").strip(),
                    chat_id_override=str(payload.get("chat_id", "")).strip(),
                )
                self._send_json({"ok": ok_ev, "message": msg_ev})
                return

            if action == "agenda_eliminar_evento":
                ok_del, msg_del = bot_agenda_cea.eliminar_evento_agenda(
                    ev_id=str(payload.get("id") or payload.get("event_id") or "").strip()
                )
                self._send_json({"ok": ok_del, "message": msg_del})
                return

            if action == "agenda_recordar_telegram":
                ok_tg, msg_tg = bot_agenda_cea.enviar_recordatorio_telegram_agenda(
                    titulo=str(payload.get("titulo", "Recordatorio Agenda CEA")).strip(),
                    fecha=str(payload.get("fecha", "")).strip(),
                    hora=str(payload.get("hora", "")).strip(),
                    responsable=str(payload.get("responsable", "CEA Auto Oriente")).strip(),
                    chat_id=str(payload.get("chat_id", "")).strip(),
                )
                self._send_json({"ok": ok_tg, "message": msg_tg})
                return

            if action == "baul_guardar_item":
                ok_b, msg_b = bot_agenda_cea.guardar_item_baul(
                    titulo=str(payload.get("titulo", "")).strip(),
                    contenido=str(payload.get("contenido", "")).strip(),
                    categoria=str(payload.get("categoria", "NOTA")).strip(),
                    origen="Dashboard Web",
                    notificar_tg=True,
                    chat_id=str(payload.get("chat_id", "")).strip(),
                )
                self._send_json({"ok": ok_b, "message": msg_b})
                return

            if action == "baul_eliminar_item":
                ok_bd, msg_bd = bot_agenda_cea.eliminar_item_baul(
                    item_id=str(payload.get("id", "")).strip()
                )
                self._send_json({"ok": ok_bd, "message": msg_bd})
                return

            if action == "baul_enviar_telegram":
                ok_bt, msg_bt = bot_agenda_cea.enviar_item_baul_telegram(
                    item_id=str(payload.get("id", "")).strip(),
                    chat_id=str(payload.get("chat_id", "")).strip(),
                )
                self._send_json({"ok": ok_bt, "message": msg_bt})
                return

            if action == "wa_registrar_lead":
                ok_wl, msg_wl = crm_whatsapp_cea.registrar_o_actualizar_lead(payload)
                self._send_json({"ok": ok_wl, "message": msg_wl})
                return

            if action == "wa_cambiar_estado":
                ok_st, msg_st = crm_whatsapp_cea.cambiar_estado_lead(
                    lead_id=str(payload.get("id", "")).strip(),
                    nuevo_estado=str(payload.get("estado_ia") or payload.get("estado") or "REMARKETING_ENVIADO").strip(),
                    incrementar_remarketing=bool(payload.get("incrementar_remarketing", False)),
                )
                self._send_json({"ok": ok_st, "message": msg_st})
                return

            if action == "wa_importar_lote":
                cnt_imp, msg_imp = crm_whatsapp_cea.importar_lote_whatsapp_texto(
                    texto_bruto=str(payload.get("texto") or payload.get("raw_text") or ""),
                    categoria_default=str(payload.get("categoria_default", "A2")).strip(),
                )
                self._send_json({"ok": cnt_imp > 0, "count": cnt_imp, "message": msg_imp})
                return

            if action == "wa_guardar_plantilla":
                cod = str(payload.get("codigo", "IA_RECUPERACION_50")).strip()
                txt = str(payload.get("texto", "")).strip()
                nom = str(payload.get("nombre") or crm_whatsapp_cea.PLANTILLAS_DEFAULT.get(cod, {}).get("nombre") or cod)
                crm_whatsapp_cea.guardar_plantilla(codigo=cod, nombre=nom, texto=txt)
                self._send_json({"ok": True, "message": f"Plantilla '{nom}' actualizada y lista para remarketing."})
                return

            if action == "runt_guardar_practica":
                res_rp = runt_control_cea.runt_service.guardar_o_actualizar_practica_runt(payload)
                _registrar_bitacora_empleado(
                    u_sesion["username"], u_sesion["nombre"], u_sesion["rol"],
                    "REPORTE_RUNT_DIGITAL",
                    res_rp.get("message", "Actualizó conteo RUNT")
                )
                self._send_json(res_rp)
                return

            if action == "runt_cambiar_no_runt":
                fila_nr = int(payload.get("fila") or 0)
                est_nr = str(payload.get("estado_registro", "REGISTRADO")).strip()
                nov_nr = str(payload.get("novedad", "")).strip()
                res_nr = runt_control_cea.runt_service.cambiar_estado_no_runt(fila_nr, est_nr, nov_nr)
                self._send_json(res_nr)
                return

            if action == "runt_sincronizar_bot":
                res_sb = runt_control_cea.runt_service.sincronizar_clases_bot_a_runt()
                _registrar_bitacora_empleado(
                    u_sesion["username"], u_sesion["nombre"], u_sesion["rol"],
                    "SINCRONIZAR_BOT_A_RUNT",
                    res_sb.get("message", "Sincronizó horas prácticas del Bot hacia hoja RUNT")
                )
                self._send_json(res_sb)
                return

            if action == "runt_configurar_sheet":
                nuevo_id = str(payload.get("sheet_id", "")).strip()
                res_cs = runt_control_cea.runt_service.cambiar_sheet_id(nuevo_id)
                self._send_json(res_cs)
                return

            if action == "runt_cargue_masivo":
                res_cm = runt_control_cea.runt_service.cargue_masivo_teoria_o_taller(payload)
                _registrar_bitacora_empleado(
                    u_sesion["username"], u_sesion["nombre"], u_sesion["rol"],
                    "CARGUE_MASIVO_RUNT",
                    res_cm.get("message", "Cargue masivo Teoría/Taller RUNT")
                )
                self._send_json(res_cm)
                return

            if action == "runt_deshacer_masivo":
                res_dm = runt_control_cea.runt_service.deshacer_o_restaurar_cargue_masivo(payload)
                _registrar_bitacora_empleado(
                    u_sesion["username"], u_sesion["nombre"], u_sesion["rol"],
                    "DESHACER_MASIVO_RUNT",
                    res_dm.get("message", "Revirtió cargue masivo RUNT")
                )
                self._send_json(res_dm)
                return

            if action == "runt_sumar_practica_individual":
                res_sp = runt_control_cea.runt_service.sumar_lote_practica_individual(payload)
                _registrar_bitacora_empleado(
                    u_sesion["username"], u_sesion["nombre"], u_sesion["rol"],
                    "REPORTE_PRACTICA_RUNT",
                    res_sp.get("message", "Sumó horas prácticas individuales RUNT")
                )
                self._send_json(res_sp)
                return

            if action == "runt_reiniciar_vencido":
                res_rv = runt_control_cea.runt_service.reiniciar_curso_vencido_con_sancion(payload)
                _registrar_bitacora_empleado(
                    u_sesion["username"], u_sesion["nombre"], u_sesion["rol"],
                    "REINICIO_RUNT_SANCION",
                    res_rv.get("message", "Reinició curso vencido con sanción económica")
                )
                self._send_json(res_rv)
                return

            self._send_json({"ok": False, "message": f"Acción desconocida: {action}"}, 400)
            return

        if self.path == "/api/bot/robot_upload":
            try:
                content_type = self.headers.get("Content-Type", "")
                length = int(self.headers.get("Content-Length", "0"))
                raw_body = self.rfile.read(length) if length > 0 else b""
                parts = _parse_multipart_binary(content_type, raw_body)
                if "file" not in parts:
                    self._send_json({"ok": False, "message": "No se recibió ningún archivo."}, 400)
                    return
                fn, fbytes = parts["file"]
                filename = fn or "documento.pdf"
                canal = "SOLO_CLASIFICAR"
                if "canal" in parts and parts["canal"][1]:
                    canal = parts["canal"][1].decode("utf-8", errors="ignore").strip() or "SOLO_CLASIFICAR"
                chat_id = ""
                if "chat_id" in parts and parts["chat_id"][1]:
                    chat_id = parts["chat_id"][1].decode("utf-8", errors="ignore").strip()
                wa_phone = ""
                if "wa_phone" in parts and parts["wa_phone"][1]:
                    wa_phone = parts["wa_phone"][1].decode("utf-8", errors="ignore").strip()

                res_up = _clasificar_y_respaldar_archivo(
                    filename=filename,
                    file_bytes=fbytes,
                    canal=canal,
                    chat_id=chat_id,
                    wa_phone=wa_phone,
                )
                self._send_json(res_up)
            except Exception as exc:
                self._send_json({"ok": False, "message": f"Error clasificando archivo: {exc}"}, 500)
            return

        self.send_response(404)
        self.end_headers()

    def log_message(self, format: str, *args: Any) -> None:
        return


def _loop_keep_alive() -> None:
    while True:
        time.sleep(600)
        if RENDER_EXTERNAL_URL:
            try:
                urllib.request.urlopen(f"{RENDER_EXTERNAL_URL.rstrip('/')}/health", timeout=15).read()
            except Exception:
                pass


def main() -> None:
    server = ThreadingHTTPServer(("0.0.0.0", PORT), AuditorWebHandler)
    threading.Thread(target=_loop_keep_alive, daemon=True).start()
    logger.info("🌐 Dashboard Web Multi-Curso de Auditoría CEA activo en http://0.0.0.0:%s", PORT)
    server.serve_forever()


if __name__ == "__main__":
    main()
