"""
Módulo de Evaluación Docente (Google Form en Vivo) y Seguimiento Pedagógico de Alumnos — CEA Auto Oriente.

1. Lectura, normalización y analítica en tiempo real del Google Form de Evaluación Docente:
   - Sheet ID: 1F6Sxe_csVal4EIOOBGXZlryxSPXutbRDJpcIJczPdtU
   - Form URL: https://docs.google.com/forms/d/1z_pYRCxH0GYgJbUjHmz8CaJ6b3T9g2l-HeXD9QKMhNI/viewform
   - Normaliza nombres de instructores con variantes históricas, puntúa automáticamente respuestas
     numéricas (1-5) y cualitativas (10 preguntas actuales, sin duplicar col 5/21), detecta alertas
     de calidad docente y genera ranking por instructor.

2. Observaciones Pedagógicas por Botones (cero escritura) y Consulta de Progreso del Alumno:
   - Guarda en Google Sheets (pestaña 'Observaciones_Alumnos') y respaldo local JSON.
   - Cruza el avance real del alumno con el temario oficial (FORMATOS_OFICIALES) para mostrar al
     instructor: (a) objetivo oficial de la próxima clase y (b) recomendaciones de repaso dejadas
     en clases anteriores (ej. Reversa y parqueo, Arranque en pendiente).
"""

import datetime
import json
import logging
import os
import re
import threading
import time
import unicodedata
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import clases_core

logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent

EVAL_SHEET_ID_DEFAULT = os.getenv(
    "EVAL_DOCENTE_SHEET_ID",
    "1F6Sxe_csVal4EIOOBGXZlryxSPXutbRDJpcIJczPdtU",
).strip()

EVAL_FORM_URL_DEFAULT = os.getenv(
    "EVAL_DOCENTE_FORM_URL",
    "https://docs.google.com/forms/d/1z_pYRCxH0GYgJbUjHmz8CaJ6b3T9g2l-HeXD9QKMhNI/viewform",
).strip()

EVAL_CACHE_LOCAL_FILE = BASE_DIR / "cache_evaluacion_docente_cea.json"
OBS_LOCAL_FILE = BASE_DIR / "observaciones_pedagogicas_cea.json"
INST_ESTADO_LOCAL_FILE = BASE_DIR / "instructores_estado_cea.json"

SHEET_OBSERVACIONES = "Observaciones_Alumnos"
SHEET_ESTADO_INSTRUCTORES = "Estado_Instructores_CEA"
HEADERS_ESTADO_INSTRUCTORES = ["Instructor", "Estado", "Tipo", "Fecha_Cambio"]
HEADERS_OBSERVACIONES = [
    "ID_Obs",
    "Timestamp",
    "Fecha",
    "Documento",
    "Nombre_Alumno",
    "Categoria",
    "Clase_Num",
    "Instructor",
    "Nivel_Avance",
    "Recomendaciones_Proxima_Clase",
    "Fortalezas",
    "Estado_Refuerzo",
]

# =============================================================================
# CATÁLOGOS DE SELECCIÓN RÁPIDA (CERO ESCRITURA PARA INSTRUCTORES)
# =============================================================================
CATALOGO_NIVELES = [
    {"id": "AVANZA_BIEN", "label": "🟢 Avanza bien (Cumplió objetivo)", "color": "success"},
    {"id": "REFORZAR", "label": "🟡 En proceso (Reforzar próxima clase)", "color": "warning"},
    {"id": "PRIORITARIO", "label": "🔴 Requiere repaso prioritario", "color": "danger"},
]

CATALOGO_RECOMENDACIONES_CARRO = [
    "🔁 Reversa y parqueo",
    "⛰️ Arranque en pendiente",
    "⚙️ Embrague y cambio de marchas",
    "🚦 Señales de tránsito y pares",
    "🪞 Uso de espejos y cambio de carril",
    "🌀 Glorietas, cruces y tráfico denso",
    "🛑 Frenado y control de velocidad",
    "📐 Giros, trayectoria y distancia de seguridad",
    "🧭 Dominio del volante y punto de contacto",
    "✅ Sin pendientes (Listo para siguiente tema)",
]

CATALOGO_RECOMENDACIONES_MOTO = [
    "⚖️ Equilibrio y baja velocidad",
    "⚙️ Embrague y cambio de marchas",
    "🛑 Frenado controlado y de emergencia",
    "⛰️ Arranque en pendiente",
    "🐍 Eslalon, conos y giros en U",
    "🚦 Señales de tránsito y tráfico urbano",
    "🪞 Espejos, postura y cambio de carril",
    "✅ Sin pendientes (Listo para siguiente tema)",
]

CATALOGO_FORTALEZAS = [
    "👍 Buena postura y calma",
    "👍 Buen manejo de espejos",
    "👍 Respeta señales y pares",
    "👍 Buen control de embrague/freno",
    "👍 Excelente actitud y atención",
]


def _norm(txt: Any) -> str:
    s = str(txt or "").strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s)


def _normalizar_nombre_profesor(raw: Any) -> str:
    """Unifica variantes y errores tipográficos históricos del Google Form en nombres canónicos."""
    s_raw = str(raw or "").strip()
    if not s_raw:
        return "SIN ESPECIFICAR"
    n = _norm(s_raw)
    if n in ("profesor", "exelente", "excelente", "ninguno", "n/a", "na", ".", "-"):
        return "SIN ESPECIFICAR"

    # Si vienen 2 profesores separados por coma, tomar el primero para el cruce principal o dejarlos claros
    if "dayron" in n and "carlos" in n:
        return "DAYRON ANDRES SERRANO LOAIZA"

    reglas = [
        (["carreno", "jhoan sebastian", "jhon carreno"], "JHOAN SEBASTIAN CARREÑO SIERRA"),
        (["oscar mantilla", "mantilla"], "OSCAR MANTILLA"),
        (["yulian", "camacho"], "YULIAN CAMACHO"),
        (["carlos", "carlo", "caros", "gras", "grass"], "CARLOS ARTURO GOMEZ GRASS"),
        (["medaglia", "cesar"], "CESAR MAURICIO MEDAGLIA APARICIO"),
        (["dayron", "dairon", "dairo", "dayro", "dario", "serrano", "cerrano", "laoiza", "loaiza"], "DAYRON ANDRES SERRANO LOAIZA"),
        (["marco", "marcos", "villamizar", "villamisar"], "MARCO ANTONIO VILLAMIZAR"),
        (["rolando", "robando", "ramirez", "ramires"], "ROLANDO ANTONIO RAMIREZ GIRALDO"),
        (["martin", "parra"], "MARTIN FERNANDO SANCHEZ PARRA"),
        (["jesus", "sanchez", "sanches"], "JESUS SANCHEZ"),
        (["jhon alexander", "gomez sierra"], "JHON ALEXANDER GOMEZ SIERRA"),
        (["juan david", "mejia", "meca"], "JUAN DAVID MEJIA"),
        (["hermes", "becerra"], "HERMES GIOVANNY BECERRA MONTAÑEZ"),
        (["alcibiades"], "ALCIBIADES BUENO RIVERO"),
        (["leidy", "salamanca"], "LEIDY SALAMANCA"),
        (["segundo", "murillo"], "SEGUNDO MURILLO"),
        (["victor", "anaya"], "VICTOR ANAYA"),
        (["mauricio castro"], "MAURICIO CASTRO"),
    ]
    # Caso especial: "carlos medaglia" -> Cesar Mauricio Medaglia
    if "medaglia" in n:
        return "CESAR MAURICIO MEDAGLIA APARICIO"
    if "mauricio" in n and "castro" not in n:
        return "CESAR MAURICIO MEDAGLIA APARICIO"
    if n == "juan":
        return "JUAN DAVID MEJIA"

    for claves, canonico in reglas:
        if any(cl in n for cl in claves):
            return canonico

    return s_raw.upper()


def _parsear_fecha_eval(raw_fecha: str, raw_ts: str) -> str:
    for cand in (raw_fecha, raw_ts):
        s = str(cand or "").strip()
        if not s:
            continue
        m_iso = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
        if m_iso:
            return f"{int(m_iso.group(1)):04d}-{int(m_iso.group(2)):02d}-{int(m_iso.group(3)):02d}"
        m_dmy = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})", s)
        if m_dmy:
            d, m, y = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
            if m > 12 and d <= 12:
                d, m = m, d
            if y < 100:
                y += 2000
            return f"{y:04d}-{m:02d}-{d:02d}"
    return ""


def _puntuar_respuesta_criterio(valor: Any) -> Tuple[Optional[float], bool]:
    """
    Convierte una respuesta numérica ('1'-'5') o cualitativa del Google Form a escala 1.0-5.0.
    Retorna (puntaje, es_alerta_critica).
    """
    s = str(valor or "").strip()
    if not s:
        return None, False

    # 1. Respuesta numérica directa (histórico)
    if re.match(r"^[1-5]([.,]\d+)?$", s):
        val = float(s.replace(",", "."))
        return val, (val <= 3.0)

    n = _norm(s)

    # 2. Respuestas cualitativas actuales (si hay selección múltiple separada por coma, toma la más baja/conservadora)
    # Niveles críticos (1.0 - 2.0) -> Alerta roja
    criticos_1 = [
        "sin dar ninguna instruccion",
        "asuntos ajenos a la clase",
        "uso varias veces para asuntos ajenos",
        "nada claro",
        "nada respetuoso",
        "nunca",
    ]
    if any(k in n for k in criticos_1):
        return 1.0, True

    criticos_2 = [
        "distraido ocasionalmente",
        "se perdio una parte considerable",
        "solo improvisacion",
        "confuso;",
        "poco respetuoso",
        "poco claro",
    ]
    if any(k in n for k in criticos_2):
        return 2.0, True

    # Niveles regulares (3.0) -> Alerta de mejora
    regulares_3 = [
        "silencio prolongado",
        "regular;",
        "regular:",
        "frio o cortante",
        "a veces; responde",
        "solo se cubrio la mitad",
    ]
    if any(k in n for k in regulares_3):
        return 3.0, True

    # Niveles buenos (4.0)
    buenos_4 = [
        "bueno;",
        "bueno:",
        "buena calidad",
        "casi siempre",
        "casi todo el tiempo",
        "en su mayoria",
        "distracciones minimas",
        "estrictamente necesario",
    ]
    if any(k in n for k in buenos_4):
        return 4.0, False

    # Niveles excelentes (5.0)
    excelentes_5 = [
        "excelente",
        "muy alta calidad",
        "siempre;",
        "si, se aprovecho cada minuto",
        "totalmente, el contenido",
        "100% enfocado",
        "no utilizo el celular en absoluto",
        "si, me dio instrucciones",
    ]
    if any(k in n for k in excelentes_5):
        return 5.0, False

    return None, False


def _es_comentario_mejora_real(txt: Any) -> bool:
    s = str(txt or "").strip()
    if len(s) < 4:
        return False
    n = _norm(s)
    ruido = {
        "no", "n/a", "na", "ninguna", "ninguno", "no ninguna", "ningun", "nada",
        "todo bien", "todo excelente", "excelente", "muy bien", "muy bueno",
        "asi esta bien", "ninguna todo bien", "ninguna, todo bien", "10/10", "100",
        "ok", "bien", "super", "gracias", "todo perfecto", "no tengo", "sin novedad",
        "ninguna observacion", "no aplica",
    }
    if n in ruido or re.match(r"^[\W\d_]+$", n):
        return False
    if any(n.startswith(p) for p in ("ninguna", "todo bien", "todo excelente", "no hay", "no ninguna")):
        return False
    return True


def _total_clases_cat(categoria: str) -> int:
    cat_norm = clases_core.normalizar_categoria(categoria)
    cfg = clases_core.FORMATOS_OFICIALES.get(cat_norm) or clases_core.FORMATOS_OFICIALES.get("B1") or {}
    return int(cfg.get("total_clases") or 20)


def _obtener_objetivo_proxima_clase(categoria: str, clases_completadas: int) -> Dict[str, Any]:
    """Consulta FORMATOS_OFICIALES para indicar el tema oficial exacto que corresponde a la siguiente clase."""
    cat_norm = clases_core.normalizar_categoria(categoria)
    cfg = clases_core.FORMATOS_OFICIALES.get(cat_norm) or clases_core.FORMATOS_OFICIALES.get("B1") or {}
    filas = cfg.get("filas") or []
    total_req = _total_clases_cat(cat_norm)

    if clases_completadas >= total_req:
        return {
            "siguiente_num": total_req,
            "completado": True,
            "etapa": "finalizado",
            "tema": "✅ Curso práctico completado al 100% (listo para certificación / refuerzo final si aplica).",
            "codigo_formato": cfg.get("codigo_formato", ""),
        }

    siguiente = clases_completadas + 1
    # Para RECAT_C1 las clases oficiales van 21..30 (o 1..8 según conteo)
    for fila in filas:
        cl_list = fila.get("clases") or []
        if siguiente in cl_list:
            return {
                "siguiente_num": siguiente,
                "completado": False,
                "etapa": fila.get("etapa", "via"),
                "tema": str(fila.get("tema", "")).replace("\n", " "),
                "codigo_formato": cfg.get("codigo_formato", ""),
            }

    # Si el índice no coincidió (ej. RECAT_C1 con numeración 21..30), mapear por posición
    if filas:
        pos = min(len(filas) - 1, max(0, clases_completadas))
        f_pos = filas[pos]
        return {
            "siguiente_num": siguiente,
            "completado": False,
            "etapa": f_pos.get("etapa", "via"),
            "tema": str(f_pos.get("tema", "")).replace("\n", " "),
            "codigo_formato": cfg.get("codigo_formato", ""),
        }

    return {
        "siguiente_num": siguiente,
        "completado": False,
        "etapa": "via",
        "tema": f"Clase práctica #{siguiente} — Categoría {cat_norm}",
        "codigo_formato": "",
    }


class EvaluacionPedagogicaService:
    def __init__(self) -> None:
        self.sheet_id = EVAL_SHEET_ID_DEFAULT
        self.form_url = EVAL_FORM_URL_DEFAULT
        self._lock = threading.RLock()
        self._spreadsheet: Optional[Any] = None
        self._cache_eval: Optional[Tuple[float, List[List[str]]]] = None
        self._ultimo_fallo_ts: float = 0.0
        self._ultimo_error_msg: str = ""
        self._obs_cache: Optional[Tuple[float, List[Dict[str, Any]]]] = None

    # -------------------------------------------------------------------------
    # A. LECTURA Y ANALÍTICA DE EVALUACIONES DOCENTES (GOOGLE FORM EN VIVO)
    # -------------------------------------------------------------------------
    def _leer_filas_evaluacion(self, forzar: bool = False) -> Tuple[List[List[str]], bool]:
        ahora = time.time()
        with self._lock:
            if not forzar and self._cache_eval and (ahora - self._cache_eval[0] < 120.0):
                return [list(r) for r in self._cache_eval[1]], (self._spreadsheet is not None)

        try:
            with self._lock:
                if self._spreadsheet is None:
                    if not forzar and self._ultimo_fallo_ts > 0 and (ahora - self._ultimo_fallo_ts < 25.0):
                        raise ConnectionError(self._ultimo_error_msg or "Reintento en espera.")
                    creds = clases_core.obtener_credenciales_google(clases_core.SCOPES)
                    client = clases_core.gspread.authorize(creds)
                    self._spreadsheet = client.open_by_key(self.sheet_id)
                ws = self._spreadsheet.worksheets()[0]
            vals = ws.get_all_values()
            if vals:
                with self._lock:
                    self._cache_eval = (time.time(), vals)
                    self._ultimo_fallo_ts = 0.0
                    self._ultimo_error_msg = ""
                try:
                    EVAL_CACHE_LOCAL_FILE.write_text(
                        json.dumps({"ts": time.time(), "rows": vals}, ensure_ascii=False),
                        encoding="utf-8",
                    )
                except Exception:
                    pass
                return [list(r) for r in vals], True
        except Exception as exc:
            with self._lock:
                self._ultimo_fallo_ts = time.time()
                self._ultimo_error_msg = str(exc)
                if self._cache_eval:
                    return [list(r) for r in self._cache_eval[1]], False

        # Respaldo local si está offline
        if EVAL_CACHE_LOCAL_FILE.exists():
            try:
                saved = json.loads(EVAL_CACHE_LOCAL_FILE.read_text(encoding="utf-8"))
                rows = saved.get("rows") or []
                if rows:
                    with self._lock:
                        self._cache_eval = (time.time(), rows)
                    return [list(r) for r in rows], False
            except Exception:
                pass
        return [], False

    def obtener_evaluaciones_docentes(self, forzar: bool = False) -> Dict[str, Any]:
        filas, en_vivo = self._leer_filas_evaluacion(forzar=forzar)
        if not filas or len(filas) < 2:
            return {
                "ok": True,
                "conectado_en_vivo": en_vivo,
                "form_url": self.form_url,
                "sheet_url": f"https://docs.google.com/spreadsheets/d/{self.sheet_id}/edit",
                "kpis": {
                    "total_respuestas": 0,
                    "promedio_global": 0.0,
                    "pct_excelencia": 0,
                    "total_alertas": 0,
                    "total_instructores": 0,
                },
                "ranking_instructores": [],
                "respuestas": [],
                "alertas_recientes": [],
                "error_conexion": self._ultimo_error_msg,
            }

        # Mapeo de criterios evaluados (evitando duplicar col 5 y col 21)
        criterios_cols = [
            ("trato", "Trato y Actitud", [16]),
            ("tiempo", "Cumplimiento Tiempo", [17, 10]),
            ("objetivos", "Objetivos del Día", [18]),
            ("atencion", "Atención en Clase", [19]),
            ("celular", "Uso de Celular", [20]),
            ("ensenanza", "Enseñanza Constante", [5, 21]),
            ("claridad", "Claridad / Explicación", [7, 6]),
            ("metodologia", "Orden y Metodología", [22]),
            ("disposicion", "Disposición a Dudas", [11]),
            ("apoyo", "Material / Apoyo", [9, 8, 12]),
        ]

        respuestas: List[Dict[str, Any]] = []
        por_prof: Dict[str, Dict[str, Any]] = {}

        for idx_f, r in enumerate(filas[1:], start=2):
            ts_raw = str(r[0]).strip() if len(r) > 0 else ""
            fecha_raw = str(r[2]).strip() if len(r) > 2 else ""
            alumno = str(r[3]).strip() if len(r) > 3 else ""
            prof_raw = str(r[4]).strip() if len(r) > 4 else ""
            if not ts_raw and not alumno and not prof_raw:
                continue

            prof = _normalizar_nombre_profesor(prof_raw)
            fecha_iso = _parsear_fecha_eval(fecha_raw, ts_raw)
            comentario = str(r[15]).strip() if len(r) > 15 else ""
            habeas = str(r[14]).strip() if len(r) > 14 else ""

            puntajes_crit: Dict[str, float] = {}
            alertas_detalle: List[str] = []

            for c_key, c_lbl, col_indices in criterios_cols:
                vals_col = []
                for ci in col_indices:
                    if len(r) > ci and str(r[ci]).strip():
                        sc, is_al = _puntuar_respuesta_criterio(r[ci])
                        if sc is not None:
                            vals_col.append((sc, is_al, str(r[ci]).strip()))
                if vals_col:
                    # Si había col 5 y col 21, tomar el menor (más conservador)
                    sc_min, al_any, txt_min = min(vals_col, key=lambda x: x[0])
                    puntajes_crit[c_key] = sc_min
                    if al_any:
                        alertas_detalle.append(f"{c_lbl}: {txt_min[:70]}")

            promedio_resp = round(sum(puntajes_crit.values()) / len(puntajes_crit), 2) if puntajes_crit else 5.0
            tiene_coment_mejora = _es_comentario_mejora_real(comentario)
            tiene_alerta = bool(alertas_detalle) or promedio_resp < 4.0 or tiene_coment_mejora

            item_resp = {
                "fila": idx_f,
                "timestamp": ts_raw,
                "fecha": fecha_iso,
                "alumno": alumno or "Anónimo",
                "instructor": prof,
                "instructor_raw": prof_raw,
                "promedio": promedio_resp,
                "criterios": puntajes_crit,
                "comentario": comentario if _es_comentario_mejora_real(comentario) else "",
                "comentario_raw": comentario,
                "tiene_alerta": tiene_alerta,
                "alertas_detalle": alertas_detalle,
                "habeas_data": habeas or "Sí",
            }
            respuestas.append(item_resp)

            if prof != "SIN ESPECIFICAR":
                agg = por_prof.setdefault(prof, {
                    "instructor": prof,
                    "total": 0,
                    "suma_prom": 0.0,
                    "excelentes": 0,
                    "alertas": 0,
                    "ultima_fecha": "",
                    "crit_sum": {},
                    "crit_cnt": {},
                })
                agg["total"] += 1
                agg["suma_prom"] += promedio_resp
                if promedio_resp >= 4.5:
                    agg["excelentes"] += 1
                if tiene_alerta:
                    agg["alertas"] += 1
                if fecha_iso and fecha_iso > agg["ultima_fecha"]:
                    agg["ultima_fecha"] = fecha_iso
                for ck, cv in puntajes_crit.items():
                    agg["crit_sum"][ck] = agg["crit_sum"].get(ck, 0.0) + cv
                    agg["crit_cnt"][ck] = agg["crit_cnt"].get(ck, 0) + 1

        ranking: List[Dict[str, Any]] = []
        for prof, agg in por_prof.items():
            tot = max(1, agg["total"])
            prom = round(agg["suma_prom"] / tot, 2)
            desglose = {
                ck: round(agg["crit_sum"][ck] / max(1, agg["crit_cnt"][ck]), 2)
                for ck in agg["crit_sum"]
            }
            ranking.append({
                "instructor": prof,
                "total_evaluaciones": agg["total"],
                "promedio": prom,
                "pct_excelencia": int(round((agg["excelentes"] / tot) * 100)),
                "alertas": agg["alertas"],
                "ultima_fecha": agg["ultima_fecha"],
                "criterios": desglose,
            })

        ranking.sort(key=lambda x: (0 if x["total_evaluaciones"] >= 5 else 1, -x["promedio"], -x["total_evaluaciones"], x["instructor"]))
        respuestas_ord = list(reversed(respuestas))

        total_r = len(respuestas)
        prom_global = round(sum(r["promedio"] for r in respuestas) / total_r, 2) if total_r else 0.0
        exc_global = sum(1 for r in respuestas if r["promedio"] >= 4.5)
        alertas_list = [r for r in respuestas_ord if r["tiene_alerta"]]

        return {
            "ok": True,
            "conectado_en_vivo": en_vivo,
            "form_url": self.form_url,
            "sheet_url": f"https://docs.google.com/spreadsheets/d/{self.sheet_id}/edit",
            "kpis": {
                "total_respuestas": total_r,
                "promedio_global": prom_global,
                "pct_excelencia": int(round((exc_global / max(1, total_r)) * 100)),
                "total_alertas": len(alertas_list),
                "total_instructores": len(ranking),
            },
            "ranking_instructores": ranking,
            "respuestas": respuestas_ord[:350],
            "alertas_recientes": alertas_list[:60],
            "error_conexion": self._ultimo_error_msg if not en_vivo else "",
        }

    # -------------------------------------------------------------------------
    # B. OBSERVACIONES PEDAGÓGICAS DE INSTRUCTORES (CERO ESCRITURA)
    # -------------------------------------------------------------------------
    def _cargar_obs_locales(self) -> List[Dict[str, Any]]:
        if not OBS_LOCAL_FILE.exists():
            return []
        try:
            data = json.loads(OBS_LOCAL_FILE.read_text(encoding="utf-8"))
            return data if isinstance(data, list) else []
        except Exception:
            return []

    def _guardar_obs_locales(self, items: List[Dict[str, Any]]) -> None:
        try:
            OBS_LOCAL_FILE.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception as exc:
            logger.warning("Aviso guardando observaciones locales JSON: %s", exc)

    def listar_observaciones(self, forzar: bool = False) -> List[Dict[str, Any]]:
        ahora = time.time()
        with self._lock:
            if not forzar and self._obs_cache and (ahora - self._obs_cache[0] < 90.0):
                return [dict(x) for x in self._obs_cache[1]]

        items: List[Dict[str, Any]] = []
        leido_sheets = False
        try:
            ws = clases_core.sheets_service.obtener_o_crear_hoja(SHEET_OBSERVACIONES, HEADERS_OBSERVACIONES)
            if forzar and hasattr(ws, "_invalidate"):
                ws._invalidate()
            rows = ws.get_all_values()
            if rows and len(rows) >= 2:
                leido_sheets = True
                for idx_f, r in enumerate(rows[1:], start=2):
                    if not any(str(c).strip() for c in r):
                        continue
                    recs_raw = str(r[9]).strip() if len(r) > 9 else ""
                    fort_raw = str(r[10]).strip() if len(r) > 10 else ""
                    recs = [x.strip() for x in recs_raw.split("|") if x.strip()]
                    forts = [x.strip() for x in fort_raw.split("|") if x.strip()]
                    items.append({
                        "fila": idx_f,
                        "id_obs": str(r[0]).strip() if len(r) > 0 else f"OBS-{idx_f}",
                        "timestamp": str(r[1]).strip() if len(r) > 1 else "",
                        "fecha": str(r[2]).strip() if len(r) > 2 else "",
                        "documento": str(r[3]).strip() if len(r) > 3 else "",
                        "nombre_alumno": str(r[4]).strip() if len(r) > 4 else "",
                        "categoria": clases_core.normalizar_categoria(str(r[5]).strip() if len(r) > 5 else "B1"),
                        "clase_num": str(r[6]).strip() if len(r) > 6 else "",
                        "instructor": str(r[7]).strip() if len(r) > 7 else "",
                        "nivel_avance": str(r[8]).strip() if len(r) > 8 else "AVANZA_BIEN",
                        "recomendaciones": recs,
                        "fortalezas": forts,
                        "estado_refuerzo": (str(r[11]).strip().upper() if len(r) > 11 else "") or ("PENDIENTE" if recs and not all("Sin pendientes" in x for x in recs) else "SUPERADO"),
                    })
                self._guardar_obs_locales(items)
        except Exception as exc:
            logger.warning("Aviso leyendo pestaña Observaciones_Alumnos: %s", exc)

        if not leido_sheets:
            items = self._cargar_obs_locales()

        items_ord = list(reversed(items))
        with self._lock:
            self._obs_cache = (time.time(), items_ord)
        return [dict(x) for x in items_ord]

    def guardar_observacion(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        nombre = str(payload.get("nombre_alumno", "")).strip().upper()
        documento = re.sub(r"[^\dA-Za-z]", "", str(payload.get("documento", "")).strip())
        if not nombre and not documento:
            return {"ok": False, "message": "Selecciona un alumno para registrar la observación pedagógica."}

        categoria = clases_core.normalizar_categoria(str(payload.get("categoria", "B1")))
        clase_num = str(payload.get("clase_num", "")).strip()
        instructor = str(payload.get("instructor", "")).strip().upper() or "INSTRUCTOR CEA"
        nivel = str(payload.get("nivel_avance", "AVANZA_BIEN")).strip().upper()
        if nivel not in ("AVANZA_BIEN", "REFORZAR", "PRIORITARIO"):
            nivel = "AVANZA_BIEN"

        recs_in = payload.get("recomendaciones") or []
        if isinstance(recs_in, str):
            recs_in = [x.strip() for x in recs_in.split("|") if x.strip()]
        recs = [str(x).strip() for x in recs_in if str(x).strip()]

        fort_in = payload.get("fortalezas") or []
        if isinstance(fort_in, str):
            fort_in = [x.strip() for x in fort_in.split("|") if x.strip()]
        forts = [str(x).strip() for x in fort_in if str(x).strip()]

        if not recs:
            recs = ["✅ Sin pendientes (Listo para siguiente tema)"]

        solo_sin_pendientes = all("Sin pendientes" in x for x in recs)
        estado_refuerzo = "SUPERADO" if (solo_sin_pendientes and nivel == "AVANZA_BIEN") else "PENDIENTE"

        ahora = datetime.datetime.now(clases_core.TZ_COLOMBIA)
        ts_str = ahora.strftime("%Y-%m-%d %H:%M:%S")
        fecha_str = str(payload.get("fecha", "")).strip() or ahora.strftime("%Y-%m-%d")
        id_obs = f"OBS-{ahora.strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:4].upper()}"

        fila_vals = [
            id_obs,
            ts_str,
            fecha_str,
            documento,
            nombre,
            categoria,
            clase_num,
            instructor,
            nivel,
            " | ".join(recs),
            " | ".join(forts),
            estado_refuerzo,
        ]

        guardado_sheets = False
        try:
            ws = clases_core.sheets_service.obtener_o_crear_hoja(SHEET_OBSERVACIONES, HEADERS_OBSERVACIONES)
            ws.append_row(fila_vals, value_input_option="USER_ENTERED")
            guardado_sheets = True
        except Exception as exc:
            logger.warning("Aviso guardando observación en Google Sheets: %s", exc)

        nuevo_item = {
            "fila": 0,
            "id_obs": id_obs,
            "timestamp": ts_str,
            "fecha": fecha_str,
            "documento": documento,
            "nombre_alumno": nombre,
            "categoria": categoria,
            "clase_num": clase_num,
            "instructor": instructor,
            "nivel_avance": nivel,
            "recomendaciones": recs,
            "fortalezas": forts,
            "estado_refuerzo": estado_refuerzo,
        }

        locales = self._cargar_obs_locales()
        locales.append(nuevo_item)
        self._guardar_obs_locales(locales)

        with self._lock:
            self._obs_cache = None

        return {
            "ok": True,
            "guardado_google_sheets": guardado_sheets,
            "observacion": nuevo_item,
            "message": (
                f"✅ Observación pedagógica guardada para {nombre or documento} ({categoria})"
                + (" en Google Sheets." if guardado_sheets else " en respaldo local.")
            ),
        }

    def marcar_refuerzo_superado(self, id_obs: str = "", documento: str = "", nombre_alumno: str = "", categoria: str = "") -> Dict[str, Any]:
        """Marca como SUPERADO una observación puntual (por id_obs) o todos los refuerzos pendientes de un alumno."""
        id_obj = str(id_obs or "").strip()
        doc_n = _norm(documento)
        nom_n = _norm(nombre_alumno)
        cat_n = clases_core.normalizar_categoria(categoria) if categoria else ""

        actualizados = 0
        try:
            ws = clases_core.sheets_service.obtener_o_crear_hoja(SHEET_OBSERVACIONES, HEADERS_OBSERVACIONES)
            rows = ws.get_all_values()
            for idx_f in range(2, len(rows) + 1):
                r = rows[idx_f - 1]
                r_id = str(r[0]).strip() if len(r) > 0 else ""
                r_doc = _norm(r[3] if len(r) > 3 else "")
                r_nom = _norm(r[4] if len(r) > 4 else "")
                r_cat = clases_core.normalizar_categoria(str(r[5]).strip() if len(r) > 5 else "B1")
                r_est = str(r[11]).strip().upper() if len(r) > 11 else "PENDIENTE"

                coincide = False
                if id_obj and r_id == id_obj:
                    coincide = True
                elif not id_obj and r_est == "PENDIENTE":
                    if (not cat_n or r_cat == cat_n) and ((doc_n and r_doc == doc_n) or (nom_n and r_nom == nom_n)):
                        coincide = True

                if coincide and r_est != "SUPERADO":
                    ws.update_cell(idx_f, 12, "SUPERADO")
                    actualizados += 1
        except Exception as exc:
            logger.warning("Aviso actualizando Estado_Refuerzo en Google Sheets: %s", exc)

        # Actualizar también respaldo local
        locales = self._cargar_obs_locales()
        for it in locales:
            r_id = str(it.get("id_obs", "")).strip()
            r_doc = _norm(it.get("documento", ""))
            r_nom = _norm(it.get("nombre_alumno", ""))
            r_cat = clases_core.normalizar_categoria(it.get("categoria", "B1"))
            if (id_obj and r_id == id_obj) or (
                not id_obj
                and (not cat_n or r_cat == cat_n)
                and ((doc_n and r_doc == doc_n) or (nom_n and r_nom == nom_n))
            ):
                if it.get("estado_refuerzo") != "SUPERADO":
                    it["estado_refuerzo"] = "SUPERADO"
                    if actualizados == 0:
                        actualizados += 1
        self._guardar_obs_locales(locales)

        with self._lock:
            self._obs_cache = None

        return {
            "ok": True,
            "actualizados": actualizados,
            "message": f"✅ {actualizados} refuerzo(s) marcado(s) como superado(s).",
        }

    # -------------------------------------------------------------------------
    # C. DIRECTORIO DE INSTRUCTORES (ACTIVOS / RETIRADOS / NUEVOS)
    # -------------------------------------------------------------------------
    def _cargar_estados_instructores(self, forzar: bool = False) -> Dict[str, Dict[str, str]]:
        """Lee overrides de estado (ACTIVO/RETIRADO) y nuevos instructores desde Sheets + respaldo JSON."""
        ahora = time.time()
        with self._lock:
            cache_est = getattr(self, "_est_inst_cache", None)
            if not forzar and cache_est and (ahora - cache_est[0] < 90.0):
                return {k: dict(v) for k, v in cache_est[1].items()}

        mapa: Dict[str, Dict[str, str]] = {}
        if INST_ESTADO_LOCAL_FILE.exists():
            try:
                raw_loc = json.loads(INST_ESTADO_LOCAL_FILE.read_text(encoding="utf-8"))
                if isinstance(raw_loc, dict):
                    for k, v in raw_loc.items():
                        if isinstance(v, dict) and v.get("instructor"):
                            canon = _normalizar_nombre_profesor(v["instructor"])
                            mapa[canon] = {
                                "instructor": canon,
                                "estado": str(v.get("estado", "ACTIVO")).strip().upper(),
                                "tipo": str(v.get("tipo", "PRÁCTICO / TEÓRICO")).strip().upper(),
                                "fecha_cambio": str(v.get("fecha_cambio", "")).strip(),
                            }
            except Exception:
                pass

        try:
            ws = clases_core.sheets_service.obtener_o_crear_hoja(SHEET_ESTADO_INSTRUCTORES, HEADERS_ESTADO_INSTRUCTORES)
            if forzar and hasattr(ws, "_invalidate"):
                ws._invalidate()
            rows = ws.get_all_values()
            if rows and len(rows) >= 2:
                for r in rows[1:]:
                    nom_r = str(r[0]).strip() if len(r) > 0 else ""
                    if not nom_r:
                        continue
                    canon = _normalizar_nombre_profesor(nom_r)
                    est = str(r[1]).strip().upper() if len(r) > 1 and str(r[1]).strip() else "ACTIVO"
                    tipo = str(r[2]).strip().upper() if len(r) > 2 and str(r[2]).strip() else "PRÁCTICO / TEÓRICO"
                    f_chg = str(r[3]).strip() if len(r) > 3 else ""
                    mapa[canon] = {
                        "instructor": canon,
                        "estado": "RETIRADO" if "RETIR" in est else "ACTIVO",
                        "tipo": tipo,
                        "fecha_cambio": f_chg,
                    }
                try:
                    INST_ESTADO_LOCAL_FILE.write_text(json.dumps(mapa, ensure_ascii=False, indent=2), encoding="utf-8")
                except Exception:
                    pass
        except Exception as exc:
            logger.debug("Aviso leyendo Estado_Instructores_CEA: %s", exc)

        with self._lock:
            self._est_inst_cache = (time.time(), mapa)
        return {k: dict(v) for k, v in mapa.items()}

    def guardar_o_cambiar_estado_instructor(
        self,
        nombre: str,
        estado: str = "ACTIVO",
        tipo: str = "PRÁCTICO / TEÓRICO",
    ) -> Dict[str, Any]:
        """Agrega un instructor nuevo o cambia su estado entre ACTIVO y RETIRADO (sin borrar su histórico)."""
        nom_in = str(nombre or "").strip()
        if not nom_in:
            return {"ok": False, "message": "Ingresa el nombre del instructor."}

        canon = _normalizar_nombre_profesor(nom_in)
        if canon == "SIN ESPECIFICAR":
            canon = nom_in.upper()
        est_norm = "RETIRADO" if "RETIR" in str(estado or "").upper() else "ACTIVO"
        tipo_norm = str(tipo or "PRÁCTICO / TEÓRICO").strip().upper() or "PRÁCTICO / TEÓRICO"
        fecha_hoy = datetime.datetime.now(clases_core.TZ_COLOMBIA).strftime("%Y-%m-%d")

        guardado_sheets = False
        try:
            ws = clases_core.sheets_service.obtener_o_crear_hoja(SHEET_ESTADO_INSTRUCTORES, HEADERS_ESTADO_INSTRUCTORES)
            rows = ws.get_all_values()
            fila_encontrada = 0
            for idx_f in range(2, len(rows) + 1):
                r = rows[idx_f - 1]
                r_nom = _normalizar_nombre_profesor(r[0] if len(r) > 0 else "")
                if r_nom == canon:
                    fila_encontrada = idx_f
                    break
            nueva_fila = [canon, est_norm, tipo_norm, fecha_hoy]
            if fila_encontrada >= 2:
                ws.update(values=[nueva_fila], range_name=f"A{fila_encontrada}:D{fila_encontrada}", value_input_option="USER_ENTERED")
            else:
                ws.append_row(nueva_fila, value_input_option="USER_ENTERED")
            guardado_sheets = True
        except Exception as exc:
            logger.warning("Aviso guardando Estado_Instructores_CEA en Sheets: %s", exc)

        mapa = self._cargar_estados_instructores(forzar=False)
        mapa[canon] = {
            "instructor": canon,
            "estado": est_norm,
            "tipo": tipo_norm,
            "fecha_cambio": fecha_hoy,
        }
        try:
            INST_ESTADO_LOCAL_FILE.write_text(json.dumps(mapa, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass
        with self._lock:
            self._est_inst_cache = (time.time(), mapa)

        ico = "🟢 ACTIVO" if est_norm == "ACTIVO" else "🚫 RETIRADO"
        return {
            "ok": True,
            "instructor": canon,
            "estado": est_norm,
            "guardado_google_sheets": guardado_sheets,
            "message": f"✅ Instructor {canon} actualizado a {ico}" + (" en Google Sheets." if guardado_sheets else " en respaldo local."),
        }

    def _construir_directorio_instructores(self, eval_data: Dict[str, Any], forzar: bool = False) -> Tuple[List[Dict[str, Any]], List[str]]:
        """
        Consolida instructores del Bot + Google Form + nuevos agregados manualmente,
        marcando estado ACTIVO / RETIRADO (por defecto activos los del Bot y evaluados en 2026+,
        retirados los históricos anteriores a 2026 que no están en el Bot, siempre editables en 1 clic).
        """
        overrides = self._cargar_estados_instructores(forzar=forzar)
        en_bot_set: set = set()
        try:
            for inst_b in clases_core.sheets_service.listar_instructores_registrados():
                nom_b = _normalizar_nombre_profesor(inst_b.get("nombre", ""))
                if nom_b and nom_b != "SIN ESPECIFICAR":
                    en_bot_set.add(nom_b)
        except Exception:
            pass

        eval_map: Dict[str, Dict[str, Any]] = {}
        for r in (eval_data.get("ranking_instructores") or []):
            eval_map[r["instructor"]] = r

        todos_noms = set(en_bot_set) | set(eval_map.keys()) | set(overrides.keys())
        directorio: List[Dict[str, Any]] = []
        for nom in sorted(todos_noms):
            if not nom or nom == "SIN ESPECIFICAR":
                continue
            ev_r = eval_map.get(nom)
            ov = overrides.get(nom)
            en_bot = nom in en_bot_set
            ult_fecha = (ev_r or {}).get("ultima_fecha", "")
            anio_ult = int(ult_fecha[:4]) if len(ult_fecha) >= 4 and ult_fecha[:4].isdigit() else 0

            if ov:
                estado = ov["estado"]
                tipo = ov.get("tipo") or "PRÁCTICO / TEÓRICO"
            else:
                estado = "ACTIVO" if (en_bot or anio_ult >= 2026) else "RETIRADO"
                tipo = "PRÁCTICO / BOT" if en_bot else "TEÓRICO / FORM"

            if ev_r is not None:
                ev_r["estado"] = estado
                ev_r["tipo"] = tipo

            directorio.append({
                "instructor": nom,
                "estado": estado,
                "tipo": tipo,
                "en_bot": en_bot,
                "total_evaluaciones": int((ev_r or {}).get("total_evaluaciones") or 0),
                "promedio": float((ev_r or {}).get("promedio") or 0.0),
                "alertas": int((ev_r or {}).get("alertas") or 0),
                "ultima_fecha": ult_fecha,
            })

        directorio.sort(key=lambda x: (0 if x["estado"] == "ACTIVO" else 1, -x["total_evaluaciones"], x["instructor"]))
        activos = sorted([d["instructor"] for d in directorio if d["estado"] == "ACTIVO"])
        return directorio, activos

    def obtener_panorama_pedagogico(self, forzar: bool = False) -> Dict[str, Any]:
        """
        Consolida para el Dashboard Web:
        1. Analítica de Evaluación Docente del Google Form en vivo (ranking, alertas, link).
        2. Catálogos de selección rápida por botones (cero escritura).
        3. Progreso pedagógico de alumnos activos + tema oficial de su próxima clase + refuerzos pendientes.
        """
        eval_data = self.obtener_evaluaciones_docentes(forzar=forzar)
        directorio_inst, instructores_activos = self._construir_directorio_instructores(eval_data, forzar=forzar)
        observaciones = self.listar_observaciones(forzar=forzar)

        # Agrupar observaciones por (documento, categoria) y (nombre, categoria)
        obs_por_doc: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
        obs_por_nom: Dict[Tuple[str, str], List[Dict[str, Any]]] = {}
        for o in observaciones:
            c_o = clases_core.normalizar_categoria(o.get("categoria", "B1"))
            d_o = _norm(o.get("documento", ""))
            n_o = _norm(o.get("nombre_alumno", ""))
            if d_o:
                obs_por_doc.setdefault((d_o, c_o), []).append(o)
            if n_o:
                obs_por_nom.setdefault((n_o, c_o), []).append(o)

        alumnos_progreso: List[Dict[str, Any]] = []
        vistos = set()
        try:
            alumnos_bot = clases_core.sheets_service.listar_alumnos_activos("TODAS", modo="todos")
        except Exception:
            alumnos_bot = []

        for a in alumnos_bot:
            nom = str(a.get("Nombre_Completo", "")).strip()
            doc = str(a.get("Documento", "")).strip()
            cat = clases_core.normalizar_categoria(a.get("Categoria", "B1"))
            clave = (_norm(doc) or _norm(nom), cat)
            if not clave[0] or clave in vistos:
                continue
            vistos.add(clave)

            comp = int(a.get("Clases_Completadas") or 0)
            req = int(a.get("Total_Requerido") or _total_clases_cat(cat))
            pct = min(100, int(round((comp / max(1, req)) * 100)))
            est_curso = str(a.get("Estado_Curso", "EN CURSO")).strip()
            obj_sig = _obtener_objetivo_proxima_clase(cat, comp)

            hist_obs = obs_por_doc.get((_norm(doc), cat), []) or obs_por_nom.get((_norm(nom), cat), [])
            # Recomendaciones activas (solo de observaciones PENDIENTE)
            conteo_recs: Dict[str, int] = {}
            ultimo_nivel = ""
            ultimo_inst = ""
            ultima_fecha_obs = ""
            if hist_obs:
                ultimo_nivel = hist_obs[0].get("nivel_avance", "")
                ultimo_inst = hist_obs[0].get("instructor", "")
                ultima_fecha_obs = hist_obs[0].get("fecha", "")
            for ob in hist_obs:
                if ob.get("estado_refuerzo") == "PENDIENTE":
                    for rec in ob.get("recomendaciones") or []:
                        if "Sin pendientes" in rec:
                            continue
                        conteo_recs[rec] = conteo_recs.get(rec, 0) + 1

            recs_pendientes = [
                {"tema": k, "veces": v}
                for k, v in sorted(conteo_recs.items(), key=lambda kv: (-kv[1], kv[0]))
            ]
            alumnos_progreso.append({
                "documento": doc,
                "nombre": nom,
                "categoria": cat,
                "es_moto": cat in ("A1", "A2"),
                "clases_completadas": comp,
                "total_requerido": req,
                "porcentaje": pct,
                "estado_curso": est_curso,
                "siguiente_clase": obj_sig,
                "recomendaciones_pendientes": recs_pendientes,
                "tiene_refuerzo_pendiente": len(recs_pendientes) > 0,
                "ultimo_nivel": ultimo_nivel,
                "ultimo_instructor_obs": ultimo_inst,
                "ultima_fecha_obs": ultima_fecha_obs,
                "total_observaciones": len(hist_obs),
                "historial_observaciones": hist_obs[:15],
            })

        # Incluir también alumnos que tengan observaciones registradas aunque aún no estén en listar_alumnos_activos
        for o in observaciones:
            cat = clases_core.normalizar_categoria(o.get("categoria", "B1"))
            doc = str(o.get("documento", "")).strip()
            nom = str(o.get("nombre_alumno", "")).strip()
            clave = (_norm(doc) or _norm(nom), cat)
            if not clave[0] or clave in vistos:
                continue
            vistos.add(clave)
            hist_obs = obs_por_doc.get((_norm(doc), cat), []) or obs_por_nom.get((_norm(nom), cat), [])
            conteo_recs = {}
            for ob in hist_obs:
                if ob.get("estado_refuerzo") == "PENDIENTE":
                    for rec in ob.get("recomendaciones") or []:
                        if "Sin pendientes" not in rec:
                            conteo_recs[rec] = conteo_recs.get(rec, 0) + 1
            recs_pendientes = [{"tema": k, "veces": v} for k, v in sorted(conteo_recs.items(), key=lambda kv: (-kv[1], kv[0]))]
            req = _total_clases_cat(cat)
            alumnos_progreso.append({
                "documento": doc,
                "nombre": nom,
                "categoria": cat,
                "es_moto": cat in ("A1", "A2"),
                "clases_completadas": 0,
                "total_requerido": req,
                "porcentaje": 0,
                "estado_curso": "EN CURSO",
                "siguiente_clase": _obtener_objetivo_proxima_clase(cat, 0),
                "recomendaciones_pendientes": recs_pendientes,
                "tiene_refuerzo_pendiente": len(recs_pendientes) > 0,
                "ultimo_nivel": hist_obs[0].get("nivel_avance", "") if hist_obs else "",
                "ultimo_instructor_obs": hist_obs[0].get("instructor", "") if hist_obs else "",
                "ultima_fecha_obs": hist_obs[0].get("fecha", "") if hist_obs else "",
                "total_observaciones": len(hist_obs),
                "historial_observaciones": hist_obs[:15],
            })

        # Priorizar alumnos con refuerzo pendiente y en curso
        alumnos_progreso.sort(
            key=lambda x: (
                0 if x["tiene_refuerzo_pendiente"] else 1,
                0 if x["estado_curso"] != "FINALIZADO" else 1,
                -x["clases_completadas"],
                x["nombre"],
            )
        )

        return {
            "ok": True,
            "evaluacion_docente": eval_data,
            "directorio_instructores": directorio_inst,
            "instructores_activos": instructores_activos,
            "catalogos": {
                "niveles": CATALOGO_NIVELES,
                "recomendaciones_carro": CATALOGO_RECOMENDACIONES_CARRO,
                "recomendaciones_moto": CATALOGO_RECOMENDACIONES_MOTO,
                "fortalezas": CATALOGO_FORTALEZAS,
            },
            "kpis_pedagogicos": {
                "total_alumnos_seguimiento": len(alumnos_progreso),
                "alumnos_con_refuerzo": sum(1 for a in alumnos_progreso if a["tiene_refuerzo_pendiente"]),
                "total_observaciones": len(observaciones),
                "refuerzos_superados": sum(1 for o in observaciones if o.get("estado_refuerzo") == "SUPERADO"),
            },
            "alumnos_progreso": alumnos_progreso,
            "observaciones_recientes": observaciones[:120],
        }


evaluacion_service = EvaluacionPedagogicaService()
