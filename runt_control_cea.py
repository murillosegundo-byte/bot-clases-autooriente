"""
Servicio de Integración en Tiempo Real con la Hoja de Cálculo de Control Digital RUNT
ID: 1XHbqfDa7dK5CWcgWKNgJZiXN_OSrjJCEJJu8F4z-u2M
Pestañas gestionadas:
1. Practica (Control de Teoría, Taller, Práctica, Clases Cargadas, Vigencia 90 Días y Estado Cargue RUNT)
2. No runt (Alumnos pendientes de inscripción / registro en RUNT)
3. Inicio Practica (Horarios de práctica de alumnos y novedades)
4. Codigo de Runt (Códigos RUNT asignados)
5. Codigo toma de foro (Códigos SISEC / Toma de Foto)
"""

import datetime
import logging
import os
import re
import time
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import clases_core

logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
RUNT_SHEET_ID_DEFAULT = "1XHbqfDa7dK5CWcgWKNgJZiXN_OSrjJCEJJu8F4z-u2M"
RUNT_LOCAL_XLSX = BASE_DIR / "control_runt_digital.xlsx"

TAB_PRACTICA = "Practica"
TAB_NO_RUNT = "No runt"
TAB_INICIO_PRACTICA = "Inicio Practica"
TAB_CODIGO_RUNT = "Codigo de Runt"
TAB_CODIGO_FOTO = "Codigo toma de foro "

HEADERS_PRACTICA = [
    "Nombre y apellido",
    "Cedula",
    "CATEGORIA",
    "Fecha Matricula",
    "Prioridad",
    "Vence 90 dias",
    "Fecha registro Runt",
    "Vigente / Vencido",
    "Teoria",
    "Taller",
    "Practica",
    "Clases Cargadas",
    "INSTRUCTOR",
    "HORA REGISTRO",
    "Actividad Estado",
    "Actividad Texto",
    "Fecha cargue",
    "Cargue Diario",
]


def _norm(txt: Any) -> str:
    s = str(txt or "").strip().lower()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s)


def _limpiar_cedula(val: Any) -> str:
    s = str(val or "").strip()
    if not s:
        return ""
    # Handle scientific notation like '1.094278698E9' or float '1094278698.0'
    if re.match(r"^\d+(\.\d+)?[eE][+-]?\d+$", s):
        try:
            return str(int(round(float(s))))
        except Exception:
            pass
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    digits = re.sub(r"[^\d]", "", s)
    return digits if len(digits) >= 4 else s


def _limpiar_entero(val: Any, default: int = 0) -> int:
    s = str(val or "").strip()
    if not s:
        return default
    try:
        return int(round(float(s.replace(",", "."))))
    except Exception:
        digits = re.sub(r"[^\d]", "", s)
        return int(digits) if digits else default


def _excel_serial_a_iso(val: Any) -> str:
    s = str(val or "").strip()
    if not s:
        return ""
    # Check if already YYYY-MM-DD
    m_iso = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        if y < 1900 and y > 200:
            y = 2026
        return f"{y:04d}-{m:02d}-{d:02d}"
    # Check DD/MM/YYYY
    m_dmy = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2,4})", s)
    if m_dmy:
        d, m, y = int(m_dmy.group(1)), int(m_dmy.group(2)), int(m_dmy.group(3))
        if y < 100:
            y += 2000
        elif y < 1900:
            y = 2026
        return f"{y:04d}-{m:02d}-{d:02d}"
    # Check Excel serial number (e.g. 46268.0)
    if re.match(r"^\d{5}(\.\d+)?$", s):
        try:
            days = int(float(s))
            if 35000 <= days <= 60000:
                base = datetime.date(1899, 12, 30)
                return (base + datetime.timedelta(days=days)).strftime("%Y-%m-%d")
        except Exception:
            pass
    return s


def _normalizar_cat_runt(cat_raw: Any) -> str:
    s = str(cat_raw or "").strip().upper()
    if not s:
        return "A2"
    if "RAC" in s or "REC" in s or "C1- R" in s or "C1-R" in s:
        return "RECAT_C1"
    if "A1" in s and "A2" not in s:
        return "A1"
    if "A2" in s and "B1" not in s and "C1" not in s:
        return "A2"
    if "B1" in s and "A2" not in s:
        return "B1"
    if "C1" in s and "A2" not in s:
        return "C1"
    return s


def _col_to_idx(col_str: str) -> int:
    res = 0
    for c in col_str:
        res = res * 26 + (ord(c.upper()) - ord("A") + 1)
    return res - 1


def cargar_hojas_desde_xlsx_local(xlsx_path: Path) -> Dict[str, List[List[str]]]:
    """Lee todas las pestañas de control_runt_digital.xlsx sin depender de openpyxl."""
    resultado: Dict[str, List[List[str]]] = {}
    if not xlsx_path.exists():
        return resultado
    try:
        with zipfile.ZipFile(xlsx_path, "r") as z:
            ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
            shared: List[str] = []
            if "xl/sharedStrings.xml" in z.namelist():
                ss_root = ET.fromstring(z.read("xl/sharedStrings.xml"))
                for si in ss_root.findall("s:si", ns):
                    texts = [t.text or "" for t in si.findall(".//s:t", ns)]
                    shared.append("".join(texts))

            wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
            rels_xml = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
            rel_map = {r.attrib["Id"]: r.attrib["Target"] for r in rels_xml}

            for sh in wb_xml.findall(".//s:sheet", ns):
                name = sh.attrib.get("name", "")
                rid = sh.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id", "")
                if not rid or rid not in rel_map:
                    continue
                target = "xl/" + rel_map[rid].lstrip("/")
                if target not in z.namelist():
                    continue
                ws_xml = ET.fromstring(z.read(target))
                rows_out: List[List[str]] = []
                for row in ws_xml.findall(".//s:sheetData/s:row", ns):
                    vals: List[str] = []
                    for c in row.findall("s:c", ns):
                        ref = c.attrib.get("r", "A1")
                        m = re.match(r"([A-Za-z]+)", ref)
                        c_idx = _col_to_idx(m.group(1)) if m else len(vals)
                        while len(vals) <= c_idx:
                            vals.append("")
                        t = c.attrib.get("t", "")
                        v_el = c.find("s:v", ns)
                        is_el = c.find("s:is", ns)
                        val = ""
                        if t == "s" and v_el is not None and v_el.text is not None:
                            s_idx = int(v_el.text)
                            val = shared[s_idx] if 0 <= s_idx < len(shared) else ""
                        elif t == "inlineStr" and is_el is not None:
                            val = "".join(x.text or "" for x in is_el.findall(".//s:t", ns))
                        elif v_el is not None and v_el.text is not None:
                            val = v_el.text
                        vals[c_idx] = val.replace("\n", " ").strip()
                    if any(vals):
                        rows_out.append(vals)
                resultado[name.strip()] = rows_out
    except Exception as exc:
        logger.warning("No se pudo leer control_runt_digital.xlsx local: %s", exc)
    return resultado


class RuntControlService:
    """
    Gestiona la hoja de cálculo de Control Digital RUNT (1XHbqfDa7dK5CWcgWKNgJZiXN_OSrjJCEJJu8F4z-u2M)
    con caché inteligente de 90s y respaldo local automático desde control_runt_digital.xlsx.
    """

    def __init__(self) -> None:
        self.sheet_id: str = os.getenv("GOOGLE_RUNT_SHEET_ID", RUNT_SHEET_ID_DEFAULT).strip() or RUNT_SHEET_ID_DEFAULT
        self._spreadsheet: Optional[Any] = None
        self._worksheets_cache: Dict[str, Any] = {}
        self._cache_lecturas: Dict[str, Tuple[float, List[List[str]]]] = {}
        self._memoria_hojas: Dict[str, List[List[str]]] = {}
        self._ultimo_fallo_ts: float = 0.0
        self._ultimo_error_msg: str = ""
        self._historial_lotes_masivos: List[List[Dict[str, Any]]] = []
        self._inicializar_memoria_local()

    def _inicializar_memoria_local(self) -> None:
        hojas_locales = cargar_hojas_desde_xlsx_local(RUNT_LOCAL_XLSX)
        for k, rows in hojas_locales.items():
            self._memoria_hojas[k.strip()] = rows

    @property
    def conectado_en_vivo(self) -> bool:
        if self._spreadsheet is None:
            try:
                self.conectar()
            except Exception:
                pass
        return self._spreadsheet is not None

    def cambiar_sheet_id(self, nuevo_id: str) -> Dict[str, Any]:
        limpio = str(nuevo_id or "").strip()
        m = re.search(r"/d/([a-zA-Z0-9-_]+)", limpio)
        if m:
            limpio = m.group(1)
        limpio = limpio.strip("/")
        if not limpio:
            return {"ok": False, "message": "ID de hoja de cálculo inválido."}
        self.sheet_id = limpio
        self._spreadsheet = None
        self._worksheets_cache.clear()
        self._cache_lecturas.clear()
        self._ultimo_fallo_ts = 0.0
        try:
            self.conectar()
            return {"ok": True, "message": f"Conectado en vivo a la hoja RUNT ({self.sheet_id}).", "conectado": True}
        except Exception as exc:
            return {
                "ok": True,
                "message": f"ID RUNT configurado ({self.sheet_id}). Modo local activo mientras compartes permisos con la cuenta de servicio ({exc}).",
                "conectado": False,
            }

    def conectar(self) -> Any:
        if self._spreadsheet is None:
            ahora = time.time()
            if self._ultimo_fallo_ts > 0 and (ahora - self._ultimo_fallo_ts < 25.0):
                raise ConnectionError(self._ultimo_error_msg or "Reintento en espera.")
            try:
                creds = clases_core.obtener_credenciales_google(clases_core.SCOPES)
                client = clases_core.gspread.authorize(creds)
                self._spreadsheet = client.open_by_key(self.sheet_id)
                self._ultimo_fallo_ts = 0.0
                self._ultimo_error_msg = ""
            except Exception as exc:
                self._ultimo_fallo_ts = ahora
                self._ultimo_error_msg = str(exc)
                raise
        return self._spreadsheet

    def _obtener_filas_pestana(self, nombre_pestana: str, forzar: bool = False) -> List[List[str]]:
        key = nombre_pestana.strip()
        ahora = time.time()
        if not forzar and key in self._cache_lecturas and (ahora - self._cache_lecturas[key][0] < 90.0):
            return [list(r) for r in self._cache_lecturas[key][1]]

        try:
            book = self.conectar()
            ws = self._worksheets_cache.get(key)
            if ws is None:
                # Buscar coincidencia exacta o ignorando espacios al final
                for w_cand in book.worksheets():
                    if w_cand.title.strip().lower() == key.lower():
                        ws = w_cand
                        break
                if ws is None:
                    ws = book.worksheet(nombre_pestana)
                self._worksheets_cache[key] = ws
            vals = ws.get_all_values()
            if vals:
                self._cache_lecturas[key] = (ahora, vals)
                self._memoria_hojas[key] = vals
                return [list(r) for r in vals]
        except Exception as exc:
            self._ultimo_error_msg = str(exc)
            if key in self._cache_lecturas:
                return [list(r) for r in self._cache_lecturas[key][1]]

        return [list(r) for r in self._memoria_hojas.get(key, [])]

    def _obtener_ws_en_vivo(self, nombre_pestana: str) -> Optional[Any]:
        key = nombre_pestana.strip()
        try:
            book = self.conectar()
            if key in self._worksheets_cache:
                return self._worksheets_cache[key]
            for w_cand in book.worksheets():
                if w_cand.title.strip().lower() == key.lower():
                    self._worksheets_cache[key] = w_cand
                    return w_cand
        except Exception as exc:
            logger.warning("Aviso obteniendo worksheet en vivo '%s': %s", nombre_pestana, exc)
        return None

    def listar_practica_runt(self, forzar: bool = False) -> List[Dict[str, Any]]:
        """
        Lee la pestaña principal 'Practica' donde se lleva el conteo de horas cargadas en RUNT:
        Teoría, Taller, Práctica, Total Clases Cargadas, Vigencia 90 días, Estado RUNT y Cargue Diario.
        Además cruza con las clases reales dictadas en el Bot de Clases (Registro_Clases).
        """
        filas = self._obtener_filas_pestana(TAB_PRACTICA, forzar=forzar)
        if not filas or len(filas) < 2:
            return []

        # Mapa de clases reales dictadas en el Bot + alumnos activos de Imagen 1 (Hoja Alumnos, Pendientes Cargue RUNT)
        por_doc_bot: Dict[Tuple[str, str], set] = {}
        por_nom_bot: Dict[Tuple[str, str], set] = {}
        bot_activos_doc: Dict[Tuple[str, str], Dict[str, Any]] = {}
        bot_activos_nom: Dict[Tuple[str, str], Dict[str, Any]] = {}
        bot_cargados_doc: set = set()
        bot_cargados_nom: set = set()
        try:
            por_doc_bot, por_nom_bot = clases_core.sheets_service._calcular_mapa_clases_por_alumno()
        except Exception:
            pass
        try:
            alumnos_bot_todos = clases_core.sheets_service.listar_alumnos_activos("TODAS", modo="todos")
            for al_b in alumnos_bot_todos:
                c_b = _normalizar_cat_runt(al_b.get("Categoria", "A2"))
                d_b = _norm(_limpiar_cedula(al_b.get("Documento", "")))
                n_b = _norm(clases_core.limpiar_solo_nombre_persona(str(al_b.get("Nombre_Completo", ""))) or str(al_b.get("Nombre_Completo", "")))
                est_r_b = str(al_b.get("Estado_RUNT", "")).strip().lower()
                if "cargue" in est_r_b or "cargado" in est_r_b:
                    if d_b:
                        bot_cargados_doc.add((d_b, c_b))
                    if n_b:
                        bot_cargados_nom.add((n_b, c_b))
                else:
                    if d_b:
                        bot_activos_doc[(d_b, c_b)] = al_b
                    if n_b:
                        bot_activos_nom[(n_b, c_b)] = al_b
        except Exception as exc_ab:
            logger.warning("Aviso cruzando alumnos activos del Bot en RUNT: %s", exc_ab)

        hoy = datetime.datetime.now(clases_core.TZ_COLOMBIA).date()
        registros: List[Dict[str, Any]] = []

        for idx_f, r in enumerate(filas[1:], start=2):
            nom = str(r[0]).strip() if len(r) > 0 else ""
            ced = _limpiar_cedula(r[1] if len(r) > 1 else "")
            if not nom and not ced:
                continue
            cat_raw = str(r[2]).strip().upper() if len(r) > 2 else "A2"
            cat = _normalizar_cat_runt(cat_raw)
            f_mat = _excel_serial_a_iso(r[3] if len(r) > 3 else "")
            prioridad = str(r[4]).strip() if len(r) > 4 else ""
            f_vence = _excel_serial_a_iso(r[5] if len(r) > 5 else "")
            f_reg_runt = _excel_serial_a_iso(r[6] if len(r) > 6 else "")
            vigente_txt = str(r[7]).strip() if len(r) > 7 else ""

            teoria = _limpiar_entero(r[8] if len(r) > 8 else 0, 0)
            taller = _limpiar_entero(r[9] if len(r) > 9 else 0, 0)
            practica = _limpiar_entero(r[10] if len(r) > 10 else 0, 0)
            cargadas_raw = _limpiar_entero(r[11] if len(r) > 11 else 0, 0)
            total_cargadas = teoria + taller + practica if (teoria or taller or practica) else cargadas_raw

            instructor = str(r[12]).strip() if len(r) > 12 else ""
            hora_reg = str(r[13]).strip() if len(r) > 13 else ""
            act_1 = str(r[14]).strip() if len(r) > 14 else ""
            act_2 = str(r[15]).strip() if len(r) > 15 else ""
            f_cargue = _excel_serial_a_iso(r[16] if len(r) > 16 else "")
            cargue_diario = str(r[17]).strip().upper() if len(r) > 17 else ""

            es_cargado = "CARGADO" in act_1.upper() or "CARGADO" in act_2.upper() or "CARGUE" in act_1.upper()
            estado_runt = "CARGADO RUNT" if es_cargado else "PENDIENTE RUNT"

            # Metas reglamentarias por categoría (Teoría + Taller + Práctica)
            metas_cat = {
                "A1": {"teoria": 25, "taller": 3, "practica": 8, "total": 36},
                "A2": {"teoria": 25, "taller": 3, "practica": 15, "total": 43},
                "B1": {"teoria": 25, "taller": 5, "practica": 20, "total": 50},
                "C1": {"teoria": 30, "taller": 5, "practica": 30, "total": 65},
                "RECAT_C1": {"teoria": 8, "taller": 0, "practica": 8, "total": 16},
            }.get(cat, {"teoria": 25, "taller": 5, "practica": 20, "total": 50})

            # Cruce con clases prácticas dictadas en el Bot de Telegram y con la lista activa del Bot (Imagen 1)
            doc_norm = _norm(ced)
            nom_norm = _norm(clases_core.limpiar_solo_nombre_persona(nom) or nom)
            clases_bot_set: set = set()
            if doc_norm and (doc_norm, cat) in por_doc_bot:
                clases_bot_set.update(por_doc_bot[(doc_norm, cat)])
            if nom_norm and (nom_norm, cat) in por_nom_bot:
                clases_bot_set.update(por_nom_bot[(nom_norm, cat)])
            practica_bot = len({c for c in clases_bot_set if 1 <= c <= metas_cat["practica"]})
            diferencia_practica_bot = max(0, practica_bot - practica)

            match_bot_activo = (
                (bot_activos_doc.get((doc_norm, cat)) if doc_norm else None)
                or (bot_activos_nom.get((nom_norm, cat)) if nom_norm else None)
            )
            clases_bot_al = _limpiar_entero(match_bot_activo.get("Clases_Completadas") if match_bot_activo else 0, 0)
            if match_bot_activo:
                practica_bot = max(practica_bot, clases_bot_al)
                diferencia_practica_bot = max(0, practica_bot - practica)
            else:
                practica_bot = 0
                diferencia_practica_bot = 0

            # Estricto: SOLO alumnos que están activos en la tabla del Bot (Imagen 1: En Curso / Finalizaron sin RUNT)
            en_bot_practica = bool(match_bot_activo)
            bot_avance_pct = min(100, int(round((practica_bot / max(1, metas_cat["practica"])) * 100))) if en_bot_practica else 0
            bot_termino_practica = en_bot_practica and (
                practica_bot >= metas_cat["practica"]
                or bool(match_bot_activo and match_bot_activo.get("Curso_Completado"))
                or str((match_bot_activo or {}).get("Estado_Curso", "")).strip().upper() == "FINALIZADO"
            )

            # Cálculo automático de Vence 90 días y Semáforo si tiene Fecha Matrícula
            if f_mat and len(f_mat) == 10 and not f_vence:
                try:
                    d_m = datetime.datetime.strptime(f_mat, "%Y-%m-%d").date()
                    f_vence = (d_m + datetime.timedelta(days=90)).strftime("%Y-%m-%d")
                except Exception:
                    pass

            dias_restantes_90: Optional[int] = None
            semaforo_90 = "SIN_FECHA"
            if f_vence and len(f_vence) == 10:
                try:
                    d_v = datetime.datetime.strptime(f_vence, "%Y-%m-%d").date()
                    dias_restantes_90 = (d_v - hoy).days
                    if es_cargado or total_cargadas >= metas_cat["total"]:
                        semaforo_90 = "COMPLETADO"
                        if not vigente_txt:
                            vigente_txt = "✅ VIGENTE"
                    elif dias_restantes_90 < 0:
                        semaforo_90 = "VENCIDO"
                        vigente_txt = "❌ VENCIDO"
                    elif dias_restantes_90 <= 15:
                        semaforo_90 = "POR_VENCER"
                        vigente_txt = "⚠️ POR VENCER"
                    else:
                        semaforo_90 = "VIGENTE"
                        if not vigente_txt:
                            vigente_txt = "✅ VIGENTE"
                except Exception:
                    pass
            elif "VENCIDO" in vigente_txt.upper():
                semaforo_90 = "VENCIDO"
            elif "POR VENCER" in vigente_txt.upper():
                semaforo_90 = "POR_VENCER"
            elif "VIGENTE" in vigente_txt.upper():
                semaforo_90 = "VIGENTE"

            pct_runt = min(100, int(round((total_cargadas / max(1, metas_cat["total"])) * 100)))
            teoria_faltante = max(0, metas_cat["teoria"] - teoria)
            taller_faltante = max(0, metas_cat["taller"] - taller)
            practica_real = max(practica, practica_bot)
            practica_faltante = max(0, metas_cat["practica"] - practica)
            falta_algo_runt = (teoria_faltante > 0) or (taller_faltante > 0) or (practica_faltante > 0)

            # Prioridad basada en alumnos ACTIVOS EN PRÁCTICA DEL BOT (Imagen 1) que no tienen 100% RUNT:
            # 1: Activo en Bot que YA TERMINÓ práctica (o ya tiene 100% horas) y falta cargue manual RUNT
            # 2: Activo en Bot en Práctica que aún NO tiene 100% de Teoría, Taller o Práctica (PRIORIDAD CARGUE)
            # 3: Pendientes en hoja RUNT sin actividad reciente en el Bot (inactivos / revisados / no volvieron)
            # 4: Solo teoría/taller sin práctica ni actividad Bot
            # 5: Ya cargados 100% en RUNT
            es_terminado_sin_cargue = (
                not es_cargado
                and en_bot_practica
                and (bot_termino_practica or not falta_algo_runt)
            )
            es_urgente_en_practica = (
                not es_cargado
                and en_bot_practica
                and not es_terminado_sin_cargue
                and falta_algo_runt
            )
            es_prioridad_bot = es_terminado_sin_cargue or es_urgente_en_practica

            prio_low = prioridad.strip().lower()
            tiene_prio_hoja = bool(prio_low) and prio_low not in ("revisado", "sin prioridad", "-")

            if es_cargado:
                orden_prioridad = 5
                etiqueta_prioridad_op = "✅ 5. YA CERTIFICADO / CARGADO RUNT"
            elif es_terminado_sin_cargue:
                orden_prioridad = 1
                etiqueta_prioridad_op = "🔥 1. BOT TERMINÓ PRÁCTICA (PRIORIDAD #1)"
            elif es_urgente_en_practica:
                orden_prioridad = 2
                etiqueta_prioridad_op = "🤖 2. EN PRÁCTICA BOT (PRIORIDAD <100%)"
            elif tiene_prio_hoja:
                orden_prioridad = 3
                etiqueta_prioridad_op = "⏳ 3. PENDIENTE RUNT (SIN ACTIVIDAD BOT)"
            else:
                orden_prioridad = 4
                etiqueta_prioridad_op = "⚪ 4. SIN PRIORIDAD / INACTIVO (SIN BOT)"

            # Sugerencia de lote respetando el tope de máximo 8 clases por día en RUNT
            sug_teoria_8 = min(8, teoria_faltante) if teoria_faltante > 0 else 0
            sug_taller_8 = min(8, taller_faltante) if taller_faltante > 0 else 0
            sug_practica_8 = min(8, practica_faltante) if practica_faltante > 0 else 0

            registros.append({
                "fila": idx_f,
                "nombre": nom,
                "cedula": ced,
                "categoria": cat,
                "categoria_raw": cat_raw or cat,
                "fecha_matricula": f_mat,
                "prioridad": prioridad,
                "orden_prioridad": orden_prioridad,
                "etiqueta_prioridad_op": etiqueta_prioridad_op,
                "en_bot_practica": en_bot_practica,
                "bot_termino_practica": bot_termino_practica,
                "bot_avance_pct": bot_avance_pct,
                "es_prioridad_bot": es_prioridad_bot,
                "es_terminado_sin_cargue": es_terminado_sin_cargue,
                "es_urgente_en_practica": es_urgente_en_practica,
                "vence_90_dias": f_vence,
                "dias_restantes_90": dias_restantes_90,
                "fecha_registro_runt": f_reg_runt,
                "vigente_texto": vigente_txt or "✅ VIGENTE",
                "semaforo_90": semaforo_90,
                "teoria": teoria,
                "meta_teoria": metas_cat["teoria"],
                "teoria_faltante": teoria_faltante,
                "sug_teoria_8": sug_teoria_8,
                "taller": taller,
                "meta_taller": metas_cat["taller"],
                "taller_faltante": taller_faltante,
                "sug_taller_8": sug_taller_8,
                "practica": practica,
                "meta_practica": metas_cat["practica"],
                "practica_faltante": practica_faltante,
                "sug_practica_8": sug_practica_8,
                "practica_bot": practica_bot,
                "pendiente_subir_bot": diferencia_practica_bot,
                "clases_cargadas": total_cargadas,
                "meta_total": metas_cat["total"],
                "porcentaje_runt": pct_runt,
                "instructor": instructor,
                "hora_registro": hora_reg,
                "estado_runt": estado_runt,
                "fecha_cargue": f_cargue,
                "cargue_diario": cargue_diario,
            })

        return registros

    def listar_no_runt(self) -> List[Dict[str, Any]]:
        filas = self._obtener_filas_pestana(TAB_NO_RUNT)
        if not filas or len(filas) < 2:
            return []
        out: List[Dict[str, Any]] = []
        for idx_f, r in enumerate(filas[1:], start=2):
            ts = _excel_serial_a_iso(r[0] if len(r) > 0 else "")
            nom = str(r[1]).strip() if len(r) > 1 else ""
            ced = _limpiar_cedula(r[2] if len(r) > 2 else "")
            if not nom and not ced:
                continue
            cat = str(r[3]).strip().upper() if len(r) > 3 else "A2"
            f_mat = _excel_serial_a_iso(r[4] if len(r) > 4 else "")
            tipo_doc = str(r[5]).strip().upper() if len(r) > 5 else "CC"
            f_runt = _excel_serial_a_iso(r[6] if len(r) > 6 else "")
            novedad = str(r[7]).strip() if len(r) > 7 else ""
            est = str(r[8]).strip().upper() if len(r) > 8 else "SIN REGISTRO"
            out.append({
                "fila": idx_f,
                "timestamp": ts,
                "nombre": nom,
                "cedula": ced,
                "categoria": cat,
                "fecha_matricula": f_mat,
                "tipo_documento": tipo_doc or "CC",
                "fecha_registro_runt": f_runt,
                "novedad": novedad,
                "estado_registro": est,
            })
        return out

    def listar_inicio_practica(self) -> List[Dict[str, Any]]:
        filas = self._obtener_filas_pestana(TAB_INICIO_PRACTICA)
        if not filas or len(filas) < 2:
            return []
        out: List[Dict[str, Any]] = []
        for idx_f, r in enumerate(filas[1:], start=2):
            ts = _excel_serial_a_iso(r[0] if len(r) > 0 else "")
            nom = str(r[1]).strip() if len(r) > 1 else ""
            ced = _limpiar_cedula(r[2] if len(r) > 2 else "")
            if not nom and not ced:
                continue
            nov = str(r[3]).strip() if len(r) > 3 else ""
            cat = str(r[4]).strip() if len(r) > 4 else ""
            horarios = str(r[5]).strip() if len(r) > 5 else ""
            out.append({
                "fila": idx_f,
                "fecha": ts,
                "nombre": nom,
                "cedula": ced,
                "novedades": nov,
                "categoria": cat,
                "horarios_practica": horarios,
            })
        return list(reversed(out))

    def listar_codigos_runt(self) -> List[Dict[str, Any]]:
        filas = self._obtener_filas_pestana(TAB_CODIGO_RUNT)
        if not filas or len(filas) < 2:
            return []
        out: List[Dict[str, Any]] = []
        for idx_f, r in enumerate(filas[1:], start=2):
            nom = str(r[0]).strip() if len(r) > 0 else ""
            ced = _limpiar_cedula(r[1] if len(r) > 1 else "")
            fecha = _excel_serial_a_iso(r[2] if len(r) > 2 else "")
            codigo = str(r[3]).strip() if len(r) > 3 else ""
            num = _limpiar_entero(r[4] if len(r) > 4 else 0, 0)
            resp = str(r[5]).strip() if len(r) > 5 else ""
            if not nom and not ced and not codigo:
                continue
            out.append({
                "fila": idx_f,
                "numeracion": num,
                "nombre": nom,
                "cedula": ced,
                "fecha": fecha,
                "codigo_runt": codigo,
                "responsable": resp,
            })
        return out[:250]

    def guardar_o_actualizar_practica_runt(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Actualiza o inserta una fila en la pestaña 'Practica' del Google Sheet RUNT
        (y también sincroniza el Estado_RUNT en la hoja 'Alumnos' del Bot).
        """
        fila_idx = _limpiar_entero(payload.get("fila", 0), 0)
        nombre = str(payload.get("nombre", "")).strip().upper()
        cedula = _limpiar_cedula(payload.get("cedula", ""))
        categoria = _normalizar_cat_runt(payload.get("categoria", "A2"))
        fecha_mat = _excel_serial_a_iso(payload.get("fecha_matricula", ""))
        prioridad = str(payload.get("prioridad", "")).strip()
        fecha_reg_runt = _excel_serial_a_iso(payload.get("fecha_registro_runt", "")) or datetime.datetime.now(clases_core.TZ_COLOMBIA).strftime("%Y-%m-%d")
        teoria = _limpiar_entero(payload.get("teoria", 0), 0)
        taller = _limpiar_entero(payload.get("taller", 0), 0)
        practica = _limpiar_entero(payload.get("practica", 0), 0)
        total_cargadas = teoria + taller + practica
        instructor = str(payload.get("instructor", "")).strip().upper()
        estado_raw = str(payload.get("estado_runt", "PENDIENTE RUNT")).strip().upper()
        es_cargado = "CARGADO" in estado_raw or "CARGUE" in estado_raw
        act_1 = "✅ CARGADO RUNT" if es_cargado else "⚠️ PENDIENTE RUNT"
        act_2 = "CARGADO RUNT" if es_cargado else "PENDIENTE RUNT"
        cargue_diario = str(payload.get("cargue_diario", "")).strip().upper()
        ahora_col = datetime.datetime.now(clases_core.TZ_COLOMBIA)
        fecha_cargue = str(payload.get("fecha_cargue", "")).strip() or ahora_col.strftime("%Y-%m-%d")
        hora_reg = str(payload.get("hora_registro", "")).strip() or ahora_col.strftime("%H:%M:%S")

        vence_90 = _excel_serial_a_iso(payload.get("vence_90_dias", ""))
        if fecha_mat and len(fecha_mat) == 10 and not vence_90:
            try:
                d0 = datetime.datetime.strptime(fecha_mat, "%Y-%m-%d").date()
                vence_90 = (d0 + datetime.timedelta(days=90)).strftime("%Y-%m-%d")
            except Exception:
                pass

        vigente_txt = "✅ VIGENTE"
        if vence_90 and len(vence_90) == 10:
            try:
                dv = datetime.datetime.strptime(vence_90, "%Y-%m-%d").date()
                d_rest = (dv - ahora_col.date()).days
                if d_rest < 0 and not es_cargado:
                    vigente_txt = "❌ VENCIDO"
                elif d_rest <= 15 and not es_cargado:
                    vigente_txt = "⚠️ POR VENCER"
            except Exception:
                pass

        nueva_fila = [
            nombre,
            cedula,
            categoria,
            fecha_mat,
            prioridad,
            vence_90,
            fecha_reg_runt,
            vigente_txt,
            str(teoria) if teoria > 0 else "",
            str(taller) if taller > 0 else "",
            str(practica) if practica > 0 else "",
            str(total_cargadas),
            instructor,
            hora_reg,
            act_1,
            act_2,
            fecha_cargue,
            cargue_diario,
        ]

        filas = self._obtener_filas_pestana(TAB_PRACTICA)
        if not filas:
            filas = [list(HEADERS_PRACTICA)]

        # Si no viene fila_idx, buscar por (cedula, categoria) o (nombre, categoria)
        if fila_idx < 2:
            for idx_cand in range(2, len(filas) + 1):
                fc = filas[idx_cand - 1]
                nom_c = _norm(fc[0] if len(fc) > 0 else "")
                ced_c = _limpiar_cedula(fc[1] if len(fc) > 1 else "")
                cat_c = _normalizar_cat_runt(fc[2] if len(fc) > 2 else "A2")
                if cat_c == categoria and ((cedula and ced_c == cedula) or (nombre and nom_c == _norm(nombre))):
                    fila_idx = idx_cand
                    # Conservar fecha_matricula o fecha_registro_runt previas si no se enviaron
                    if not fecha_mat and len(fc) > 3 and fc[3]:
                        nueva_fila[3] = _excel_serial_a_iso(fc[3])
                    if len(fc) > 6 and fc[6] and not payload.get("fecha_registro_runt"):
                        nueva_fila[6] = _excel_serial_a_iso(fc[6])
                    break

        guardado_sheets = False
        ws = self._obtener_ws_en_vivo(TAB_PRACTICA)
        if fila_idx >= 2 and fila_idx <= len(filas):
            filas[fila_idx - 1] = nueva_fila
            if ws is not None:
                try:
                    ws.update(
                        values=[nueva_fila],
                        range_name=f"A{fila_idx}:R{fila_idx}",
                        value_input_option="USER_ENTERED",
                    )
                    guardado_sheets = True
                except Exception as exc:
                    logger.warning("No se pudo actualizar fila %d en hoja Practica RUNT: %s", fila_idx, exc)
        else:
            filas.append(nueva_fila)
            fila_idx = len(filas)
            if ws is not None:
                try:
                    if fila_idx > ws.row_count:
                        ws.add_rows(100)
                    ws.update(
                        values=[nueva_fila],
                        range_name=f"A{fila_idx}:R{fila_idx}",
                        value_input_option="USER_ENTERED",
                    )
                    guardado_sheets = True
                except Exception as exc:
                    logger.warning("No se pudo insertar fila nueva en hoja Practica RUNT: %s", exc)

        self._memoria_hojas[TAB_PRACTICA] = filas
        self._cache_lecturas[TAB_PRACTICA] = (time.time(), filas)

        # Sincronizar bidireccionalmente con la hoja 'Alumnos' del Bot (Módulo 2)
        try:
            est_bot = "Cargue RUNT" if es_cargado else "No Registrado RUNT"
            clases_core.sheets_service.actualizar_estado_runt_alumno(
                documento_o_nombre=cedula or nombre,
                categoria=categoria,
                estado_runt=est_bot,
                nombre_alumno=nombre,
            )
        except Exception as exc_sync:
            logger.warning("Aviso sincronizando estado RUNT hacia hoja Alumnos del Bot: %s", exc_sync)

        return {
            "ok": True,
            "fila": fila_idx,
            "guardado_google_sheets": guardado_sheets,
            "message": (
                f"✅ {nombre or cedula} ({categoria}) actualizado: Teoría={teoria}, Taller={taller}, Práctica={practica} "
                f"(Total: {total_cargadas} h) • {act_1}"
                + (" [Sincronizado en Google Sheets RUNT]" if guardado_sheets else " [Guardado en base activa y sincronizado con Módulo 2]")
            ),
        }

    def cambiar_estado_no_runt(self, fila_idx: int, nuevo_estado: str = "REGISTRADO", novedad: str = "") -> Dict[str, Any]:
        filas = self._obtener_filas_pestana(TAB_NO_RUNT)
        if fila_idx < 2 or fila_idx > len(filas):
            return {"ok": False, "message": f"Fila #{fila_idx} no encontrada en pestaña 'No runt'."}
        fila = list(filas[fila_idx - 1])
        while len(fila) < 9:
            fila.append("")
        hoy_str = datetime.datetime.now(clases_core.TZ_COLOMBIA).strftime("%Y-%m-%d")
        if nuevo_estado.upper() == "REGISTRADO" and not fila[6]:
            fila[6] = hoy_str
        if novedad:
            fila[7] = novedad
        fila[8] = nuevo_estado.upper()
        filas[fila_idx - 1] = fila
        self._memoria_hojas[TAB_NO_RUNT] = filas
        self._cache_lecturas[TAB_NO_RUNT] = (time.time(), filas)

        ws = self._obtener_ws_en_vivo(TAB_NO_RUNT)
        if ws is not None:
            try:
                ws.update(values=[[fila[6], fila[7], fila[8]]], range_name=f"G{fila_idx}:I{fila_idx}", value_input_option="USER_ENTERED")
            except Exception as exc:
                logger.warning("Aviso actualizando pestaña 'No runt': %s", exc)

        return {
            "ok": True,
            "message": f"Alumno {fila[1] or fila[2]} actualizado a '{nuevo_estado.upper()}' en pestaña 'No runt'.",
        }

    def sincronizar_clases_bot_a_runt(self) -> Dict[str, Any]:
        """
        Toma las clases prácticas reales reportadas por los instructores en el Bot
        y actualiza el conteo de 'Practica' y 'Clases Cargadas' en la hoja RUNT
        para todos los alumnos donde el Bot tenga más horas registradas.
        """
        practica_list = self.listar_practica_runt(forzar=False)
        actualizados = 0
        for item in practica_list:
            if item["practica_bot"] > item["practica"]:
                self.guardar_o_actualizar_practica_runt({
                    "fila": item["fila"],
                    "nombre": item["nombre"],
                    "cedula": item["cedula"],
                    "categoria": item["categoria"],
                    "fecha_matricula": item["fecha_matricula"],
                    "prioridad": "Termino" if item["practica_bot"] >= item["meta_practica"] else item["prioridad"],
                    "vence_90_dias": item["vence_90_dias"],
                    "fecha_registro_runt": item["fecha_registro_runt"],
                    "teoria": item["teoria"] or item["meta_teoria"],
                    "taller": item["taller"] or item["meta_taller"],
                    "practica": item["practica_bot"],
                    "instructor": item["instructor"],
                    "estado_runt": item["estado_runt"],
                    "cargue_diario": "PRACTICA",
                })
                actualizados += 1
        return {
            "ok": True,
            "actualizados": actualizados,
            "message": f"🔄 Sincronización completada: {actualizados} alumno(s) actualizados con las horas prácticas dictadas en el Bot.",
        }

    def cargue_masivo_teoria_o_taller(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Realiza el cargue masivo sumando horas de Teoría, Taller o Práctica en la pestaña 'Practica'
        para las filas seleccionadas (☑️) o los alumnos pendientes de una categoría.
        Acepta `horas_sumar` personalizada (por defecto 8) y suma automáticamente sobre lo que ya
        tiene cada alumno, respetando el tope reglamentario de su categoría.
        """
        tipo = str(payload.get("tipo", "TEORIA")).strip().upper()  # 'TEORIA', 'TALLER', 'PRACTICA' o 'TEORIA_Y_TALLER'
        categoria_filtro = str(payload.get("categoria", "A2")).strip().upper()
        modo_cupo = str(payload.get("modo_cupo", "MAX_8")).strip().upper()  # 'MAX_8', 'SUMA_EXACTA' o 'COMPLETAR_META'
        raw_horas = payload.get("horas_sumar", None)
        try:
            horas_inc = int(raw_horas) if raw_horas is not None and str(raw_horas).strip() != "" else 8
        except Exception:
            horas_inc = 8
        horas_inc = max(1, min(40, horas_inc))

        fecha_reporte = _excel_serial_a_iso(payload.get("fecha_cargue", "")) or datetime.datetime.now(clases_core.TZ_COLOMBIA).strftime("%Y-%m-%d")
        filas_objetivo = payload.get("filas") or []
        filas_set = {int(x) for x in filas_objetivo if str(x).isdigit()}

        practica_list = self.listar_practica_runt(forzar=False)
        actualizados = 0
        nombres_act: List[str] = []
        snapshot_lote: List[Dict[str, Any]] = []

        for item in practica_list:
            if item["estado_runt"] == "CARGADO RUNT":
                continue
            if filas_set:
                if item["fila"] not in filas_set:
                    continue
            else:
                if categoria_filtro != "TODAS" and item["categoria"] != categoria_filtro:
                    continue

            teo_actual = item["teoria"]
            tal_actual = item["taller"]
            prac_actual = item["practica"]
            meta_teo = item["meta_teoria"]
            meta_tal = item["meta_taller"]
            meta_prac = item["meta_practica"]

            nuevo_teo = teo_actual
            nuevo_tal = tal_actual
            nuevo_prac = prac_actual
            etiqueta_diaria = item["cargue_diario"]

            if tipo in ("TEORIA", "TEORIA_Y_TALLER") and teo_actual < meta_teo:
                if modo_cupo == "COMPLETAR_META":
                    nuevo_teo = meta_teo
                else:
                    nuevo_teo = min(meta_teo, teo_actual + horas_inc)
                etiqueta_diaria = "TEORIA"

            if tipo in ("TALLER", "TEORIA_Y_TALLER") and tal_actual < meta_tal:
                if modo_cupo == "COMPLETAR_META":
                    nuevo_tal = meta_tal
                else:
                    cupo_restante = max(0, horas_inc - (nuevo_teo - teo_actual))
                    inc_tal = min(meta_tal - tal_actual, cupo_restante if tipo == "TEORIA_Y_TALLER" else horas_inc)
                    nuevo_tal = tal_actual + inc_tal
                etiqueta_diaria = "TALLER" if nuevo_teo == teo_actual else "TEORIA+TALLER"

            if tipo == "PRACTICA" and prac_actual < meta_prac:
                if modo_cupo == "COMPLETAR_META":
                    nuevo_prac = meta_prac
                else:
                    nuevo_prac = min(meta_prac, prac_actual + horas_inc)
                etiqueta_diaria = "PRACTICA"

            if nuevo_teo != teo_actual or nuevo_tal != tal_actual or nuevo_prac != prac_actual:
                snapshot_lote.append({
                    "fila": item["fila"],
                    "nombre": item["nombre"],
                    "cedula": item["cedula"],
                    "categoria": item["categoria"],
                    "fecha_matricula": item["fecha_matricula"],
                    "prioridad": item["prioridad"],
                    "vence_90_dias": item["vence_90_dias"],
                    "fecha_registro_runt": item["fecha_registro_runt"],
                    "teoria": teo_actual,
                    "taller": tal_actual,
                    "practica": prac_actual,
                    "instructor": item["instructor"],
                    "estado_runt": item["estado_runt"],
                    "fecha_cargue": item["fecha_cargue"],
                    "cargue_diario": item["cargue_diario"],
                })
                self.guardar_o_actualizar_practica_runt({
                    "fila": item["fila"],
                    "nombre": item["nombre"],
                    "cedula": item["cedula"],
                    "categoria": item["categoria"],
                    "fecha_matricula": item["fecha_matricula"],
                    "prioridad": item["prioridad"],
                    "vence_90_dias": item["vence_90_dias"],
                    "fecha_registro_runt": item["fecha_registro_runt"],
                    "teoria": nuevo_teo,
                    "taller": nuevo_tal,
                    "practica": nuevo_prac,
                    "instructor": item["instructor"],
                    "estado_runt": item["estado_runt"],
                    "fecha_cargue": fecha_reporte,
                    "cargue_diario": etiqueta_diaria,
                })
                actualizados += 1
                if len(nombres_act) < 6:
                    if tipo == "TEORIA":
                        nombres_act.append(f"{item['nombre']} ({teo_actual}→{nuevo_teo}/{meta_teo}h Teo)")
                    elif tipo == "TALLER":
                        nombres_act.append(f"{item['nombre']} ({tal_actual}→{nuevo_tal}/{meta_tal}h Tal)")
                    elif tipo == "PRACTICA":
                        nombres_act.append(f"{item['nombre']} ({prac_actual}→{nuevo_prac}/{meta_prac}h Prác)")
                    else:
                        nombres_act.append(f"{item['nombre']} ({nuevo_teo}/{meta_teo}T, {nuevo_tal}/{meta_tal}Tl)")

        if snapshot_lote:
            self._historial_lotes_masivos.append(snapshot_lote)

        resumen_noms = ", ".join(nombres_act) + (f" y {actualizados - len(nombres_act)} más" if actualizados > len(nombres_act) else "")
        etq_modo = "Completar Meta" if modo_cupo == "COMPLETAR_META" else f"+{horas_inc}h"
        return {
            "ok": True,
            "actualizados": actualizados,
            "message": (
                f"✅ Sumado {tipo} ({etq_modo}) con fecha {fecha_reporte} a {actualizados} alumno(s)"
                + (f": {resumen_noms}." if resumen_noms else ".")
            ),
        }

    def deshacer_o_restaurar_cargue_masivo(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Permite revertir un cargue masivo realizado por error:
        - modo == 'DESHACER_ULTIMO': revierte el último lote guardado en _historial_lotes_masivos
          (y si no hay historial en memoria porque ocurrió antes de reiniciar, restaura los 14 alumnos A2 desde el archivo base original).
        - modo == 'RESTAURAR_14_A2_ORIGINAL': restaura exactamente los 14 alumnos A2 de Teoría a sus valores originales del archivo subido (16h, 8h, 24h, 0h).
        - modo == 'RESTAR_8_TEORIA': resta 8 horas de teoría a las filas seleccionadas (o a los 14 alumnos A2).
        """
        modo = str(payload.get("modo", "DESHACER_ULTIMO")).strip().upper()
        filas_objetivo = payload.get("filas") or []
        filas_set = {int(x) for x in filas_objetivo if str(x).isdigit()}

        if modo == "DESHACER_ULTIMO" and self._historial_lotes_masivos:
            lote_previo = self._historial_lotes_masivos.pop()
            restaurados = 0
            noms: List[str] = []
            for snap in lote_previo:
                self.guardar_o_actualizar_practica_runt(snap)
                restaurados += 1
                if len(noms) < 6:
                    noms.append(f"{snap['nombre']} ({snap['teoria']}T/{snap['taller']}Tl)")
            resumen = ", ".join(noms) + (f" y {restaurados - len(noms)} más" if restaurados > len(noms) else "")
            return {
                "ok": True,
                "restaurados": restaurados,
                "message": f"↩️ Se deshizo el último cargue masivo y se restauraron {restaurados} alumno(s) a sus horas anteriores: {resumen}.",
            }

        # Respaldo original desde control_runt_digital.xlsx (copia intacta de BD_Americana_2__(2).xlsx)
        hojas_base = cargar_hojas_desde_xlsx_local(RUNT_LOCAL_XLSX)
        rows_base = hojas_base.get("Practica", [])
        mapa_base_por_fila: Dict[int, Dict[str, Any]] = {}
        mapa_base_por_ced_cat: Dict[Tuple[str, str], Dict[str, Any]] = {}
        for idx, r in enumerate(rows_base[1:], start=2):
            nom_b = str(r[0]).strip() if len(r) > 0 else ""
            ced_b = _limpiar_cedula(r[1] if len(r) > 1 else "")
            cat_b = _normalizar_cat_runt(str(r[2]).strip().upper() if len(r) > 2 else "A2")
            if not nom_b and not ced_b:
                continue
            info_b = {
                "fila": idx,
                "nombre": nom_b,
                "cedula": ced_b,
                "categoria": cat_b,
                "teoria": _limpiar_entero(r[8] if len(r) > 8 else 0, 0),
                "taller": _limpiar_entero(r[9] if len(r) > 9 else 0, 0),
                "practica": _limpiar_entero(r[10] if len(r) > 10 else 0, 0),
                "fecha_cargue": _excel_serial_a_iso(r[16] if len(r) > 16 else ""),
                "cargue_diario": str(r[17]).strip().upper() if len(r) > 17 else "",
            }
            mapa_base_por_fila[idx] = info_b
            if ced_b:
                mapa_base_por_ced_cat[(ced_b, cat_b)] = info_b

        # Los 14 alumnos A2 afectados por el cargue masivo de Teoría por error:
        cedulas_14_a2 = {
            "1098649376", "1096702152", "1003251307", "1007899397", "1098660472",
            "1097194141", "71790854", "1099743289", "1095302560", "1099742651",
            "1098644892", "1005324619", "1043653731", "1116508657",
        }
        filas_14_a2 = {6, 7, 8, 10, 12, 15, 16, 17, 18, 19, 42, 46, 88, 132}

        practica_actual = self.listar_practica_runt(forzar=False)
        restaurados = 0
        noms_rest: List[str] = []

        for item in practica_actual:
            if filas_set:
                if item["fila"] not in filas_set:
                    continue
            else:
                if item["fila"] not in filas_14_a2 and item["cedula"] not in cedulas_14_a2:
                    continue
                if item["categoria"] != "A2" or item["estado_runt"] == "CARGADO RUNT":
                    continue

            orig = mapa_base_por_fila.get(item["fila"]) or mapa_base_por_ced_cat.get((item["cedula"], item["categoria"]))
            if modo == "RESTAR_8_TEORIA":
                nuevo_teo = max(0, item["teoria"] - 8)
                nuevo_tal = item["taller"]
                nueva_fecha = item["fecha_cargue"]
                nuevo_diario = item["cargue_diario"]
            elif orig:
                nuevo_teo = orig["teoria"]
                nuevo_tal = orig["taller"]
                nueva_fecha = orig["fecha_cargue"]
                nuevo_diario = orig["cargue_diario"]
            else:
                nuevo_teo = max(0, item["teoria"] - 8)
                nuevo_tal = item["taller"]
                nueva_fecha = item["fecha_cargue"]
                nuevo_diario = item["cargue_diario"]

            self.guardar_o_actualizar_practica_runt({
                "fila": item["fila"],
                "nombre": item["nombre"],
                "cedula": item["cedula"],
                "categoria": item["categoria"],
                "fecha_matricula": item["fecha_matricula"],
                "prioridad": item["prioridad"],
                "vence_90_dias": item["vence_90_dias"],
                "fecha_registro_runt": item["fecha_registro_runt"],
                "teoria": nuevo_teo,
                "taller": nuevo_tal,
                "practica": item["practica"],
                "instructor": item["instructor"],
                "estado_runt": item["estado_runt"],
                "fecha_cargue": nueva_fecha,
                "cargue_diario": nuevo_diario,
            })
            restaurados += 1
            if len(noms_rest) < 7:
                noms_rest.append(f"{item['nombre']} ({nuevo_teo}/{item['meta_teoria']}T)")

        resumen = ", ".join(noms_rest) + (f" y {restaurados - len(noms_rest)} más" if restaurados > len(noms_rest) else "")
        return {
            "ok": True,
            "restaurados": restaurados,
            "message": (
                f"✅ Se corrigieron y restauraron a su valor original {restaurados} alumno(s) en Google Sheets y en el sistema: {resumen}."
            ),
        }

    def sumar_lote_practica_individual(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Suma un lote individual de horas de Práctica (ej. +2h, +4h, +8h o completar al 100%)
        a un alumno específico, registrando la Fecha de Cargue (hoy, ayer, mañana, etc.).
        Si con el lote completa Teoría + Taller + Práctica y marcar_cargado=True, lo marca ✅ CARGADO RUNT.
        """
        fila_idx = _limpiar_entero(payload.get("fila", 0), 0)
        horas_sumar = _limpiar_entero(payload.get("horas_sumar", 8), 8)
        completar_todo = bool(payload.get("completar_todo", False))
        fecha_reporte = _excel_serial_a_iso(payload.get("fecha_cargue", "")) or datetime.datetime.now(clases_core.TZ_COLOMBIA).strftime("%Y-%m-%d")

        practica_list = self.listar_practica_runt(forzar=False)
        item = next((x for x in practica_list if x["fila"] == fila_idx), None)
        if not item:
            return {"ok": False, "message": f"No se encontró la fila #{fila_idx} en la pestaña Practica."}

        teo = item["teoria"] if item["teoria"] > 0 else item["meta_teoria"]
        tal = item["taller"] if item["taller"] > 0 else item["meta_taller"]
        if completar_todo:
            teo = item["meta_teoria"]
            tal = item["meta_taller"]
            nueva_prac = item["meta_practica"]
        else:
            nueva_prac = min(item["meta_practica"], item["practica"] + max(1, horas_sumar))

        llego_al_100 = (teo >= item["meta_teoria"] and tal >= item["meta_taller"] and nueva_prac >= item["meta_practica"])
        nuevo_estado = "CARGADO RUNT" if (completar_todo or (llego_al_100 and payload.get("auto_certificar", False))) else item["estado_runt"]
        nueva_prioridad = "Termino" if nueva_prac >= item["meta_practica"] else "EN PRACTICA"

        res = self.guardar_o_actualizar_practica_runt({
            "fila": item["fila"],
            "nombre": item["nombre"],
            "cedula": item["cedula"],
            "categoria": item["categoria"],
            "fecha_matricula": item["fecha_matricula"],
            "prioridad": nueva_prioridad,
            "vence_90_dias": item["vence_90_dias"],
            "fecha_registro_runt": item["fecha_registro_runt"],
            "teoria": teo,
            "taller": tal,
            "practica": nueva_prac,
            "instructor": item["instructor"],
            "estado_runt": nuevo_estado,
            "fecha_cargue": fecha_reporte,
            "cargue_diario": f"PRACTICA (+{nueva_prac - item['practica']}h)",
        })
        return res

    def reiniciar_curso_vencido_con_sancion(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """
        Cuando a un alumno se le vencen los 90 días del curso en RUNT:
        1. Reinicia su Fecha de Matrícula / Inicio RUNT a la nueva fecha indicada (por defecto Hoy),
           recalculando un nuevo plazo de 90 días (✅ VIGENTE).
        2. Reinicia el conteo de horas reportadas en RUNT a 0 (o conserva Teoría si se indica) para
           volver a subirlas desde cero en la plataforma RUNT.
        3. Registra automáticamente la Sanción Económica por Vencimiento de Curso (90 días) en la hoja
           de Inasistencias / Multas para control de caja y cobro al alumno.
        """
        fila_idx = _limpiar_entero(payload.get("fila", 0), 0)
        valor_sancion = _limpiar_entero(payload.get("valor_sancion", 50000), 50000)
        nueva_fecha_inicio = _excel_serial_a_iso(payload.get("nueva_fecha", "")) or datetime.datetime.now(clases_core.TZ_COLOMBIA).strftime("%Y-%m-%d")
        conservar_teoria = bool(payload.get("conservar_teoria", False))

        practica_list = self.listar_practica_runt(forzar=False)
        item = next((x for x in practica_list if x["fila"] == fila_idx), None)
        if not item:
            return {"ok": False, "message": f"No se encontró la fila #{fila_idx} en la pestaña Practica."}

        d0 = datetime.datetime.strptime(nueva_fecha_inicio, "%Y-%m-%d").date()
        nuevo_vence_90 = (d0 + datetime.timedelta(days=90)).strftime("%Y-%m-%d")

        res_runt = self.guardar_o_actualizar_practica_runt({
            "fila": item["fila"],
            "nombre": item["nombre"],
            "cedula": item["cedula"],
            "categoria": item["categoria"],
            "fecha_matricula": nueva_fecha_inicio,
            "prioridad": "EN PRACTICA",
            "vence_90_dias": nuevo_vence_90,
            "fecha_registro_runt": nueva_fecha_inicio,
            "teoria": item["teoria"] if conservar_teoria else 0,
            "taller": 0,
            "practica": 0,
            "instructor": f"REINICIO RUNT (SANCIÓN ${valor_sancion:,})",
            "estado_runt": "PENDIENTE RUNT",
            "fecha_cargue": nueva_fecha_inicio,
            "cargue_diario": f"REINICIO SANCIÓN ${valor_sancion:,}",
        })

        # Registrar también el cobro de la sanción económica en la hoja de Inasistencias / Multas
        consec_sancion = ""
        try:
            res_in = clases_core.sheets_service.registrar_inasistencia(
                documento_alumno=item["cedula"],
                nombre_alumno=item["nombre"],
                categoria=item["categoria"],
                clases_faltadas="SANCIÓN VENCIMIENTO 90 DÍAS RUNT (REINICIO CURSO)",
                nombre_instructor="CONTROL RUNT / ADMINISTRACIÓN",
                no_horas=1,
                valor_hora=valor_sancion,
                estado="Pendiente Pago",
                fecha=d0.strftime("%d/%m/%Y"),
                hora="REINICIO RUNT",
            )
            consec_sancion = str(res_in.get("Consecutivo", ""))
        except Exception as exc_s:
            logger.warning("Aviso registrando sanción económica por vencimiento RUNT: %s", exc_s)

        return {
            "ok": True,
            "fila": item["fila"],
            "message": (
                f"🔄 Curso RUNT reiniciado para {item['nombre']} ({item['categoria']}): "
                f"Nueva vigencia de 90 días hasta {nuevo_vence_90} (Horas reiniciadas a 0 para nuevo cargue RUNT) "
                f"y Sanción Económica por ${valor_sancion:,} generada en Control de Pagos"
                + (f" (Recibo/Multa #{consec_sancion})." if consec_sancion else ".")
            ),
            "detalle": res_runt,
        }

    def obtener_panorama_runt_web(self, forzar: bool = False) -> Dict[str, Any]:
        practica = self.listar_practica_runt(forzar=forzar)
        no_runt = self.listar_no_runt()
        inicio_practica = self.listar_inicio_practica()
        codigos_runt = self.listar_codigos_runt()

        cargados_cnt = sum(1 for x in practica if x["estado_runt"] == "CARGADO RUNT")
        pendientes_cnt = len(practica) - cargados_cnt
        en_bot_prioridad_cnt = sum(1 for x in practica if x.get("es_prioridad_bot"))
        terminaron_sin_cargue_cnt = sum(1 for x in practica if x["es_terminado_sin_cargue"])
        urgentes_en_practica_cnt = sum(1 for x in practica if x["es_urgente_en_practica"])
        moto_sin_25_teoria_cnt = sum(1 for x in practica if x["estado_runt"] != "CARGADO RUNT" and x["categoria"] in ("A1", "A2") and x["teoria"] < x["meta_teoria"])
        moto_sin_3_taller_cnt = sum(1 for x in practica if x["estado_runt"] != "CARGADO RUNT" and x["categoria"] in ("A1", "A2") and x["taller"] < x["meta_taller"])
        carro_sin_teoria_taller_cnt = sum(1 for x in practica if x["estado_runt"] != "CARGADO RUNT" and x["categoria"] in ("B1", "C1", "RECAT_C1") and (x["teoria"] < x["meta_teoria"] or x["taller"] < x["meta_taller"]))
        por_vencer_cnt = sum(1 for x in practica if x["semaforo_90"] in ("POR_VENCER", "VENCIDO") and x["estado_runt"] != "CARGADO RUNT")
        vencidos_sancion_cnt = sum(1 for x in practica if x["semaforo_90"] == "VENCIDO" and x["estado_runt"] != "CARGADO RUNT")
        con_desfase_bot_cnt = sum(1 for x in practica if x["pendiente_subir_bot"] > 0 and x["estado_runt"] != "CARGADO RUNT")
        sin_registro_no_runt_cnt = sum(1 for x in no_runt if "SIN" in x["estado_registro"].upper())

        return {
            "ok": True,
            "sheet_id": self.sheet_id,
            "sheet_url": f"https://docs.google.com/spreadsheets/d/{self.sheet_id}/edit",
            "conectado_en_vivo": self.conectado_en_vivo,
            "ultimo_error": self._ultimo_error_msg,
            "kpis": {
                "total_practica": len(practica),
                "cargados_runt": cargados_cnt,
                "pendientes_runt": pendientes_cnt,
                "en_bot_prioridad": en_bot_prioridad_cnt,
                "terminaron_sin_cargue": terminaron_sin_cargue_cnt,
                "urgentes_en_practica": urgentes_en_practica_cnt,
                "moto_sin_25_teoria": moto_sin_25_teoria_cnt,
                "moto_sin_3_taller": moto_sin_3_taller_cnt,
                "carro_sin_teoria_taller": carro_sin_teoria_taller_cnt,
                "por_vencer_o_vencidos": por_vencer_cnt,
                "vencidos_sancion": vencidos_sancion_cnt,
                "con_clases_bot_por_subir": con_desfase_bot_cnt,
                "no_runt_pendientes": sin_registro_no_runt_cnt,
                "total_no_runt": len(no_runt),
                "total_inicio_practica": len(inicio_practica),
                "total_codigos_runt": len(codigos_runt),
            },
            "practica": practica,
            "no_runt": no_runt,
            "inicio_practica": inicio_practica,
            "codigos_runt": codigos_runt,
        }


runt_service = RuntControlService()
