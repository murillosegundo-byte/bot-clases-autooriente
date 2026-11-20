"""
=============================================================================
MOTOR DE AUDITORÍA CONTABLE, CONSOLIDACIÓN MULTI-CURSO Y CONTROL COMERCIAL
CEA AUTO ORIENTE — Autor: Segundo Murillo
=============================================================================
Incluye:
1. Lector y normalizador nativo de archivos Excel (.xlsx) y CSV:
   - Soporta el archivo maestro de Matrículas (`matriculas.xlsx` / `convenios.xlsx`).
   - Soporta opcionalmente un segundo archivo de Recibos de Caja (`ingresos.xlsx`)
     para cruzar recibos con o sin número de matrícula (pagos huérfanos por cédula).
2. Motor de Consolidación por Cliente / Cédula (231+ alumnos Multi-Curso):
   - Agrupa todas las matrículas hermanas de una misma cédula (ej. Combo A2 + B1 o C1).
   - Calcula Total Contratado Cliente, Total Abonado Cliente y Saldo Real Cliente.
   - Compensa automáticamente sobrecobros (saldos negativos en una matrícula) y
     distingue entre:
     * 🚨 RIESGO CERO REAL ($0 Abono Global en todas sus matrículas -> Suspensión inmediata).
     * 🔄 COMBO MULTI-CURSO ($0 en la 2da matrícula, pero ya abonó en la 1ra matrícula).
3. Parametrización editable de Tarifas Base por Categoría (A1, A2, B1, C1, C2).
4. Generador de Excel Multi-Pestaña (.xlsx) estilizado (9 hojas ejecutivas) y PDF Ejecutivo.
5. Despacho de Alertas por Correo (SMTP) y Bot de Telegram.
"""

import csv
import datetime
import io
import json
import os
import smtplib
import urllib.request
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from xml.sax.saxutils import escape as xml_escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image as RLImage,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

BASE_DIR = Path(__file__).resolve().parent

TARIFAS_BASE_DEFAULT: Dict[str, float] = {
    "A1": 800000.0,
    "A2": 1000000.0,
    "B1": 1250000.0,
    "C1": 1450000.0,
    "C2": 1600000.0,
}

CONFIG_DEFAULT = {
    "META_MENSUAL_TOTAL": int(os.getenv("META_MENSUAL_TOTAL", "100")),
    "META_MENSUAL_JAVIER_TORRES": int(os.getenv("META_MENSUAL_JAVIER_TORRES", "30")),
    "DIAS_ALERTA_MORA_TEMPRANA": int(os.getenv("DIAS_ALERTA_MORA_TEMPRANA", "30")),
    "DIAS_ALERTA_MORA_MEDIA": int(os.getenv("DIAS_ALERTA_MORA_MEDIA", "60")),
    "DIAS_ALERTA_MORA_CRITICA": int(os.getenv("DIAS_ALERTA_MORA_CRITICA", "90")),
    "TARIFAS_BASE": dict(TARIFAS_BASE_DEFAULT),
    "SMTP_SERVER": os.getenv("SMTP_SERVER", "smtp.gmail.com"),
    "SMTP_PORT": int(os.getenv("SMTP_PORT", "587")),
    "EMAIL_REMITENTE": os.getenv("EMAIL_REMITENTE", "tu_correo@gmail.com"),
    "EMAIL_PASSWORD": os.getenv("EMAIL_PASSWORD", ""),
    "EMAIL_DESTINATARIOS": [
        e.strip()
        for e in os.getenv(
            "EMAIL_DESTINATARIOS",
            "gerencia@autooriente.com,contabilidad@autooriente.com",
        ).split(",")
        if e.strip()
    ],
    "TELEGRAM_BOT_TOKEN": os.getenv("TELEGRAM_BOT_TOKEN", ""),
    "TELEGRAM_CHAT_ID": os.getenv("ADMIN_CHAT_ID", ""),
}

MESES_ES = {
    "01": "Enero",
    "02": "Febrero",
    "03": "Marzo",
    "04": "Abril",
    "05": "Mayo",
    "06": "Junio",
    "07": "Julio",
    "08": "Agosto",
    "09": "Septiembre",
    "10": "Octubre",
    "11": "Noviembre",
    "12": "Diciembre",
}


def _to_float(val: Any) -> float:
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace("$", "").replace(" ", "")
    if not s:
        return 0.0
    if s.count(".") > 1:
        s = s.replace(".", "")
    elif s.count(",") > 1:
        s = s.replace(",", "")
    elif "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts[-1]) == 3 and len(parts[0]) >= 1:
            s = s.replace(",", "")
        else:
            s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _parse_date(val: Any) -> Optional[datetime.datetime]:
    if not val:
        return None
    if isinstance(val, datetime.datetime):
        return val
    if isinstance(val, datetime.date):
        return datetime.datetime(val.year, val.month, val.day)
    s = str(val).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d", "%d/%m/%y"):
        try:
            return datetime.datetime.strptime(s[:10], fmt)
        except ValueError:
            continue
    try:
        serial = float(s)
        if 30000 < serial < 60000:
            return datetime.datetime(1899, 12, 30) + datetime.timedelta(days=serial)
    except ValueError:
        pass
    return None


def _norm_header(h: Any) -> str:
    """Normaliza encabezados quitando tildes, espacios, puntos, guiones y caracteres especiales."""
    import unicodedata

    s = str(h or "").strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    for ch in (" ", "_", "-", ".", "/", "(", ")", "°", "º", "#", ":"):
        s = s.replace(ch, "")
    return s


def _leer_filas_html_xls(file_bytes: bytes) -> List[List[str]]:
    """Soporta archivos .xls/.xlsx que en realidad son tablas HTML o XML SpreadsheetML 2003 exportadas por software web."""
    import html as html_lib
    import re

    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            text = file_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = file_bytes.decode("utf-8", errors="ignore")

    rows_out: List[List[str]] = []
    # Intentar primero filas HTML <tr...>...</tr> o XML <Row...>...<Cell><Data>...</Data></Cell></Row>
    tr_matches = re.findall(r"<(?:tr|row)\b[^>]*>(.*?)</(?:tr|row)>", text, flags=re.IGNORECASE | re.DOTALL)
    for tr in tr_matches:
        cells = re.findall(r"<(?:td|th|data|cell)\b[^>]*>(.*?)</(?:td|th|data|cell)>", tr, flags=re.IGNORECASE | re.DOTALL)
        if cells:
            clean_row = []
            for c in cells:
                c_txt = re.sub(r"<[^>]+>", "", c)
                c_txt = html_lib.unescape(c_txt).strip()
                clean_row.append(c_txt)
            if any(clean_row):
                rows_out.append(clean_row)
    return rows_out


def _leer_hoja_xml(sheet_xml: bytes, shared_strings: List[str]) -> List[List[str]]:
    root = ET.fromstring(sheet_xml)

    def local_tag(el: ET.Element) -> str:
        return el.tag.split("}")[-1] if "}" in el.tag else el.tag

    def col_to_idx(col_str: str) -> int:
        idx = 0
        for ch in col_str:
            idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
        return idx - 1

    rows_out: List[List[str]] = []
    for row in root.iter():
        if local_tag(row) != "row":
            continue
        row_cells: Dict[int, str] = {}
        max_col = -1
        for c in row:
            if local_tag(c) != "c":
                continue
            ref = c.attrib.get("r", "")
            col_letters = "".join([ch for ch in ref if ch.isalpha()])
            c_idx = col_to_idx(col_letters) if col_letters else (max_col + 1)
            max_col = max(max_col, c_idx)
            t = c.attrib.get("t", "")
            val = ""
            if t == "inlineStr":
                texts = [el.text or "" for el in c.iter() if local_tag(el) == "t"]
                val = "".join(texts)
            else:
                v_el = next((el for el in c if local_tag(el) == "v"), None)
                val = v_el.text if v_el is not None and v_el.text is not None else ""
                if t == "s" and val != "":
                    try:
                        val = shared_strings[int(val)]
                    except Exception:
                        pass
            row_cells[c_idx] = str(val).strip()
        if max_col >= 0:
            fila = [row_cells.get(i, "") for i in range(max_col + 1)]
            rows_out.append(fila)
    return rows_out


def _leer_filas_xlsx_nativo(
    file_bytes: bytes,
    prefer_keywords: Tuple[str, ...] = (
        "NUMEROMATRICULA",
        "MATRICULA",
        "VALORCURSO",
        "VALOR",
        "NUMERODOCUMENTO",
        "IDENTIFICACION",
        "NOMBRETERCERO",
        "COMPROBANTE",
        "CODIGOCONTABLE",
        "CREDITO",
        "DEBITO",
    ),
) -> List[List[str]]:
    # Si no empieza con firma ZIP ('PK'), verificar si es HTML/XML exportado como .xls/.xlsx
    if not file_bytes.startswith(b"PK"):
        head_strip = file_bytes[:512].lstrip().lower()
        if head_strip.startswith((b"<html", b"<!doctype", b"<table", b"<?xml", b"<workbook")):
            filas_html = _leer_filas_html_xls(file_bytes)
            if filas_html:
                return filas_html
        return _leer_filas_csv(file_bytes)

    with zipfile.ZipFile(io.BytesIO(file_bytes), "r") as z:
        shared_strings: List[str] = []
        ss_candidates = [n for n in z.namelist() if n.lower().endswith("sharedstrings.xml")]
        if ss_candidates:
            ss_xml = z.read(ss_candidates[0])
            root_ss = ET.fromstring(ss_xml)
            for si in root_ss.iter():
                tag_si = si.tag.split("}")[-1] if "}" in si.tag else si.tag
                if tag_si == "si":
                    texts = [
                        (t.text or "")
                        for t in si.iter()
                        if (t.tag.split("}")[-1] if "}" in t.tag else t.tag) == "t"
                    ]
                    shared_strings.append("".join(texts))

        sheets = sorted([
            n for n in z.namelist()
            if "worksheets/" in n.lower() and n.lower().endswith(".xml")
        ])
        if not sheets:
            raise ValueError("No se encontraron hojas de cálculo dentro del archivo Excel.")

        first_non_empty: List[List[str]] = []
        for s_name in sheets:
            try:
                rows_sheet = _leer_hoja_xml(z.read(s_name), shared_strings)
            except Exception:
                continue
            if not rows_sheet:
                continue
            if not first_non_empty:
                first_non_empty = rows_sheet
            # Verificar si esta hoja contiene alguno de los encabezados buscados en sus primeras 25 filas
            for fila in rows_sheet[:25]:
                norm_cells = [_norm_header(c) for c in fila]
                if any(k in norm_cells for k in prefer_keywords):
                    return rows_sheet

        return first_non_empty


def _leer_filas_csv(file_bytes: bytes) -> List[List[str]]:
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            text = file_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = file_bytes.decode("utf-8", errors="ignore")

    sample = text[:4096]
    delimiter = ";" if sample.count(";") > sample.count(",") else ","
    reader = csv.reader(io.StringIO(text), delimiter=delimiter)
    return [[str(c).strip() for c in row] for row in reader]


def _extraer_bytes(origen: Any) -> Tuple[bytes, bool]:
    if isinstance(origen, (str, Path)):
        ruta = Path(origen)
        b = ruta.read_bytes()
        return b, ruta.suffix.lower() == ".csv"
    elif isinstance(origen, bytes):
        return origen, not origen.startswith(b"PK")
    elif hasattr(origen, "read"):
        b = origen.read()
        return b, not b.startswith(b"PK")
    raise ValueError("Origen de datos no soportado.")


# =============================================================================
# 1. CAPA DE INGESTIÓN: MATRÍCULAS + RECIBOS DE CAJA (OPCIONAL)
# =============================================================================
def _inferir_categoria_por_monto_contable(monto: float, codigo_contable: str = "", comprobante: str = "") -> str:
    """Infiere la categoría o tipo de servicio a partir del valor facturado en contabilidad (Siigo FV-2 / NC-2)."""
    comp_up = str(comprobante or "").upper()
    if comp_up.startswith("NC") or str(codigo_contable).startswith("4175"):
        return "DEVOLUCION_NC"
    m = abs(monto)
    if m <= 190000:
        return "CERTIFICADO_ADICIONAL"
    if 600000 <= m <= 990000:
        return "A2"
    if 990001 <= m <= 1480000:
        return "B1"
    if m > 1480000:
        return "C1"
    return "A2"


def cargar_recibos_ingresos(origen_ingresos: Any) -> List[Dict[str, Any]]:
    """
    Lee un archivo opcional de recibos de caja (`ingresos.xlsx` o `.csv`) O un Auxiliar Contable
    (`Movimiento auxiliar de tercero por cuenta contable` con `FV-2-...` y `NC-2-...`).
    """
    if not origen_ingresos:
        return []
    file_bytes, is_csv = _extraer_bytes(origen_ingresos)
    filas = _leer_filas_csv(file_bytes) if is_csv else _leer_filas_xlsx_nativo(file_bytes)

    header_idx = -1
    headers: List[str] = []
    es_auxiliar_contable = False
    for idx, fila in enumerate(filas):
        fila_norm = [_norm_header(c) for c in fila]
        if any(k in fila_norm for k in ("CODIGOCONTABLE", "COMPROBANTE", "NOMBRETERCERO", "SALDOMOVIMIENTO")) and any(
            k in fila_norm for k in ("CREDITO", "DEBITO", "IDENTIFICACION")
        ):
            header_idx = idx
            headers = fila_norm
            es_auxiliar_contable = True
            break
        if any(k in fila_norm for k in ("VALOR", "ABONO", "MONTO", "TOTAL", "CREDITO")) and any(
            k in fila_norm for k in ("NUMERODOCUMENTO", "DOCUMENTO", "CEDULA", "NUMEROMATRICULA", "MATRICULA", "IDENTIFICACION")
        ) and not any(k in fila_norm for k in ("VALORCURSO", "VALORPENDIENTE", "CLASESCONTRATADAS")):
            header_idx = idx
            headers = fila_norm
            break

    if header_idx == -1:
        return []

    recibos: List[Dict[str, Any]] = []
    for i, fila in enumerate(filas[header_idx + 1 :], start=1):
        if not any(str(c).strip() for c in fila):
            continue
        raw = {headers[c_i]: (fila[c_i].strip() if c_i < len(fila) else "") for c_i in range(len(headers)) if headers[c_i]}

        if es_auxiliar_contable:
            doc = (raw.get("IDENTIFICACION") or raw.get("NUMERODOCUMENTO") or raw.get("DOCUMENTO") or "").strip()
            tercero = (raw.get("NOMBRETERCERO") or raw.get("TERCERO") or "").strip()
            comprobante = (raw.get("COMPROBANTE") or f"MOV-{i:04d}").strip()
            cod_cta = (raw.get("CODIGOCONTABLE") or "").strip()
            nom_cta = (raw.get("CUENTACONTABLE") or "Servicios").strip()
            debito = _to_float(raw.get("DEBITO", 0))
            credito = _to_float(raw.get("CREDITO", 0))
            if not doc and not tercero and debito == 0 and credito == 0:
                continue
            if "PROCESADOEN" in _norm_header(doc):
                continue
            es_nc = comprobante.upper().startswith("NC") or cod_cta.startswith("4175") or (debito > 0 and credito == 0)
            val_neto = -debito if es_nc else credito
            fecha_dt = _parse_date(raw.get("FECHAELABORACION") or raw.get("FECHA") or "")
            fecha_str = fecha_dt.strftime("%Y-%m-%d") if fecha_dt else str(raw.get("FECHAELABORACION", "")).strip()
            recibos.append({
                "RECIBO": comprobante,
                "NUMERODOCUMENTO": doc,
                "TERCERO": tercero,
                "NUMEROMATRICULA": "",
                "VALOR": val_neto,
                "DEBITO": debito,
                "CREDITO": credito,
                "CODIGOCONTABLE": cod_cta,
                "FECHA": fecha_str,
                "TIPOPAGO": "NOTA CRÉDITO / DEVOLUCIÓN" if es_nc else f"FACTURA ELECTRÓNICA ({cod_cta})",
                "OBSERVACION": f"{nom_cta} ({comprobante}) - {tercero}",
                "ESTADO": "DEVOLUCION" if es_nc else "FACTURADO DIAN",
                "HUERFANO": True,
                "ES_CONTABLE": True,
            })
            continue

        estado = str(raw.get("ESTADO", "ACTIVO")).strip().upper()
        if estado in ("ANULADO", "ANULADA", "CANCELADO", "RECHAZADO"):
            continue

        doc = (
            raw.get("NUMERODOCUMENTO")
            or raw.get("DOCUMENTO")
            or raw.get("CEDULA")
            or raw.get("IDENTIFICACION")
            or ""
        ).strip()
        mat = (raw.get("NUMEROMATRICULA") or raw.get("MATRICULA") or "").strip()
        if mat.upper() in ("NAN", "NONE", "NULL", "0", "-"):
            mat = ""

        val = _to_float(raw.get("VALOR") or raw.get("MONTO") or raw.get("ABONO") or raw.get("TOTAL") or 0)
        fecha_dt = _parse_date(raw.get("FECHA") or raw.get("FECHAPAGO") or raw.get("FECHARECIBO") or "")
        fecha_str = fecha_dt.strftime("%Y-%m-%d") if fecha_dt else str(raw.get("FECHA", "")).strip()
        num_recibo = (raw.get("NUMERORECIBO") or raw.get("RECIBO") or raw.get("ID") or f"RC-{i:04d}").strip()
        tipo_pago = (raw.get("TIPOPAGO") or raw.get("MEDIOPAGO") or raw.get("FORMA_PAGO") or "CAJA").strip().upper()
        obs = (raw.get("OBSERVACION") or raw.get("OBSERVACIONES") or raw.get("CONCEPTO") or "").strip()

        if not doc and not mat and val == 0:
            continue

        recibos.append({
            "RECIBO": num_recibo,
            "NUMERODOCUMENTO": doc,
            "NUMEROMATRICULA": mat,
            "VALOR": val,
            "FECHA": fecha_str,
            "TIPOPAGO": tipo_pago,
            "OBSERVACION": obs,
            "ESTADO": estado or "APROBADO",
            "HUERFANO": bool(doc and not mat),
        })
    return recibos


def cargar_y_limpiar_datos(
    origen: Any,
    fecha_corte: Optional[datetime.datetime] = None,
    origen_ingresos: Any = None,
    tarifas_base: Optional[Dict[str, float]] = None,
) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """
    Motor Flexible Multi-Formato:
    1. Soporta Reporte Oficial de Matrículas / Convenios SISEC (`NUMEROMATRICULA`, `VALORCURSO`, `ABONOS`, etc.).
    2. Soporta Movimiento Auxiliar Contable Siigo/DIAN (`Identificación`, `Nombre tercero`, `Código contable`, `Comprobante` FV-2/NC-2, `Débito`, `Crédito`, `Saldo Movimiento`).
    3. Soporta cualquier otro archivo Excel/CSV administrativo o contable detectando dinámicamente sus columnas.
    """
    file_bytes, is_csv = _extraer_bytes(origen)
    filas = _leer_filas_csv(file_bytes) if is_csv else _leer_filas_xlsx_nativo(file_bytes)
    tarifas = {**TARIFAS_BASE_DEFAULT, **(tarifas_base or {})}

    metadata: Dict[str, Any] = {
        "sede": "CEA AUTO ORIENTE",
        "tipo_archivo_detectado": "MATRICULAS_SISEC",
        "fecha_inicial": "",
        "fecha_final": "",
        "registros_declarados": 0,
        "recibos_externos_cargados": 0,
        "recibos_huerfanos_cruzados": 0,
        "contabilidad_resumen": None,
    }

    header_idx = -1
    headers: List[str] = []
    es_auxiliar_contable = False

    alias_map = {
        "MATRICULA": "NUMEROMATRICULA",
        "NROMATRICULA": "NUMEROMATRICULA",
        "NOMATRICULA": "NUMEROMATRICULA",
        "NUMMATRICULA": "NUMEROMATRICULA",
        "IDMATRICULA": "NUMEROMATRICULA",
        "COMPROBANTE": "NUMEROMATRICULA",
        "FACTURA": "NUMEROMATRICULA",
        "CONSECUTIVO": "NUMEROMATRICULA",
        "DOCUMENTO": "NUMERODOCUMENTO",
        "CEDULA": "NUMERODOCUMENTO",
        "NRODOCUMENTO": "NUMERODOCUMENTO",
        "NODOCUMENTO": "NUMERODOCUMENTO",
        "IDENTIFICACION": "NUMERODOCUMENTO",
        "NIT": "NUMERODOCUMENTO",
        "TERCEROID": "NUMERODOCUMENTO",
        "NOMBRE": "NOMBRES",
        "NOMBRETERCERO": "NOMBRES",
        "TERCERO": "NOMBRES",
        "RAZONSOCIAL": "NOMBRES",
        "CLIENTE": "NOMBRES",
        "APELLIDO": "APELLIDOS",
        "ALUMNO": "NOMBRES",
        "ESTUDIANTE": "NOMBRES",
        "VALOR": "VALORCURSO",
        "VALORTOTAL": "VALORCURSO",
        "PRECIO": "VALORCURSO",
        "MONTO": "VALORCURSO",
        "TOTAL": "VALORCURSO",
        "ABONO": "ABONOS",
        "TOTALABONADO": "ABONOS",
        "PAGADO": "ABONOS",
        "SALDO": "VALORPENDIENTE",
        "PENDIENTE": "VALORPENDIENTE",
        "SALDOPENDIENTE": "VALORPENDIENTE",
        "FECHAMATRICULA": "FECHAINSCRIPCION",
        "FECHAELABORACION": "FECHAINSCRIPCION",
        "FECHAFACTURA": "FECHAINSCRIPCION",
        "FECHAMOVIMIENTO": "FECHAINSCRIPCION",
        "FECHA": "FECHAINSCRIPCION",
    }

    for idx, fila in enumerate(filas):
        fila_upper = [str(c).strip().upper() for c in fila]
        fila_norm = [_norm_header(c) for c in fila]
        texto_fila = " ".join(c for c in fila_upper if c)
        if "MOVIMIENTO AUXILIAR" in texto_fila or "CUENTA CONTABLE" in texto_fila:
            metadata["sede"] = "CEA AUTO ORIENTE SAS (CONTABILIDAD 901517095-8)"
        if "DE ENERO" in texto_fila or ("DE " in texto_fila and " A " in texto_fila and "202" in texto_fila):
            metadata["fecha_inicial"] = texto_fila[:45]

        for i, cell in enumerate(fila_upper):
            if "SEDE:" in cell and i + 1 < len(fila):
                metadata["sede"] = fila[i + 1].strip() or "CEA"
            elif "FECHA INICIAL:" in cell and i + 1 < len(fila):
                metadata["fecha_inicial"] = fila[i + 1].strip()
            elif "FECHA FINAL:" in cell and i + 1 < len(fila):
                metadata["fecha_final"] = fila[i + 1].strip()
            elif "CANTIDAD DE REGISTROS:" in cell and i + 1 < len(fila):
                metadata["registros_declarados"] = int(_to_float(fila[i + 1]))

        # 1) Detectar Auxiliar Contable (Siigo / DIAN)
        if any(c in ("CODIGOCONTABLE", "CUENTACONTABLE", "SALDOMOVIMIENTO") for c in fila_norm) and any(
            c in ("IDENTIFICACION", "NOMBRETERCERO", "COMPROBANTE", "CREDITO", "DEBITO") for c in fila_norm
        ):
            header_idx = idx
            headers = [alias_map.get(c, c) for c in fila_norm]
            es_auxiliar_contable = True
            metadata["tipo_archivo_detectado"] = "AUXILIAR_CONTABLE_SIIGO"
            break

        # 2) Detectar Matrículas / Convenios o tabla con Documento + Nombre
        tiene_mat = any(c in ("NUMEROMATRICULA", "MATRICULA", "NROMATRICULA", "NOMATRICULA", "COMPROBANTE") for c in fila_norm)
        tiene_doc = any(c in ("NUMERODOCUMENTO", "DOCUMENTO", "CEDULA", "IDENTIFICACION", "NIT") for c in fila_norm)
        tiene_nom = any(c in ("NOMBRES", "NOMBRE", "ALUMNO", "ESTUDIANTE", "NOMBRETERCERO", "TERCERO", "CLIENTE", "RAZONSOCIAL") for c in fila_norm)
        if tiene_mat or (tiene_doc and tiene_nom):
            header_idx = idx
            headers = [alias_map.get(c, c) for c in fila_norm]
            break

    # 3) Fallback Universal Flexible: si ningún encabezado estándar coincidió, tomar la primera fila con >= 3 columnas con texto
    if header_idx == -1:
        for idx, fila in enumerate(filas[:35]):
            non_empty = [c for c in fila if str(c).strip()]
            if len(non_empty) >= 3:
                header_idx = idx
                headers = [alias_map.get(_norm_header(c), _norm_header(c) or f"COL_{i}") for i, c in enumerate(fila)]
                metadata["tipo_archivo_detectado"] = "TABLA_FLEXIBLE_UNIVERSAL"
                break

    if header_idx == -1:
        raise ValueError("El archivo subido está vacío o no contiene filas legibles.")

    if fecha_corte is None:
        dt_final = _parse_date(metadata.get("fecha_final"))
        fecha_corte = dt_final if dt_final else datetime.datetime.now()

    registros: List[Dict[str, Any]] = []

    if es_auxiliar_contable:
        total_credito_4180 = 0.0
        total_debito_4175 = 0.0
        total_facturas_fv = 0
        total_notas_nc = 0
        fechas_vistas: List[str] = []

        for idx_r, fila in enumerate(filas[header_idx + 1 :], start=1):
            if not any(str(c).strip() for c in fila):
                continue
            raw = {}
            for c_i, col_name in enumerate(headers):
                if not col_name:
                    continue
                raw[col_name] = fila[c_i].strip() if c_i < len(fila) else ""

            num_doc = str(raw.get("NUMERODOCUMENTO", "")).strip()
            nombres = str(raw.get("NOMBRES", "")).strip()
            comprobante = str(raw.get("NUMEROMATRICULA", "")).strip()
            if not num_doc and not nombres:
                continue
            if "PROCESADO" in num_doc.upper() or "MOVIMIENTO" in num_doc.upper():
                continue

            cod_contable = str(raw.get("CODIGOCONTABLE", "41800101")).strip()
            cta_contable = str(raw.get("CUENTACONTABLE", "Servicios")).strip()
            debito = _to_float(raw.get("DEBITO", 0))
            credito = _to_float(raw.get("CREDITO", 0))
            saldo_mov = _to_float(raw.get("SALDOMOVIMIENTO", 0))

            es_nc = comprobante.upper().startswith("NC") or cod_contable.startswith("4175") or (debito > 0 and credito == 0)
            if es_nc:
                total_debito_4175 += debito
                total_notas_nc += 1
            else:
                total_credito_4180 += credito
                total_facturas_fv += 1

            cat_inf = _inferir_categoria_por_monto_contable(credito if not es_nc else debito, cod_contable, comprobante)
            tarifa_estandar = tarifas.get(cat_inf, credito if credito > 0 else debito)

            # En Auxiliar Contable:
            # - Para Factura FV-2 (41800101): Crédito es lo facturado y pagado/causado.
            # - Para Nota Crédito NC-2 (41750502): Débito es devolución en servicios.
            if es_nc:
                valor_curso = debito
                abonos = 0.0
                valor_pendiente = debito
                convenio_norm = f"DEVOLUCIÓN NC ({cod_contable})"
            else:
                valor_curso = credito
                abonos = credito
                valor_pendiente = 0.0
                convenio_norm = f"FACTURACIÓN DIAN ({cod_contable})"

            fecha_insc_dt = _parse_date(raw.get("FECHAINSCRIPCION", ""))
            fecha_insc_str = fecha_insc_dt.strftime("%Y-%m-%d") if fecha_insc_dt else ""
            if fecha_insc_str:
                fechas_vistas.append(fecha_insc_str)
            mes_inscripcion = fecha_insc_dt.strftime("%Y-%m") if fecha_insc_dt else "SIN_FECHA"
            dias_antiguedad = max(0, (fecha_corte - fecha_insc_dt).days) if fecha_insc_dt else 0

            reg = {
                "NUMEROMATRICULA": comprobante or f"MOV-{idx_r:04d}",
                "NUMERODOCUMENTO": num_doc,
                "TIPODOCUMENTO": "CC/NIT",
                "NOMBRES": nombres,
                "APELLIDOS": "",
                "ALUMNO": nombres,
                "CONVENIO": cta_contable.upper(),
                "CONVENIO_NORM": convenio_norm,
                "CONVENIO_VACIO": False,
                "PUBLICIDAD": cod_contable,
                "CATEGORIA": cat_inf,
                "FECHAINSCRIPCION": fecha_insc_str,
                "FECHAINSCRIPCION_RAW": str(raw.get("FECHAINSCRIPCION", "")).strip(),
                "MES_INSCRIPCION": mes_inscripcion,
                "VALORHORA": 0.0,
                "CLASESCONTRATADAS": 0,
                "CLASESRECIBIDAS": 0,
                "CLASESPENDIENTES": 0,
                "VALORCURSO": valor_curso,
                "TARIFA_BASE_CAT": tarifa_estandar,
                "ABONOS_ORIGINAL": abonos,
                "ABONOS": abonos,
                "VALORPENDIENTE_ORIGINAL": valor_pendiente,
                "VALORPENDIENTE": valor_pendiente,
                "SALDO_CALCULADO": valor_curso - abonos,
                "ERROR_CUADRE": 0.0,
                "DIAS_ANTIGUEDAD": dias_antiguedad,
                "ESTADO": "NOTA_CREDITO" if es_nc else "FACTURADO_DIAN",
                "FECHAEXAMENMEDICO": "",
                "CODIGOCONTABLE": cod_contable,
                "DEBITO": debito,
                "CREDITO": credito,
                "SALDO_MOVIMIENTO_CONTABLE": saldo_mov,
                "RECIBOS_CAJA": [{
                    "RECIBO": comprobante,
                    "NUMERODOCUMENTO": num_doc,
                    "NUMEROMATRICULA": comprobante,
                    "VALOR": -debito if es_nc else credito,
                    "FECHA": fecha_insc_str,
                    "TIPOPAGO": f"{cod_contable} - {cta_contable}",
                    "OBSERVACION": f"{'Devolución Nota Crédito' if es_nc else 'Factura Electrónica de Venta'} ({comprobante})",
                    "ESTADO": "DEVOLUCIÓN NC" if es_nc else "CONTABILIZADO",
                    "HUERFANO": False,
                }],
            }
            registros.append(reg)

        if fechas_vistas:
            metadata["fecha_inicial"] = min(fechas_vistas)
            metadata["fecha_final"] = max(fechas_vistas)
        metadata["contabilidad_resumen"] = {
            "es_contabilidad": True,
            "total_credito_servicios_4180": round(total_credito_4180, 2),
            "total_debito_devoluciones_4175": round(total_debito_4175, 2),
            "ingreso_neto_contable": round(total_credito_4180 - total_debito_4175, 2),
            "total_facturas_fv": total_facturas_fv,
            "total_notas_credito_nc": total_notas_nc,
        }
        metadata["fecha_corte_usada"] = fecha_corte.strftime("%Y-%m-%d")
        metadata["total_procesados"] = len(registros)
        return registros, metadata

    for idx_r, fila in enumerate(filas[header_idx + 1 :], start=1):
        if not any(str(c).strip() for c in fila):
            continue
        raw = {}
        for c_i, col_name in enumerate(headers):
            if not col_name:
                continue
            raw[col_name] = fila[c_i].strip() if c_i < len(fila) else ""

        num_mat = str(raw.get("NUMEROMATRICULA", "")).strip()
        nombres = str(raw.get("NOMBRES", "")).strip()
        apellidos = str(raw.get("APELLIDOS", "")).strip()
        num_doc = str(raw.get("NUMERODOCUMENTO", "")).strip()

        # Fallback flexible si la tabla tiene columnas con otros nombres
        if not num_doc:
            for k_col, v_col in raw.items():
                if any(w in k_col for w in ("DOC", "CED", "IDENT", "NIT", "CC")) and v_col:
                    num_doc = v_col
                    break
        if not nombres:
            for k_col, v_col in raw.items():
                if any(w in k_col for w in ("NOM", "TERCERO", "CLIEN", "ALUM", "ESTUD", "DESCRIP")) and v_col:
                    nombres = v_col
                    break
        if not num_mat and (num_doc or nombres):
            num_mat = f"REG-{idx_r:04d}"

        if not num_mat and not nombres and not num_doc:
            continue

        categoria = str(raw.get("CATEGORIA", "")).strip().upper() or "GENERAL"
        valor_curso_raw = _to_float(
            raw.get("VALORCURSO")
            or raw.get("CREDITO")
            or raw.get("DEBITO")
            or raw.get("SALDOMOVIMIENTO")
            or 0
        )
        tarifa_estandar = tarifas.get(categoria, abs(valor_curso_raw) if valor_curso_raw != 0 else 1000000.0)
        valor_curso = abs(valor_curso_raw) if valor_curso_raw != 0 else tarifa_estandar

        abonos = _to_float(raw.get("ABONOS", raw.get("CREDITO", 0)))
        valor_pendiente = _to_float(raw.get("VALORPENDIENTE", 0))
        valor_hora = _to_float(raw.get("VALORHORA", 0))
        clases_contratadas = int(_to_float(raw.get("CLASESCONTRATADAS", 0)))
        clases_recibidas = int(_to_float(raw.get("CLASESRECIBIDAS", 0)))
        clases_pendientes = int(_to_float(raw.get("CLASESPENDIENTES", 0)))

        convenio_original = str(raw.get("CONVENIO", "")).strip().upper()
        convenio_vacio = convenio_original == ""
        convenio_norm = convenio_original if not convenio_vacio else "CEA AUTO ORIENTE"

        fecha_insc_dt = _parse_date(raw.get("FECHAINSCRIPCION", ""))
        fecha_insc_str = fecha_insc_dt.strftime("%Y-%m-%d") if fecha_insc_dt else ""
        mes_inscripcion = fecha_insc_dt.strftime("%Y-%m") if fecha_insc_dt else "SIN_FECHA"
        dias_antiguedad = max(0, (fecha_corte - fecha_insc_dt).days) if fecha_insc_dt else 0

        saldo_calculado = valor_curso - abonos
        error_cuadre = valor_pendiente - saldo_calculado

        reg = {
            "NUMEROMATRICULA": num_mat,
            "NUMERODOCUMENTO": num_doc,
            "TIPODOCUMENTO": str(raw.get("TIPODOCUMENTO", "")).strip(),
            "NOMBRES": nombres,
            "APELLIDOS": apellidos,
            "ALUMNO": f"{nombres} {apellidos}".strip(),
            "CONVENIO": convenio_original,
            "CONVENIO_NORM": convenio_norm,
            "CONVENIO_VACIO": convenio_vacio,
            "PUBLICIDAD": str(raw.get("PUBLICIDAD", "")).strip().upper() or "NO ESPECIFICADO",
            "CATEGORIA": categoria,
            "FECHAINSCRIPCION": fecha_insc_str,
            "FECHAINSCRIPCION_RAW": str(raw.get("FECHAINSCRIPCION", "")).strip(),
            "MES_INSCRIPCION": mes_inscripcion,
            "VALORHORA": valor_hora,
            "CLASESCONTRATADAS": clases_contratadas,
            "CLASESRECIBIDAS": clases_recibidas,
            "CLASESPENDIENTES": clases_pendientes,
            "VALORCURSO": valor_curso,
            "TARIFA_BASE_CAT": tarifa_estandar,
            "ABONOS_ORIGINAL": abonos,
            "ABONOS": abonos,
            "VALORPENDIENTE_ORIGINAL": valor_pendiente,
            "VALORPENDIENTE": valor_pendiente,
            "SALDO_CALCULADO": saldo_calculado,
            "ERROR_CUADRE": error_cuadre,
            "DIAS_ANTIGUEDAD": dias_antiguedad,
            "ESTADO": str(raw.get("ESTADO", "ACTIVA")).strip().upper() or "ACTIVA",
            "FECHAEXAMENMEDICO": str(raw.get("FECHAEXAMENMEDICO", "")).strip(),
            "RECIBOS_CAJA": [],
        }
        registros.append(reg)

    # Si se subió archivo de recibos de caja (ingresos.xlsx), cruzar por matrícula y por cédula
    recibos_ext = cargar_recibos_ingresos(origen_ingresos) if origen_ingresos else []
    if recibos_ext:
        metadata["recibos_externos_cargados"] = len(recibos_ext)
        por_mat = {r["NUMEROMATRICULA"]: r for r in registros if r["NUMEROMATRICULA"]}
        por_doc = defaultdict(list)
        for r in registros:
            if r["NUMERODOCUMENTO"]:
                por_doc[r["NUMERODOCUMENTO"]].append(r)

        # Reiniciar abonos para calcular desde recibos si el archivo de ingresos trae la caja detallada
        tiene_pagos_mat = any(rc["NUMEROMATRICULA"] in por_mat for rc in recibos_ext)
        if tiene_pagos_mat:
            for r in registros:
                r["ABONOS"] = 0.0

        huerfanos_cruzados = 0
        for rc in recibos_ext:
            mat_id = rc["NUMEROMATRICULA"]
            doc_id = rc["NUMERODOCUMENTO"]
            if mat_id and mat_id in por_mat:
                target = por_mat[mat_id]
                target["RECIBOS_CAJA"].append(rc)
                if tiene_pagos_mat:
                    target["ABONOS"] += rc["VALOR"]
            elif doc_id and doc_id in por_doc:
                # Pago Huérfano (sin matrícula o matrícula no encontrada, pero cédula sí existe)
                huerfanos_cruzados += 1
                hermanas = por_doc[doc_id]
                # Acreditar empezando por la matrícula con mayor saldo pendiente
                restante = rc["VALOR"]
                for h in hermanas:
                    h["RECIBOS_CAJA"].append({**rc, "NOTA_CRUCE": "Acreditado por Cédula (Recibo sin matrícula)"})
                    if tiene_pagos_mat and restante > 0:
                        falta = max(0.0, h["VALORCURSO"] - h["ABONOS"])
                        aplicar = min(restante, falta) if falta > 0 else restante
                        h["ABONOS"] += aplicar
                        restante -= aplicar
        metadata["recibos_huerfanos_cruzados"] = huerfanos_cruzados
    else:
        # Generar recibo sintético a partir del abono consolidado para visualización en el Expander
        for r in registros:
            if r["ABONOS"] > 0:
                r["RECIBOS_CAJA"].append({
                    "RECIBO": f"AB-{r['NUMEROMATRICULA']}",
                    "NUMERODOCUMENTO": r["NUMERODOCUMENTO"],
                    "NUMEROMATRICULA": r["NUMEROMATRICULA"],
                    "VALOR": r["ABONOS"],
                    "FECHA": r["FECHAINSCRIPCION"],
                    "TIPOPAGO": "ABONO REGISTRADO EN SISTEMA",
                    "OBSERVACION": f"Abono acumulado matrícula {r['CATEGORIA']}",
                    "ESTADO": "APLICADO",
                    "HUERFANO": False,
                })

    metadata["fecha_corte_usada"] = fecha_corte.strftime("%Y-%m-%d")
    metadata["total_procesados"] = len(registros)
    return registros, metadata


# =============================================================================
# 2. MOTOR DE AUDITORÍA, COMPENSACIÓN MULTI-CURSO Y 4 BANDERAS ROJAS
# =============================================================================
def auditar_riesgos(
    registros: List[Dict[str, Any]],
    metadata: Optional[Dict[str, Any]] = None,
    config: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    cfg = {**CONFIG_DEFAULT, **(config or {})}
    meta_total = int(cfg["META_MENSUAL_TOTAL"])
    meta_jt = int(cfg["META_MENSUAL_JAVIER_TORRES"])

    # -------------------------------------------------------------------------
    # PASO 1: AGRUPACIÓN Y CONSOLIDACIÓN POR CLIENTE / CÉDULA (MULTI-CURSO)
    # -------------------------------------------------------------------------
    por_cedula: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    for r in registros:
        key_doc = r["NUMERODOCUMENTO"] or f"MAT_{r['NUMEROMATRICULA']}"
        por_cedula[key_doc].append(r)

    clientes_consolidados: List[Dict[str, Any]] = []
    clientes_multicurso_count = 0
    matriculas_en_multicurso_count = 0
    compensaciones_realizadas = 0

    for doc, regs_cliente in por_cedula.items():
        # Ordenar matrículas del cliente por fecha de inscripción y número
        regs_cliente.sort(key=lambda x: (x["FECHAINSCRIPCION"] or "9999-99-99", x["NUMEROMATRICULA"]))
        es_multicurso = len(regs_cliente) >= 2
        if es_multicurso:
            clientes_multicurso_count += 1
            matriculas_en_multicurso_count += len(regs_cliente)

        total_contratado_cliente = sum(x["VALORCURSO"] for x in regs_cliente)
        total_pagado_cliente = sum(x["ABONOS"] for x in regs_cliente)
        saldo_real_cliente = total_contratado_cliente - total_pagado_cliente

        categorias_cliente = [x["CATEGORIA"] for x in regs_cliente]
        resumen_cursos = [
            {
                "NUMEROMATRICULA": x["NUMEROMATRICULA"],
                "CATEGORIA": x["CATEGORIA"],
                "FECHAINSCRIPCION": x["FECHAINSCRIPCION"],
                "CONVENIO": x["CONVENIO_NORM"],
                "VALORCURSO": x["VALORCURSO"],
                "ABONOS_ORIGINAL": x["ABONOS_ORIGINAL"],
                "VALORPENDIENTE_ORIGINAL": x["VALORPENDIENTE_ORIGINAL"],
                "ESTADO": x["ESTADO"],
            }
            for x in regs_cliente
        ]
        todos_recibos_cliente = []
        for x in regs_cliente:
            todos_recibos_cliente.extend(x.get("RECIBOS_CAJA", []))

        # Redistribución / Compensación de saldos entre matrículas hermanas:
        # Si una matrícula tiene sobrecobro (ABONOS > VALORCURSO) y otra tiene deuda,
        # el excedente se redistribuye automáticamente.
        excedente_pool = sum(max(0.0, x["ABONOS"] - x["VALORCURSO"]) for x in regs_cliente)
        hubo_compensacion = False
        for x in regs_cliente:
            ab_ajustado = min(x["VALORCURSO"], x["ABONOS"]) if excedente_pool > 0 and len(regs_cliente) > 1 else x["ABONOS"]
            x["ABONO_COMPENSADO"] = ab_ajustado

        if excedente_pool > 0 and len(regs_cliente) > 1:
            for x in regs_cliente:
                falta = max(0.0, x["VALORCURSO"] - x["ABONO_COMPENSADO"])
                if falta > 0 and excedente_pool > 0:
                    aplicar = min(falta, excedente_pool)
                    x["ABONO_COMPENSADO"] += aplicar
                    excedente_pool -= aplicar
                    hubo_compensacion = True
            if excedente_pool > 0:
                regs_cliente[0]["ABONO_COMPENSADO"] += excedente_pool
            if hubo_compensacion:
                compensaciones_realizadas += 1

        # Etiquetar cada matrícula con el contexto consolidado de su Cédula
        for x in regs_cliente:
            x["ES_MULTICURSO"] = es_multicurso
            x["CANTIDAD_CURSOS_CLIENTE"] = len(regs_cliente)
            x["CATEGORIAS_CLIENTE"] = " + ".join(categorias_cliente)
            x["TOTAL_CONTRATADO_CLIENTE"] = total_contratado_cliente
            x["TOTAL_ABONADO_CLIENTE"] = total_pagado_cliente
            x["SALDO_REAL_CLIENTE"] = saldo_real_cliente
            x["SALDO_COMPENSADO"] = x["VALORCURSO"] - x["ABONO_COMPENSADO"]
            x["CURSOS_HERMANOS"] = resumen_cursos
            x["RECIBOS_CLIENTE"] = todos_recibos_cliente

            # Clasificación inteligente del Abono $0:
            if x["ABONOS"] == 0 and x["VALORCURSO"] > 0:
                if total_pagado_cliente == 0:
                    x["TIPO_RIESGO_CERO"] = "RIESGO_CERO_REAL"
                    x["ETIQUETA_MULTICURSO"] = "🚨 $0 Global (Sin abono en ningún curso)"
                else:
                    x["TIPO_RIESGO_CERO"] = "MULTICURSO_CON_ABONO"
                    x["ETIQUETA_MULTICURSO"] = (
                        f"🔄 Combo Multi-Curso ({x['CATEGORIAS_CLIENTE']}) • Abonó ${total_pagado_cliente:,.0f} en otra matrícula"
                    )
            elif es_multicurso:
                x["TIPO_RIESGO_CERO"] = "NORMAL_MULTICURSO"
                x["ETIQUETA_MULTICURSO"] = f"👥 Cliente Multi-Curso ({x['CATEGORIAS_CLIENTE']})"
            else:
                x["TIPO_RIESGO_CERO"] = "CURSO_UNICO"
                x["ETIQUETA_MULTICURSO"] = "Curso Único"

        clientes_consolidados.append({
            "NUMERODOCUMENTO": doc,
            "ALUMNO": regs_cliente[0]["ALUMNO"],
            "CONVENIO_NORM": regs_cliente[0]["CONVENIO_NORM"],
            "CANTIDAD_CURSOS": len(regs_cliente),
            "CATEGORIAS": " + ".join(categorias_cliente),
            "MATRICULAS": ", ".join(x["NUMEROMATRICULA"] for x in regs_cliente),
            "FECHA_PRIMERA_INSC": regs_cliente[0]["FECHAINSCRIPCION"],
            "DIAS_ANTIGUEDAD_MAX": max(x["DIAS_ANTIGUEDAD"] for x in regs_cliente),
            "TOTAL_CONTRATADO": total_contratado_cliente,
            "TOTAL_ABONADO": total_pagado_cliente,
            "SALDO_REAL_CLIENTE": saldo_real_cliente,
            "TIENE_MATRICULA_EN_CERO": any(x["ABONOS"] == 0 and x["VALORCURSO"] > 0 for x in regs_cliente),
            "ABONO_GLOBAL_CERO": total_pagado_cliente == 0 and total_contratado_cliente > 0,
            "CURSOS_DETALLE": resumen_cursos,
            "RECIBOS_DETALLE": todos_recibos_cliente,
        })

    # -------------------------------------------------------------------------
    # PASO 2: CLASIFICACIÓN EN BANDERAS ROJAS Y CONSOLIDADOS
    # -------------------------------------------------------------------------
    abono_cero_todos: List[Dict[str, Any]] = []
    abono_cero_real_global: List[Dict[str, Any]] = []
    abono_cero_falso_multicurso: List[Dict[str, Any]] = []

    mora_temprana_30: List[Dict[str, Any]] = []   # 31 a 60 días
    mora_media_60: List[Dict[str, Any]] = []      # 61 a 90 días
    cartera_critica_90: List[Dict[str, Any]] = []  # > 90 días
    cartera_al_dia: List[Dict[str, Any]] = []      # <= 30 días con saldo > 0

    errores_digitacion: List[Dict[str, Any]] = []
    descuadres: List[Dict[str, Any]] = []

    jt_todos: List[Dict[str, Any]] = []
    jt_cartera: List[Dict[str, Any]] = []
    jt_abono_cero: List[Dict[str, Any]] = []
    jt_abono_cero_real: List[Dict[str, Any]] = []

    por_mes: Dict[str, Dict[str, Any]] = defaultdict(
        lambda: {
            "mes": "",
            "nombre_mes": "",
            "total_matriculas": 0,
            "meta_empresa": meta_total,
            "cumplimiento_empresa_pct": 0.0,
            "matriculas_jt": 0,
            "meta_jt": meta_jt,
            "cumplimiento_jt_pct": 0.0,
            "facturado": 0.0,
            "recaudado": 0.0,
            "pendiente": 0.0,
            "abono_cero_count": 0,
            "abono_cero_real_count": 0,
        }
    )

    por_convenio: Dict[str, Dict[str, Any]] = defaultdict(
        lambda: {
            "convenio": "",
            "matriculas": 0,
            "clientes_unicos": set(),
            "facturado": 0.0,
            "recaudado": 0.0,
            "pendiente": 0.0,
            "con_saldo": 0,
            "abono_cero": 0,
            "abono_cero_real": 0,
            "tasa_recaudo": 0.0,
            "tasa_morosidad": 0.0,
        }
    )

    por_categoria: Dict[str, Dict[str, Any]] = defaultdict(
        lambda: {
            "categoria": "",
            "matriculas": 0,
            "facturado": 0.0,
            "recaudado": 0.0,
            "pendiente": 0.0,
        }
    )

    por_publicidad: Counter = Counter()

    total_facturado = 0.0
    total_abonos = 0.0
    total_pendiente = 0.0
    alumnos_con_saldo = 0
    convenios_vacios_count = 0

    for r in registros:
        vc = r["VALORCURSO"]
        ab = r["ABONOS"]
        vp = r["VALORPENDIENTE"]
        vh = r["VALORHORA"]
        cc = r["CLASESCONTRATADAS"]
        dias = r["DIAS_ANTIGUEDAD"]
        conv = r["CONVENIO_NORM"]
        cat = r["CATEGORIA"]
        mes = r["MES_INSCRIPCION"]

        total_facturado += vc
        total_abonos += ab
        total_pendiente += vp
        if r["CONVENIO_VACIO"]:
            convenios_vacios_count += 1

        if vp > 0:
            alumnos_con_saldo += 1

        # Bandera 1: Matrículas en $0 (Separando Riesgo Real vs Combo Multi-Curso)
        if ab == 0 and vc > 0:
            abono_cero_todos.append(r)
            if r["TIPO_RIESGO_CERO"] == "RIESGO_CERO_REAL":
                abono_cero_real_global.append(r)
            else:
                abono_cero_falso_multicurso.append(r)

        # Bandera 2: Antigüedad de Cartera
        if vp > 0:
            if dias > cfg["DIAS_ALERTA_MORA_CRITICA"]:
                cartera_critica_90.append(r)
            elif dias > cfg["DIAS_ALERTA_MORA_MEDIA"]:
                mora_media_60.append(r)
            elif dias > cfg["DIAS_ALERTA_MORA_TEMPRANA"]:
                mora_temprana_30.append(r)
            else:
                cartera_al_dia.append(r)

        # Bandera 3: Inconsistencias contables y errores de digitación
        motivos_error = []
        if vc < 100000:
            motivos_error.append(f"Valor Curso irrisorio (${vc:,.0f})")
        if vp < 0:
            if r["ES_MULTICURSO"]:
                motivos_error.append(f"Saldo negativo (${vp:,.0f}) en cliente Multi-Curso ({r['CATEGORIAS_CLIENTE']})")
            else:
                motivos_error.append(f"Saldo Pendiente negativo (${vp:,.0f} / sobrecobro)")
        if cc > 60:
            motivos_error.append(f"Clases contratadas atípicas ({cc} clases)")
        if 0 < vh < 10000:
            motivos_error.append(f"Valor Hora atípico (${vh:,.0f})")
        if abs(r["ERROR_CUADRE"]) > 1:
            motivos_error.append(f"Descuadre matemático (${r['ERROR_CUADRE']:,.0f})")
            descuadres.append(r)

        if motivos_error:
            r_err = {**r, "MOTIVO_ERROR": " | ".join(motivos_error)}
            errores_digitacion.append(r_err)

        # Convenio Javier Torres
        if conv == "JAVIER TORRES":
            jt_todos.append(r)
            if vp > 0:
                jt_cartera.append(r)
            if ab == 0 and vc > 0:
                jt_abono_cero.append(r)
                if r["TIPO_RIESGO_CERO"] == "RIESGO_CERO_REAL":
                    jt_abono_cero_real.append(r)

        # Agrupación Mensual (Bandera 4)
        m_item = por_mes[mes]
        m_item["mes"] = mes
        if "-" in mes:
            yyyy, mm = mes.split("-")
            m_item["nombre_mes"] = f"{MESES_ES.get(mm, mm)} {yyyy}"
        else:
            m_item["nombre_mes"] = mes
        m_item["total_matriculas"] += 1
        if conv == "JAVIER TORRES":
            m_item["matriculas_jt"] += 1
        m_item["facturado"] += vc
        m_item["recaudado"] += ab
        m_item["pendiente"] += vp
        if ab == 0 and vc > 0:
            m_item["abono_cero_count"] += 1
            if r["TIPO_RIESGO_CERO"] == "RIESGO_CERO_REAL":
                m_item["abono_cero_real_count"] += 1

        # Agrupación por Convenio
        c_item = por_convenio[conv]
        c_item["convenio"] = conv
        c_item["matriculas"] += 1
        if r["NUMERODOCUMENTO"]:
            c_item["clientes_unicos"].add(r["NUMERODOCUMENTO"])
        c_item["facturado"] += vc
        c_item["recaudado"] += ab
        c_item["pendiente"] += vp
        if vp > 0:
            c_item["con_saldo"] += 1
        if ab == 0 and vc > 0:
            c_item["abono_cero"] += 1
            if r["TIPO_RIESGO_CERO"] == "RIESGO_CERO_REAL":
                c_item["abono_cero_real"] += 1

        # Agrupación por Categoría
        cat_item = por_categoria[cat]
        cat_item["categoria"] = cat
        cat_item["matriculas"] += 1
        cat_item["facturado"] += vc
        cat_item["recaudado"] += ab
        cat_item["pendiente"] += vp

        por_publicidad[r["PUBLICIDAD"]] += 1

    lista_meses = []
    for k in sorted(por_mes.keys()):
        m = por_mes[k]
        m["cumplimiento_empresa_pct"] = round(m["total_matriculas"] / meta_total * 100, 1) if meta_total else 0.0
        m["cumplimiento_jt_pct"] = round(m["matriculas_jt"] / meta_jt * 100, 1) if meta_jt else 0.0
        m["estado_meta_empresa"] = "CUMPLIDA" if m["total_matriculas"] >= meta_total else "BAJO META"
        m["estado_meta_jt"] = "CUMPLIDA" if m["matriculas_jt"] >= meta_jt else "BAJO META"
        lista_meses.append(m)

    lista_convenios = []
    for k, c in sorted(por_convenio.items(), key=lambda x: x[1]["matriculas"], reverse=True):
        c_out = {**c, "clientes_unicos": len(c["clientes_unicos"])}
        c_out["tasa_recaudo"] = round(c["recaudado"] / c["facturado"] * 100, 1) if c["facturado"] > 0 else 0.0
        c_out["tasa_morosidad"] = round(c["pendiente"] / c["facturado"] * 100, 1) if c["facturado"] > 0 else 0.0
        lista_convenios.append(c_out)

    lista_categorias = sorted(por_categoria.values(), key=lambda x: x["matriculas"], reverse=True)

    # Ordenamientos: por defecto de los más recientes a los más antiguos (FECHAINSCRIPCION descendente, luego NUMEROMATRICULA descendente)
    def _key_reciente(x: Dict[str, Any]):
        f_str = str(x.get("FECHAINSCRIPCION", "") or "")
        try:
            mat_num = int("".join(ch for ch in str(x.get("NUMEROMATRICULA", "")) if ch.isdigit()) or 0)
        except Exception:
            mat_num = 0
        return (f_str, mat_num, float(x.get("VALORPENDIENTE", 0) or 0))

    registros.sort(key=_key_reciente, reverse=True)
    jt_todos.sort(key=_key_reciente, reverse=True)
    jt_cartera.sort(key=_key_reciente, reverse=True)
    abono_cero_todos.sort(key=_key_reciente, reverse=True)
    abono_cero_real_global.sort(key=_key_reciente, reverse=True)
    abono_cero_falso_multicurso.sort(key=_key_reciente, reverse=True)
    cartera_critica_90.sort(key=_key_reciente, reverse=True)
    mora_media_60.sort(key=_key_reciente, reverse=True)
    mora_temprana_30.sort(key=_key_reciente, reverse=True)

    clientes_multicurso_lista = [c for c in clientes_consolidados if c["CANTIDAD_CURSOS"] >= 2]
    clientes_multicurso_lista.sort(key=lambda x: (str(x.get("FECHA_PRIMERA_INSC", "") or ""), x["SALDO_REAL_CLIENTE"]), reverse=True)

    tasa_recaudo = round((total_abonos / total_facturado * 100), 2) if total_facturado > 0 else 0.0
    tasa_morosidad = round((total_pendiente / total_facturado * 100), 2) if total_facturado > 0 else 0.0

    num_meses_activos = len([m for m in lista_meses if m["mes"] != "SIN_FECHA"])
    promedio_mensual_emp = round(len(registros) / max(1, num_meses_activos), 1)
    promedio_mensual_jt = round(len(jt_todos) / max(1, num_meses_activos), 1)

    kpis = {
        "total_matriculas": len(registros),
        "total_clientes_unicos": len(clientes_consolidados),
        "clientes_multicurso_count": clientes_multicurso_count,
        "matriculas_en_multicurso_count": matriculas_en_multicurso_count,
        "compensaciones_multicurso": compensaciones_realizadas,
        "total_facturado": total_facturado,
        "total_abonos": total_abonos,
        "total_pendiente": total_pendiente,
        "tasa_recaudo_pct": tasa_recaudo,
        "tasa_morosidad_pct": tasa_morosidad,
        "alumnos_con_saldo": alumnos_con_saldo,
        "alumnos_paz_y_salvo": len(registros) - alumnos_con_saldo,
        "abono_cero_count": len(abono_cero_todos),
        "abono_cero_monto": sum(x["VALORPENDIENTE"] for x in abono_cero_todos),
        "abono_cero_real_count": len(abono_cero_real_global),
        "abono_cero_real_clientes": len(set(x["NUMERODOCUMENTO"] for x in abono_cero_real_global)),
        "abono_cero_real_monto": sum(x["VALORPENDIENTE"] for x in abono_cero_real_global),
        "abono_cero_multicurso_count": len(abono_cero_falso_multicurso),
        "abono_cero_multicurso_monto": sum(x["VALORPENDIENTE"] for x in abono_cero_falso_multicurso),
        "mora_90_count": len(cartera_critica_90),
        "mora_90_monto": sum(x["VALORPENDIENTE"] for x in cartera_critica_90),
        "mora_60_count": len(mora_media_60),
        "mora_60_monto": sum(x["VALORPENDIENTE"] for x in mora_media_60),
        "mora_30_count": len(mora_temprana_30),
        "mora_30_monto": sum(x["VALORPENDIENTE"] for x in mora_temprana_30),
        "mora_aldia_count": len(cartera_al_dia),
        "mora_aldia_monto": sum(x["VALORPENDIENTE"] for x in cartera_al_dia),
        "jt_total_matriculas": len(jt_todos),
        "jt_facturado": sum(x["VALORCURSO"] for x in jt_todos),
        "jt_recaudado": sum(x["ABONOS"] for x in jt_todos),
        "jt_pendiente": sum(x["VALORPENDIENTE"] for x in jt_cartera),
        "jt_con_saldo_count": len(jt_cartera),
        "jt_abono_cero_count": len(jt_abono_cero),
        "jt_abono_cero_monto": sum(x["VALORPENDIENTE"] for x in jt_abono_cero),
        "jt_abono_cero_real_count": len(jt_abono_cero_real),
        "jt_abono_cero_real_monto": sum(x["VALORPENDIENTE"] for x in jt_abono_cero_real),
        "errores_digitacion_count": len(errores_digitacion),
        "descuadres_count": len(descuadres),
        "convenios_vacios_count": convenios_vacios_count,
        "promedio_mensual_empresa": promedio_mensual_emp,
        "promedio_mensual_jt": promedio_mensual_jt,
        "meta_mensual_empresa": meta_total,
        "meta_mensual_jt": meta_jt,
    }

    return {
        "metadata": metadata or {},
        "kpis": kpis,
        "abono_cero": abono_cero_todos,
        "abono_cero_real": abono_cero_real_global,
        "abono_cero_multicurso": abono_cero_falso_multicurso,
        "clientes_multicurso": clientes_multicurso_lista,
        "cartera_critica_90": cartera_critica_90,
        "mora_media_60": mora_media_60,
        "mora_temprana_30": mora_temprana_30,
        "cartera_al_dia": cartera_al_dia,
        "errores_digitacion": errores_digitacion,
        "descuadres": descuadres,
        "jt_cartera": jt_cartera,
        "jt_abono_cero": jt_abono_cero,
        "jt_todos": jt_todos,
        "mensual": lista_meses,
        "convenios": lista_convenios,
        "categorias": lista_categorias,
        "publicidad": dict(por_publicidad.most_common()),
        "todas_matriculas": registros,
    }


# =============================================================================
# 3. GENERADOR DE EXCEL MULTI-PESTAÑA NATIVO (.XLSX CON 9 HOJAS)
# =============================================================================
def _col_letter(n: int) -> str:
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def _construir_xlsx_multipestana(sheets_data: List[Tuple[str, List[str], List[List[Any]]]]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        sheet_overrides = "".join(
            f'<Override PartName="/xl/worksheets/sheet{i+1}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(len(sheets_data))
        )
        content_types = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  {sheet_overrides}
</Types>"""
        z.writestr("[Content_Types].xml", content_types)

        rels_root = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
        z.writestr("_rels/.rels", rels_root)

        wb_rels_items = "".join(
            f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i+1}.xml"/>'
            for i in range(len(sheets_data))
        )
        style_rid = f"rId{len(sheets_data)+1}"
        wb_rels = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {wb_rels_items}
  <Relationship Id="{style_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)

        sheets_xml = "".join(
            f'<sheet name="{xml_escape(name[:31])}" sheetId="{i+1}" r:id="rId{i+1}"/>'
            for i, (name, _, _) in enumerate(sheets_data)
        )
        workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>{sheets_xml}</sheets>
</workbook>"""
        z.writestr("xl/workbook.xml", workbook_xml)

        styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="$#,##0;($#,##0);&quot;$0&quot;"/>
  </numFmts>
  <fonts count="3">
    <font><sz val="10"/><color theme="1"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF9C0006"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF0D47A1"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border>
      <left style="thin"><color rgb="FFD9D9D9"/></left>
      <right style="thin"><color rgb="FFD9D9D9"/></right>
      <top style="thin"><color rgb="FFD9D9D9"/></top>
      <bottom style="thin"><color rgb="FFD9D9D9"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="3" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
    <xf numFmtId="164" fontId="2" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
  </cellXfs>
</styleSheet>"""
        z.writestr("xl/styles.xml", styles_xml)

        for idx, (sheet_name, headers, rows) in enumerate(sheets_data):
            lines = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
            ]
            num_cols = max(len(headers), max((len(r) for r in rows), default=1))
            lines.append("<cols>")
            for c_i in range(num_cols):
                hdr_len = len(str(headers[c_i])) if c_i < len(headers) else 12
                max_len = max([hdr_len] + [len(str(r[c_i])) for r in rows[:200] if c_i < len(r)])
                width = min(max(max_len + 4, 12), 44)
                lines.append(f'<col min="{c_i+1}" max="{c_i+1}" width="{width}" customWidth="1"/>')
            lines.append("</cols>")
            lines.append("<sheetData>")

            lines.append('<row r="1" ht="22" customHeight="1">')
            for c_i, h_val in enumerate(headers):
                col_ref = f"{_col_letter(c_i+1)}1"
                lines.append(
                    f'<c r="{col_ref}" s="1" t="inlineStr"><is><t>{xml_escape(str(h_val))}</t></is></c>'
                )
            lines.append("</row>")

            for r_i, row_vals in enumerate(rows, start=2):
                lines.append(f'<row r="{r_i}">')
                for c_i, val in enumerate(row_vals):
                    col_ref = f"{_col_letter(c_i+1)}{r_i}"
                    h_name = headers[c_i].upper() if c_i < len(headers) else ""
                    if isinstance(val, (int, float)) and not isinstance(val, bool):
                        is_money = any(
                            k in h_name
                            for k in ("VALOR", "ABONO", "SALDO", "FACTURADO", "RECAUDADO", "PENDIENTE", "MONTO", "CONTRATADO")
                        )
                        if is_money:
                            style_id = "4" if ("PENDIENTE" in h_name and val > 0) or ("SALDO_REAL" in h_name and val > 0) or val < 0 else "2"
                        else:
                            style_id = "3"
                        lines.append(f'<c r="{col_ref}" s="{style_id}"><v>{val}</v></c>')
                    else:
                        s_val = "" if val is None else str(val)
                        lines.append(
                            f'<c r="{col_ref}" s="0" t="inlineStr"><is><t>{xml_escape(s_val)}</t></is></c>'
                        )
                lines.append("</row>")

            lines.append("</sheetData>")
            if headers:
                last_col = _col_letter(len(headers))
                lines.append(f'<autoFilter ref="A1:{last_col}{len(rows)+1}"/>')
            lines.append("</worksheet>")
            z.writestr(f"xl/worksheets/sheet{idx+1}.xml", "\n".join(lines))

    return buf.getvalue()


def generar_reporte_excel_bytes(hallazgos: Dict[str, Any], tipo_informe: str = "ANUAL") -> bytes:
    k = hallazgos["kpis"]
    meta = hallazgos.get("metadata", {})

    resumen_headers = ["INDICADOR CONTABLE / COMERCIAL", "VALOR CONSOLIDADO", "ANÁLISIS Y CONTEXTO MULTI-CURSO"]
    resumen_rows = [
        ["Sede Auditada", meta.get("sede", "CEA AUTO ORIENTE"), f"Informe {tipo_informe}"],
        ["Periodo Analizado", f"{meta.get('fecha_inicial', '01/01/2026')} al {meta.get('fecha_final', '11/11/2026')}", f"Corte: {meta.get('fecha_corte_usada', '')}"],
        ["Matrículas Totales Registradas", k["total_matriculas"], f"Corresponden a {k['total_clientes_unicos']} clientes únicos (Cédulas)"],
        ["Clientes Multi-Curso (2 o más categorías)", k["clientes_multicurso_count"], f"Agrupan {k['matriculas_en_multicurso_count']} matrículas simultáneas (ej. Combo A2 + B1/C1)"],
        ["Total Facturado (Valor Cursos)", k["total_facturado"], "100% de ventas brutas del periodo"],
        ["Total Abonos Recaudados (Caja Efectiva)", k["total_abonos"], f"Tasa de Recaudo Efectivo: {k['tasa_recaudo_pct']}%"],
        ["Cartera Pendiente Total por Cobrar", k["total_pendiente"], f"Índice de Morosidad Global: {k['tasa_morosidad_pct']}% ({k['alumnos_con_saldo']} matrículas)"],
        ["[BANDERA 1 REAL] Alumnos con $0 en Caja Global (Suspensión)", k["abono_cero_real_count"], f"{k['abono_cero_real_clientes']} clientes con $0 en todos sus cursos (${k['abono_cero_real_monto']:,.0f})"],
        ["[FALSOS $0 MULTI-CURSO] 2da Matrícula en $0 pero abonó en la 1ra", k["abono_cero_multicurso_count"], f"Ya pagaron cuota inicial en su otra matrícula (${k['abono_cero_multicurso_monto']:,.0f})"],
        ["[BANDERA 2] Cartera Crítica (> 90 días)", k["mora_90_count"], f"Saldo vencido >90d: ${k['mora_90_monto']:,.0f}"],
        ["[BANDERA 2] Cartera Mora Media (61 - 90 días)", k["mora_60_count"], f"Saldo vencido 61-90d: ${k['mora_60_monto']:,.0f}"],
        ["[BANDERA 2] Cartera Mora Temprana (31 - 60 días)", k["mora_30_count"], f"Saldo vencido 31-60d: ${k['mora_30_monto']:,.0f}"],
        ["[BANDERA 3] Errores de Digitación / Saldos Negativos", k["errores_digitacion_count"], "Cursos en $930, sobrecobros negativos o clases atípicas"],
        ["[BANDERA 4] Ventas Totales Convenio Javier Torres", k["jt_total_matriculas"], f"Promedio: {k['promedio_mensual_jt']} mat./mes (Meta: {k['meta_mensual_jt']})"],
        ["[BANDERA 4] Cartera Pendiente Javier Torres", k["jt_pendiente"], f"{k['jt_con_saldo_count']} matrículas ({k['jt_abono_cero_real_count']} con $0 global real)"],
    ]

    cols_detalle = [
        "NUMEROMATRICULA",
        "NUMERODOCUMENTO",
        "ALUMNO",
        "CONVENIO_NORM",
        "FECHAINSCRIPCION",
        "CATEGORIA",
        "VALORCURSO",
        "ABONOS",
        "VALORPENDIENTE",
        "CATEGORIAS_CLIENTE",
        "TOTAL_CONTRATADO_CLIENTE",
        "TOTAL_ABONADO_CLIENTE",
        "SALDO_REAL_CLIENTE",
        "ETIQUETA_MULTICURSO",
        "DIAS_ANTIGUEDAD",
        "ESTADO",
    ]

    def extraer_filas(lista_regs: List[Dict[str, Any]], incluir_motivo: bool = False) -> Tuple[List[str], List[List[Any]]]:
        hdrs = list(cols_detalle) + (["MOTIVO_ERROR"] if incluir_motivo else [])
        filas_out = []
        for r in lista_regs:
            fila = [r.get(c, "") for c in cols_detalle]
            if incluir_motivo:
                fila.append(r.get("MOTIVO_ERROR", ""))
            filas_out.append(fila)
        return hdrs, filas_out

    jt_hdrs, jt_rows = extraer_filas(hallazgos["jt_cartera"])
    cero_real_hdrs, cero_real_rows = extraer_filas(hallazgos["abono_cero_real"])
    cero_mc_hdrs, cero_mc_rows = extraer_filas(hallazgos["abono_cero_multicurso"])

    # Consolidado de Clientes Multi-Curso
    mc_hdrs = [
        "NUMERODOCUMENTO",
        "ALUMNO",
        "CONVENIO_NORM",
        "CANTIDAD_CURSOS",
        "CATEGORIAS",
        "MATRICULAS",
        "FECHA_PRIMERA_INSC",
        "TOTAL_CONTRATADO",
        "TOTAL_ABONADO",
        "SALDO_REAL_CLIENTE",
        "DIAS_ANTIGUEDAD_MAX",
    ]
    mc_rows = [
        [c.get(k_col, "") for k_col in mc_hdrs]
        for c in hallazgos["clientes_multicurso"]
    ]

    mora90_hdrs, mora90_rows = extraer_filas(hallazgos["cartera_critica_90"])
    err_hdrs, err_rows = extraer_filas(hallazgos["errores_digitacion"], incluir_motivo=True)

    metas_hdrs = [
        "MES",
        "PERIODO",
        "MATRÍCULAS EMPRESA",
        "META EMPRESA",
        "% CUMPLIMIENTO EMPRESA",
        "ESTADO EMPRESA",
        "MATRÍCULAS JAVIER TORRES",
        "META JAVIER TORRES",
        "% CUMPLIMIENTO JT",
        "ESTADO JT",
        "TOTAL FACTURADO MES",
        "TOTAL RECAUDADO MES",
        "SALDO PENDIENTE MES",
        "ABONO $0 REAL",
    ]
    metas_rows = [
        [
            m["mes"],
            m["nombre_mes"],
            m["total_matriculas"],
            m["meta_empresa"],
            f"{m['cumplimiento_empresa_pct']}%",
            m["estado_meta_empresa"],
            m["matriculas_jt"],
            m["meta_jt"],
            f"{m['cumplimiento_jt_pct']}%",
            m["estado_meta_jt"],
            m["facturado"],
            m["recaudado"],
            m["pendiente"],
            m["abono_cero_real_count"],
        ]
        for m in hallazgos["mensual"]
    ]

    conv_hdrs = [
        "CONVENIO",
        "MATRÍCULAS",
        "CLIENTES ÚNICOS",
        "TOTAL FACTURADO",
        "TOTAL RECAUDADO",
        "% RECAUDO",
        "CARTERA PENDIENTE",
        "% MOROSIDAD",
        "DEUDORES ACTIVOS",
        "ABONO $0 REAL",
    ]
    conv_rows = [
        [
            c["convenio"],
            c["matriculas"],
            c["clientes_unicos"],
            c["facturado"],
            c["recaudado"],
            f"{c['tasa_recaudo']}%",
            c["pendiente"],
            f"{c['tasa_morosidad']}%",
            c["con_saldo"],
            c["abono_cero_real"],
        ]
        for c in hallazgos["convenios"]
    ]

    sheets = [
        ("Dashboard_Resumen", resumen_headers, resumen_rows),
        ("Cartera_Javier_Torres", jt_hdrs, jt_rows),
        ("Riesgo_Cero_Real_Suspender", cero_real_hdrs, cero_real_rows),
        ("Combos_MultiCurso_En_Cero", cero_mc_hdrs, cero_mc_rows),
        ("Consolidado_231_MultiCurso", mc_hdrs, mc_rows),
        ("Mora_Mayor_90_Dias", mora90_hdrs, mora90_rows),
        ("Errores_Digitacion", err_hdrs, err_rows),
        ("Cumplimiento_Metas_Mes", metas_hdrs, metas_rows),
        ("Consolidado_Convenios", conv_hdrs, conv_rows),
    ]
    return _construir_xlsx_multipestana(sheets)


def generar_reporte_excel(
    registros: List[Dict[str, Any]],
    hallazgos: Dict[str, Any],
    tipo_informe: str = "ANUAL",
    output_filename: str = "auditoria_cea.xlsx",
) -> str:
    xlsx_bytes = generar_reporte_excel_bytes(hallazgos, tipo_informe=tipo_informe)
    Path(output_filename).write_bytes(xlsx_bytes)
    return output_filename


# =============================================================================
# 4. GENERADOR DE INFORME EJECUTIVO EN PDF (REPORTLAB)
# =============================================================================
def generar_reporte_pdf_bytes(hallazgos: Dict[str, Any], tipo_informe: str = "ANUAL") -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=landscape(letter),
        leftMargin=1.2 * cm,
        rightMargin=1.2 * cm,
        topMargin=1.0 * cm,
        bottomMargin=1.0 * cm,
        title=f"Informe de Auditoría Contable y Comercial - {tipo_informe}",
    )

    styles = getSampleStyleSheet()
    subtitle_style = ParagraphStyle(
        "AuditSub",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=12,
        textColor=colors.HexColor("#37474F"),
    )
    sec_style = ParagraphStyle(
        "SecHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=11.0,
        leading=14,
        textColor=colors.HexColor("#0D47A1"),
        spaceBefore=8,
        spaceAfter=4,
    )
    cell_style = ParagraphStyle(
        "CellText",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.2,
        leading=8.8,
    )

    story = []
    k = hallazgos["kpis"]
    meta = hallazgos.get("metadata", {})

    logo_path = BASE_DIR / "logo_autooriente.png"
    if not logo_path.exists():
        logo_path = BASE_DIR / "SUBIR_A_GITHUB" / "logo_autooriente.png"
    if not logo_path.exists():
        logo_path = BASE_DIR / "assets" / "logo_autooriente.png"

    header_text = Paragraph(
        f"<b>CENTRO DE ENSEÑANZA AUTOMOVILÍSTICA AUTO ORIENTE</b><br/>"
        f"<font size=12 color='#1565C0'><b>AUDITORÍA CONTABLE, CRUCE MULTI-CURSO Y CONTROL DE CARTERA ({tipo_informe})</b></font><br/>"
        f"<font size=8.5 color='#546E7A'>Periodo: {meta.get('fecha_inicial', '01/01/2026')} al {meta.get('fecha_final', '11/11/2026')} | "
        f"Matrículas: {k['total_matriculas']} ({k['total_clientes_unicos']} Clientes Únicos | {k['clientes_multicurso_count']} Clientes Multi-Curso)</font>",
        subtitle_style,
    )

    if logo_path.exists():
        try:
            img = RLImage(str(logo_path), width=3.8 * cm, height=1.8 * cm, kind="proportional")
            hdr_table = Table([[img, header_text]], colWidths=[4.2 * cm, 21.0 * cm])
            hdr_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))
            story.append(hdr_table)
        except Exception:
            story.append(header_text)
    else:
        story.append(header_text)

    story.append(Spacer(1, 0.2 * cm))

    story.append(Paragraph("1. DASHBOARD EJECUTIVO Y CRUCE DE CLIENTES MULTI-CURSO", sec_style))
    kpi_data = [
        [
            "TOTAL FACTURADO",
            "RECAUDO EN CAJA",
            "CARTERA PENDIENTE",
            "RIESGO $0 REAL (GLOBAL)",
            "COMBOS MULTI-CURSO",
            "CARTERA JAVIER TORRES",
        ],
        [
            f"${k['total_facturado']:,.0f}\n({k['total_matriculas']} mat. / {k['total_clientes_unicos']} clientes)",
            f"${k['total_abonos']:,.0f}\n(Tasa: {k['tasa_recaudo_pct']}%)",
            f"${k['total_pendiente']:,.0f}\n({k['alumnos_con_saldo']} matrículas)",
            f"{k['abono_cero_real_count']} matrículas\n(${k['abono_cero_real_monto']:,.0f} sin abono)",
            f"{k['clientes_multicurso_count']} clientes (2+ cursos)\n({k['abono_cero_multicurso_count']} con 2da mat. en $0)",
            f"${k['jt_pendiente']:,.0f}\n({k['jt_con_saldo_count']} de {k['jt_total_matriculas']} mat.)",
        ],
    ]
    t_kpi = Table(kpi_data, colWidths=[4.2 * cm] * 6)
    t_kpi.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0D47A1")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, 0), 8.0),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F5F7FA")),
            ("TEXTCOLOR", (2, 1), (3, 1), colors.HexColor("#B71C1C")),
            ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
            ("FONTSIZE", (0, 1), (-1, 1), 8.8),
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#90A4AE")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CFD8DC")),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ])
    )
    story.append(t_kpi)
    story.append(Spacer(1, 0.25 * cm))

    story.append(
        Paragraph(
            f"2. ALUMNOS EN RIESGO CERO REAL ($0 ABONADO EN TODAS SUS MATRÍCULAS - BLOQUEO INMEDIATO: {len(hallazgos['abono_cero_real'])} Matrículas)",
            sec_style,
        )
    )
    cero_rows = [["MATRÍCULA", "CÉDULA", "ALUMNO", "CONVENIO", "FECHA INSC.", "CURSOS CLIENTE", "SALDO CLIENTE", "DÍAS", "ESTADO"]]
    for r in hallazgos["abono_cero_real"]:
        cero_rows.append([
            r["NUMEROMATRICULA"],
            r["NUMERODOCUMENTO"],
            Paragraph(r["ALUMNO"][:32], cell_style),
            r["CONVENIO_NORM"],
            r["FECHAINSCRIPCION"],
            r["CATEGORIAS_CLIENTE"],
            f"${r['SALDO_REAL_CLIENTE']:,.0f}",
            f"{r['DIAS_ANTIGUEDAD']} d",
            r["ESTADO"],
        ])
    t_cero = Table(
        cero_rows,
        colWidths=[2.3 * cm, 2.4 * cm, 5.8 * cm, 3.4 * cm, 2.2 * cm, 2.5 * cm, 2.6 * cm, 1.7 * cm, 2.4 * cm],
        repeatRows=1,
    )
    t_cero.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#B71C1C")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.2),
            ("ALIGN", (6, 1), (6, -1), "RIGHT"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#FFF8F8"), colors.white]),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#EF9A9A")),
        ])
    )
    story.append(t_cero)

    story.append(PageBreak())
    story.append(
        Paragraph(
            f"3. CARTERA PENDIENTE CONVENIO JAVIER TORRES CON CONTEXTO MULTI-CURSO ({len(hallazgos['jt_cartera'])} Matrículas - ${k['jt_pendiente']:,.0f})",
            sec_style,
        )
    )
    jt_rows = [["MATRÍCULA", "CÉDULA", "ALUMNO", "CAT.", "ABONO MAT.", "SALDO MAT.", "COMBO CLIENTE", "ABONO GLOBAL", "SALDO GLOBAL", "DÍAS"]]
    for r in hallazgos["jt_cartera"]:
        jt_rows.append([
            r["NUMEROMATRICULA"],
            r["NUMERODOCUMENTO"],
            Paragraph(r["ALUMNO"][:30], cell_style),
            r["CATEGORIA"],
            f"${r['ABONOS']:,.0f}",
            f"${r['VALORPENDIENTE']:,.0f}",
            r["CATEGORIAS_CLIENTE"],
            f"${r['TOTAL_ABONADO_CLIENTE']:,.0f}",
            f"${r['SALDO_REAL_CLIENTE']:,.0f}",
            f"{r['DIAS_ANTIGUEDAD']} d",
        ])
    t_jt = Table(
        jt_rows,
        colWidths=[2.2 * cm, 2.3 * cm, 5.5 * cm, 1.2 * cm, 2.3 * cm, 2.3 * cm, 2.5 * cm, 2.5 * cm, 2.5 * cm, 1.8 * cm],
        repeatRows=1,
    )
    t_jt.setStyle(
        TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0D47A1")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.0),
            ("ALIGN", (4, 1), (5, -1), "RIGHT"),
            ("ALIGN", (7, 1), (8, -1), "RIGHT"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F5F7FA")]),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#CFD8DC")),
        ])
    )
    story.append(t_jt)

    doc.build(story)
    return buf.getvalue()


# =============================================================================
# 5. CAPA DE DESPACHO: CORREO ELECTRÓNICO (SMTP) Y BOT DE TELEGRAM
# =============================================================================
def enviar_correo_auditoria(
    hallazgos: Dict[str, Any],
    excel_bytes: bytes,
    pdf_bytes: bytes,
    periodicidad: str = "ANUAL",
    destinatarios_extra: Optional[List[str]] = None,
    config: Optional[Dict[str, Any]] = None,
) -> Tuple[bool, str]:
    cfg = {**CONFIG_DEFAULT, **(config or {})}
    remitente = cfg["EMAIL_REMITENTE"]
    password = cfg["EMAIL_PASSWORD"]
    destinatarios = list(cfg["EMAIL_DESTINATARIOS"])
    if destinatarios_extra:
        for d in destinatarios_extra:
            if d and d not in destinatarios:
                destinatarios.append(d)

    if not password or password == "tu_contraseña_de_aplicacion" or remitente == "tu_correo@gmail.com":
        return False, "Credenciales SMTP no configuradas (define EMAIL_REMITENTE y EMAIL_PASSWORD en Variables de Entorno)."

    k = hallazgos["kpis"]
    fecha_str = datetime.date.today().strftime("%Y-%m-%d")

    try:
        msg = MIMEMultipart()
        msg["From"] = remitente
        msg["To"] = ", ".join(destinatarios)
        msg["Subject"] = f"🚨 INFORME DE AUDITORÍA Y CRUCE MULTI-CURSO ({periodicidad}) - CEA AUTO ORIENTE - {fecha_str}"

        cuerpo = f"""
        <html>
        <body style="font-family: 'Segoe UI', Arial, sans-serif; color: #1e293b; background-color: #f8fafc; padding: 20px;">
            <div style="max-width: 680px; margin: auto; background: #ffffff; border-radius: 10px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.08);">
                <div style="background: linear-gradient(135deg, #0d47a1, #1976d2); color: #ffffff; padding: 20px 24px;">
                    <h2 style="margin: 0; font-size: 20px;">Auditoría Contable y Consolidación Multi-Curso ({periodicidad})</h2>
                    <p style="margin: 5px 0 0; opacity: 0.9; font-size: 13px;">CEA Auto Oriente | Corte: {fecha_str}</p>
                </div>
                <div style="padding: 24px;">
                    <p>Resumen consolidado por <b>Matrícula ({k['total_matriculas']})</b> y por <b>Cédula / Cliente Único ({k['total_clientes_unicos']})</b>:</p>
                    <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; border-color: #cbd5e1; font-size: 14px;">
                        <tr style="background-color: #f1f5f9; text-align: left;"><th>Indicador Auditado</th><th>Resultado Consolidado</th></tr>
                        <tr><td><b>Total Facturado</b></td><td>${k['total_facturado']:,.0f} ({k['total_matriculas']} matrículas / {k['total_clientes_unicos']} clientes)</td></tr>
                        <tr><td><b>Clientes Multi-Curso (Combos A2+B1/C1)</b></td><td><b>{k['clientes_multicurso_count']} clientes</b> ({k['matriculas_en_multicurso_count']} matrículas simultáneas)</td></tr>
                        <tr><td><b>Recaudo Efectivo en Caja</b></td><td style="color: #15803d; font-weight: bold;">${k['total_abonos']:,.0f} ({k['tasa_recaudo_pct']}%)</td></tr>
                        <tr style="background-color: #fef2f2;"><td style="color: #b91c1c;"><b>Cartera Pendiente por Cobrar</b></td><td style="color: #b91c1c; font-weight: bold;">${k['total_pendiente']:,.0f} ({k['alumnos_con_saldo']} matrículas)</td></tr>
                        <tr><td><b>🚨 Riesgo Cero REAL ($0 en todos sus cursos)</b></td><td><b>{k['abono_cero_real_count']} matrículas ({k['abono_cero_real_clientes']} clientes)</b> (${k['abono_cero_real_monto']:,.0f})</td></tr>
                        <tr><td><b>🔄 Falsos $0 (Combos Multi-Curso con abono en 1ra mat.)</b></td><td><b>{k['abono_cero_multicurso_count']} matrículas</b> (${k['abono_cero_multicurso_monto']:,.0f})</td></tr>
                        <tr><td><b>⏳ Cartera Crítica (&gt; 90 días)</b></td><td><b>{k['mora_90_count']} matrículas</b> (${k['mora_90_monto']:,.0f})</td></tr>
                        <tr><td><b>🤝 Convenio Javier Torres</b></td><td><b>{k['jt_total_matriculas']} matrículas</b> | Cartera: <b>${k['jt_pendiente']:,.0f}</b> ({k['jt_abono_cero_real_count']} en $0 real)</td></tr>
                    </table>
                </div>
            </div>
        </body>
        </html>
        """
        msg.attach(MIMEText(cuerpo, "html", "utf-8"))

        part_xlsx = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        part_xlsx.set_payload(excel_bytes)
        encoders.encode_base64(part_xlsx)
        part_xlsx.add_header("Content-Disposition", f'attachment; filename="Auditoria_CEA_{periodicidad}_{fecha_str}.xlsx"')
        msg.attach(part_xlsx)

        part_pdf = MIMEBase("application", "pdf")
        part_pdf.set_payload(pdf_bytes)
        encoders.encode_base64(part_pdf)
        part_pdf.add_header("Content-Disposition", f'attachment; filename="Resumen_Ejecutivo_CEA_{periodicidad}_{fecha_str}.pdf"')
        msg.attach(part_pdf)

        server = smtplib.SMTP(cfg["SMTP_SERVER"], cfg["SMTP_PORT"], timeout=20)
        server.starttls()
        server.login(remitente, password)
        server.sendmail(remitente, destinatarios, msg.as_string())
        server.quit()
        return True, f"Correo enviado exitosamente a: {', '.join(destinatarios)}"
    except Exception as exc:
        return False, f"Error al enviar correo SMTP: {exc}"


def formatear_resumen_telegram(hallazgos: Dict[str, Any], periodicidad: str = "ANUAL") -> str:
    k = hallazgos["kpis"]
    meta = hallazgos.get("metadata", {})
    return (
        f"📊 *AUDITORÍA CONTABLE Y CRUCE MULTI-CURSO ({periodicidad})*\n"
        f"🏫 *CEA Auto Oriente* | `{meta.get('fecha_inicial', '01/01/2026')}` - `{meta.get('fecha_final', '11/11/2026')}`\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"👥 *Matrículas:* `{k['total_matriculas']}` (`{k['total_clientes_unicos']}` clientes únicos)\n"
        f"🔄 *Clientes Multi-Curso (2+ cursos):* `{k['clientes_multicurso_count']}` clientes\n"
        f"💰 *Total Facturado:* `${k['total_facturado']:,.0f}`\n"
        f"✅ *Recaudo en Caja:* `${k['total_abonos']:,.0f}` (`{k['tasa_recaudo_pct']}%`)\n"
        f"⚠️ *Cartera Pendiente:* `${k['total_pendiente']:,.0f}` (`{k['alumnos_con_saldo']}` mat.)\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"🚩 *BANDERAS ROJAS (CON CRUCE POR CÉDULA):*\n"
        f"1️⃣ *Riesgo $0 REAL (Suspensión):* `{k['abono_cero_real_count']}` mat. / `{k['abono_cero_real_clientes']}` clientes (`${k['abono_cero_real_monto']:,.0f}`)\n"
        f"   • _Combos Multi-Curso (2da mat. en $0 pero abonó en la 1ra):_ `{k['abono_cero_multicurso_count']}` alumnos\n"
        f"2️⃣ *Mora Crítica (>90 días):* `{k['mora_90_count']}` alumnos (`${k['mora_90_monto']:,.0f}`)\n"
        f"3️⃣ *Errores Digitación / Negativos:* `{k['errores_digitacion_count']}` registros\n"
        f"4️⃣ *Convenio Javier Torres:* `{k['jt_total_matriculas']}` mat.\n"
        f"   • _Cartera JT:_ `${k['jt_pendiente']:,.0f}` (`{k['jt_con_saldo_count']}` mat. | `{k['jt_abono_cero_real_count']}` en $0 real)\n"
        f"━━━━━━━━━━━━━━━━━━━━━━\n"
        f"📎 _Se adjuntan Excel Multi-Pestaña (9 hojas) y PDF Ejecutivo._"
    )


def enviar_telegram_auditoria_sync(
    hallazgos: Dict[str, Any],
    excel_bytes: bytes,
    pdf_bytes: bytes,
    periodicidad: str = "ANUAL",
    chat_id: Optional[str] = None,
    bot_token: Optional[str] = None,
) -> Tuple[bool, str]:
    token = (bot_token or CONFIG_DEFAULT["TELEGRAM_BOT_TOKEN"]).strip()
    destino = str(chat_id or CONFIG_DEFAULT["TELEGRAM_CHAT_ID"]).strip()

    if not token or token == "TU_TELEGRAM_BOT_TOKEN":
        return False, "TELEGRAM_BOT_TOKEN no configurado."
    if not destino:
        return False, "Chat ID de destino no especificado."

    resumen_md = formatear_resumen_telegram(hallazgos, periodicidad=periodicidad)
    fecha_str = datetime.date.today().strftime("%Y%m%d")

    try:
        url_msg = f"https://api.telegram.org/bot{token}/sendMessage"
        payload = json.dumps({"chat_id": destino, "text": resumen_md, "parse_mode": "Markdown"}).encode("utf-8")
        req = urllib.request.Request(url_msg, data=payload, headers={"Content-Type": "application/json"})
        urllib.request.urlopen(req, timeout=15).read()

        def _send_doc(file_bytes: bytes, filename: str, mime: str, caption: str) -> None:
            boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
            body = io.BytesIO()
            for k_f, v_f in [("chat_id", destino), ("caption", caption)]:
                body.write(f"--{boundary}\r\n".encode())
                body.write(f'Content-Disposition: form-data; name="{k_f}"\r\n\r\n{v_f}\r\n'.encode())
            body.write(f"--{boundary}\r\n".encode())
            body.write(
                f'Content-Disposition: form-data; name="document"; filename="{filename}"\r\n'
                f"Content-Type: {mime}\r\n\r\n".encode()
            )
            body.write(file_bytes)
            body.write(f"\r\n--{boundary}--\r\n".encode())
            req_doc = urllib.request.Request(
                f"https://api.telegram.org/bot{token}/sendDocument",
                data=body.getvalue(),
                headers={"Content-Type": f"multipart/form-data; boundary={boundary}"},
            )
            urllib.request.urlopen(req_doc, timeout=30).read()

        _send_doc(
            excel_bytes,
            f"Auditoria_CEA_{periodicidad}_{fecha_str}.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            f"📊 Excel Multi-Pestaña de Auditoría ({periodicidad})",
        )
        _send_doc(
            pdf_bytes,
            f"Resumen_Ejecutivo_CEA_{periodicidad}_{fecha_str}.pdf",
            "application/pdf",
            f"📄 Informe Ejecutivo en PDF ({periodicidad})",
        )
        return True, f"Informe enviado exitosamente a Telegram (Chat ID: {destino})."
    except Exception as exc:
        return False, f"Error al enviar por Telegram: {exc}"


def ejecutar(ruta_excel: str, periodicidad: str = "ANUAL", ruta_ingresos: Optional[str] = None) -> Dict[str, Any]:
    registros, metadata = cargar_y_limpiar_datos(ruta_excel, origen_ingresos=ruta_ingresos)
    hallazgos = auditar_riesgos(registros, metadata=metadata)
    fecha_str = datetime.datetime.now().strftime("%Y%m%d")
    archivo_xlsx = f"Auditoria_CEA_{periodicidad}_{fecha_str}.xlsx"
    archivo_pdf = f"Auditoria_CEA_{periodicidad}_{fecha_str}.pdf"

    xlsx_bytes = generar_reporte_excel_bytes(hallazgos, tipo_informe=periodicidad)
    pdf_bytes = generar_reporte_pdf_bytes(hallazgos, tipo_informe=periodicidad)

    Path(archivo_xlsx).write_bytes(xlsx_bytes)
    Path(archivo_pdf).write_bytes(pdf_bytes)
    print(f"[OK] Reporte Excel Multi-Pestaña generado: {archivo_xlsx}")
    print(f"[OK] Reporte PDF Ejecutivo generado: {archivo_pdf}")

    ok_mail, msg_mail = enviar_correo_auditoria(hallazgos, xlsx_bytes, pdf_bytes, periodicidad=periodicidad)
    print(f"[{'OK' if ok_mail else 'AVISO'}] {msg_mail}")
    return hallazgos


if __name__ == "__main__":
    import sys

    archivo = sys.argv[1] if len(sys.argv) > 1 else "matriculas11112026.xlsx"
    periodo = sys.argv[2] if len(sys.argv) > 2 else "ANUAL"
    ingresos = sys.argv[3] if len(sys.argv) > 3 else None
    ejecutar(archivo, periodicidad=periodo, ruta_ingresos=ingresos)
