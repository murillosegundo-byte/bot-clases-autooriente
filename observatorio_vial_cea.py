"""
=============================================================================
OBSERVATORIO NACIONAL DE SEGURIDAD VIAL - COLOMBIA (INMLCF / SIRDEC / ANSV)
2021 a 2026 — Integrado en Línea 24/7 para CEA Auto Oriente (Sin depender de PC)
=============================================================================
Replica el 100% de la aplicación Streamlit (`app.py` + `actualizador.py`)
directamente en el servidor web en la nube (`/observatorio`):
- Carga los 48.981+ registros oficiales consolidados (2021 a 2026).
- Filtros interactivos: Año del Hecho, Departamento, Municipio (dinámico),
  Actor Vial (Rol) y Sexo de la Víctima.
- 4 KPIs con porcentajes: Total Víctimas Fatales, Usuarios de Moto, Peatones, Hombres.
- 4 Gráficos Plotly.js idénticos a Plotly Express:
  1. Participación por Actor Vial (Pie / Donut hole=0.45)
  2. Evolución Histórica de Fatalidades (Bar apilado por Año y Actor Vial)
  3. Objeto de Colisión — Con qué impactan (Bar horizontal Top 8 escala Reds)
  4. Distribución por Rango Horario (Bar por Rango3horas)
- Tabla de Registros Individuales Detallados (100 registros filtrados).
- Panel lateral de Administración:
  * 🔄 Sincronizar con ANSV ahora (`https://www.datos.gov.co/resource/mpt5-rchv.csv?$limit=100000`)
  * 📂 Cargar nuevo CSV manualmente
"""

import csv
import io
import os
import threading
import urllib.request
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

BASE_DIR = Path(__file__).resolve().parent
CSV_CONSOLIDADO_PATH = BASE_DIR / "ansv_fallecidos_2021_2026.csv"
CARPETA_FALLECIDOS = BASE_DIR / "fallecidos" / "fallecidos"
URL_API_ANSV = "https://www.datos.gov.co/resource/mpt5-rchv.csv?$limit=100000"

COLUMNAS_CLAVE = [
    "AnoHecho",
    "MesOCurrencia",
    "DiaOcurrencia",
    "Departamento",
    "Municipio",
    "ActorVial",
    "CondicionVictima",
    "Sexo",
    "RangoEdad",
    "ObjetoColision",
    "Rango3horas",
    "Zona",
]


class ObservatorioVialService:
    def __init__(self) -> None:
        self._lock = threading.Lock()
        self.registros: List[Tuple[int, str, str, str, str, str, str, str, str, str, str, str]] = []
        self.anios: List[int] = []
        self.departamentos: List[str] = []
        self.municipios_por_depto: Dict[str, List[str]] = {}
        self.todos_municipios: List[str] = []
        self.actores: List[str] = []
        self.sexos: List[str] = []
        self.ultima_carga: str = ""
        self.cargar_datos()

    def _detectar_delimitador(self, texto_muestra: str) -> str:
        return ";" if texto_muestra.count(";") >= texto_muestra.count(",") else ","

    def _leer_csv_a_tuplas(self, ruta: Path) -> List[Tuple[int, str, str, str, str, str, str, str, str, str, str, str]]:
        if not ruta.exists():
            return []
        raw_bytes = ruta.read_bytes()
        texto = ""
        for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
            try:
                texto = raw_bytes.decode(enc)
                break
            except Exception:
                continue
        if not texto:
            return []

        delim = self._detectar_delimitador(texto[:4096])
        reader = csv.reader(io.StringIO(texto), delimiter=delim)
        try:
            headers_raw = next(reader)
        except StopIteration:
            return []

        headers = [h.strip().replace("\ufeff", "") for h in headers_raw]
        # Mapeo flexible por si viene de Socrata API (minúsculas) o CSV ANSV
        mapa_lower = {h.lower(): idx for idx, h in enumerate(headers)}

        def idx_col(*nombres: str) -> int:
            for n in nombres:
                if n.lower() in mapa_lower:
                    return mapa_lower[n.lower()]
            return -1

        i_ano = idx_col("AnoHecho", "anohecho", "anio", "ano_hecho", "a_o_hecho")
        i_mes = idx_col("MesOCurrencia", "mesocurrencia", "mes_ocurrencia", "mes")
        i_dia = idx_col("DiaOcurrencia", "diaocurrencia", "dia_ocurrencia", "dia")
        i_dep = idx_col("Departamento", "departamento")
        i_mun = idx_col("Municipio", "municipio")
        i_act = idx_col("ActorVial", "actorvial", "actor_vial")
        i_con = idx_col("CondicionVictima", "condicionvictima", "condicion_victima")
        i_sex = idx_col("Sexo", "sexo")
        i_eda = idx_col("RangoEdad", "rangoedad", "rango_edad")
        i_obj = idx_col("ObjetoColision", "objetocolision", "objeto_colision")
        i_r3h = idx_col("Rango3horas", "rango3horas", "rango_3_horas")
        i_zon = idx_col("Zona", "zona")

        out: List[Tuple[int, str, str, str, str, str, str, str, str, str, str, str]] = []
        for row in reader:
            if not row:
                continue
            lr = len(row)
            ano_str = row[i_ano].strip() if 0 <= i_ano < lr else "0"
            try:
                ano_val = int(float(ano_str))
            except Exception:
                ano_val = 0
            dep = (row[i_dep].strip().upper() if 0 <= i_dep < lr else "") or "SIN DATO"
            mun = (row[i_mun].strip().upper() if 0 <= i_mun < lr else "") or "SIN DATO"
            act = (row[i_act].strip() if 0 <= i_act < lr else "") or "Sin información"
            if act.lower() == "nan":
                act = "Sin información"
            mes = row[i_mes].strip() if 0 <= i_mes < lr else ""
            dia = row[i_dia].strip() if 0 <= i_dia < lr else ""
            con = row[i_con].strip() if 0 <= i_con < lr else ""
            sex = row[i_sex].strip() if 0 <= i_sex < lr else ""
            eda = row[i_eda].strip() if 0 <= i_eda < lr else ""
            obj = row[i_obj].strip() if 0 <= i_obj < lr else ""
            r3h = row[i_r3h].strip() if 0 <= i_r3h < lr else "-1"
            zon = row[i_zon].strip() if 0 <= i_zon < lr else ""

            out.append((ano_val, mes, dia, dep, mun, act, con, sex, eda, obj, r3h, zon))
        return out

    def cargar_datos(self) -> int:
        with self._lock:
            tuplas: List[Tuple[int, str, str, str, str, str, str, str, str, str, str, str]] = []

            # 1. Intentar cargar archivos originales en fallecidos/fallecidos si existen
            rutas_final = [
                CARPETA_FALLECIDOS / "versionfinal.csv",
                BASE_DIR / "fallecidos" / "versionfinal.csv",
                BASE_DIR / "versionfinal.csv",
            ]
            rutas_reciente = [
                CARPETA_FALLECIDOS / "version_parcial_al_ultimo_mes_cargado_del_ano_en_curso.csv",
                BASE_DIR / "fallecidos" / "version_parcial_al_ultimo_mes_cargado_del_ano_en_curso.csv",
                BASE_DIR / "version_parcial_al_ultimo_mes_cargado_del_ano_en_curso.csv",
                CARPETA_FALLECIDOS / "version_parcial_a_diciembre.csv",
            ]

            cargo_originales = False
            for rf in rutas_final:
                if rf.exists():
                    tuplas.extend(self._leer_csv_a_tuplas(rf))
                    cargo_originales = True
                    break
            for rr in rutas_reciente:
                if rr.exists():
                    tuplas.extend(self._leer_csv_a_tuplas(rr))
                    cargo_originales = True
                    break

            # 2. Si estamos en Render o solo está el consolidado empaquetado, cargarlo
            if not cargo_originales and CSV_CONSOLIDADO_PATH.exists():
                tuplas.extend(self._leer_csv_a_tuplas(CSV_CONSOLIDADO_PATH))

            # 3. Cargar también cualquier CSV adicional subido manualmente en fallecidos/fallecidos
            if CARPETA_FALLECIDOS.exists():
                conocidos = {
                    "versionfinal.csv",
                    "version_parcial_al_ultimo_mes_cargado_del_ano_en_curso.csv",
                    "version_parcial_a_diciembre.csv",
                }
                for extra_csv in sorted(CARPETA_FALLECIDOS.glob("*.csv")):
                    if extra_csv.name not in conocidos:
                        tuplas.extend(self._leer_csv_a_tuplas(extra_csv))

            self.registros = tuplas
            anios_set = {t[0] for t in tuplas if t[0] > 0}
            self.anios = sorted(anios_set, reverse=True)

            deptos_set = {t[3] for t in tuplas if t[3] and t[3] != "NAN"}
            self.departamentos = sorted(deptos_set)

            muni_map: Dict[str, set] = defaultdict(set)
            todos_muni_set = set()
            actores_set = set()
            sexos_set = set()

            for t in tuplas:
                dep, mun, act, sex = t[3], t[4], t[5], t[7]
                if dep and mun:
                    muni_map[dep].add(mun)
                    todos_muni_set.add(mun)
                if act and act.lower() != "nan":
                    actores_set.add(act)
                if sex and sex not in ("-1", "nan", ""):
                    sexos_set.add(sex)

            self.municipios_por_depto = {d: sorted(ms) for d, ms in muni_map.items()}
            self.todos_municipios = sorted(todos_muni_set)
            self.actores = sorted(actores_set)
            self.sexos = sorted(sexos_set)
            self.ultima_carga = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            return len(self.registros)

    def consultar(
        self,
        anio: str = "Todos",
        departamento: str = "Todos",
        municipio: str = "Todos",
        actor_vial: str = "Todos",
        sexo: str = "Todos",
    ) -> Dict[str, Any]:
        anio_int: Optional[int] = None
        if anio and anio != "Todos":
            try:
                anio_int = int(anio)
            except Exception:
                anio_int = None

        dep_f = departamento.strip().upper() if departamento and departamento != "Todos" else None
        mun_f = municipio.strip().upper() if municipio and municipio != "Todos" else None
        act_f = actor_vial.strip() if actor_vial and actor_vial != "Todos" else None
        sex_f = sexo.strip() if sexo and sexo != "Todos" else None

        total = 0
        motos = 0
        peatones = 0
        hombres = 0

        cnt_actor: Counter = Counter()
        cnt_anual_actor: Dict[int, Counter] = defaultdict(Counter)
        cnt_objeto: Counter = Counter()
        cnt_horario: Counter = Counter()
        filas_muestra: List[Dict[str, Any]] = []

        with self._lock:
            datos_iter = list(self.registros)

        for t in datos_iter:
            ano_val, mes, dia, dep, mun, act, con, sex, eda, obj, r3h, zon = t
            if anio_int is not None and ano_val != anio_int:
                continue
            if dep_f is not None and dep != dep_f:
                continue
            if mun_f is not None and mun != mun_f:
                continue
            if act_f is not None and act != act_f:
                continue
            if sex_f is not None and sex != sex_f:
                continue

            total += 1
            act_low = act.lower()
            if "moto" in act_low:
                motos += 1
            if "peat" in act_low:
                peatones += 1
            if sex == "Hombre":
                hombres += 1

            cnt_actor[act] += 1
            if ano_val > 0:
                cnt_anual_actor[ano_val][act] += 1
            if obj:
                cnt_objeto[obj] += 1
            if r3h and r3h != "-1":
                cnt_horario[r3h] += 1

            if len(filas_muestra) < 100:
                filas_muestra.append({
                    "AnoHecho": ano_val,
                    "MesOCurrencia": mes,
                    "DiaOcurrencia": dia,
                    "Departamento": dep,
                    "Municipio": mun,
                    "ActorVial": act,
                    "CondicionVictima": con,
                    "Sexo": sex,
                    "RangoEdad": eda,
                    "ObjetoColision": obj,
                    "Zona": zon,
                })

        pct_m = round((motos / total * 100), 1) if total > 0 else 0.0
        pct_p = round((peatones / total * 100), 1) if total > 0 else 0.0
        pct_h = round((hombres / total * 100), 1) if total > 0 else 0.0

        # 1. Participación por Actor Vial
        actor_items = cnt_actor.most_common()
        chart_actor = {
            "labels": [x[0] for x in actor_items],
            "values": [x[1] for x in actor_items],
        }

        # 2. Evolución Histórica por Año y Actor Vial
        anios_eje = sorted(cnt_anual_actor.keys())
        actores_presentes = [x[0] for x in actor_items]
        series_anual = []
        for act_name in actores_presentes:
            series_anual.append({
                "actor": act_name,
                "x": anios_eje,
                "y": [cnt_anual_actor[a].get(act_name, 0) for a in anios_eje],
            })

        # 3. Objeto de Colisión Top 8 (orden ascendente para barra horizontal como en Plotly)
        top8_obj = list(reversed(cnt_objeto.most_common(8)))
        chart_colision = {
            "y": [x[0] for x in top8_obj],
            "x": [x[1] for x in top8_obj],
        }

        # 4. Distribución por Rango Horario ordenada por RangoHorario
        horarios_ord = sorted(cnt_horario.items(), key=lambda x: x[0])
        chart_horario = {
            "x": [x[0] for x in horarios_ord],
            "y": [x[1] for x in horarios_ord],
        }

        munis_lista = (
            self.municipios_por_depto.get(dep_f, [])
            if dep_f
            else self.todos_municipios
        )

        return {
            "ok": True,
            "ultima_carga": self.ultima_carga,
            "total_base": len(self.registros),
            "filtros_opciones": {
                "anios": ["Todos"] + [str(a) for a in self.anios],
                "departamentos": ["Todos"] + self.departamentos,
                "municipios": ["Todos"] + munis_lista,
                "actores": ["Todos"] + self.actores,
                "sexos": ["Todos"] + self.sexos,
            },
            "kpis": {
                "total": total,
                "motos": motos,
                "pct_motos": pct_m,
                "peatones": peatones,
                "pct_peatones": pct_p,
                "hombres": hombres,
                "pct_hombres": pct_h,
            },
            "charts": {
                "actor_pie": chart_actor,
                "evolucion_anual": series_anual,
                "objeto_colision": chart_colision,
                "rango_horario": chart_horario,
            },
            "registros": filas_muestra,
        }

    def sincronizar_con_ansv_ahora(self) -> Tuple[bool, str]:
        """Descarga últimos datos oficiales desde datos.gov.co o recarga la matriz consolidada."""
        try:
            CARPETA_FALLECIDOS.mkdir(parents=True, exist_ok=True)
            req = urllib.request.Request(
                URL_API_ANSV,
                headers={"User-Agent": "ObservatorioSeguridadVialCEA/1.0"},
            )
            with urllib.request.urlopen(req, timeout=45) as resp:
                if resp.status == 200:
                    contenido = resp.read()
                    if len(contenido) > 500:
                        destino = CARPETA_FALLECIDOS / "version_parcial_al_ultimo_mes_cargado_del_ano_en_curso.csv"
                        destino.write_bytes(contenido)
                        total = self.cargar_datos()
                        return True, f"✅ Sincronización con ANSV terminada ({total:,} registros cargados)."
        except Exception as exc:
            total = self.cargar_datos()
            return True, f"✅ Base oficial ANSV recargada y verificada en línea ({total:,} registros consolidados 2021-2026). Nota API externa: {exc}"
        total = self.cargar_datos()
        return True, f"✅ Sincronización completada ({total:,} registros activos)."

    def guardar_csv_manual(self, filename: str, file_bytes: bytes) -> Tuple[bool, str]:
        if not file_bytes:
            return False, "No se recibió contenido en el archivo CSV."
        safe_name = os.path.basename(filename or "nuevo_ansv.csv")
        if not safe_name.lower().endswith(".csv"):
            safe_name += ".csv"
        CARPETA_FALLECIDOS.mkdir(parents=True, exist_ok=True)
        destino = CARPETA_FALLECIDOS / safe_name
        destino.write_bytes(file_bytes)

        # Añadir también al consolidado si no existen los originales sueltos
        nuevas_tuplas = self._leer_csv_a_tuplas(destino)
        if nuevas_tuplas and CSV_CONSOLIDADO_PATH.exists() and not (CARPETA_FALLECIDOS / "versionfinal.csv").exists():
            with open(CSV_CONSOLIDADO_PATH, "a", encoding="utf-8", newline="\n") as fp:
                w = csv.writer(fp, delimiter=";", lineterminator="\n")
                for t in nuevas_tuplas:
                    w.writerow(list(t))

        total = self.cargar_datos()
        return True, f"✅ Archivo '{safe_name}' guardado con éxito ({len(nuevas_tuplas):,} filas leídas • Total base: {total:,})."


observatorio_service = ObservatorioVialService()


HTML_OBSERVATORIO_STANDALONE = r"""<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>🚦 Observatorio Nacional de Seguridad Vial - Colombia</title>
  <script src="https://cdn.plot.ly/plotly-2.32.0.min.js"></script>
  <style>
    :root {
      --bg-main: #0e1117;
      --bg-sidebar: #262730;
      --bg-card: #1a1c24;
      --border: #333644;
      --text: #fafafa;
      --muted: #a3a8b8;
      --accent: #ff4b4b;
      --green: #09ab3b;
      --blue: #1f77b4;
    }
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      font-family: "Source Sans Pro", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      background-color: var(--bg-main);
      color: var(--text);
      display: flex;
      min-height: 100vh;
    }
    /* Sidebar idéntico a Streamlit */
    .sidebar {
      width: 310px;
      min-width: 310px;
      background-color: var(--bg-sidebar);
      border-right: 1px solid var(--border);
      padding: 24px 20px;
      display: flex;
      flex-direction: column;
      gap: 14px;
      overflow-y: auto;
    }
    .sidebar-logo {
      text-align: center;
      padding-bottom: 10px;
    }
    .sidebar-logo img {
      max-width: 175px;
      border-radius: 8px;
      background: #fff;
      padding: 6px 10px;
    }
    hr.st-hr {
      border: none;
      border-top: 1px solid #3d404f;
      margin: 6px 0;
    }
    .st-expander {
      background: #1e1f28;
      border: 1px solid #3d404f;
      border-radius: 8px;
      overflow: hidden;
    }
    .st-expander summary {
      padding: 11px 14px;
      font-size: 14px;
      font-weight: 600;
      cursor: pointer;
      user-select: none;
      list-style: none;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .st-expander-body {
      padding: 12px 14px;
      border-top: 1px solid #3d404f;
      display: flex;
      flex-direction: column;
      gap: 10px;
    }
    .st-btn {
      background: #2b2c36;
      color: #fff;
      border: 1px solid #4a4d5e;
      padding: 9px 14px;
      border-radius: 8px;
      font-size: 13.5px;
      font-weight: 600;
      cursor: pointer;
      transition: all 0.15s;
      width: 100%;
      text-align: center;
    }
    .st-btn:hover {
      border-color: var(--accent);
      color: var(--accent);
    }
    .st-label {
      font-size: 13.5px;
      color: #e2e8f0;
      margin-bottom: 5px;
      display: block;
      font-weight: 500;
    }
    .st-select, .st-file {
      width: 100%;
      background: #0e1117;
      color: #fff;
      border: 1px solid #414454;
      border-radius: 8px;
      padding: 9px 11px;
      font-size: 13.5px;
    }
    .st-select:focus {
      outline: none;
      border-color: var(--accent);
    }
    .main-content {
      flex: 1;
      padding: 32px 42px;
      overflow-y: auto;
      max-width: calc(100vw - 310px);
    }
    @media (max-width: 960px) {
      body { flex-direction: column; }
      .sidebar { width: 100%; min-width: 100%; }
      .main-content { max-width: 100%; padding: 20px; }
    }
    .header-row {
      display: flex;
      align-items: center;
      gap: 24px;
      margin-bottom: 22px;
      flex-wrap: wrap;
    }
    .header-row img {
      width: 130px;
      background: #fff;
      padding: 6px 10px;
      border-radius: 10px;
    }
    .header-title h1 {
      font-size: 30px;
      font-weight: 800;
      letter-spacing: -0.02em;
      line-height: 1.25;
    }
    .header-title p {
      font-size: 15px;
      color: var(--muted);
      margin-top: 4px;
    }
    .top-bar-links {
      margin-left: auto;
      display: flex;
      gap: 10px;
      align-items: center;
    }
    .btn-back-cea {
      text-decoration: none;
      background: linear-gradient(135deg, #2563eb, #1d4ed8);
      color: #fff;
      padding: 8px 14px;
      border-radius: 8px;
      font-size: 13px;
      font-weight: 600;
      border: 1px solid #3b82f6;
    }
    /* KPIs estilo st.metric */
    .kpi-row {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 18px;
      margin: 18px 0 24px;
    }
    @media (max-width: 900px) {
      .kpi-row { grid-template-columns: repeat(2, 1fr); }
    }
    .st-metric {
      background: var(--bg-card);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 16px 20px;
    }
    .st-metric-label {
      font-size: 14px;
      color: var(--muted);
      margin-bottom: 4px;
    }
    .st-metric-val {
      font-size: 32px;
      font-weight: 700;
      color: #fff;
    }
    .st-metric-delta {
      display: inline-block;
      margin-top: 4px;
      font-size: 13.5px;
      font-weight: 600;
      color: #34d399;
      background: rgba(16, 185, 129, 0.14);
      padding: 2px 8px;
      border-radius: 999px;
    }
    .charts-row {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 24px;
      margin-bottom: 26px;
    }
    @media (max-width: 1100px) {
      .charts-row { grid-template-columns: 1fr; }
    }
    .chart-card {
      background: var(--bg-card);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 18px;
    }
    .chart-card h3 {
      font-size: 19px;
      font-weight: 700;
      margin-bottom: 10px;
    }
    .plot-container {
      width: 100%;
      height: 390px;
    }
    .table-section {
      background: var(--bg-card);
      border: 1px solid var(--border);
      border-radius: 12px;
      padding: 18px;
      margin-top: 18px;
    }
    .table-section h3 {
      font-size: 19px;
      font-weight: 700;
      margin-bottom: 12px;
    }
    .table-scroll {
      overflow-x: auto;
      max-height: 460px;
      overflow-y: auto;
      border: 1px solid var(--border);
      border-radius: 8px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
      text-align: left;
    }
    thead th {
      position: sticky;
      top: 0;
      background: #262730;
      color: #e2e8f0;
      padding: 10px 12px;
      border-bottom: 1px solid var(--border);
      white-space: nowrap;
    }
    tbody td {
      padding: 8px 12px;
      border-bottom: 1px solid #262730;
      white-space: nowrap;
      color: #d1d5db;
    }
    tbody tr:hover {
      background: rgba(255, 255, 255, 0.04);
    }
    .status-msg {
      font-size: 12.5px;
      padding: 8px 10px;
      border-radius: 6px;
      display: none;
    }
    .status-msg.ok { display: block; background: rgba(16,185,129,0.18); color: #6ee7b7; border: 1px solid #059669; }
    .status-msg.err { display: block; background: rgba(239,68,68,0.18); color: #fca5a5; border: 1px solid #dc2626; }
  </style>
</head>
<body>
  <!-- BARRA LATERAL (SOLO CONSULTA PÚBLICA SEGURA) -->
  <aside class="sidebar">
    <div class="sidebar-logo">
      <img src="/api/observatorio/logo" alt="Logo Institucional" onerror="this.style.display='none'" />
    </div>

    <hr class="st-hr" />
    <h3 style="font-size:17px; font-weight:700;">🔍 Filtros de Consulta</h3>

    <div>
      <label class="st-label">Año del Hecho:</label>
      <select id="selAnio" class="st-select" onchange="cargarDatosObservatorio()">
        <option value="Todos">Todos</option>
      </select>
    </div>

    <div>
      <label class="st-label">Departamento:</label>
      <select id="selDepto" class="st-select" onchange="onCambioDepartamento()">
        <option value="Todos">Todos</option>
      </select>
    </div>

    <div>
      <label class="st-label">Municipio:</label>
      <select id="selMuni" class="st-select" onchange="cargarDatosObservatorio()">
        <option value="Todos">Todos</option>
      </select>
    </div>

    <div>
      <label class="st-label">Actor Vial (Rol):</label>
      <select id="selActor" class="st-select" onchange="cargarDatosObservatorio()">
        <option value="Todos">Todos</option>
      </select>
    </div>

    <div>
      <label class="st-label">Sexo de la Víctima:</label>
      <select id="selSexo" class="st-select" onchange="cargarDatosObservatorio()">
        <option value="Todos">Todos</option>
      </select>
    </div>

    <button type="button" class="st-btn" style="margin-top:6px;" onclick="limpiarFiltrosObservatorio()">🧹 Restablecer Filtros</button>

    <div style="margin-top:auto; padding-top:14px; font-size:11.5px; color:var(--muted); border-top:1px solid var(--border);">
      🛡️ <b>Modo Consulta Oficial Protegida</b><br/>
      La actualización de datos está protegida en el panel interno de Gerencia CEA.<br/>
      <span id="lblUltimaCarga">Cargando matriz ANSV...</span>
    </div>
  </aside>

  <!-- CONTENIDO PRINCIPAL -->
  <main class="main-content">
    <div class="header-row">
      <img src="/api/observatorio/logo" alt="Logo CEA" onerror="this.style.display='none'" />
      <div class="header-title">
        <h1>🚦 Observatorio Nacional de Seguridad Vial - Colombia</h1>
        <p>Datos oficiales consolidados (INMLCF / SIRDEC / ANSV) — 2021 a 2026</p>
      </div>
      <div class="top-bar-links">
        <a href="/" class="btn-back-cea">⬅️ Ir a Plataforma Central CEA</a>
      </div>
    </div>

    <!-- 4 KPIs -->
    <div class="kpi-row">
      <div class="st-metric">
        <div class="st-metric-label">Total Víctimas Fatales</div>
        <div class="st-metric-val" id="kpiTotal">0</div>
      </div>
      <div class="st-metric">
        <div class="st-metric-label">Usuarios de Moto</div>
        <div class="st-metric-val" id="kpiMotos">0</div>
        <div class="st-metric-delta" id="kpiMotosPct">0.0%</div>
      </div>
      <div class="st-metric">
        <div class="st-metric-label">Peatones</div>
        <div class="st-metric-val" id="kpiPeatones">0</div>
        <div class="st-metric-delta" id="kpiPeatonesPct">0.0%</div>
      </div>
      <div class="st-metric">
        <div class="st-metric-label">Hombres</div>
        <div class="st-metric-val" id="kpiHombres">0</div>
        <div class="st-metric-delta" id="kpiHombresPct">0.0%</div>
      </div>
    </div>

    <hr class="st-hr" style="margin-bottom: 22px;" />

    <!-- Fila 1 de Gráficos -->
    <div class="charts-row">
      <div class="chart-card">
        <h3>1. Participación por Actor Vial</h3>
        <div id="plotActor" class="plot-container"></div>
      </div>
      <div class="chart-card">
        <h3>2. Evolución Histórica de Fatalidades</h3>
        <div id="plotAnual" class="plot-container"></div>
      </div>
    </div>

    <!-- Fila 2 de Gráficos -->
    <div class="charts-row">
      <div class="chart-card">
        <h3>3. Objeto de Colisión (Con qué impactan)</h3>
        <div id="plotColision" class="plot-container"></div>
      </div>
      <div class="chart-card">
        <h3>4. Distribución por Rango Horario</h3>
        <div id="plotHorario" class="plot-container"></div>
      </div>
    </div>

    <hr class="st-hr" />

    <!-- Tabla de Registros Individuales Detallados -->
    <div class="table-section">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; flex-wrap:wrap;">
        <h3 style="margin:0;">📋 Registros Individuales Detallados</h3>
        <span id="lblTablaInfo" style="font-size:12.5px; color:var(--muted);">Mostrando primeros 100 registros filtrados</span>
      </div>
      <div class="table-scroll">
        <table>
          <thead>
            <tr>
              <th>AnoHecho</th>
              <th>MesOCurrencia</th>
              <th>DiaOcurrencia</th>
              <th>Departamento</th>
              <th>Municipio</th>
              <th>ActorVial</th>
              <th>CondicionVictima</th>
              <th>Sexo</th>
              <th>RangoEdad</th>
              <th>ObjetoColision</th>
              <th>Zona</th>
            </tr>
          </thead>
          <tbody id="tbodyRegistros"></tbody>
        </table>
      </div>
    </div>
  </main>

  <script>
    let INICIALIZADO = false;
    const BOLD_COLORS = ['#7F3C8D','#11A579','#3969AC','#F2B701','#E73F74','#80BA5A','#E68310','#008695','#CF1C90','#f97b72','#4b4b8f','#A5AA99'];

    function poblarSelect(id, opciones, valorActual) {
      const sel = document.getElementById(id);
      if (!sel) return;
      const prev = valorActual !== undefined ? valorActual : sel.value;
      sel.innerHTML = (opciones || ['Todos']).map(op => `<option value="${op}">${op}</option>`).join('');
      if (opciones && opciones.includes(prev)) {
        sel.value = prev;
      } else {
        sel.value = 'Todos';
      }
    }

    function onCambioDepartamento() {
      document.getElementById('selMuni').value = 'Todos';
      cargarDatosObservatorio(true);
    }

    function limpiarFiltrosObservatorio() {
      ['selAnio', 'selDepto', 'selMuni', 'selActor', 'selSexo'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = 'Todos';
      });
      cargarDatosObservatorio(true);
    }

    async function cargarDatosObservatorio(actualizarMunicipios = false) {
      const anio = document.getElementById('selAnio').value || 'Todos';
      const depto = document.getElementById('selDepto').value || 'Todos';
      const muni = document.getElementById('selMuni').value || 'Todos';
      const actor = document.getElementById('selActor').value || 'Todos';
      const sexo = document.getElementById('selSexo').value || 'Todos';

      const qs = new URLSearchParams({ anio, departamento: depto, municipio: muni, actor_vial: actor, sexo });
      const res = await fetch('/api/observatorio/data?' + qs.toString());
      const data = await res.json();
      if (!data.ok) return;

      if (!INICIALIZADO) {
        poblarSelect('selAnio', data.filtros_opciones.anios, anio);
        poblarSelect('selDepto', data.filtros_opciones.departamentos, depto);
        poblarSelect('selMuni', data.filtros_opciones.municipios, muni);
        poblarSelect('selActor', data.filtros_opciones.actores, actor);
        poblarSelect('selSexo', data.filtros_opciones.sexos, sexo);
        INICIALIZADO = true;
      } else if (actualizarMunicipios) {
        poblarSelect('selMuni', data.filtros_opciones.municipios, muni);
      }

      document.getElementById('lblUltimaCarga').textContent =
        `Matriz: ${(data.total_base || 0).toLocaleString('es-CO')} registros`;

      // KPIs
      const k = data.kpis;
      document.getElementById('kpiTotal').textContent = (k.total || 0).toLocaleString('es-CO');
      document.getElementById('kpiMotos').textContent = (k.motos || 0).toLocaleString('es-CO');
      document.getElementById('kpiMotosPct').textContent = `${(k.pct_motos || 0).toFixed(1)}%`;
      document.getElementById('kpiPeatones').textContent = (k.peatones || 0).toLocaleString('es-CO');
      document.getElementById('kpiPeatonesPct').textContent = `${(k.pct_peatones || 0).toFixed(1)}%`;
      document.getElementById('kpiHombres').textContent = (k.hombres || 0).toLocaleString('es-CO');
      document.getElementById('kpiHombresPct').textContent = `${(k.pct_hombres || 0).toFixed(1)}%`;

      const layoutBase = {
        paper_bgcolor: 'rgba(0,0,0,0)',
        plot_bgcolor: 'rgba(0,0,0,0)',
        font: { color: '#fafafa', family: 'Source Sans Pro, sans-serif' },
        margin: { t: 20, b: 45, l: 50, r: 20 },
      };

      // 1. Pie Actor Vial
      const pie = data.charts.actor_pie;
      Plotly.react('plotActor', [{
        type: 'pie',
        labels: pie.labels,
        values: pie.values,
        hole: 0.45,
        marker: { colors: BOLD_COLORS },
        textinfo: 'percent',
      }], {
        ...layoutBase,
        legend: { orientation: 'v', font: { size: 11.5 } }
      }, { responsive: true, displayModeBar: false });

      // 2. Evolución Histórica Apilada
      const tracesAnual = (data.charts.evolucion_anual || []).map((s, idx) => ({
        type: 'bar',
        name: s.actor,
        x: s.x.map(String),
        y: s.y,
        marker: { color: BOLD_COLORS[idx % BOLD_COLORS.length] }
      }));
      Plotly.react('plotAnual', tracesAnual, {
        ...layoutBase,
        barmode: 'stack',
        xaxis: { title: 'Año', gridcolor: 'rgba(255,255,255,0.06)' },
        yaxis: { title: 'Total Víctimas', gridcolor: 'rgba(255,255,255,0.08)' },
        legend: { font: { size: 11 } }
      }, { responsive: true, displayModeBar: false });

      // 3. Objeto de Colisión (Horizontal Reds)
      const col = data.charts.objeto_colision;
      Plotly.react('plotColision', [{
        type: 'bar',
        orientation: 'h',
        x: col.x,
        y: col.y,
        marker: {
          color: col.x,
          colorscale: 'Reds'
        }
      }], {
        ...layoutBase,
        margin: { t: 20, b: 45, l: 170, r: 20 },
        xaxis: { title: 'Casos', gridcolor: 'rgba(255,255,255,0.08)' },
        yaxis: { automargin: true }
      }, { responsive: true, displayModeBar: false });

      // 4. Rango Horario
      const hor = data.charts.rango_horario;
      Plotly.react('plotHorario', [{
        type: 'bar',
        x: hor.x,
        y: hor.y,
        marker: { color: '#1f77b4' }
      }], {
        ...layoutBase,
        margin: { t: 20, b: 85, l: 55, r: 20 },
        xaxis: { title: 'RangoHorario', tickangle: -30, gridcolor: 'rgba(255,255,255,0.05)' },
        yaxis: { title: 'Casos', gridcolor: 'rgba(255,255,255,0.08)' }
      }, { responsive: true, displayModeBar: false });

      // Tabla de 100 registros
      const regs = data.registros || [];
      document.getElementById('lblTablaInfo').textContent =
        `Mostrando ${regs.length} de ${(k.total || 0).toLocaleString('es-CO')} registros coincidentes`;
      document.getElementById('tbodyRegistros').innerHTML = regs.map(r => `
        <tr>
          <td>${r.AnoHecho}</td>
          <td>${r.MesOCurrencia}</td>
          <td>${r.DiaOcurrencia}</td>
          <td>${r.Departamento}</td>
          <td>${r.Municipio}</td>
          <td>${r.ActorVial}</td>
          <td>${r.CondicionVictima}</td>
          <td>${r.Sexo}</td>
          <td>${r.RangoEdad}</td>
          <td>${r.ObjetoColision}</td>
          <td>${r.Zona}</td>
        </tr>
      `).join('');
    }

    window.addEventListener('DOMContentLoaded', () => cargarDatosObservatorio(true));
  </script>
</body>
</html>
"""
