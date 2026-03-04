# app.py ✅ BOSS-READY + INYECCIÓN DIRECTA DE RASTREO
# ✅ Fix definitivo: Python ahora extrae "En preparacion" y "En entrega" directamente
#    de la tabla dbo.pedido_telefonia_rastreo y hace un JOIN (merge) automático.

import streamlit as st
import pandas as pd
import numpy as np
import unicodedata
from datetime import date, datetime, time, timedelta
import pyodbc
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from openpyxl.utils import get_column_letter

# -------------------------------------------------
# GLOBAL CONSTANTS
# -------------------------------------------------
EXCLUDED_VENDOR = "ABASTECEDORA Y SUMINISTROS ORTEGA/ISABEL VALDEZ JIMENEZ"
DEFAULT_LOOKBACK_DAYS = 30

FLOW_ORDER = [
    "Total",
    "Nuevo",
    "Back Office",
    "Solicitado",
    "En preparacion",
    "En entrega",
    "Reprogramado",
    "Entregado",
]
FLOW_STAGES_NO_TOTAL = ["Nuevo", "Back Office", "Solicitado", "En preparacion", "En entrega", "Reprogramado", "Entregado"]

# -------------------------------------------------
# STREAMLIT CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="Órdenes — Flujo & Tiempos",
    page_icon="⏱️",
    layout="wide",
)

# -------------------------------------------------
# STYLES
# -------------------------------------------------
st.markdown(
    """
<style>
html, body, [class*="css"] { font-family: "Segoe UI", system-ui, sans-serif; }
.block-container { padding-top: 1.1rem; padding-bottom: 2rem; }
h1 { font-weight: 900 !important; letter-spacing: -0.3px; }
h2, h3 { font-weight: 800 !important; }

/* --- Flow pills --- */
.flow-wrap {
  display:flex;
  gap: 10px;
  flex-wrap: wrap;
  align-items: center;
  padding: 10px 8px;
  border-radius: 16px;
  border: 1px solid rgba(148,163,184,0.35);
  background: rgba(148,163,184,0.06);
}
.flow-pill{
  min-width: 120px;
  text-align:center;
  padding: 10px 14px;
  border-radius: 12px;
  background: rgba(148,163,184,0.45);
  color: white;
  font-weight: 800;
  line-height: 1.05;
  box-shadow: 0 8px 16px rgba(15,23,42,0.10);
}
.flow-pill small{
  display:block;
  font-weight: 800;
  opacity: 0.95;
}
@media (prefers-color-scheme: dark) {
  .flow-wrap { background: rgba(15,23,42,0.55); border-color: rgba(148,163,184,0.45); }
  .flow-pill { background: rgba(148,163,184,0.35); box-shadow: 0 10px 22px rgba(0,0,0,0.45); }
}

/* KPI cards */
.kpi-row { display:flex; gap:12px; flex-wrap:wrap; margin-top: 10px; }
.kpi {
  background: rgba(255,255,255,0.92);
  border: 1px solid rgba(15,23,42,0.10);
  box-shadow: 0 10px 26px rgba(15,23,42,0.10);
  border-radius: 16px;
  padding: 14px 16px;
  min-width: 220px;
  flex: 1;
}
.kpi .label { font-size: 0.92rem; opacity:0.85; }
.kpi .value { font-size: 1.7rem; font-weight: 900; margin-top:6px; }
.kpi .sub { font-size: 0.85rem; opacity:0.75; margin-top:4px; }

@media (prefers-color-scheme: dark) {
  .kpi{
    background: rgba(15,23,42,0.92);
    border-color: rgba(148,163,184,0.55);
    box-shadow: 0 12px 34px rgba(0,0,0,0.55);
  }
}

/* Download button */
div[data-testid="stDownloadButton"] > button {
    border-radius: 999px;
    background: linear-gradient(90deg,#0ea5e9,#6366f1);
    color: #f9fafb;
    border: none;
    padding: 0.45rem 1.2rem;
    font-weight: 800;
}
div[data-testid="stDownloadButton"] > button:hover { filter: brightness(1.05); }

/* Plotly transparent */
.js-plotly-plot .plotly .main-svg { background-color: rgba(0,0,0,0) !important; }
</style>
""",
    unsafe_allow_html=True,
)

# -------------------------------------------------
# SESSION STATE
# -------------------------------------------------
if "last_refresh" not in st.session_state:
    st.session_state["last_refresh"] = datetime.now()

# -------------------------------------------------
# EXCEL EXPORT
# -------------------------------------------------
def _safe_sheet_name(name: str) -> str:
    invalid = ['\\', '/', '*', '[', ']', ':', '?']
    out = str(name)
    for ch in invalid:
        out = out.replace(ch, "_")
    out = out.strip() or "Sheet"
    return out[:31]

def dfs_to_excel_bytes(sheets: dict) -> bytes:
    output = BytesIO()
    used_names = set()

    def unique_sheet_name(base: str) -> str:
        base = _safe_sheet_name(base)
        if base not in used_names:
            used_names.add(base)
            return base
        i = 2
        while True:
            suffix = f"_{i}"
            cand = (base[: (31 - len(suffix))] + suffix) if len(base) + len(suffix) > 31 else (base + suffix)
            if cand not in used_names:
                used_names.add(cand)
                return cand
            i += 1

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for raw_name, df in sheets.items():
            sheet_name = unique_sheet_name(raw_name)
            if df is None:
                df = pd.DataFrame()
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            max_row = ws.max_row
            max_col = ws.max_column
            if max_col > 0 and max_row > 0:
                ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"
                for col_idx in range(1, max_col + 1):
                    col_letter = get_column_letter(col_idx)
                    max_length = 0
                    for cell in ws[col_letter]:
                        if cell.value is not None:
                            max_length = max(max_length, len(str(cell.value)))
                    ws.column_dimensions[col_letter].width = max_length + 2

    output.seek(0)
    return output.getvalue()

# -------------------------------------------------
# DISPLAY HELPERS
# -------------------------------------------------
def fmt_int(x):
    try:
        return f"{int(x):,}"
    except Exception:
        return "0"

def fmt_pct(x):
    try:
        return f"{float(x) * 100:.1f}%"
    except Exception:
        return "0.0%"

def fmt_timedelta(td) -> str:
    if td is None or pd.isna(td):
        return "—"
    if isinstance(td, (float, int)):
        td = pd.to_timedelta(td, unit="s")
    if not isinstance(td, pd.Timedelta):
        try:
            td = pd.to_timedelta(td)
        except Exception:
            return "—"
    if td < pd.Timedelta(0):
        return "—"

    days = td.days
    secs = int(td.seconds)
    hh = secs // 3600
    mm = (secs % 3600) // 60
    return f"{days}d {hh:02d}:{mm:02d}" if days > 0 else f"{hh:02d}:{mm:02d}"

def fmt_done_or_in_process(td_done, td_age):
    if td_done is not None and pd.notna(td_done):
        return fmt_timedelta(td_done)
    if td_age is not None and pd.notna(td_age):
        return f"En proceso · {fmt_timedelta(td_age)}"
    return "—"

def td_to_hours(td) -> float:
    if td is None or pd.isna(td):
        return np.nan
    try:
        return float(pd.to_timedelta(td).total_seconds() / 3600.0)
    except Exception:
        return np.nan

def kpi_card(label, value, sub=None):
    sub_html = f'<div class="sub">{sub}</div>' if sub else ""
    st.markdown(
        f"""
        <div class="kpi">
          <div class="label">{label}</div>
          <div class="value">{value}</div>
          {sub_html}
        </div>
        """,
        unsafe_allow_html=True,
    )

def render_flow_pills(counts: dict):
    html = '<div class="flow-wrap">'
    for k in FLOW_ORDER:
        v = counts.get(k, 0)
        html += f'<div class="flow-pill">{k}<small>({int(v)})</small></div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)

# -------------------------------------------------
# TEXT NORMALIZATION & SANITIZATION (ANTI-1900)
# -------------------------------------------------
def _norm_col(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf-8", "ignore")
    s = s.replace("_", " ")
    s = " ".join(s.split())
    return s

def canon_estatus(x: str) -> str:
    s = str(x).strip()
    sn = _norm_col(s)
    if sn == "nuevo":
        return "Nuevo"
    if sn in ["back office", "backoffice"]:
        return "Back Office"
    if sn in ["solicitado", "solicitada"]:
        return "Solicitado"
    if sn in ["en preparacion", "en preparación", "en preparacion."]:
        return "En preparacion"
    if sn == "en entrega":
        return "En entrega"
    if sn in ["reprogramado", "reprogramada"]:
        return "Reprogramado"
    if sn in ["entregado", "entregada"]:
        return "Entregado"
    return str(x).strip()

def sanitize_dates(dt_series: pd.Series) -> pd.Series:
    if not pd.api.types.is_datetime64_any_dtype(dt_series):
        dt_series = pd.to_datetime(dt_series, errors="coerce")
    return dt_series.where(dt_series >= pd.Timestamp('2000-01-01'), pd.NaT)

def _extract_datetime_text(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace({
        "nan": "", "none": "", "nat": "", "NaN": "", "None": "", "NaT": "", "<NA>": "", "null": "", "Null": "",
        "1900-01-01 00:00:00": "", "1900-01-01 00:00:00.000": "", "1900-01-01": "", "1900-01-01T00:00:00": ""
    })
    s = s.where(s != "", np.nan)

    if s.notna().any():
        pat = r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)|(\d{4}[/-]\d{1,2}[/-]\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)"
        ext = s.astype(str).str.extract(pat)
        ext = ext[0].fillna(ext[1])
        s2 = ext.where(ext.notna(), s)
        return s2
    return s

def parse_dt_both(series: pd.Series) -> tuple:
    dt_native = pd.to_datetime(series, errors='coerce')
    s2 = _extract_datetime_text(series)
    dt_df = pd.to_datetime(s2, errors="coerce", dayfirst=True)
    dt_mf = pd.to_datetime(s2, errors="coerce", dayfirst=False)

    dt_df = dt_df.where(dt_df.notna(), dt_native)
    dt_mf = dt_mf.where(dt_mf.notna(), dt_native)
    return sanitize_dates(dt_df), sanitize_dates(dt_mf)

def choose_dt_rowwise(dt_df: pd.Series, dt_mf: pd.Series, created: pd.Series | None=None, bo: pd.Series | None=None) -> pd.Series:
    out = dt_df.copy()
    out = out.where(~(dt_df.isna() & dt_mf.notna()), dt_mf)
    return sanitize_dates(out)

# ✅ FIX NUEVO (mínimo): selector robusto para ACTIVACIÓN usando BO como referencia + anti-futuro
def choose_dt_activation_rowwise(dt_df: pd.Series, dt_mf: pd.Series, bo: pd.Series | None = None) -> pd.Series:
    """
    Selección robusta para ACTIVACIÓN:
    - Descarta fechas en el futuro (tolerancia 1 día)
    - Descarta fechas < BO (si BO existe)
    - Si ambas (DF/MF) son válidas, elige la más cercana a BO
    """
    now_ts = pd.Timestamp(datetime.now())
    max_ts = now_ts + pd.Timedelta(days=1)  # tolerancia pequeña por TZ / retrasos

    out = dt_df.copy()
    out = out.where(~(dt_df.isna() & dt_mf.notna()), dt_mf)

    valid_df = dt_df.notna() & (dt_df <= max_ts)
    valid_mf = dt_mf.notna() & (dt_mf <= max_ts)

    if bo is not None:
        bo_dt = pd.to_datetime(bo, errors="coerce")

        valid_df &= bo_dt.isna() | (dt_df >= bo_dt)
        valid_mf &= bo_dt.isna() | (dt_mf >= bo_dt)

        out = out.where(~(valid_mf & ~valid_df), dt_mf)

        both = valid_df & valid_mf & bo_dt.notna()
        if both.any():
            diff_df = (dt_df - bo_dt).abs()
            diff_mf = (dt_mf - bo_dt).abs()
            out = out.where(~(both & (diff_mf < diff_df)), dt_mf)
    else:
        out = out.where(~(valid_mf & ~valid_df), dt_mf)

    out = out.where(out <= max_ts, pd.NaT)
    return sanitize_dates(out)

# ✅ FIX NUEVO (mínimo): selector robusto para FECHA CREACIÓN usando BO como referencia + anti-futuro
def choose_dt_created_rowwise(
    dt_df: pd.Series,
    dt_mf: pd.Series,
    bo: pd.Series | None = None,
    window_start: date | None = None,
    window_end: date | None = None,
) -> pd.Series:
    """
    Selección robusta para FECHA CREACIÓN:
    - Descarta fechas en el futuro (tolerancia 1 día)
    - Si BO existe, creación debe ser <= BO
    - Si ambas (DF/MF) son válidas, elige la más cercana a BO (la mayor pero <= BO)
    - Si no hay BO, usa ventana para preferir el parse dentro del periodo
    """
    now_ts = pd.Timestamp(datetime.now())
    max_ts = now_ts + pd.Timedelta(days=1)

    out = dt_df.copy()
    out = out.where(~(dt_df.isna() & dt_mf.notna()), dt_mf)

    valid_df = dt_df.notna() & (dt_df <= max_ts)
    valid_mf = dt_mf.notna() & (dt_mf <= max_ts)

    if bo is not None:
        bo_dt = pd.to_datetime(bo, errors="coerce")
        valid_df &= bo_dt.isna() | (dt_df <= bo_dt)
        valid_mf &= bo_dt.isna() | (dt_mf <= bo_dt)

        out = out.where(~(valid_mf & ~valid_df), dt_mf)

        both = valid_df & valid_mf & bo_dt.notna()
        if both.any():
            # elige la mayor (más cercana a BO) pero <= BO
            out = out.where(~(both & (dt_mf > dt_df)), dt_mf)

    else:
        if window_start is not None and window_end is not None:
            w0 = pd.Timestamp(window_start)
            w1 = pd.Timestamp(window_end) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)
            in_df = dt_df.between(w0, w1)
            in_mf = dt_mf.between(w0, w1)
            out = out.where(~(in_mf & ~in_df), dt_mf)

        out = out.where(~(valid_mf & ~valid_df), dt_mf)

    out = out.where(out <= max_ts, pd.NaT)
    return sanitize_dates(out)

def parse_backoffice_datetime(series: pd.Series, window_start: date | None = None, window_end: date | None = None) -> pd.Series:
    dt_df, dt_mf = parse_dt_both(series)

    if window_start is None or window_end is None:
        return dt_df

    w0 = pd.Timestamp(window_start)
    w1 = pd.Timestamp(window_end) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)

    in1 = dt_df.between(w0, w1)
    in2 = dt_mf.between(w0, w1)

    out = dt_df.copy()
    out = out.where(~(in2 & ~in1), dt_mf)
    out = out.where(~(dt_df.isna() & dt_mf.notna()), dt_mf)
    return out

def choose_backoffice_dt(df: pd.DataFrame, window_start: date, window_end: date) -> pd.Series:
    if "BO_DT_DF" in df.columns and "BO_DT_MF" in df.columns:
        dt_dayfirst = df["BO_DT_DF"]
        dt_monthfirst = df["BO_DT_MF"]

        w0 = pd.Timestamp(window_start)
        w1 = pd.Timestamp(window_end) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)

        in1 = dt_dayfirst.between(w0, w1)
        in2 = dt_monthfirst.between(w0, w1)

        out = dt_dayfirst.copy()
        out = out.where(~(in2 & ~in1), dt_monthfirst)
        out = out.where(~(dt_dayfirst.isna() & dt_monthfirst.notna()), dt_monthfirst)
        return sanitize_dates(out)
    return parse_backoffice_datetime(df["Back Office"], window_start=window_start, window_end=window_end)

def _reference_end_dt(fecha_fin: date) -> datetime:
    return datetime.now() if fecha_fin == date.today() else datetime.combine(fecha_fin, time(23, 59, 59))

def pick_activation_dt(df: pd.DataFrame) -> tuple:
    if df is None or df.empty:
        return pd.Series(pd.NaT, index=df.index), None

    cols_norm = {_norm_col(c): c for c in df.columns}
    created = df["CREATED_DT"] if "CREATED_DT" in df.columns else None
    bo = df["BO_DT"] if "BO_DT" in df.columns else None

    def _try_col(colname: str):
        if colname not in df.columns:
            return None
        dt_df, dt_mf = parse_dt_both(df[colname])
        chosen = choose_dt_activation_rowwise(dt_df, dt_mf, bo=bo)  # ✅ CAMBIO MÍNIMO AQUÍ
        return chosen if chosen.notna().any() else None

    for key in ["fecha activacion", "fecha de activacion", "fecha activación"]:
        col = cols_norm.get(key)
        if col:
            chosen = _try_col(col)
            if chosen is not None:
                return chosen, col

    for key in ["fecha venta", "fecha de venta", "fecha_venta"]:
        col = cols_norm.get(key)
        if col:
            chosen = _try_col(col)
            if chosen is not None:
                return chosen, col

    col = cols_norm.get("venta")
    if col:
        chosen = _try_col(col)
        if chosen is not None:
            return chosen, col

    return pd.Series(pd.NaT, index=df.index), None

def pick_activation_dt(df: pd.DataFrame) -> tuple:
    if df is None or df.empty:
        return pd.Series(pd.NaT, index=df.index), None

    cols_norm = {_norm_col(c): c for c in df.columns}
    bo = df["BO_DT"] if "BO_DT" in df.columns else None

    def _try_col(colname: str):
        if colname not in df.columns:
            return None
        dt_df, dt_mf = parse_dt_both(df[colname])
        chosen = choose_dt_activation_rowwise(dt_df, dt_mf, bo=bo)
        return chosen if chosen.notna().any() else None

    for key in ["fecha activacion", "fecha de activacion", "fecha activación"]:
        col = cols_norm.get(key)
        if col:
            chosen = _try_col(col)
            if chosen is not None:
                return chosen, col

    for key in ["fecha venta", "fecha de venta", "fecha_venta"]:
        col = cols_norm.get(key)
        if col:
            chosen = _try_col(col)
            if chosen is not None:
                return chosen, col

    col = cols_norm.get("venta")
    if col:
        chosen = _try_col(col)
        if chosen is not None:
            return chosen, col

    return pd.Series(pd.NaT, index=df.index), None

def pick_stage_dt_from_columns(df: pd.DataFrame, stage: str, created: pd.Series, bo: pd.Series) -> tuple:
    if df is None or df.empty:
        return pd.Series(pd.NaT, index=df.index), None

    stage_n = _norm_col(stage)

    kw = {
        "solicitado": ["solicitado", "solicit", "solic"],
        "en preparacion": ["preparacion", "preparación", "prep", "armado", "empaque"],
        "en entrega": ["en entrega", "entrega", "ruta", "transito", "trayecto"],
        "reprogramado": ["reprogramado", "reprog"],
        "entregado": ["entregado", "fecha entrega", "fecha entregado", "fin"],
    }
    keys = kw.get(stage_n, [stage_n])

    cols = []
    norm_cols = {c: _norm_col(c) for c in df.columns}

    for c, nc in norm_cols.items():
        if stage_n == "en entrega" and ("fecha entrega" in nc or "entregado" in nc):
            continue
        if stage_n == "entregado" and nc == "en entrega":
            continue

        if any(k in nc for k in keys):
            cols.append(c)

    best_col = None
    best_dt = pd.Series(pd.NaT, index=df.index)
    best_nonnull = -1

    for c in cols:
        dt_df, dt_mf = parse_dt_both(df[c])
        chosen = choose_dt_rowwise(dt_df, dt_mf, created=created, bo=bo)
        n = int(chosen.notna().sum())
        if n > best_nonnull:
            best_nonnull = n
            best_dt = chosen
            best_col = c

    return best_dt, best_col

# -------------------------------------------------
# FLOW COUNTS
# -------------------------------------------------
def compute_flow_counts(df: pd.DataFrame) -> dict:
    counts = {k: 0 for k in FLOW_ORDER}
    if df is None or df.empty or "Estatus" not in df.columns:
        return counts
    E = df["Estatus"].astype(str)
    counts["Total"] = int(len(df))
    for stg in FLOW_STAGES_NO_TOTAL:
        counts[stg] = int((E == stg).sum())
    return counts

# -------------------------------------------------
# DB CONNECTION
# -------------------------------------------------
@st.cache_resource
def get_connection():
    cfg = st.secrets["sql"]
    conn_str = (
        f"DRIVER={{{cfg['driver']}}};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['user']};"
        f"PWD={cfg['password']};"
        "Encrypt=yes;"
        "TrustServerCertificate=yes;"
        "MARS_Connection=yes;"
    )
    return pyodbc.connect(conn_str, autocommit=True)

# -------------------------------------------------
# LOAD DATA FROM SQL
# -------------------------------------------------
@st.cache_data
def load_hoja1():
    sql = """
    SELECT DISTINCT
        e.[Nombre Completo] AS NombreCompleto,
        e.[Jefe Inmediato]  AS JefeDirecto,
        e.[Region],
        e.[SubRegion],
        e.[Plaza],
        e.[Tienda],
        e.[Puesto],
        e.[Canal de Venta],
        e.[Tipo Tienda],
        e.[Operacion],
        e.[Estatus]
    FROM reporte_empleado('EMPRESA_MAESTRA',1,'','') AS e
    WHERE
        e.[Canal de Venta] = 'ATT'
        AND e.[Operacion]   = 'CONTACT CENTER'
        AND e.[Tipo Tienda] = 'VIRTUAL'
        AND e.[Puesto] IN (
            'ASESOR TELEFONICO',
            'ASESOR TELEFONICO 7500',
            'EJECUTIVO TELEFONICO 6500 AM',
            'EJECUTIVO TELEFONICO 6500 PM',
            'SUPERVISOR DE CONTACT CENTER'
        )
        AND e.[Estatus] = 'ACTIVO';
    """
    conn = get_connection()
    df = pd.read_sql(sql, conn)

    text_cols = [
        "NombreCompleto","JefeDirecto","Region","SubRegion","Plaza","Tienda",
        "Puesto","Canal de Venta","Tipo Tienda","Operacion","Estatus"
    ]
    for col in text_cols:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": np.nan, "None": np.nan})

    df["JefeDirecto"] = df["JefeDirecto"].fillna("").str.strip().replace("", "ENCUBADORA")
    df["Coordinador"] = df["JefeDirecto"]
    df = df[df["NombreCompleto"].str.upper() != EXCLUDED_VENDOR].copy()
    return df

@st.cache_data
def load_consulta1(fecha_ini: date, fecha_fin: date) -> pd.DataFrame:
    fi = fecha_ini.strftime("%Y%m%d")
    ff = fecha_fin.strftime("%Y%m%d")

    sql = f"""
    SELECT
        *,
        [Tienda solicita] AS Centro
    FROM reporte_programacion_entrega('empresa_maestra', 4, '{fi}', '{ff}')
    WHERE
        [Tienda solicita] LIKE 'EXP ATT C CENTER%' AND
        [Estatus] IN ('Nuevo','Back Office','Solicitado','En preparacion','En entrega','Reprogramado','Entregado','Canc Error');
    """
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SET NOCOUNT ON; SET ANSI_WARNINGS OFF;")
    df = pd.read_sql(sql, conn)
    cur.execute("SET ANSI_WARNINGS ON;")
    cur.close()
    return df

# ✅ FIX NUEVO: Función para conectarse directo a la tabla de bitácora
@st.cache_data
def load_rastreo_extra(fecha_ini: date, fecha_fin: date) -> pd.DataFrame:
    fi = fecha_ini.strftime("%Y%m%d")
    ff = fecha_fin.strftime("%Y%m%d")

    # Esta consulta cruza la tabla de bitácora con los datos filtrados para sacar SOLO lo que nos importa
    sql = f"""
    SELECT 
        r.id_pedido_telefonia AS Programacion, 
        r.accion, 
        MAX(r.fecha) AS fecha_rastreo
    FROM dbo.pedido_telefonia_rastreo r
    INNER JOIN reporte_programacion_entrega('empresa_maestra', 4, '{fi}', '{ff}') t
        ON r.id_pedido_telefonia = t.Programacion
    WHERE r.accion IN ('En preparacion', 'En entrega', 'Reprogramado')
    GROUP BY r.id_pedido_telefonia, r.accion
    """
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("SET NOCOUNT ON; SET ANSI_WARNINGS OFF;")
    df = pd.read_sql(sql, conn)
    cur.execute("SET ANSI_WARNINGS ON;")
    cur.close()
    return df

# -------------------------------------------------
# TRANSFORM
# -------------------------------------------------
def transform_consulta1(df_raw: pd.DataFrame, hoja: pd.DataFrame, rastreo_extra: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()

    # ✅ FIX NUEVO: Pegamos las columnas faltantes directamente desde la bitácora
    if rastreo_extra is not None and not rastreo_extra.empty:
        piv = rastreo_extra.pivot_table(index="Programacion", columns="accion", values="fecha_rastreo", aggfunc="max").reset_index()

        rename_map = {
            "En preparacion": "Fecha En preparacion Exacta",
            "En entrega": "Fecha En entrega Exacta",
            "Reprogramado": "Fecha Reprogramado Exacta"
        }
        piv.rename(columns=rename_map, inplace=True)

        if "Programacion" in df.columns:
            df = df.merge(piv, on="Programacion", how="left")

    clean_cols = [
        "Centro", "Estatus", "Back Office", "Vendedor", "Cliente",
        "Nuevo", "Solicitado", "En preparacion", "En preparación",
        "En entrega", "Reprogramado", "Entregado", "Fecha creacion", "Venta"
    ]
    for col in df.columns:
        if col in clean_cols or "fecha" in col.lower() or any(stg in col for stg in ["Solicitado", "preparacion", "entrega", "Reprogramado", "Entregado"]):
            df[col] = df[col].astype(str).str.strip().replace({
                "nan": np.nan, "None": np.nan, "NaT": np.nan, "nat": np.nan, "none": np.nan, "<NA>": np.nan, "null": np.nan,
                "1900-01-01 00:00:00": np.nan, "1900-01-01 00:00:00.000": np.nan, "1900-01-01": np.nan, "1900-01-01T00:00:00": np.nan
            })

    if "Vendedor" in df.columns:
        df = df[df["Vendedor"].astype(str).str.upper() != EXCLUDED_VENDOR].copy()

    if "Estatus" in df.columns:
        df["Estatus"] = df["Estatus"].astype(str).map(canon_estatus)

    df["Centro Original"] = pd.Series(pd.NA, index=df.index, dtype="object")
    mask_cc2 = df["Centro"].astype(str).str.contains("EXP ATT C CENTER 2", na=False)
    mask_jv = df["Centro"].astype(str).str.contains("EXP ATT C CENTER JUAREZ", na=False)
    df.loc[mask_cc2, "Centro Original"] = "CC2"
    df.loc[mask_jv, "Centro Original"] = "CC JV"

    empleados_join = hoja[hoja["Puesto"].isin(["ASESOR TELEFONICO 7500", "EJECUTIVO TELEFONICO 6500 AM"])].copy()
    empleados_join = empleados_join[empleados_join["JefeDirecto"] != "ENCUBADORA"].drop_duplicates(subset=["NombreCompleto"])
    empleados_join = empleados_join[empleados_join["NombreCompleto"].str.upper() != EXCLUDED_VENDOR]

    hoja_join = empleados_join.rename(columns={"NombreCompleto": "Nombre Completo", "JefeDirecto": "Jefe directo"})
    df = df.merge(
        hoja_join[["Nombre Completo", "Jefe directo", "Coordinador"]],
        how="left",
        left_on="Vendedor",
        right_on="Nombre Completo",
    )
    df.drop(columns=["Nombre Completo"], inplace=True, errors="ignore")
    df["Jefe directo"] = df["Jefe directo"].fillna("").astype(str).str.strip().replace("", "ENCUBADORA")

    # (se mantiene) parse previo de Fecha creacion, build_view decide DF/MF final
    if "Fecha creacion" in df.columns:
        df["Fecha creacion"] = pd.to_datetime(df["Fecha creacion"], errors="coerce", dayfirst=True)
        df["Fecha creacion"] = sanitize_dates(df["Fecha creacion"])

    if "Back Office" in df.columns:
        s = df["Back Office"].astype(str).str.strip()
        s = s.replace({"nan": "", "None": "", "NaT": "", "<NA>": "", "null": ""})
        s = s.where(s != "", np.nan)
        if s.notna().any():
            pat = r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?)|(\d{4}[/-]\d{1,2}[/-]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?)"
            ext = s.astype(str).str.extract(pat)
            ext = ext[0].fillna(ext[1])
            s2 = ext.where(ext.notna(), s)
        else:
            s2 = s

        df["BO_DT_DF"] = pd.to_datetime(s2, errors="coerce", dayfirst=True)
        df["BO_DT_DF"] = sanitize_dates(df["BO_DT_DF"])

        df["BO_DT_MF"] = pd.to_datetime(s2, errors="coerce", dayfirst=False)
        df["BO_DT_MF"] = sanitize_dates(df["BO_DT_MF"])

    return df

# -------------------------------------------------
# BUILD VIEW
# -------------------------------------------------
def build_view(df_ctx: pd.DataFrame, fecha_ini: date, fecha_fin: date):
    meta = {
        "activation_col": None,
        "has_activation_dt": False,
        "stage_sources": {}
    }
    df = df_ctx.copy()

    # ✅ BO primero (para poder decidir CREATED_DT contra BO)
    df["BO_DT"] = choose_backoffice_dt(df, window_start=fecha_ini, window_end=fecha_fin) if "Back Office" in df.columns else pd.NaT

    # ✅ CREATED_DT robusto (DF/MF) usando BO como referencia
    if "Fecha creacion" in df.columns:
        c_df, c_mf = parse_dt_both(df["Fecha creacion"])
        df["CREATED_DT"] = choose_dt_created_rowwise(c_df, c_mf, bo=df["BO_DT"], window_start=fecha_ini, window_end=fecha_fin)
    else:
        df["CREATED_DT"] = pd.NaT

    meta["stage_sources"]["Nuevo"] = "Fecha creacion" if "Fecha creacion" in df.columns else None
    meta["stage_sources"]["Back Office"] = "Back Office" if "Back Office" in df.columns else None

    df["STG_Nuevo_DT"] = df["CREATED_DT"]
    df["STG_BackOffice_DT"] = df["BO_DT"]

    for stage, outcol in [
        ("Solicitado", "STG_Solicitado_DT"),
        ("En preparacion", "STG_EnPreparacion_DT"),
        ("En entrega", "STG_EnEntrega_DT"),
        ("Reprogramado", "STG_Reprogramado_DT"),
        ("Entregado", "STG_Entregado_DT"),
    ]:
        dt_stage, src = pick_stage_dt_from_columns(df, stage, created=df["CREATED_DT"], bo=df["BO_DT"])
        df[outcol] = sanitize_dates(dt_stage)
        meta["stage_sources"][stage] = src

    act_dt, act_col = pick_activation_dt(df)
    df["ACT_DT"] = sanitize_dates(act_dt)

    # ✅ Guardrail mínimo: una activación no puede estar en el futuro
    _now = pd.Timestamp(datetime.now()) + pd.Timedelta(days=1)
    df.loc[df["ACT_DT"] > _now, "ACT_DT"] = pd.NaT

    meta["activation_col"] = act_col
    meta["has_activation_dt"] = bool(df["ACT_DT"].notna().any())

    if "Venta" in df.columns:
        venta = df["Venta"]
        venta_ok = ~(venta.isna() | venta.astype(str).str.strip().eq(""))
    else:
        venta_ok = pd.Series(False, index=df.index)

    df["IS_ACTIVADA_COMPLETA"] = venta_ok | df["ACT_DT"].notna()
    df["ENTREGADO_SIN_VENTA"] = (df["Estatus"].astype(str).eq("Entregado")) & (~venta_ok)

    df["TD_Creacion_a_BO"] = df["BO_DT"] - df["CREATED_DT"]
    df["TD_BO_a_Act"] = df["ACT_DT"] - df["BO_DT"]
    df["TD_Creacion_a_Act"] = df["ACT_DT"] - df["CREATED_DT"]

    def safe_td(a, b):
        td = a - b
        td = td.where(td >= pd.Timedelta(0))
        return td

    df["TD_BO_a_Solicitado"] = safe_td(df["STG_Solicitado_DT"], df["STG_BackOffice_DT"])
    df["TD_Solicitado_a_Preparacion"] = safe_td(df["STG_EnPreparacion_DT"], df["STG_Solicitado_DT"])
    df["TD_Preparacion_a_EnEntrega"] = safe_td(df["STG_EnEntrega_DT"], df["STG_EnPreparacion_DT"])
    df["TD_EnEntrega_a_Entregado"] = safe_td(df["STG_Entregado_DT"], df["STG_EnEntrega_DT"])

    ref_dt = pd.Timestamp(_reference_end_dt(fecha_fin))
    df["TD_Age_Desde_Creacion"] = ref_dt - df["CREATED_DT"]
    df["TD_Age_Desde_BO"] = ref_dt - df["BO_DT"]

    for c in ["TD_Creacion_a_BO","TD_BO_a_Act","TD_Creacion_a_Act","TD_Age_Desde_Creacion","TD_Age_Desde_BO"]:
        df.loc[df[c] < pd.Timedelta(0), c] = pd.NaT

    df["H_Creacion_a_BO"] = df["TD_Creacion_a_BO"].apply(td_to_hours)
    df["H_BO_a_Act"] = df["TD_BO_a_Act"].apply(td_to_hours)
    df["H_Creacion_a_Act"] = df["TD_Creacion_a_Act"].apply(td_to_hours)
    df["H_Age_Desde_Creacion"] = df["TD_Age_Desde_Creacion"].apply(td_to_hours)
    df["H_Age_Desde_BO"] = df["TD_Age_Desde_BO"].apply(td_to_hours)

    df["CREATED_DATE"] = df["CREATED_DT"].dt.date
    df["CREATED_HOUR"] = df["CREATED_DT"].dt.hour
    df["CREATED_DOW"] = df["CREATED_DT"].dt.day_name()

    return df, meta

# -------------------------------------------------
# CHARTS
# -------------------------------------------------
def make_funnel(counts: dict) -> go.Figure:
    stages = ["Nuevo", "Back Office", "Solicitado", "En preparacion", "En entrega", "Reprogramado", "Entregado"]
    values = [int(counts.get(s, 0)) for s in stages]
    fig = go.Figure(go.Funnel(y=stages, x=values))
    fig.update_layout(title="Funnel del flujo operativo (conteos por etapa)", margin=dict(l=40, r=20, t=60, b=20))
    return fig

def make_flow_bar(counts: dict) -> go.Figure:
    stages = ["Nuevo", "Back Office", "Solicitado", "En preparacion", "En entrega", "Reprogramado", "Entregado"]
    values = [int(counts.get(s, 0)) for s in stages]
    fig = px.bar(pd.DataFrame({"Etapa": stages, "Total": values}), x="Etapa", y="Total", title="Conteos por etapa")
    fig.update_xaxes(type="category")
    fig.update_layout(margin=dict(l=40, r=20, t=60, b=20))
    return fig

def make_backlog_over_time(view: pd.DataFrame) -> go.Figure | None:
    if view.empty or view["CREATED_DT"].isna().all():
        return None
    dfp = view.copy()
    dfp["Fecha"] = dfp["CREATED_DT"].dt.date
    grp = dfp.groupby(["Fecha", "Estatus"]).size().reset_index(name="Total")
    grp = grp[grp["Estatus"].isin(FLOW_STAGES_NO_TOTAL)].copy()
    return px.area(grp, x="Fecha", y="Total", color="Estatus", title="Backlog por etapa a través del tiempo (creación)")

def make_bottleneck_chart(view: pd.DataFrame) -> go.Figure | None:
    if view.empty:
        return None
    open_mask = ~view["Estatus"].astype(str).eq("Entregado")
    dfp = view.loc[open_mask].copy()
    if dfp.empty:
        return None

    dfp["Aging_h"] = np.where(
        dfp["Estatus"].astype(str).eq("Back Office"),
        dfp["H_Age_Desde_BO"],
        dfp["H_Age_Desde_Creacion"],
    )

    grp = dfp.groupby("Estatus", as_index=False)["Aging_h"].median()
    grp = grp[grp["Estatus"].isin(FLOW_STAGES_NO_TOTAL)].copy()
    grp["Estatus"] = pd.Categorical(grp["Estatus"], categories=FLOW_STAGES_NO_TOTAL, ordered=True)
    grp = grp.sort_values("Estatus")
    return px.bar(grp, x="Aging_h", y="Estatus", orientation="h")

def make_trends(view: pd.DataFrame, meta: dict) -> go.Figure | None:
    if view.empty or view["CREATED_DT"].isna().all():
        return None
    created = (view.assign(Fecha=view["CREATED_DT"].dt.date).groupby("Fecha", as_index=False).size().rename(columns={"size": "Creadas"}))
    if meta.get("has_activation_dt", False) and view["ACT_DT"].notna().any():
        activated = (view[view["ACT_DT"].notna()].assign(Fecha=view["ACT_DT"].dt.date).groupby("Fecha", as_index=False).size().rename(columns={"size": "Activadas"}))
    else:
        activated = pd.DataFrame({"Fecha": [], "Activadas": []})
    trend = created.merge(activated, on="Fecha", how="left")
    trend["Activadas"] = trend["Activadas"].fillna(0).astype(int)
    trend_long = trend.melt(id_vars="Fecha", var_name="Tipo", value_name="Total")

    fig = px.line(trend_long, x="Fecha", y="Total", color="Tipo", markers=True, title="Tendencia: Creadas vs Activadas (por día)")
    fig.update_layout(margin=dict(l=40, r=20, t=60, b=20))
    return fig

def make_sla_gauge(sla_rate: float | None, sla_h: int) -> go.Figure:
    val = 0 if sla_rate is None or np.isnan(sla_rate) else float(sla_rate) * 100.0
    fig = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=val,
            number={"suffix": "%"},
            title={"text": f"Cumplimiento SLA (BO → Activación ≤ {sla_h}h)"},
            gauge={
                "axis": {"range": [0, 100]},
                "steps": [{"range": [0, 60]}, {"range": [60, 80]}, {"range": [80, 100]}],
                "threshold": {"line": {"width": 4}, "thickness": 0.75, "value": 80},
            },
        )
    )
    fig.update_layout(margin=dict(l=20, r=20, t=70, b=20), height=320)
    return fig

def make_sla_trend(view: pd.DataFrame, sla_h: int, meta: dict) -> go.Figure | None:
    if not meta.get("has_activation_dt", False):
        return None
    if view.empty or view["ACT_DT"].isna().all() or view["H_BO_a_Act"].dropna().empty:
        return None
    dfp = view[view["H_BO_a_Act"].notna()].copy()
    dfp["Fecha"] = dfp["ACT_DT"].dt.date
    grp = dfp.groupby("Fecha").agg(
        Total=("H_BO_a_Act", "size"),
        Dentro=("H_BO_a_Act", lambda s: int((s <= float(sla_h)).sum())),
    ).reset_index()
    grp["Cumplimiento_%"] = np.where(grp["Total"] > 0, grp["Dentro"] / grp["Total"] * 100.0, np.nan)
    fig = px.line(grp, x="Fecha", y="Cumplimiento_%", markers=True, title="Cumplimiento SLA por día (BO → Activación)")
    fig.update_yaxes(range=[0, 100])
    return fig

def make_heatmap_created(view: pd.DataFrame) -> go.Figure | None:
    if view.empty or view["CREATED_DT"].isna().all():
        return None
    tmp = view.copy()
    tmp["DOW"] = tmp["CREATED_DT"].dt.day_name()
    tmp["HOUR"] = tmp["CREATED_DT"].dt.hour
    piv = tmp.pivot_table(index="DOW", columns="HOUR", values="Estatus", aggfunc="count", fill_value=0)
    order = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]
    piv = piv.reindex([d for d in order if d in piv.index])
    fig = px.imshow(piv, title="Mapa de calor: órdenes creadas (día vs hora)", aspect="auto")
    fig.update_layout(margin=dict(l=40, r=20, t=60, b=20))
    return fig

def make_stage_waterfall(view: pd.DataFrame) -> go.Figure | None:
    cols = [
        ("Nuevo → Back Office", "TD_Creacion_a_BO"),
        ("BO → Solicitado", "TD_BO_a_Solicitado"),
        ("Solicitado → Preparación", "TD_Solicitado_a_Preparacion"),
        ("Preparación → En entrega", "TD_Preparacion_a_EnEntrega"),
        ("En entrega → Entregado", "TD_EnEntrega_a_Entregado"),
    ]

    labels, meds = [], []
    for lab, c in cols:
        labels.append(lab)
        if c in view.columns and view[c].notna().any():
            meds.append(view[c].dropna().median())
        else:
            meds.append(pd.NaT)

    if all(pd.isna(x) for x in meds):
        return None

    hours = [td_to_hours(x) for x in meds]
    hours_disp = [0 if (h is None or (isinstance(h, float) and np.isnan(h))) else float(h) for h in hours]

    fig = go.Figure(go.Waterfall(
        name="Mediana",
        orientation="v",
        x=labels,
        y=hours_disp,
        measure=["relative"] * len(labels),
        text=[fmt_timedelta(m) for m in meds],
        textposition="outside",
    ))
    fig.update_layout(title="Waterfall: tiempo mediano por etapa (si existen fechas por etapa)", showlegend=False)
    fig.update_yaxes(title="Horas (mediana)")
    return fig

# -------------------------------------------------
# MAIN
# -------------------------------------------------
def main():
    st.title("⏱️ Órdenes — Flujo y Tiempo de Activación")

    st.sidebar.header("Panel de control")

    if st.sidebar.button("🔄 Actualizar"):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.session_state["last_refresh"] = datetime.now()
        st.rerun()

    preset = st.sidebar.radio("Periodo", ["Hoy", "Últimos 7 días", "Últimos 30 días", "Personalizado"], index=2)

    today = date.today()
    if preset == "Hoy":
        fecha_ini = today
        fecha_fin = today
    elif preset == "Últimos 7 días":
        fecha_ini = today - timedelta(days=6)
        fecha_fin = today
    elif preset == "Últimos 30 días":
        fecha_ini = today - timedelta(days=29)
        fecha_fin = today
    else:
        fecha_ini = st.sidebar.date_input("Fecha inicio", today - timedelta(days=DEFAULT_LOOKBACK_DAYS))
        fecha_fin = st.sidebar.date_input("Fecha fin", today)

    if fecha_ini > fecha_fin:
        st.sidebar.error("La fecha inicio no puede ser mayor que la fecha fin.")
        return

    st.sidebar.markdown("---")
    sla_h = st.sidebar.number_input("SLA objetivo BO → Activación (horas)", 1, 240, 24, 1)

    st.sidebar.subheader("Alertas (horas) — Pendientes críticos")
    alert_map = {
        "Nuevo": st.sidebar.number_input("Nuevo >", 1, 720, 24, 1),
        "Back Office": st.sidebar.number_input("Back Office >", 1, 720, 24, 1),
        "Solicitado": st.sidebar.number_input("Solicitado >", 1, 720, 24, 1),
        "En preparacion": st.sidebar.number_input("En preparacion >", 1, 720, 24, 1),
        "En entrega": st.sidebar.number_input("En entrega >", 1, 720, 48, 1),
        "Reprogramado": st.sidebar.number_input("Reprogramado >", 1, 720, 48, 1),
    }

    with st.spinner("Cargando datos..."):
        hoja = load_hoja1()
        raw = load_consulta1(fecha_ini, fecha_fin)
        rastreo_extra = load_rastreo_extra(fecha_ini, fecha_fin)
        consulta = transform_consulta1(raw, hoja, rastreo_extra)

    # Optional filters
    with st.sidebar.expander("Filtros (opcional)"):
        centros = ["All"] + sorted([c for c in consulta.get("Centro Original", pd.Series(dtype="object")).dropna().unique().tolist()])
        supervisores = ["All"] + sorted([s for s in consulta.get("Jefe directo", pd.Series(dtype="object")).dropna().unique().tolist()])
        centro_sel = st.selectbox("Centro", centros, 0)
        supervisor_sel = st.selectbox("Supervisor", supervisores, 0)

        df_for_exec = consulta.copy()
        if centro_sel != "All" and "Centro Original" in df_for_exec.columns:
            df_for_exec = df_for_exec[df_for_exec["Centro Original"] == centro_sel]
        if supervisor_sel != "All" and "Jefe directo" in df_for_exec.columns:
            df_for_exec = df_for_exec[df_for_exec["Jefe directo"] == supervisor_sel]

        if "Vendedor" in df_for_exec.columns:
            df_for_exec = df_for_exec[df_for_exec["Vendedor"].astype(str).str.upper() != EXCLUDED_VENDOR]
            ejecutivos = ["All"] + sorted(df_for_exec["Vendedor"].dropna().unique().tolist())
        else:
            ejecutivos = ["All"]
        ejecutivo_sel = st.selectbox("Ejecutivo", ejecutivos, 0)

    # Apply filters
    df = consulta.copy()
    if "Centro Original" in df.columns and centro_sel != "All":
        df = df[df["Centro Original"] == centro_sel]
    if "Jefe directo" in df.columns and supervisor_sel != "All":
        df = df[df["Jefe directo"] == supervisor_sel]
    if "Vendedor" in df.columns and ejecutivo_sel != "All":
        df = df[df["Vendedor"] == ejecutivo_sel]

    tabs = st.tabs(["Resumen Ejecutivo", "Gráficas", "Pendientes a Recuperar", "Detalle / Export"])

    # ============================
    # TAB 0: RESUMEN
    # ============================
    with tabs[0]:
        if df.empty:
            st.info("No hay datos para los filtros actuales.")
        else:
            view, meta = build_view(df, fecha_ini, fecha_fin)
            counts = compute_flow_counts(view)

            st.subheader("Flujo de Órdenes (operación)")
            render_flow_pills(counts)

            total = int(len(view))
            entregado = int((view["Estatus"].astype(str).eq("Entregado")).sum())
            activadas_completas = int(view["IS_ACTIVADA_COMPLETA"].sum())
            entregado_sin_venta = int(view["ENTREGADO_SIN_VENTA"].sum())

            if meta["has_activation_dt"] and view["H_BO_a_Act"].notna().any():
                ba = view[view["H_BO_a_Act"].notna()].copy()
                sla_ok = int((ba["H_BO_a_Act"] <= float(sla_h)).sum())
                sla_total = int(ba.shape[0])
                sla_rate = (sla_ok / sla_total) if sla_total else np.nan
            else:
                sla_rate = np.nan

            med_td_cb = view["TD_Creacion_a_BO"].dropna()
            med_nuevo_bo = med_td_cb.median() if not med_td_cb.empty else pd.NaT

            med_bo_act = view.loc[view["TD_BO_a_Act"].notna(), "TD_BO_a_Act"].median() if view["TD_BO_a_Act"].notna().any() else pd.NaT
            med_age_bo = view.loc[view["TD_Age_Desde_BO"].notna(), "TD_Age_Desde_BO"].median() if view["TD_Age_Desde_BO"].notna().any() else pd.NaT

            st.markdown('<div class="kpi-row">', unsafe_allow_html=True)
            kpi_card("Órdenes en el periodo", fmt_int(total), sub=f"Del {fecha_ini} al {fecha_fin}")
            kpi_card("Entregado (estatus)", fmt_int(entregado))
            kpi_card("Activadas completas", fmt_int(activadas_completas), sub="Con Venta o con fecha de activación/venta")
            kpi_card("Mediana Nuevo → Back Office", fmt_timedelta(med_nuevo_bo), sub="Back Office = Rastreo (col Back Office)")

            if meta["has_activation_dt"]:
                kpi_card("Mediana BO → Activación", fmt_timedelta(med_bo_act) if pd.notna(med_bo_act) else "—", sub=f"SLA ≤ {sla_h}h")
                kpi_card("Cumplimiento SLA", fmt_pct(sla_rate) if not np.isnan(sla_rate) else "—", sub=f"BO → Activación (≤ {sla_h}h)")
            else:
                kpi_card("BO → Activación", "En proceso", sub=f"Mediana antigüedad desde BO: {fmt_timedelta(med_age_bo)}")
                kpi_card("Cumplimiento SLA", "—", sub="Se habilita al tener Fecha activación/venta")
            st.markdown("</div>", unsafe_allow_html=True)

            if entregado_sin_venta > 0:
                st.warning(f"⚠️ Hay **{entregado_sin_venta}** órdenes en **Entregado** pero **sin Venta** (revisar / recuperar).")
                with st.expander("Ver cuáles son Entregado sin Venta", expanded=True):
                    df_esv = view[view["ENTREGADO_SIN_VENTA"]].copy()
                    df_esv["Antigüedad"] = df_esv["TD_Age_Desde_Creacion"].apply(fmt_timedelta)

                    cols_esv = [c for c in [
                        "Antigüedad", "Jefe directo", "Vendedor", "Cliente", "Telefono", "Folio",
                        "Centro", "Estatus", "Venta", "Fecha creacion", "Back Office"
                    ] if c in df_esv.columns]

                    show_esv = df_esv[cols_esv].copy().rename(columns={"Vendedor": "Ejecutivo"})
                    show_esv = show_esv.assign(_age=df_esv["TD_Age_Desde_Creacion"]).sort_values("_age", ascending=False).drop(columns=["_age"], errors="ignore")

                    st.dataframe(show_esv, use_container_width=True)

                    st.download_button(
                        "Descargar Entregado sin Venta (Excel)",
                        data=dfs_to_excel_bytes({"Entregado_sin_Venta": show_esv}),
                        file_name=f"entregado_sin_venta_{fecha_ini}_{fecha_fin}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

            st.caption(f"Actualizado: {st.session_state['last_refresh'].strftime('%Y-%m-%d %H:%M')}")

    # ============================
    # TAB 1: GRÁFICAS
    # ============================
    with tabs[1]:
        if df.empty:
            st.info("No hay datos para los filtros actuales.")
        else:
            view, meta = build_view(df, fecha_ini, fecha_fin)
            counts = compute_flow_counts(view)

            st.subheader("📊 Gráficas principales")

            c1, c2 = st.columns(2)
            with c1:
                st.plotly_chart(make_funnel(counts), use_container_width=True)
            with c2:
                st.plotly_chart(make_flow_bar(counts), use_container_width=True)

            fig_backlog = make_backlog_over_time(view)
            if fig_backlog is not None:
                st.plotly_chart(fig_backlog, use_container_width=True)

            fig_bneck = make_bottleneck_chart(view)
            if fig_bneck is not None:
                st.plotly_chart(fig_bneck, use_container_width=True)

            fig_wf = make_stage_waterfall(view)
            if fig_wf is not None:
                st.plotly_chart(fig_wf, use_container_width=True)
            else:
                st.info("Waterfall por etapas requiere que el TVF traiga fechas/hora por etapa (además de Back Office).")

            fig_tr = make_trends(view, meta)
            if fig_tr is not None:
                st.plotly_chart(fig_tr, use_container_width=True)

            st.markdown("---")
            st.subheader("⏱️ SLA y tiempos")

            cA, cB = st.columns(2)
            with cA:
                if meta["has_activation_dt"] and view["H_BO_a_Act"].notna().any():
                    ba = view[view["H_BO_a_Act"].notna()].copy()
                    sla_ok = int((ba["H_BO_a_Act"] <= float(sla_h)).sum())
                    sla_total = int(ba.shape[0])
                    sla_rate = (sla_ok / sla_total) if sla_total else np.nan
                else:
                    sla_rate = np.nan
                st.plotly_chart(make_sla_gauge(sla_rate, int(sla_h)), use_container_width=True)

                fig_sla_tr = make_sla_trend(view, int(sla_h), meta)
                if fig_sla_tr is not None:
                    st.plotly_chart(fig_sla_tr, use_container_width=True)

            with cB:
                if meta["has_activation_dt"] and view["H_BO_a_Act"].notna().any():
                    fig = px.histogram(view[view["H_BO_a_Act"].notna()], x="H_BO_a_Act", nbins=40,
                                       title="Distribución: BO → Activación (horas exactas)")
                else:
                    fig = px.histogram(view[view["H_Age_Desde_BO"].notna()], x="H_Age_Desde_BO", nbins=40,
                                       title="Distribución: Antigüedad desde Back Office (horas exactas)")
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")
            st.subheader("🧭 Mapa de calor (cuándo se crean más órdenes)")
            fig_hm = make_heatmap_created(view)
            if fig_hm is not None:
                st.plotly_chart(fig_hm, use_container_width=True)

            st.markdown("---")
            st.subheader("👥 Pendientes por Supervisor / Ejecutivo")

            open_mask = ~view["Estatus"].astype(str).eq("Entregado")

            if "Jefe directo" in view.columns:
                by_sup = view.loc[open_mask].groupby("Jefe directo", as_index=False).size().rename(columns={"size": "Pendientes"})
                by_sup = by_sup.sort_values("Pendientes", ascending=False).head(15)
                if not by_sup.empty:
                    fig_sup = px.bar(by_sup, x="Pendientes", y="Jefe directo", orientation="h", title="Pendientes por Supervisor (Top 15)")
                    st.plotly_chart(fig_sup, use_container_width=True)

            if "Vendedor" in view.columns:
                by_exec = view.loc[open_mask].groupby("Vendedor", as_index=False).size().rename(columns={"size": "Pendientes"})
                by_exec = by_exec.sort_values("Pendientes", ascending=False).head(15)
                if not by_exec.empty:
                    fig_exec = px.bar(by_exec, x="Pendientes", y="Vendedor", orientation="h", title="Pendientes por Ejecutivo (Top 15)")
                    st.plotly_chart(fig_exec, use_container_width=True)

    # ============================
    # TAB 2: PENDIENTES A RECUPERAR
    # ============================
    with tabs[2]:
        if df.empty:
            st.info("No hay datos para los filtros actuales.")
        else:
            view, _ = build_view(df, fecha_ini, fecha_fin)
            work = view.copy()
            E = work["Estatus"].astype(str)

            work["Aging_h"] = np.where(
                E.eq("Back Office"),
                work["H_Age_Desde_BO"],
                work["H_Age_Desde_Creacion"],
            )

            work["CRITICO"] = False
            for stg, th in alert_map.items():
                work.loc[E.eq(stg) & (work["Aging_h"] > float(th)), "CRITICO"] = True

            crit = work[work["CRITICO"]].copy().sort_values("Aging_h", ascending=False)

            st.subheader("📌 Pendientes críticos (para recuperar hoy)")
            st.caption("Ordenados por antigüedad (más viejos arriba).")
            st.write(f"Total críticos: **{len(crit)}**")

            cols = [c for c in [
                "Estatus",
                "Jefe directo", "Vendedor", "Cliente", "Telefono", "Folio", "Centro", "Venta",
                "TD_Age_Desde_Creacion", "TD_Age_Desde_BO"
            ] if c in crit.columns]

            show = crit[cols].copy().rename(columns={"Vendedor": "Ejecutivo"})
            show["Antigüedad"] = np.where(
                show["Estatus"].astype(str).eq("Back Office"),
                crit["TD_Age_Desde_BO"].apply(fmt_timedelta),
                crit["TD_Age_Desde_Creacion"].apply(fmt_timedelta),
            )
            show.drop(columns=["TD_Age_Desde_Creacion", "TD_Age_Desde_BO"], inplace=True, errors="ignore")

            st.dataframe(show, use_container_width=True)

            if not crit.empty:
                by_stage = crit.groupby("Estatus", as_index=False).size().rename(columns={"size": "Críticos"})
                by_stage["Estatus"] = pd.Categorical(by_stage["Estatus"], categories=FLOW_STAGES_NO_TOTAL, ordered=True)
                by_stage = by_stage.sort_values("Estatus")
                fig = px.bar(by_stage, x="Estatus", y="Críticos", title="Críticos por etapa")
                fig.update_xaxes(type="category")
                st.plotly_chart(fig, use_container_width=True)

            st.download_button(
                "Descargar críticos (Excel)",
                data=dfs_to_excel_bytes({"Criticos": show}),
                file_name=f"pendientes_criticos_{fecha_ini}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ============================
    # TAB 3: DETALLE / EXPORT
    # ============================
    with tabs[3]:
        if df.empty:
            st.info("No hay datos para los filtros actuales.")
        else:
            view, meta = build_view(df, fecha_ini, fecha_fin)

            detail = view.copy()

            # ✅ Tiempo Nuevo→BO: si ya tiene BO => TD_Creacion_a_BO
            # ✅ si sigue en Nuevo y no tiene BO => En proceso desde Fecha Creación
            detail["Tiempo Nuevo→BO (HH:MM)"] = [
                (fmt_timedelta(done) if (done is not None and pd.notna(done))
                 else (f"En proceso · {fmt_timedelta(age)}" if (str(stg) == "Nuevo" and age is not None and pd.notna(age)) else "—"))
                for done, age, stg in zip(detail["TD_Creacion_a_BO"], detail["TD_Age_Desde_Creacion"], detail["Estatus"])
            ]

            detail["Tiempo BO→Act (HH:MM)"] = [
                fmt_done_or_in_process(done, age)
                for done, age in zip(detail["TD_BO_a_Act"], detail["TD_Age_Desde_BO"])
            ]
            detail["Tiempo Total (HH:MM)"] = [
                fmt_done_or_in_process(done, age)
                for done, age in zip(detail["TD_Creacion_a_Act"], detail["TD_Age_Desde_Creacion"])
            ]

            detail["Antigüedad desde Creación (HH:MM)"] = detail["TD_Age_Desde_Creacion"].apply(fmt_timedelta)
            detail["Antigüedad desde BO (HH:MM)"] = detail["TD_Age_Desde_BO"].apply(fmt_timedelta)

            # Volvemos a sanitizar justo antes de mostrar (doble seguridad visual)
            for c in ["STG_Solicitado_DT","STG_EnPreparacion_DT","STG_EnEntrega_DT","STG_Reprogramado_DT","STG_Entregado_DT"]:
                if c in detail.columns:
                    detail[c] = pd.to_datetime(detail[c], errors="coerce")
                    detail[c] = sanitize_dates(detail[c])

            keep = [c for c in [
                "Estatus", "Jefe directo", "Vendedor", "Cliente", "Telefono", "Folio", "Centro",
                "Venta", "CREATED_DT", "BO_DT", "ACT_DT",
                "STG_Solicitado_DT","STG_EnPreparacion_DT","STG_EnEntrega_DT","STG_Reprogramado_DT","STG_Entregado_DT",
                "Tiempo Nuevo→BO (HH:MM)", "Tiempo BO→Act (HH:MM)", "Tiempo Total (HH:MM)",
                "Antigüedad desde Creación (HH:MM)", "Antigüedad desde BO (HH:MM)",
                "ENTREGADO_SIN_VENTA"
            ] if c in detail.columns]

            rename_map = {
                "Vendedor": "Ejecutivo",
                "STG_Solicitado_DT": "Fecha Solicitado",
                "STG_EnPreparacion_DT": "Fecha En preparacion",
                "STG_EnEntrega_DT": "Fecha En entrega",
                "STG_Reprogramado_DT": "Fecha Reprogramado",
                "STG_Entregado_DT": "Fecha Entregado",
                "CREATED_DT": "Fecha Creacion",
                "BO_DT": "Fecha Back Office",
                "ACT_DT": "Fecha Activacion",
            }

            show = detail[keep].copy().rename(columns=rename_map)
            if "Fecha Creacion" in show.columns:
                show = show.sort_values("Fecha Creacion", ascending=False)

            st.subheader("📄 Detalle completo")
            st.dataframe(show, use_container_width=True)

            summary = pd.DataFrame([{
                "Periodo": f"{fecha_ini} a {fecha_fin}",
                "Órdenes total": int(len(view)),
                "Fuente Back Office": "Back Office (Rastreo)",
                "Fuente fecha activación": meta["activation_col"] if meta["has_activation_dt"] else "No disponible",
                "Actualizado": st.session_state["last_refresh"].strftime("%Y-%m-%d %H:%M"),
            }])

            st.download_button(
                "Descargar reporte (Excel)",
                data=dfs_to_excel_bytes({"Resumen": summary, "Detalle": show}),
                file_name=f"reporte_flujo_tiempos_{fecha_ini}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

if __name__ == "__main__":
    main()
