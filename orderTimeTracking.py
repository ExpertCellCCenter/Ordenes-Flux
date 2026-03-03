# app.py  ✅ BOSS-READY + MÁS GRÁFICO + TIEMPOS EXACTOS (HH:MM) — Flujo Operativo + Lead Time
# ✅ Flujo EXACTO (pills):
#    Total | Nuevo | Back Office | Solicitado | En preparacion | En entrega | Reprogramado | Entregado
# ✅ Activación/Entrega (para tiempos) se toma SIEMPRE de:
#    1) "Fecha activación"  2) "Fecha venta"  3) (fallback) "Venta" si parece fecha-hora
# ✅ Tiempos EXACTOS: se calculan con timedeltas (incluye minutos) desde las fechas de Global (SQL).
# ✅ Incluye visualización + descarga de "Entregado sin Venta".
# ✅ Gráficas interactivas (Plotly): Funnel, backlog over time, tendencias, SLA gauge, SLA trend, heatmap, bottleneck chart, top equipos.
# ✅ (Boxplot ELIMINADO)

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
DEFAULT_LOOKBACK_DAYS = 30  # boss-friendly default

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

FLOW_STAGES_NO_TOTAL = [
    "Nuevo", "Back Office", "Solicitado", "En preparacion", "En entrega", "Reprogramado", "Entregado"
]

# -------------------------------------------------
# STREAMLIT CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="Órdenes — Flujo & Tiempos (Nuevo → Activada)",
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
        return f"{float(x)*100:.1f}%"
    except Exception:
        return "0.0%"

def fmt_timedelta(td) -> str:
    """Exact display: HH:MM (or Xd HH:MM) from timedeltas."""
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

    if days > 0:
        return f"{days}d {hh:02d}:{mm:02d}"
    return f"{hh:02d}:{mm:02d}"

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
# NORMALIZATION (Estatus)
# -------------------------------------------------
def _norm_col(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf-8", "ignore")
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
    if sn in ["en preparacion", "en preparación"]:
        return "En preparacion"
    if sn == "en entrega":
        return "En entrega"
    if sn in ["reprogramado", "reprogramada"]:
        return "Reprogramado"
    if sn in ["entregado", "entregada"]:
        return "Entregado"
    return s

# -------------------------------------------------
# DATETIME PARSING (STRICTER, GLOBAL-LIKE)
# -------------------------------------------------
def _extract_datetime_text(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace({"nan": "", "none": "", "nat": ""})
    s = s.where(s != "", np.nan)

    if s.notna().any():
        pat = r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?)|(\d{4}[/-]\d{1,2}[/-]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?)"
        ext = s.astype(str).str.extract(pat)
        ext = ext[0].fillna(ext[1])
        s2 = ext.where(ext.notna(), s)
        return s2
    return s

def parse_dt_both(series: pd.Series) -> tuple[pd.Series, pd.Series]:
    s2 = _extract_datetime_text(series)
    dt_df = pd.to_datetime(s2, errors="coerce", dayfirst=True)
    dt_mf = pd.to_datetime(s2, errors="coerce", dayfirst=False)
    return dt_df, dt_mf

def choose_dt_rowwise(dt_df: pd.Series, dt_mf: pd.Series, created: pd.Series | None, bo: pd.Series | None) -> pd.Series:
    out = dt_df.copy()
    out = out.where(~(dt_df.isna() & dt_mf.notna()), dt_mf)

    if created is None and bo is None:
        return out

    if created is not None:
        c = created
        valid_df_c = dt_df.notna() & c.notna() & (dt_df >= c) & (dt_df <= c + pd.Timedelta(days=730))
        valid_mf_c = dt_mf.notna() & c.notna() & (dt_mf >= c) & (dt_mf <= c + pd.Timedelta(days=730))
    else:
        valid_df_c = pd.Series(False, index=dt_df.index)
        valid_mf_c = pd.Series(False, index=dt_df.index)

    if bo is not None:
        b = bo
        valid_df_b = dt_df.notna() & b.notna() & (dt_df >= b) & (dt_df <= b + pd.Timedelta(days=730))
        valid_mf_b = dt_mf.notna() & b.notna() & (dt_mf >= b) & (dt_mf <= b + pd.Timedelta(days=730))
    else:
        valid_df_b = pd.Series(False, index=dt_df.index)
        valid_mf_b = pd.Series(False, index=dt_df.index)

    valid_df = valid_df_c | valid_df_b
    valid_mf = valid_mf_c | valid_mf_b

    out = out.where(~(valid_mf & ~valid_df), dt_mf)
    return out

def parse_backoffice_datetime(series: pd.Series, window_start: date | None = None, window_end: date | None = None) -> pd.Series:
    s = series.astype(str).str.strip()
    s = s.replace({"nan": "", "None": "", "NaT": ""})
    s = s.where(s != "", np.nan)

    if s.notna().any():
        pat = r"(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\s+\d{1,2}:\d{2}(?::\d{2})?)|(\d{4}[/-]\d{1,2}[/-]\d{1,2}\s+\d{1,2}:\d{2}(?::\d{2})?)"
        ext = s.astype(str).str.extract(pat)
        ext = ext[0].fillna(ext[1])
        s2 = ext.where(ext.notna(), s)
    else:
        s2 = s

    dt_dayfirst = pd.to_datetime(s2, errors="coerce", dayfirst=True)
    dt_monthfirst = pd.to_datetime(s2, errors="coerce", dayfirst=False)

    if window_start is None or window_end is None:
        return dt_dayfirst

    w0 = pd.Timestamp(window_start)
    w1 = pd.Timestamp(window_end) + pd.Timedelta(days=1) - pd.Timedelta(microseconds=1)

    in1 = dt_dayfirst.between(w0, w1)
    in2 = dt_monthfirst.between(w0, w1)

    out = dt_dayfirst.copy()
    out = out.where(~(in2 & ~in1), dt_monthfirst)
    out = out.where(~(dt_dayfirst.isna() & dt_monthfirst.notna()), dt_monthfirst)
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
        return out

    return parse_backoffice_datetime(df["Back Office"], window_start=window_start, window_end=window_end)

def _reference_end_dt(fecha_fin: date) -> datetime:
    return datetime.now() if fecha_fin == date.today() else datetime.combine(fecha_fin, time(23, 59, 59))

def pick_activation_dt(df: pd.DataFrame) -> tuple[pd.Series, str | None]:
    if df is None or df.empty:
        return pd.Series(pd.NaT, index=df.index), None

    cols_norm = {_norm_col(c): c for c in df.columns}

    created = df["CREATED_DT"] if "CREATED_DT" in df.columns else None
    bo = df["BO_DT"] if "BO_DT" in df.columns else None

    def _try_col(colname: str):
        if colname not in df.columns:
            return None
        dt_df, dt_mf = parse_dt_both(df[colname])
        chosen = choose_dt_rowwise(dt_df, dt_mf, created=created, bo=bo)
        return chosen if chosen.notna().any() else None

    for key in ["fecha activacion", "fecha de activacion", "fecha activacion completa", "fecha activada", "fecha activación"]:
        col = cols_norm.get(key)
        if col:
            chosen = _try_col(col)
            if chosen is not None:
                return chosen, col

    for key in ["fecha venta", "fecha de venta", "fecha_venta", "fecha venta "]:
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

# -------------------------------------------------
# TRANSFORM
# -------------------------------------------------
def transform_consulta1(df_raw: pd.DataFrame, hoja: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()

    for col in ["Centro", "Estatus", "Back Office", "Vendedor", "Cliente"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace({"nan": np.nan, "None": np.nan})

    if "Venta" in df.columns:
        df["Venta"] = df["Venta"].replace({"nan": np.nan, "None": np.nan})

    if "Vendedor" in df.columns:
        df = df[df["Vendedor"].astype(str).str.upper() != EXCLUDED_VENDOR].copy()

    if "Estatus" in df.columns:
        df["Estatus"] = df["Estatus"].astype(str).map(canon_estatus)

    # Centro Original
    df["Centro Original"] = pd.Series(pd.NA, index=df.index, dtype="object")
    mask_cc2 = df["Centro"].astype(str).str.contains("EXP ATT C CENTER 2", na=False)
    mask_jv = df["Centro"].astype(str).str.contains("EXP ATT C CENTER JUAREZ", na=False)
    df.loc[mask_cc2, "Centro Original"] = "CC2"
    df.loc[mask_jv, "Centro Original"] = "CC JV"

    # Join supervisor
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

    # Creation datetime (exact)
    if "Fecha creacion" in df.columns:
        df["Fecha creacion"] = pd.to_datetime(df["Fecha creacion"], errors="coerce", dayfirst=True)

    # Pre-parse BO datetimes
    if "Back Office" in df.columns:
        s = df["Back Office"].astype(str).str.strip().replace({"nan": "", "None": "", "NaT": ""})
        s = s.where(s != "", np.nan)
        s2 = _extract_datetime_text(s)
        df["BO_DT_DF"] = pd.to_datetime(s2, errors="coerce", dayfirst=True)
        df["BO_DT_MF"] = pd.to_datetime(s2, errors="coerce", dayfirst=False)

    return df

# -------------------------------------------------
# BUILD VIEW (Exact timedeltas)
# -------------------------------------------------
def build_view(df_ctx: pd.DataFrame, fecha_ini: date, fecha_fin: date):
    meta = {"activation_col": None, "has_activation_dt": False}
    df = df_ctx.copy()

    df["CREATED_DT"] = pd.to_datetime(df["Fecha creacion"], errors="coerce", dayfirst=True) if "Fecha creacion" in df.columns else pd.NaT
    df["BO_DT"] = choose_backoffice_dt(df, window_start=fecha_ini, window_end=fecha_fin) if "Back Office" in df.columns else pd.NaT

    act_dt, act_col = pick_activation_dt(df)
    df["ACT_DT"] = act_dt
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

    ref_dt = pd.Timestamp(_reference_end_dt(fecha_fin))
    df["TD_Age_Desde_Creacion"] = ref_dt - df["CREATED_DT"]
    df["TD_Age_Desde_BO"] = ref_dt - df["BO_DT"]

    for c in ["TD_Creacion_a_BO", "TD_BO_a_Act", "TD_Creacion_a_Act", "TD_Age_Desde_Creacion", "TD_Age_Desde_BO"]:
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

def make_trends(view: pd.DataFrame, meta: dict) -> go.Figure | None:
    if view.empty or view["CREATED_DT"].isna().all():
        return None

    created = (
        view.assign(Fecha=view["CREATED_DT"].dt.date)
        .groupby("Fecha", as_index=False)
        .size()
        .rename(columns={"size": "Creadas"})
    )

    if meta.get("has_activation_dt", False) and view["ACT_DT"].notna().any():
        activated = (
            view[view["ACT_DT"].notna()]
            .assign(Fecha=view["ACT_DT"].dt.date)
            .groupby("Fecha", as_index=False)
            .size()
            .rename(columns={"size": "Activadas"})
        )
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
        Mediana=("H_BO_a_Act", "median"),
    ).reset_index()
    grp["Cumplimiento_%"] = np.where(grp["Total"] > 0, grp["Dentro"] / grp["Total"] * 100.0, np.nan)

    fig = px.line(grp, x="Fecha", y="Cumplimiento_%", markers=True, title="Cumplimiento SLA por día (BO → Activación)")
    fig.update_yaxes(range=[0, 100])
    return fig

def make_backlog_over_time(view: pd.DataFrame) -> go.Figure | None:
    if view.empty or view["CREATED_DT"].isna().all():
        return None
    dfp = view.copy()
    dfp["Fecha"] = dfp["CREATED_DT"].dt.date
    grp = dfp.groupby(["Fecha", "Estatus"]).size().reset_index(name="Total")
    grp = grp[grp["Estatus"].isin(FLOW_STAGES_NO_TOTAL)].copy()
    fig = px.area(grp, x="Fecha", y="Total", color="Estatus", title="Backlog por etapa a través del tiempo (creación)")
    return fig

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

    fig = px.bar(grp, x="Aging_h", y="Estatus", orientation="h")
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

    st.sidebar.subheader("Alertas por etapa (horas)")
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
        consulta = transform_consulta1(raw, hoja)

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

    df = consulta.copy()
    if "Centro Original" in df.columns and "centro_sel" in locals() and centro_sel != "All":
        df = df[df["Centro Original"] == centro_sel]
    if "Jefe directo" in df.columns and "supervisor_sel" in locals() and supervisor_sel != "All":
        df = df[df["Jefe directo"] == supervisor_sel]
    if "Vendedor" in df.columns and "ejecutivo_sel" in locals() and ejecutivo_sel != "All":
        df = df[df["Vendedor"] == ejecutivo_sel]

    tabs = st.tabs(["Resumen Ejecutivo", "Gráficas", "Pendientes a Recuperar", "Detalle / Export"])

    # ----------------------------
    # TAB 1
    # ----------------------------
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

            st.markdown('<div class="kpi-row">', unsafe_allow_html=True)
            kpi_card("Órdenes en el periodo", fmt_int(total), sub=f"Del {fecha_ini} al {fecha_fin}")
            kpi_card("Entregado (estatus)", fmt_int(entregado), sub="Según estatus operativo")
            kpi_card("Activadas completas", fmt_int(activadas_completas), sub="Con Venta o con fecha de activación/venta")
            # medians with exact format
            med_td_cb = view["TD_Creacion_a_BO"].dropna()
            kpi_card("Mediana Creación → Back Office", fmt_timedelta(med_td_cb.median() if not med_td_cb.empty else pd.NaT))
            if meta["has_activation_dt"]:
                med_td_ca = view["TD_Creacion_a_Act"].dropna()
                kpi_card("Mediana Creación → Activación", fmt_timedelta(med_td_ca.median() if not med_td_ca.empty else pd.NaT),
                         sub=f"Fuente: {meta['activation_col']}")
                # SLA
                if view["H_BO_a_Act"].notna().any():
                    ba = view[view["H_BO_a_Act"].notna()].copy()
                    sla_ok = int((ba["H_BO_a_Act"] <= float(sla_h)).sum())
                    sla_total = int(ba.shape[0])
                    sla_rate = (sla_ok / sla_total) if sla_total else np.nan
                else:
                    sla_rate = np.nan
                kpi_card("Cumplimiento SLA", fmt_pct(sla_rate) if not np.isnan(sla_rate) else "—", sub=f"BO → Activación (≤ {sla_h}h)")
            else:
                kpi_card("Mediana Creación → Activación", "—", sub="No se encontró Fecha activación/Fecha venta")
                kpi_card("Cumplimiento SLA", "—", sub="Se habilita al tener Fecha activación o Fecha venta")
            st.markdown("</div>", unsafe_allow_html=True)

            if entregado_sin_venta > 0:
                st.warning(f"⚠️ Hay **{entregado_sin_venta}** órdenes en **Entregado** pero **sin Venta** (revisar / recuperar).")

                with st.expander("Entregado sin Venta", expanded=True):
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

    # ----------------------------
    # TAB 2 (plots) — boxplot removed
    # ----------------------------
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
                    fig = px.histogram(view[view["H_Creacion_a_BO"].notna()], x="H_Creacion_a_BO", nbins=40,
                                       title="Distribución: Creación → Back Office (horas exactas)")
                st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")
            st.subheader("🧭 Mapa de calor (cuándo se crean más órdenes)")
            fig_hm = make_heatmap_created(view)
            if fig_hm is not None:
                st.plotly_chart(fig_hm, use_container_width=True)

            st.markdown("---")
            st.subheader("👥 Pendientes por Supervisor / Ejecutivo (interactivo)")

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

    # ----------------------------
    # TAB 3
    # ----------------------------
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

            st.subheader("📌 Pendientes críticos ")
            st.caption("Ordenados por antigüedad (más viejos arriba).")
            st.write(f"Total críticos: **{len(crit)}**")

            cols = [c for c in [
                "Estatus", "TD_Age_Desde_Creacion", "TD_Age_Desde_BO",
                "Jefe directo", "Vendedor", "Cliente", "Telefono", "Folio", "Centro", "Venta"
            ] if c in crit.columns]

            show = crit[cols].copy().rename(columns={"Vendedor": "Ejecutivo"})
            show["Antigüedad"] = np.where(
                show["Estatus"].astype(str).eq("Back Office"),
                crit["TD_Age_Desde_BO"].apply(fmt_timedelta),
                crit["TD_Age_Desde_Creacion"].apply(fmt_timedelta),
            )
            show.drop(columns=["TD_Age_Desde_Creacion", "TD_Age_Desde_BO"], inplace=True, errors="ignore")

            st.dataframe(show, use_container_width=True)

            st.download_button(
                "Descargar críticos (Excel)",
                data=dfs_to_excel_bytes({"Criticos": show}),
                file_name=f"pendientes_criticos_{fecha_ini}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ----------------------------
    # TAB 4
    # ----------------------------
    with tabs[3]:
        if df.empty:
            st.info("No hay datos para los filtros actuales.")
        else:
            view, meta = build_view(df, fecha_ini, fecha_fin)

            detail = view.copy()
            detail["Tiempo Creación→BO (HH:MM)"] = detail["TD_Creacion_a_BO"].apply(fmt_timedelta)
            detail["Tiempo BO→Activación (HH:MM)"] = detail["TD_BO_a_Act"].apply(fmt_timedelta)
            detail["Tiempo Total (HH:MM)"] = detail["TD_Creacion_a_Act"].apply(fmt_timedelta)
            detail["Antigüedad desde Creación (HH:MM)"] = detail["TD_Age_Desde_Creacion"].apply(fmt_timedelta)
            detail["Antigüedad desde BO (HH:MM)"] = detail["TD_Age_Desde_BO"].apply(fmt_timedelta)

            keep = [c for c in [
                "Estatus", "Jefe directo", "Vendedor", "Cliente", "Telefono", "Folio", "Centro",
                "Venta", "CREATED_DT", "BO_DT", "ACT_DT",
                "Tiempo Creación→BO (HH:MM)", "Tiempo BO→Activación (HH:MM)", "Tiempo Total (HH:MM)",
                "Antigüedad desde Creación (HH:MM)", "Antigüedad desde BO (HH:MM)",
                "ENTREGADO_SIN_VENTA"
            ] if c in detail.columns]

            show = detail[keep].copy().rename(columns={"Vendedor": "Ejecutivo"})
            if "CREATED_DT" in show.columns:
                show = show.sort_values("CREATED_DT", ascending=False)

            st.subheader("📄 Detalle completo (tiempos exactos)")
            st.dataframe(show, use_container_width=True)

            summary = pd.DataFrame([{
                "Periodo": f"{fecha_ini} a {fecha_fin}",
                "Órdenes total": int(len(view)),
                "Fuente fecha activación": meta["activation_col"] if meta["has_activation_dt"] else "No disponible (Fecha activación / Fecha venta / Venta)",
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