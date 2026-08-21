import os
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Control Gerencial Documental | Electroingeniería",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
)

TODAY = pd.Timestamp("2026-08-21")  # Fecha de corte del dashboard
EXCEL_FILE = "M-FT-1 LIBRO MAESTRO DE DOCUMENTOS v2.xlsx"
LOGO_FILE = "Logos Electroingeniería-01.jpg"

CORPORATE_BLUE = "#00369C"
CORPORATE_YELLOW = "#F6D000"
DARK = "#18212F"
TEXT = "#334155"
MUTED = "#64748B"
LIGHT = "#F5F7FA"
WHITE = "#FFFFFF"
RED = "#C62828"
ORANGE = "#EF6C00"
GREEN = "#2E7D32"

# ============================================================
# CSS
# ============================================================
st.markdown(
    f"""
    <style>
        .stApp {{ background: {LIGHT}; }}
        [data-testid="stSidebar"] {{ background: #FFFFFF; border-right: 1px solid #E2E8F0; }}
        .block-container {{ padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1600px; }}
        h1, h2, h3 {{ color: {DARK}; }}
        .section-title {{
            font-size: 1.05rem; font-weight: 800; color: {DARK};
            margin-top: 0.4rem; margin-bottom: 0.55rem;
        }}
        .section-subtitle {{ font-size: 0.84rem; color: {MUTED}; margin-bottom: 0.8rem; }}
        .hero {{
            background: linear-gradient(135deg, #FFFFFF 0%, #F7FAFF 100%);
            border: 1px solid #E2E8F0; border-radius: 14px; padding: 18px 22px;
            box-shadow: 0 4px 18px rgba(15,23,42,0.05);
        }}
        .hero-kicker {{ color: {CORPORATE_BLUE}; font-size: 0.75rem; font-weight: 800; letter-spacing: .08em; text-transform: uppercase; }}
        .hero-title {{ color: {DARK}; font-size: 1.85rem; line-height: 1.15; font-weight: 800; margin-top: .25rem; }}
        .hero-copy {{ color: {MUTED}; font-size: .92rem; margin-top: .45rem; }}
        .decision-box {{
            background: #FFFFFF; border-left: 5px solid {CORPORATE_BLUE};
            border-top: 1px solid #E2E8F0; border-right: 1px solid #E2E8F0; border-bottom: 1px solid #E2E8F0;
            padding: 14px 16px; border-radius: 8px; margin-bottom: 10px;
        }}
        .decision-label {{ color: {MUTED}; font-size: .72rem; font-weight: 800; text-transform: uppercase; letter-spacing: .06em; }}
        .decision-text {{ color: {DARK}; font-size: .96rem; font-weight: 700; margin-top: .2rem; }}
        .risk-critical {{ border-left-color: {RED}; }}
        .risk-high {{ border-left-color: {ORANGE}; }}
        .risk-medium {{ border-left-color: {CORPORATE_YELLOW}; }}
        .risk-low {{ border-left-color: {GREEN}; }}
        .mini-card {{
            background:#FFFFFF; border:1px solid #E2E8F0; border-radius:12px; padding:12px 14px;
            min-height: 110px; box-shadow:0 2px 10px rgba(15,23,42,.04);
        }}
        .mini-label {{ color:{MUTED}; font-size:.76rem; font-weight:700; }}
        .mini-value {{ color:{DARK}; font-size:1.45rem; font-weight:800; margin-top:.2rem; }}
        .mini-note {{ color:{MUTED}; font-size:.72rem; margin-top:.2rem; }}
        .status-pill {{
            display:inline-block; padding:4px 9px; border-radius:999px; font-size:.72rem; font-weight:800;
        }}
        .good {{ background:#E8F5E9; color:#1B5E20; }}
        .warning {{ background:#FFF4E5; color:#9A3412; }}
        .danger {{ background:#FDECEC; color:#991B1B; }}
        .neutral {{ background:#EEF2F7; color:#475569; }}
        div[data-testid="stMetric"] {{
            background:#FFFFFF; border:1px solid #E2E8F0; border-radius:12px;
            padding:13px 15px; box-shadow:0 2px 10px rgba(15,23,42,.04);
        }}
        div[data-testid="stMetricLabel"] p {{ color:{MUTED}; font-weight:700; font-size:.78rem; }}
        div[data-testid="stMetricValue"] {{ color:{CORPORATE_BLUE}; }}
        button[data-baseweb="tab"] {{ font-weight:700 !important; color:{MUTED} !important; }}
        button[data-baseweb="tab"][aria-selected="true"] {{ color:{CORPORATE_BLUE} !important; }}
        .small-note {{ color:{MUTED}; font-size:.76rem; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# UTILIDADES
# ============================================================
def clean_text(series: pd.Series) -> pd.Series:
    return series.astype(str).replace("nan", np.nan).str.strip()


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = (
        pd.Index(df.columns)
        .astype(str)
        .str.replace("\n", " ", regex=False)
        .str.strip()
    )
    return df


def first_existing(df: pd.DataFrame, options):
    for c in options:
        if c in df.columns:
            return c
    return None


def pct(value, denominator):
    return 0 if not denominator else value / denominator * 100


def safe_mean(s):
    return float(s.dropna().mean()) if s.notna().any() else np.nan


def safe_median(s):
    return float(s.dropna().median()) if s.notna().any() else np.nan


def risk_class(score):
    if score >= 80:
        return "BAJO"
    if score >= 65:
        return "MEDIO"
    if score >= 50:
        return "ALTO"
    return "CRÍTICO"


def risk_css(level):
    return {
        "BAJO": "risk-low",
        "MEDIO": "risk-medium",
        "ALTO": "risk-high",
        "CRÍTICO": "risk-critical",
    }.get(level, "")


def format_days(x):
    if pd.isna(x):
        return "N/D"
    return f"{int(round(x)):,} días".replace(",", ".")


def bar_threshold(x):
    if x < 50:
        return RED
    if x < 65:
        return ORANGE
    if x < 80:
        return CORPORATE_YELLOW
    return GREEN


# ============================================================
# CARGA DE DATOS
# ============================================================
@st.cache_data(show_spinner=False)
def load_workbook(path):
    master = pd.read_excel(path, sheet_name="LIBRO MAESTRO", header=2)
    master = normalize_columns(master)
    master = master.loc[:, ~master.columns.str.contains(r"^Unnamed", case=False, na=False)]
    master = master.dropna(how="all").copy()

    for c in master.select_dtypes(include="object").columns:
        master[c] = clean_text(master[c])

    date_cols = [c for c in master.columns if "FECHA" in c.upper()]
    for c in date_cols:
        master[c] = pd.to_datetime(master[c], errors="coerce")

    binaps = pd.read_excel(path, sheet_name="Binaps")
    binaps = normalize_columns(binaps)
    binaps = binaps.loc[:, ~binaps.columns.str.contains(r"^Unnamed", case=False, na=False)]
    for c in binaps.select_dtypes(include="object").columns:
        binaps[c] = clean_text(binaps[c])

    registro = pd.read_excel(path, sheet_name="Registro Binaps")
    registro = normalize_columns(registro)
    registro = registro.loc[:, ~registro.columns.str.contains(r"^Unnamed", case=False, na=False)]
    for c in registro.select_dtypes(include="object").columns:
        registro[c] = clean_text(registro[c])

    datos = pd.read_excel(path, sheet_name="DATOS")
    datos = normalize_columns(datos)

    return master, binaps, registro, datos


if not os.path.exists(EXCEL_FILE):
    st.error(f"No se encontró el archivo: {EXCEL_FILE}")
    st.stop()

master, binaps, registro_binaps, datos = load_workbook(EXCEL_FILE)

# Renombrado tolerante a espacios
col_unidad = first_existing(master, ["UNIDAD DE NEGOCIO", "UNIDAD DE NEGOCIO "])
col_macro = first_existing(master, ["MACROPROCESO"])
col_proceso = first_existing(master, ["PROCESO"])
col_tipo = first_existing(master, ["TIPO DE DOCUMENTO"])
col_estado = first_existing(master, ["ESTADO"])
col_codigo = first_existing(master, ["CODIGO DEL DOCUMENTO"])
col_nombre = first_existing(master, ["NOMBRE DEL DOCUMENTO"])
col_version = first_existing(master, ["VERSIÓN"])
col_creacion = first_existing(master, ["FECHA DE CREACIÓN"])
col_actualizacion = first_existing(master, ["FECHA DE ACTUALIZACIÓN"])
col_vigencia = first_existing(master, ["FECHA DE VIGENCIA HASTA"])
col_revision = first_existing(master, ["FECHA DE REVISIÓN"])
col_aprobacion = first_existing(master, ["FECHA DE APROBACIÓN"])

# ============================================================
# METRICAS DERIVADAS
# ============================================================
df = master.copy()

if col_estado:
    df["ESTADO_NORMALIZADO"] = df[col_estado].astype(str).str.strip().str.upper()
else:
    df["ESTADO_NORMALIZADO"] = ""

# Estados base
APPROVED = "APROBADO"
NOT_CREATED = "NO CREADO"
REJECTED = "RECHAZADO"
IN_FLOW = {"EN REVISION", "EN APROBACION", "EN ELABORACION"}
ELIMINATED = "ELIMINADO"

# Ciclo creación -> aprobación
if col_creacion and col_aprobacion:
    df["DIAS_CICLO_APROBACION"] = (df[col_aprobacion] - df[col_creacion]).dt.days
else:
    df["DIAS_CICLO_APROBACION"] = np.nan

# Vigencia respecto a fecha de corte
if col_vigencia:
    df["DIAS_PARA_VENCER"] = (df[col_vigencia] - TODAY).dt.days
else:
    df["DIAS_PARA_VENCER"] = np.nan

# Antigüedad / demora documental
if col_creacion:
    df["ANTIGUEDAD_DIAS"] = (TODAY - df[col_creacion]).dt.days
else:
    df["ANTIGUEDAD_DIAS"] = np.nan

# Definiciones de control
active_universe = df[df["ESTADO_NORMALIZADO"] != ELIMINATED].copy()
approved = active_universe[active_universe["ESTADO_NORMALIZADO"] == APPROVED].copy()
not_created = active_universe[active_universe["ESTADO_NORMALIZADO"] == NOT_CREATED].copy()
rejected = active_universe[active_universe["ESTADO_NORMALIZADO"] == REJECTED].copy()
in_flow = active_universe[active_universe["ESTADO_NORMALIZADO"].isin(IN_FLOW)].copy()

# Documentos vencidos: solo aprobados con fecha de vigencia informada
expired = approved[approved["DIAS_PARA_VENCER"].notna() & (approved["DIAS_PARA_VENCER"] < 0)].copy()
next_30 = approved[approved["DIAS_PARA_VENCER"].between(0, 30, inclusive="both")].copy()
next_90 = approved[approved["DIAS_PARA_VENCER"].between(31, 90, inclusive="both")].copy()

# Cobertura documental: aprobados / universo activo
coverage_pct = pct(len(approved), len(active_universe))
# Formalización: aprobados vigentes / universo activo
approved_current = approved[approved["DIAS_PARA_VENCER"].isna() | (approved["DIAS_PARA_VENCER"] >= 0)].copy()
formalization_pct = pct(len(approved_current), len(active_universe))
# Repositorio: cruce por código
binaps_code = first_existing(binaps, ["Codigo Documento", "Codigo_Documento", "Codigo Documento ", "CODIGO"])
master_codes = set(active_universe[col_codigo].dropna().astype(str).str.strip()) if col_codigo else set()
binaps_codes = set(binaps[binaps_code].dropna().astype(str).str.strip()) if binaps_code else set()
matched_codes = master_codes.intersection(binaps_codes)
repository_pct = pct(len(matched_codes), len(master_codes))

# Calidad del dato
required_cols = [c for c in [col_codigo, col_nombre, col_proceso, col_estado] if c]
missing_required_rows = active_universe[required_cols].isna().any(axis=1).sum() if required_cols else 0
duplicated_codes = 0
if col_codigo:
    duplicated_codes = active_universe[col_codigo].dropna().astype(str).str.strip().duplicated(keep=False).sum()

missing_approval_for_approved = int((approved[col_aprobacion].isna()).sum()) if col_aprobacion else 0
missing_expiry_for_approved = int((approved[col_vigencia].isna()).sum()) if col_vigencia else 0

# Calidad (100 = mejor)
data_quality = max(
    0,
    100
    - pct(missing_required_rows, len(active_universe)) * 0.5
    - pct(duplicated_codes, max(len(active_universe), 1)) * 0.3
    - pct(missing_approval_for_approved, max(len(approved), 1)) * 0.1
    - pct(missing_expiry_for_approved, max(len(approved), 1)) * 0.1,
)

# Ciclo de aprobación
cycle = approved["DIAS_CICLO_APROBACION"].dropna()
avg_cycle = cycle.mean() if not cycle.empty else np.nan
median_cycle = cycle.median() if not cycle.empty else np.nan
pct_cycle_le_30 = (cycle.le(30).mean() * 100) if not cycle.empty else np.nan
pct_cycle_gt_365 = (cycle.gt(365).mean() * 100) if not cycle.empty else np.nan
long_cycle = approved[approved["DIAS_CICLO_APROBACION"].gt(365)].copy()

# Índice de salud documental (0-100)
# Componentes: cobertura, vigencia, repositorio, velocidad y calidad de dato.
velocity_component = 100.0
if not cycle.empty:
    velocity_component = max(0, min(100, (pct_cycle_le_30 * 0.7) + ((100 - min(pct_cycle_gt_365, 100)) * 0.3)))
health_score = round(
    coverage_pct * 0.30
    + formalization_pct * 0.20
    + repository_pct * 0.15
    + velocity_component * 0.15
    + data_quality * 0.20,
    1,
)
health_level = risk_class(health_score)

# ============================================================
# FILTROS
# ============================================================
st.sidebar.markdown("### Filtros de control")
st.sidebar.caption(f"Fecha de corte: {TODAY.strftime('%d/%m/%Y')}")

def options_for(col, frame):
    if not col or col not in frame.columns:
        return []
    return sorted(frame[col].dropna().astype(str).str.strip().unique().tolist())

selected_unidad = st.sidebar.multiselect("Unidad de negocio", options_for(col_unidad, df))
filtered = df.copy()
if selected_unidad:
    filtered = filtered[filtered[col_unidad].isin(selected_unidad)]

selected_macro = st.sidebar.multiselect("Macroproceso", options_for(col_macro, filtered))
if selected_macro:
    filtered = filtered[filtered[col_macro].isin(selected_macro)]

selected_proceso = st.sidebar.multiselect("Proceso", options_for(col_proceso, filtered))
if selected_proceso:
    filtered = filtered[filtered[col_proceso].isin(selected_proceso)]

selected_estado = st.sidebar.multiselect("Estado", options_for(col_estado, filtered))
if selected_estado:
    filtered = filtered[filtered[col_estado].isin(selected_estado)]

selected_tipo = st.sidebar.multiselect("Tipo de documento", options_for(col_tipo, filtered))
if selected_tipo:
    filtered = filtered[filtered[col_tipo].isin(selected_tipo)]

search_text = st.sidebar.text_input("Buscar por código o nombre")
if search_text:
    q = search_text.strip().lower()
    m1 = filtered[col_codigo].astype(str).str.lower().str.contains(q, na=False) if col_codigo else False
    m2 = filtered[col_nombre].astype(str).str.lower().str.contains(q, na=False) if col_nombre else False
    filtered = filtered[m1 | m2]

# Recalcular vista filtrada
f_active = filtered[filtered["ESTADO_NORMALIZADO"] != ELIMINATED].copy()
f_approved = f_active[f_active["ESTADO_NORMALIZADO"] == APPROVED].copy()
f_not_created = f_active[f_active["ESTADO_NORMALIZADO"] == NOT_CREATED].copy()
f_rejected = f_active[f_active["ESTADO_NORMALIZADO"] == REJECTED].copy()
f_flow = f_active[f_active["ESTADO_NORMALIZADO"].isin(IN_FLOW)].copy()
f_expired = f_approved[f_approved["DIAS_PARA_VENCER"].notna() & (f_approved["DIAS_PARA_VENCER"] < 0)].copy()
f_next30 = f_approved[f_approved["DIAS_PARA_VENCER"].between(0, 30, inclusive="both")].copy()
f_next90 = f_approved[f_approved["DIAS_PARA_VENCER"].between(31, 90, inclusive="both")].copy()
f_cycle = f_approved["DIAS_CICLO_APROBACION"].dropna()

f_coverage = pct(len(f_approved), len(f_active))
f_current = f_approved[f_approved["DIAS_PARA_VENCER"].isna() | (f_approved["DIAS_PARA_VENCER"] >= 0)]
f_formalization = pct(len(f_current), len(f_active))
f_avg_cycle = f_cycle.mean() if not f_cycle.empty else np.nan
f_median_cycle = f_cycle.median() if not f_cycle.empty else np.nan
f_pct_cycle_30 = f_cycle.le(30).mean() * 100 if not f_cycle.empty else np.nan

# ============================================================
# ENCABEZADO
# ============================================================
header_left, header_right = st.columns([1.8, 5])
with header_left:
    if os.path.exists(LOGO_FILE):
        st.image(LOGO_FILE, width=170)
with header_right:
    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-kicker">Control gerencial</div>
            <div class="hero-title">Estado ejecutivo del sistema documental</div>
            <div class="hero-copy">El dashboard traduce el Libro Maestro en cobertura, riesgo, prioridad y acciones de intervención.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

st.write("")

# ============================================================
# KPI PRINCIPALES
# ============================================================
k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Índice de salud documental", f"{health_score:.1f}/100", health_level)
k2.metric("Cobertura documental", f"{f_coverage:.1f}%", f"{len(f_approved):,} aprobados / {len(f_active):,}")
k3.metric("Backlog no creado", f"{len(f_not_created):,}", f"{pct(len(f_not_created), len(f_active)):.1f}% del universo")
k4.metric("Documentos vencidos", f"{len(f_expired):,}", f"{pct(len(f_expired), len(f_approved)):.1f}% de aprobados")
k5.metric("En riesgo 90 días", f"{len(f_next30)+len(f_next90):,}", f"{len(f_next30)} en 30d / {len(f_next90)} en 31-90d")
k6.metric("Ciclo mediano aprobación", format_days(f_median_cycle), f"Promedio: {format_days(f_avg_cycle)}")

st.write("")

# ============================================================
# ALERTA EJECUTIVA
# ============================================================
alert_messages = []
if len(f_not_created) > 0:
    alert_messages.append(f"{len(f_not_created)} documentos aún no están creados")
if len(f_expired) > 0:
    alert_messages.append(f"{len(f_expired)} documentos aprobados están vencidos")
if len(f_next30) > 0:
    alert_messages.append(f"{len(f_next30)} documentos vencen en los próximos 30 días")
if not pd.isna(f_pct_cycle_30) and f_pct_cycle_30 < 50:
    alert_messages.append(f"solo {f_pct_cycle_30:.1f}% de los ciclos de aprobación termina en ≤30 días")

if alert_messages:
    st.markdown(
        f"""
        <div class="decision-box risk-critical">
            <div class="decision-label">Lectura ejecutiva</div>
            <div class="decision-text">{'; '.join(alert_messages)}.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.success("No se detectan alertas críticas con los filtros seleccionados.")

# ============================================================
# TABS
# ============================================================
(
    tab_exec,
    tab_risk,
    tab_causes,
    tab_actions,
    tab_future,
    tab_repo,
    tab_detail,
) = st.tabs(
    [
        "1. Ejecutivo",
        "2. Riesgo",
        "3. Causas",
        "4. Prioridades",
        "5. Riesgo futuro",
        "6. Binaps",
        "7. Detalle",
    ]
)

# ============================================================
# 1. EJECUTIVO
# ============================================================
with tab_exec:
    st.markdown('<div class="section-title">¿Estamos controlando el sistema documental?</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">La lectura está orientada a responder si el sistema está formalizado, vigente, disponible y gestionable.</div>', unsafe_allow_html=True)

    c1, c2 = st.columns([1.25, 1])
    with c1:
        gauge = go.Figure(go.Indicator(
            mode="gauge+number",
            value=health_score,
            number={"suffix": "/100", "font": {"size": 34, "color": DARK}},
            title={"text": "Índice de Salud Documental", "font": {"size": 18, "color": DARK}},
            gauge={
                "axis": {"range": [0, 100], "tickwidth": 1},
                "bar": {"color": CORPORATE_BLUE},
                "steps": [
                    {"range": [0, 50], "color": "#FDECEC"},
                    {"range": [50, 65], "color": "#FFF4E5"},
                    {"range": [65, 80], "color": "#FFF9DB"},
                    {"range": [80, 100], "color": "#E8F5E9"},
                ],
                "threshold": {"line": {"color": CORPORATE_YELLOW, "width": 4}, "value": 80},
            },
        ))
        gauge.update_layout(height=290, margin=dict(l=10, r=10, t=55, b=10), paper_bgcolor="white")
        st.plotly_chart(gauge, use_container_width=True, config={"displayModeBar": False})

    with c2:
        components = pd.DataFrame({
            "Componente": ["Cobertura", "Formalización", "Repositorio", "Velocidad", "Calidad del dato"],
            "Puntaje": [f_coverage, f_formalization, repository_pct, velocity_component, data_quality],
        }).sort_values("Puntaje")
        fig = px.bar(components, x="Puntaje", y="Componente", orientation="h", text="Puntaje")
        fig.update_traces(marker_color=CORPORATE_BLUE, texttemplate="%{text:.1f}", textposition="outside")
        fig.update_layout(height=290, margin=dict(l=10, r=40, t=25, b=10), xaxis=dict(range=[0, 110], title="Puntaje"), yaxis_title="")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown('<div class="section-title">¿Qué decisión debería tomar gerencia?</div>', unsafe_allow_html=True)
    decision_cols = st.columns(4)
    decisions = [
        ("Intervención inmediata", f"Atacar {len(f_expired)} documentos vencidos y {len(f_not_created)} no creados.", "danger" if (len(f_expired)+len(f_not_created)) else "good"),
        ("Velocidad", f"El ciclo mediano es {format_days(f_median_cycle)} y {f_pct_cycle_30:.1f}% termina en ≤30 días." if not pd.isna(f_pct_cycle_30) else "No hay suficientes datos de ciclo.", "warning" if not pd.isna(f_pct_cycle_30) and f_pct_cycle_30 < 70 else "good"),
        ("Disponibilidad", f"{repository_pct:.1f}% de los códigos activos aparecen en Binaps.", "warning" if repository_pct < 90 else "good"),
        ("Vigencia", f"{len(f_next30)+len(f_next90)} documentos presentan riesgo de vencimiento en 90 días.", "warning" if len(f_next30)+len(f_next90) else "good"),
    ]
    for col, (label, text, cls) in zip(decision_cols, decisions):
        with col:
            st.markdown(f'<div class="decision-box {risk_css("CRÍTICO" if cls=="danger" else "ALTO" if cls=="warning" else "BAJO")}"><div class="decision-label">{label}</div><div class="decision-text">{text}</div></div>', unsafe_allow_html=True)

# ============================================================
# 2. RIESGO
# ============================================================
with tab_risk:
    st.markdown('<div class="section-title">Mapa de riesgo por proceso</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Cada proceso obtiene un puntaje de salud considerando formalización, vencimiento, backlog y desempeño del ciclo de aprobación.</div>', unsafe_allow_html=True)

    group = col_proceso
    rows = []
    for process, g in f_active.groupby(group, dropna=False):
        approved_g = g[g["ESTADO_NORMALIZADO"] == APPROVED]
        expired_g = approved_g[approved_g["DIAS_PARA_VENCER"].notna() & (approved_g["DIAS_PARA_VENCER"] < 0)]
        pending_g = g[g["ESTADO_NORMALIZADO"] == NOT_CREATED]
        cycle_g = approved_g["DIAS_CICLO_APROBACION"].dropna()
        coverage_g = pct(len(approved_g), len(g))
        current_g = approved_g[approved_g["DIAS_PARA_VENCER"].isna() | (approved_g["DIAS_PARA_VENCER"] >= 0)]
        current_g_pct = pct(len(current_g), len(g))
        repo_g_pct = pct(sum(str(c).strip() in binaps_codes for c in g[col_codigo].dropna().astype(str)) if col_codigo else 0, len(g))
        vel_g = cycle_g.le(30).mean() * 100 if not cycle_g.empty else 50
        health_g = round(coverage_g * .30 + current_g_pct * .25 + repo_g_pct * .15 + vel_g * .10 + max(0, 100 - pct(len(pending_g), len(g)))*.20, 1)
        rows.append({
            "Proceso": process if pd.notna(process) else "Sin proceso",
            "Documentos": len(g),
            "Aprobados": len(approved_g),
            "No creados": len(pending_g),
            "Vencidos": len(expired_g),
            "Salud": health_g,
            "Riesgo": risk_class(health_g),
            "Mediana ciclo": cycle_g.median() if not cycle_g.empty else np.nan,
        })
    risk_df = pd.DataFrame(rows).sort_values(["Salud", "Vencidos", "No creados"], ascending=[True, False, False])

    rc1, rc2 = st.columns([1.5, 1])
    with rc1:
        fig = px.scatter(
            risk_df,
            x="No creados",
            y="Vencidos",
            size="Documentos",
            color="Salud",
            hover_name="Proceso",
            hover_data={"Salud": ":.1f", "No creados": True, "Vencidos": True, "Documentos": True},
            color_continuous_scale=[(0, RED), (0.35, ORANGE), (0.65, CORPORATE_YELLOW), (1, GREEN)],
            range_color=[0, 100],
        )
        fig.add_hline(y=2, line_dash="dot", line_color="#94A3B8")
        fig.add_vline(x=5, line_dash="dot", line_color="#94A3B8")
        fig.update_layout(height=430, xaxis_title="Backlog no creado", yaxis_title="Documentos vencidos", coloraxis_colorbar_title="Salud")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})
    with rc2:
        top_risk = risk_df.head(10).copy()
        top_risk["Salud"] = top_risk["Salud"].round(1)
        st.dataframe(top_risk, use_container_width=True, hide_index=True, height=430)

    st.markdown('<div class="section-title">Lectura del riesgo</div>', unsafe_allow_html=True)
    r1, r2, r3, r4 = st.columns(4)
    counts = risk_df["Riesgo"].value_counts()
    r1.metric("Críticos", int(counts.get("CRÍTICO", 0)))
    r2.metric("Altos", int(counts.get("ALTO", 0)))
    r3.metric("Medios", int(counts.get("MEDIO", 0)))
    r4.metric("Bajos", int(counts.get("BAJO", 0)))

# ============================================================
# 3. CAUSAS
# ============================================================
with tab_causes:
    st.markdown('<div class="section-title">¿Por qué el resultado es bajo?</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Separamos las causas que generan el riesgo para evitar que la gerencia se quede solo con el indicador.</div>', unsafe_allow_html=True)

    cause_data = pd.DataFrame({
        "Causa": [
            "Documentos no creados",
            "Documentos aprobados vencidos",
            "Aprobaciones >365 días",
            "Documentos en flujo",
            "Rechazados",
            "Códigos no encontrados en Binaps",
            "Filas con campos obligatorios vacíos",
            "Códigos duplicados",
        ],
        "Cantidad": [
            len(f_not_created),
            len(f_expired),
            len(f_approved[f_approved["DIAS_CICLO_APROBACION"].gt(365)]),
            len(f_flow),
            len(f_rejected),
            max(len(master_codes - binaps_codes), 0),
            int(missing_required_rows),
            int(duplicated_codes),
        ],
    }).sort_values("Cantidad", ascending=True)

    fig = px.bar(cause_data, x="Cantidad", y="Causa", orientation="h", text="Cantidad")
    fig.update_traces(marker_color=CORPORATE_BLUE, textposition="outside")
    fig.update_layout(height=430, margin=dict(l=10, r=45, t=20, b=10), xaxis_title="Casos", yaxis_title="")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    if col_macro and col_proceso:
        st.markdown('<div class="section-title">¿En qué macroprocesos se concentra el problema?</div>', unsafe_allow_html=True)
        mp = f_active.groupby(col_macro).agg(
            Documentos=(col_codigo, "count") if col_codigo else ("ESTADO_NORMALIZADO", "count"),
            No_creados=("ESTADO_NORMALIZADO", lambda x: (x == NOT_CREATED).sum()),
            Aprobados=("ESTADO_NORMALIZADO", lambda x: (x == APPROVED).sum()),
        ).reset_index()
        mp["Cumplimiento"] = (mp["Aprobados"] / mp["Documentos"] * 100).round(1)
        mp = mp.sort_values("Cumplimiento")
        fig = px.bar(mp, x="Cumplimiento", y=col_macro, orientation="h", text="Cumplimiento")
        fig.update_traces(marker_color=CORPORATE_BLUE, texttemplate="%{text:.1f}%", textposition="outside")
        fig.update_layout(height=max(350, 35 * len(mp)), xaxis=dict(range=[0, 110], title="% aprobados"), yaxis_title="")
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

# ============================================================
# 4. PRIORIDADES
# ============================================================
with tab_actions:
    st.markdown('<div class="section-title">Matriz de intervención</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">La prioridad combina tamaño del problema, vencimiento y backlog. El objetivo es definir dónde intervenir primero.</div>', unsafe_allow_html=True)

    p = risk_df.copy()
    p["Urgencia"] = (p["Vencidos"] * 2 + p["No creados"]).astype(float)
    p["Impacto"] = (p["Documentos"] * (100 - p["Salud"]) / 100).astype(float)
    p["Prioridad"] = p["Urgencia"] * 0.55 + p["Impacto"] * 0.45
    p = p.sort_values("Prioridad", ascending=False)

    fig = px.scatter(
        p.head(30),
        x="Impacto",
        y="Urgencia",
        size="Documentos",
        color="Salud",
        hover_name="Proceso",
        text="Proceso",
        color_continuous_scale=[(0, RED), (0.35, ORANGE), (0.65, CORPORATE_YELLOW), (1, GREEN)],
        range_color=[0, 100],
    )
    fig.update_traces(textposition="top center")
    fig.add_vline(x=p["Impacto"].median(), line_dash="dash", line_color="#94A3B8")
    fig.add_hline(y=p["Urgencia"].median(), line_dash="dash", line_color="#94A3B8")
    fig.update_layout(height=500, xaxis_title="Impacto", yaxis_title="Urgencia", coloraxis_colorbar_title="Salud")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown('<div class="section-title">Top de prioridades para comité gerencial</div>', unsafe_allow_html=True)
    top = p.head(10).copy()
    top["Prioridad"] = top["Prioridad"].round(1)
    top["Salud"] = top["Salud"].round(1)
    st.dataframe(top[["Proceso", "Documentos", "No creados", "Vencidos", "Salud", "Riesgo", "Prioridad"]], use_container_width=True, hide_index=True)

    # Detalle accionable
    selected_priority_process = st.selectbox("Ver documentos del proceso prioritario", top["Proceso"].tolist() if not top.empty else [])
    if selected_priority_process:
        action_docs = f_active[f_active[col_proceso].eq(selected_priority_process)].copy()
        action_docs["Prioridad documental"] = np.where(
            action_docs["ESTADO_NORMALIZADO"].eq(NOT_CREATED), "ALTA - NO CREADO",
            np.where(action_docs["DIAS_PARA_VENCER"].lt(0), "CRÍTICA - VENCIDO", "CONTROL")
        )
        show_cols = [c for c in [col_codigo, col_nombre, col_estado, col_vigencia, "DIAS_PARA_VENCER", "DIAS_CICLO_APROBACION", "Prioridad documental"] if c]
        st.dataframe(action_docs[show_cols].sort_values("Prioridad documental"), use_container_width=True, hide_index=True, height=350)

# ============================================================
# 5. RIESGO FUTURO
# ============================================================
with tab_future:
    st.markdown('<div class="section-title">¿Qué puede empeorar si no intervenimos?</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Proyección simple basada en las fechas de vigencia registradas en el Libro Maestro.</div>', unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns(4)
    f1.metric("Ya vencidos", len(f_expired))
    f2.metric("Vencen ≤30 días", len(f_next30))
    f3.metric("Vencen 31-90 días", len(f_next90))
    f4.metric("Sin fecha de vigencia", int(f_approved[col_vigencia].isna().sum()) if col_vigencia else 0)

    buckets = pd.DataFrame({
        "Horizonte": ["Vencidos", "0-30 días", "31-90 días", "91+ días"],
        "Documentos": [
            len(f_expired),
            len(f_next30),
            len(f_next90),
            len(f_approved[f_approved["DIAS_PARA_VENCER"].gt(90)]),
        ],
    })
    fig = px.bar(buckets, x="Horizonte", y="Documentos", text="Documentos")
    fig.update_traces(marker_color=CORPORATE_BLUE, textposition="outside")
    fig.update_layout(height=350, yaxis_title="Documentos", xaxis_title="")
    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    future_detail = pd.concat([f_expired, f_next30, f_next90], ignore_index=True)
    if not future_detail.empty:
        st.markdown('<div class="section-title">Documentos que requieren anticipación</div>', unsafe_allow_html=True)
        show_cols = [c for c in [col_codigo, col_nombre, col_proceso, col_estado, col_vigencia, "DIAS_PARA_VENCER"] if c]
        st.dataframe(
            future_detail.sort_values("DIAS_PARA_VENCER")[show_cols],
            use_container_width=True,
            hide_index=True,
            height=380,
        )

# ============================================================
# 6. BINAPS
# ============================================================
with tab_repo:
    st.markdown('<div class="section-title">Conciliación Libro Maestro vs Binaps</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Permite identificar si lo que el sistema considera formalizado también está disponible en el repositorio.</div>', unsafe_allow_html=True)

    b1, b2, b3, b4 = st.columns(4)
    b1.metric("Códigos activos en Maestro", len(master_codes))
    b2.metric("Códigos encontrados en Binaps", len(matched_codes))
    b3.metric("No encontrados", max(len(master_codes - binaps_codes), 0))
    b4.metric("Cobertura repositorio", f"{repository_pct:.1f}%")

    if col_codigo and binaps_code:
        maestro_repo = active_universe[[col_codigo, col_nombre, col_proceso, col_estado]].copy()
        maestro_repo["CODIGO_KEY"] = maestro_repo[col_codigo].astype(str).str.strip()
        maestro_repo["EN_BINAPS"] = maestro_repo["CODIGO_KEY"].isin(binaps_codes)
        missing_repo = maestro_repo[~maestro_repo["EN_BINAPS"]].copy()
        st.markdown('<div class="section-title">Documentos del Maestro sin coincidencia en Binaps</div>', unsafe_allow_html=True)
        st.dataframe(missing_repo.drop(columns=["CODIGO_KEY", "EN_BINAPS"]), use_container_width=True, hide_index=True, height=430)

    # Coincidencia por nombre, como señal de inconsistencias
    binaps_name = first_existing(binaps, ["Nombre del Documento", "Nombre_Documento"])
    if col_nombre and binaps_name:
        master_name_set = set(active_universe[col_nombre].dropna().astype(str).str.strip().str.lower())
        binaps_name_set = set(binaps[binaps_name].dropna().astype(str).str.strip().str.lower())
        common_names = master_name_set.intersection(binaps_name_set)
        st.caption(f"Coincidencia exacta por nombre: {pct(len(common_names), len(master_name_set)):.1f}% del conjunto activo del Maestro.")

# ============================================================
# 7. DETALLE
# ============================================================
with tab_detail:
    st.markdown('<div class="section-title">Libro Maestro filtrado</div>', unsafe_allow_html=True)
    st.markdown('<div class="section-subtitle">Esta pestaña es deliberadamente operativa: aquí se baja desde la decisión hasta el documento concreto.</div>', unsafe_allow_html=True)

    detail = filtered.copy()
    priority_flag = []
    for _, r in detail.iterrows():
        if r["ESTADO_NORMALIZADO"] == NOT_CREATED:
            priority_flag.append("ALTA - NO CREADO")
        elif r["ESTADO_NORMALIZADO"] == REJECTED:
            priority_flag.append("ALTA - RECHAZADO")
        elif r["DIAS_PARA_VENCER"] < 0:
            priority_flag.append("CRÍTICA - VENCIDO")
        elif pd.notna(r["DIAS_PARA_VENCER"]) and r["DIAS_PARA_VENCER"] <= 30:
            priority_flag.append("ALTA - VENCE ≤30D")
        elif pd.notna(r["DIAS_CICLO_APROBACION"]) and r["DIAS_CICLO_APROBACION"] > 365:
            priority_flag.append("ALTA - CICLO >365D")
        else:
            priority_flag.append("CONTROL")
    detail.insert(0, "PRIORIDAD GERENCIAL", priority_flag)

    st.dataframe(detail, use_container_width=True, height=600, hide_index=True)

    csv = detail.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Descargar vista filtrada CSV",
        data=csv,
        file_name="libro_maestro_vista_gerencial.csv",
        mime="text/csv",
    )

# ============================================================
# PIE
# ============================================================
st.divider()
st.caption(
    "Modelo gerencial: cobertura, formalización, vigencia, repositorio, velocidad y calidad del dato. "
    "Los puntajes son reglas de gestión configurables y deben validarse con el criterio de la organización."
)
