import os
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# ============================================================
# CONFIGURACIÓN
# ============================================================
st.set_page_config(
    page_title="Libro Maestro | Electroingeniería",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

BASE_DIR = Path(__file__).resolve().parent
EXCEL_FILE = BASE_DIR / "M-FT-1 LIBRO MAESTRO DE DOCUMENTOS v2.xlsx"
LOGO_FILE = BASE_DIR / "Logos Electroingeniería-01.jpg"

# Paleta
BLUE = "#00369C"
BLUE_DARK = "#00265F"
YELLOW = "#F6D000"
BG = "#F5F7FA"
TEXT = "#243447"
MUTED = "#6B7280"
GREEN = "#17803D"
RED = "#C62828"
ORANGE = "#C76A00"

# ============================================================
# CSS
# ============================================================
st.markdown(
    f"""
    <style>
        .stApp {{ background: {BG}; color: {TEXT}; }}
        #MainMenu, footer, header {{ visibility: hidden; }}

        .block-container {{
            padding-top: 1.2rem;
            padding-bottom: 2rem;
            max-width: 1600px;
        }}

        h1, h2, h3 {{
            color: {BLUE_DARK};
            letter-spacing: -0.02em;
        }}

        .dashboard-title {{
            font-size: 2.25rem;
            font-weight: 800;
            color: {BLUE_DARK};
            margin-bottom: 0.1rem;
        }}

        .dashboard-subtitle {{
            color: {MUTED};
            font-size: 0.95rem;
            margin-bottom: 1rem;
        }}

        .section-title {{
            font-size: 1.1rem;
            font-weight: 750;
            color: {BLUE_DARK};
            margin: 1rem 0 0.5rem 0;
        }}

        div[data-testid="stMetric"] {{
            background: white;
            border: 1px solid #E5E7EB;
            border-left: 5px solid {BLUE};
            border-radius: 10px;
            padding: 14px 16px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.035);
        }}

        div[data-testid="stMetric"] label {{
            color: {MUTED} !important;
            font-weight: 650 !important;
        }}

        div[data-testid="stMetricValue"] {{
            color: {BLUE} !important;
            font-weight: 800 !important;
        }}

        .alert-card {{
            background: white;
            border: 1px solid #E5E7EB;
            border-radius: 10px;
            padding: 14px 16px;
            min-height: 90px;
        }}

        .alert-card .value {{
            font-size: 1.55rem;
            font-weight: 800;
            color: {BLUE};
        }}

        .alert-card .label {{
            font-size: 0.85rem;
            color: {MUTED};
        }}

        .small-note {{
            font-size: 0.78rem;
            color: {MUTED};
        }}

        .stTabs [data-baseweb="tab"] {{
            font-weight: 700;
        }}

        .stTabs [aria-selected="true"] {{
            color: {BLUE} !important;
            border-bottom-color: {YELLOW} !important;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# FUNCIONES
# ============================================================
REQUIRED_COLUMNS = [
    "UNIDAD DE NEGOCIO",
    "MACROPROCESO",
    "PROCESO",
    "TIPO DE DOCUMENTO",
    "SIGLAS",
    "T. DOC",
    "CONSECUTIVO",
    "CODIGO DEL DOCUMENTO",
    "NOMBRE DEL DOCUMENTO",
    "VERSIÓN",
    "ESTADO",
    "FECHA DE CREACIÓN",
    "FECHA DE ACTUALIZACIÓN",
    "FECHA DE VIGENCIA HASTA",
    "FECHA DE REVISIÓN",
    "FECHA DE APROBACIÓN",
    "DOCUMENTOS RELACIONADOS",
    "NORMATIVIDAD RELACIONADA",
    "MANEJO DEL DOCUMENTO",
    "OBSERVACIONES",
]

DATE_COLUMNS = [
    "FECHA DE CREACIÓN",
    "FECHA DE ACTUALIZACIÓN",
    "FECHA DE VIGENCIA HASTA",
    "FECHA DE REVISIÓN",
    "FECHA DE APROBACIÓN",
]


def clean_text(series):
    return (
        series.astype("string")
        .str.strip()
        .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    )


@st.cache_data(show_spinner=False)
def load_data():
    if not EXCEL_FILE.exists():
        st.error(f"No se encontró el archivo: {EXCEL_FILE.name}")
        st.stop()

    df = pd.read_excel(EXCEL_FILE, sheet_name="LIBRO MAESTRO", header=2)
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed", case=False)]
    df = df.dropna(how="all").copy()

    # Normaliza columnas faltantes
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    # Texto
    text_cols = [
        c for c in REQUIRED_COLUMNS
        if c not in DATE_COLUMNS
    ]
    for col in text_cols:
        df[col] = clean_text(df[col])

    # Fechas
    for col in DATE_COLUMNS:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    # Fecha de corte dinámica
    today = pd.Timestamp.today().normalize()

    # Duraciones correctas
    df["DIAS_HASTA_APROBACION"] = (
        df["FECHA DE APROBACIÓN"] - df["FECHA DE CREACIÓN"]
    ).dt.days

    # Esto no se presenta como "tiempo de revisión".
    # Es la antigüedad desde la creación hasta la fecha de revisión registrada.
    df["DIAS_CREACION_A_REVISION"] = (
        df["FECHA DE REVISIÓN"] - df["FECHA DE CREACIÓN"]
    ).dt.days

    df["DIAS_PARA_VENCER"] = (
        df["FECHA DE VIGENCIA HASTA"] - today
    ).dt.days

    df["DIAS_DESDE_REVISION"] = (
        today - df["FECHA DE REVISIÓN"]
    ).dt.days

    # Clasificación de vigencia
    def vigencia(row):
        fecha = row["FECHA DE VIGENCIA HASTA"]
        estado = str(row["ESTADO"] or "").strip().upper()
        if pd.isna(fecha) or estado != "APROBADO":
            return "No aplica"
        dias = row["DIAS_PARA_VENCER"]
        if dias < 0:
            return "Vencido"
        if dias <= 30:
            return "0–30 días"
        if dias <= 90:
            return "31–90 días"
        if dias <= 180:
            return "91–180 días"
        if dias <= 365:
            return "181–365 días"
        return ">365 días"

    df["ESTADO_VIGENCIA"] = df.apply(vigencia, axis=1)

    # Clasificación de flujo
    def flujo(estado):
        e = str(estado or "").strip().upper()
        if e == "APROBADO":
            return "Aprobado"
        if e in {"NO CREADO", "EN ELABORACION", "EN REVISION", "EN APROBACION"}:
            return "Pendiente de gestión"
        if e == "RECHAZADO":
            return "Rechazado"
        if e == "ELIMINADO":
            return "Eliminado"
        return "Otro"

    df["GRUPO_ESTADO"] = df["ESTADO"].apply(flujo)

    # Indicadores de calidad del dato
    df["DATO_CRITICO_FALTANTE"] = False
    critical_cols = [
        "UNIDAD DE NEGOCIO",
        "MACROPROCESO",
        "PROCESO",
        "TIPO DE DOCUMENTO",
        "CODIGO DEL DOCUMENTO",
        "NOMBRE DEL DOCUMENTO",
        "ESTADO",
    ]
    df["DATO_CRITICO_FALTANTE"] = df[critical_cols].isna().any(axis=1)

    # Códigos repetidos. En este libro existen códigos plantilla como H-PR-,
    # por lo que se marcan para revisión y no se eliminan automáticamente.
    df["CODIGO_REPETIDO"] = df["CODIGO DEL DOCUMENTO"].duplicated(keep=False)

    return df


@st.cache_data(show_spinner=False)
def load_binaps():
    """Carga el inventario actual del repositorio Binaps del mismo Excel."""
    try:
        b = pd.read_excel(EXCEL_FILE, sheet_name="Binaps")
    except Exception:
        return pd.DataFrame(columns=["CODIGO DEL DOCUMENTO", "ESTADO_BINAPS", "NOMBRE_BINAPS", "VERSION_BINAPS"])

    b.columns = b.columns.astype(str).str.strip()

    rename = {
        "Codigo Documento": "CODIGO DEL DOCUMENTO",
        "Estado": "ESTADO_BINAPS",
        "Nombre del Documento": "NOMBRE_BINAPS",
        "VERSION": "VERSION_BINAPS",
        "Fecha_Creacion": "FECHA_CREACION_BINAPS",
        "Fecha Revision": "FECHA_REVISION_BINAPS",
        "Fecha_Aprobacion": "FECHA_APROBACION_BINAPS",
        "Fecha_Vigencia": "FECHA_VIGENCIA_BINAPS",
        "Tipo Documento": "TIPO_BINAPS",
        "Macroproceso": "MACROPROCESO_BINAPS",
        "Proceso": "PROCESO_BINAPS",
    }
    b = b.rename(columns=rename)

    for col in b.columns:
        if "FECHA" in col.upper():
            b[col] = pd.to_datetime(b[col], errors="coerce")
        elif col != "VERSION_BINAPS":
            b[col] = clean_text(b[col])

    wanted = [
        "CODIGO DEL DOCUMENTO",
        "ESTADO_BINAPS",
        "NOMBRE_BINAPS",
        "VERSION_BINAPS",
        "TIPO_BINAPS",
        "MACROPROCESO_BINAPS",
        "PROCESO_BINAPS",
        "FECHA_CREACION_BINAPS",
        "FECHA_REVISION_BINAPS",
        "FECHA_APROBACION_BINAPS",
        "FECHA_VIGENCIA_BINAPS",
    ]
    for col in wanted:
        if col not in b.columns:
            b[col] = pd.NA

    return b[wanted].copy()


def apply_filters(df):
    st.sidebar.markdown("## Filtros de control")

    units = sorted(df["UNIDAD DE NEGOCIO"].dropna().unique().tolist())
    selected_units = st.sidebar.multiselect(
        "Unidad de negocio",
        units,
        placeholder="Todas",
    )

    df1 = df[df["UNIDAD DE NEGOCIO"].isin(selected_units)] if selected_units else df

    macros = sorted(df1["MACROPROCESO"].dropna().unique().tolist())
    selected_macros = st.sidebar.multiselect(
        "Macroproceso",
        macros,
        placeholder="Todos",
    )

    df2 = df1[df1["MACROPROCESO"].isin(selected_macros)] if selected_macros else df1

    processes = sorted(df2["PROCESO"].dropna().unique().tolist())
    selected_processes = st.sidebar.multiselect(
        "Proceso",
        processes,
        placeholder="Todos",
    )

    df3 = (
        df2[df2["PROCESO"].isin(selected_processes)]
        if selected_processes else df2
    )

    states = sorted(df3["ESTADO"].dropna().unique().tolist())
    selected_states = st.sidebar.multiselect(
        "Estado",
        states,
        placeholder="Todos",
    )

    df4 = (
        df3[df3["ESTADO"].isin(selected_states)]
        if selected_states else df3
    )

    types = sorted(df4["TIPO DE DOCUMENTO"].dropna().unique().tolist())
    selected_types = st.sidebar.multiselect(
        "Tipo de documento",
        types,
        placeholder="Todos",
    )

    df5 = (
        df4[df4["TIPO DE DOCUMENTO"].isin(selected_types)]
        if selected_types else df4
    )

    validity = sorted(df5["ESTADO_VIGENCIA"].dropna().unique().tolist())
    selected_validity = st.sidebar.multiselect(
        "Vigencia",
        validity,
        placeholder="Todas",
    )

    if selected_validity:
        df5 = df5[df5["ESTADO_VIGENCIA"].isin(selected_validity)]

    # Buscador
    search = st.sidebar.text_input(
        "Buscar documento",
        placeholder="Código o nombre..."
    )

    if search:
        term = search.strip().lower()
        mask = (
            df5["CODIGO DEL DOCUMENTO"].fillna("").str.lower().str.contains(term, regex=False)
            |
            df5["NOMBRE DEL DOCUMENTO"].fillna("").str.lower().str.contains(term, regex=False)
        )
        df5 = df5[mask]

    st.sidebar.divider()
    st.sidebar.caption(f"Registros visibles: {len(df5):,}")
    st.sidebar.caption("Los filtros son dependientes y se actualizan entre sí.")

    return df5


def format_int(value):
    return f"{int(value):,}".replace(",", ".")


def safe_mean(series):
    x = pd.to_numeric(series, errors="coerce").dropna()
    return x.mean() if len(x) else np.nan


def pct(a, b):
    return (a / b * 100) if b else 0


def base_chart(fig, height=360):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=45, b=10),
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(color=TEXT),
        title_font=dict(color=BLUE_DARK, size=16),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
    )
    return fig


# ============================================================
# CARGA
# ============================================================
df = load_data()
df_binaps = load_binaps()
df_filtered = apply_filters(df)

today = pd.Timestamp.today().normalize()

# ============================================================
# ENCABEZADO
# ============================================================
head1, head2 = st.columns([1, 7])

with head1:
    if LOGO_FILE.exists():
        st.image(str(LOGO_FILE), width=145)

with head2:
    st.markdown('<div class="dashboard-title">Libro Maestro de Documentos</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="dashboard-subtitle">Control documental · estructura · flujo de aprobación · vigencia · calidad del dato</div>',
        unsafe_allow_html=True
    )

c1, c2, c3 = st.columns([2, 2, 4])
with c1:
    st.caption(f"Fecha de corte: {today.strftime('%d/%m/%Y')}")
with c2:
    st.caption(f"Fuente: {EXCEL_FILE.name}")
with c3:
    st.caption("Los indicadores se calculan sobre el conjunto filtrado.")

# ============================================================
# KPIs PRINCIPALES
# ============================================================
active = df_filtered[
    ~df_filtered["ESTADO"].fillna("").str.upper().isin(["ELIMINADO"])
]
approved = active[active["ESTADO"].fillna("").str.upper() == "APROBADO"]
not_created = active[active["ESTADO"].fillna("").str.upper() == "NO CREADO"]
workflow = active[
    active["ESTADO"].fillna("").str.upper().isin(
        ["EN ELABORACION", "EN REVISION", "EN APROBACION"]
    )
]
rejected = active[active["ESTADO"].fillna("").str.upper() == "RECHAZADO"]

expired = approved[approved["DIAS_PARA_VENCER"] < 0]
expiring_90 = approved[
    approved["DIAS_PARA_VENCER"].between(0, 90, inclusive="both")
]

approval_days = approved["DIAS_HASTA_APROBACION"].dropna()
approval_median = approval_days.median() if len(approval_days) else np.nan
approval_avg = approval_days.mean() if len(approval_days) else np.nan

formalization = pct(len(approved), len(active))

k1, k2, k3, k4, k5, k6 = st.columns(6)
k1.metric("Documentos activos", format_int(len(active)))
k2.metric("Aprobados", format_int(len(approved)), f"{formalization:.1f}% formalizados")
k3.metric("No creados", format_int(len(not_created)))
k4.metric("En flujo", format_int(len(workflow)))
k5.metric("Vencidos", format_int(len(expired)))
k6.metric("Vencen ≤90 días", format_int(len(expiring_90)))

st.divider()

# ============================================================
# TABS
# ============================================================
(
    tab_resumen,
    tab_flujo,
    tab_tiempos,
    tab_vigencia,
    tab_estructura,
    tab_repositorio,
    tab_libro,
    tab_calidad,
) = st.tabs(
    [
        "Resumen ejecutivo",
        "Flujo documental",
        "Tiempos",
        "Vigencia",
        "Estructura",
        "Repositorio Binaps",
        "Libro Maestro",
        "Calidad del dato",
    ]
)

# ============================================================
# 1. RESUMEN EJECUTIVO
# ============================================================
with tab_resumen:
    st.markdown('<div class="section-title">Situación general del sistema documental</div>', unsafe_allow_html=True)

    left, right = st.columns(2)

    with left:
        status_counts = active["ESTADO"].value_counts().reset_index()
        status_counts.columns = ["ESTADO", "CANTIDAD"]
        fig = px.bar(
            status_counts.sort_values("CANTIDAD"),
            x="CANTIDAD",
            y="ESTADO",
            orientation="h",
            title="Documentos por estado",
            text="CANTIDAD",
        )
        fig.update_traces(marker_color=BLUE, textposition="outside")
        base_chart(fig, 330)
        st.plotly_chart(fig, use_container_width=True)

    with right:
        type_counts = active["TIPO DE DOCUMENTO"].value_counts().reset_index()
        type_counts.columns = ["TIPO DE DOCUMENTO", "CANTIDAD"]
        fig = px.bar(
            type_counts.sort_values("CANTIDAD"),
            x="CANTIDAD",
            y="TIPO DE DOCUMENTO",
            orientation="h",
            title="Volumen por tipo de documento",
            text="CANTIDAD",
        )
        fig.update_traces(marker_color=BLUE, textposition="outside")
        base_chart(fig, 330)
        st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)

    with c1:
        by_macro = (
            active.groupby("MACROPROCESO")
            .size()
            .reset_index(name="DOCUMENTOS")
            .sort_values("DOCUMENTOS", ascending=False)
            .head(12)
        )
        fig = px.bar(
            by_macro,
            x="DOCUMENTOS",
            y="MACROPROCESO",
            orientation="h",
            title="Distribución documental por macroproceso",
            text="DOCUMENTOS",
        )
        fig.update_traces(marker_color=BLUE, textposition="outside")
        base_chart(fig, 400)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        trend = (
            active.assign(AÑO=active["FECHA DE CREACIÓN"].dt.year)
            .dropna(subset=["AÑO"])
            .groupby(["AÑO", "ESTADO"])
            .size()
            .reset_index(name="DOCUMENTOS")
        )
        fig = px.line(
            trend,
            x="AÑO",
            y="DOCUMENTOS",
            color="ESTADO",
            markers=True,
            title="Evolución de creación documental",
        )
        base_chart(fig, 400)
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# 2. FLUJO DOCUMENTAL
# ============================================================
with tab_flujo:
    st.markdown('<div class="section-title">Dónde está el atasco documental</div>', unsafe_allow_html=True)

    p = (
        active.groupby(["PROCESO", "GRUPO_ESTADO"])
        .size()
        .reset_index(name="DOCUMENTOS")
    )

    fig = px.bar(
        p,
        x="DOCUMENTOS",
        y="PROCESO",
        color="GRUPO_ESTADO",
        orientation="h",
        barmode="stack",
        title="Estado documental por proceso",
        category_orders={"PROCESO": p.groupby("PROCESO")["DOCUMENTOS"].sum().sort_values().index.tolist()},
    )
    base_chart(fig, max(420, min(900, 24 * p["PROCESO"].nunique())))
    st.plotly_chart(fig, use_container_width=True)

    c1, c2 = st.columns(2)

    with c1:
        backlog = (
            active[active["ESTADO"].fillna("").str.upper() != "APROBADO"]
            .groupby("PROCESO")
            .size()
            .reset_index(name="PENDIENTES")
            .sort_values("PENDIENTES", ascending=False)
            .head(15)
        )
        fig = px.bar(
            backlog.sort_values("PENDIENTES"),
            x="PENDIENTES",
            y="PROCESO",
            orientation="h",
            title="Procesos con mayor backlog",
            text="PENDIENTES",
        )
        fig.update_traces(marker_color=ORANGE, textposition="outside")
        base_chart(fig, 470)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        completion = (
            active.assign(
                APROBADO=active["ESTADO"].fillna("").str.upper().eq("APROBADO")
            )
            .groupby("PROCESO")
            .agg(
                DOCUMENTOS=("PROCESO", "size"),
                APROBADOS=("APROBADO", "sum"),
            )
            .reset_index()
        )
        completion["FORMALIZACION_%"] = (
            completion["APROBADOS"] / completion["DOCUMENTOS"] * 100
        ).round(1)
        completion = completion.sort_values("FORMALIZACION_%")

        fig = px.bar(
            completion,
            x="FORMALIZACION_%",
            y="PROCESO",
            orientation="h",
            title="% de documentos aprobados por proceso",
            text="FORMALIZACION_%",
        )
        fig.update_xaxes(range=[0, 100])
        fig.update_traces(marker_color=GREEN, texttemplate="%{text:.1f}%")
        base_chart(fig, 470)
        st.plotly_chart(fig, use_container_width=True)

# ============================================================
# 3. TIEMPOS
# ============================================================
with tab_tiempos:
    st.markdown('<div class="section-title">Tiempo real del flujo de aprobación</div>', unsafe_allow_html=True)

    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Promedio aprobación", f"{approval_avg:.0f} días" if pd.notna(approval_avg) else "N/A")
    a2.metric("Mediana aprobación", f"{approval_median:.0f} días" if pd.notna(approval_median) else "N/A")
    a3.metric("Aprobaciones ≤30 días", f"{pct((approval_days <= 30).sum(), len(approval_days)):.1f}%")
    a4.metric("Aprobaciones >365 días", f"{pct((approval_days > 365).sum(), len(approval_days)):.1f}%")

    st.caption(
        "La mediana se muestra junto al promedio porque el promedio está muy influenciado por documentos antiguos. "
        "La fecha de revisión del Excel no se interpreta como duración del proceso de revisión."
    )

    c1, c2 = st.columns(2)

    with c1:
        time_df = approved[["PROCESO", "DIAS_HASTA_APROBACION"]].dropna()
        if not time_df.empty:
            process_time = (
                time_df.groupby("PROCESO")["DIAS_HASTA_APROBACION"]
                .median()
                .reset_index(name="MEDIANA_DIAS")
                .sort_values("MEDIANA_DIAS")
            )
            fig = px.bar(
                process_time,
                x="MEDIANA_DIAS",
                y="PROCESO",
                orientation="h",
                title="Mediana de días hasta aprobación por proceso",
                text="MEDIANA_DIAS",
            )
            fig.update_traces(marker_color=BLUE, textposition="outside")
            base_chart(fig, max(400, min(850, 24 * len(process_time))))
            st.plotly_chart(fig, use_container_width=True)

    with c2:
        if not approval_days.empty:
            fig = px.histogram(
                approval_days,
                nbins=20,
                title="Distribución de días hasta aprobación",
                labels={"value": "Días"},
            )
            fig.update_traces(marker_color=BLUE)
            base_chart(fig, 420)
            st.plotly_chart(fig, use_container_width=True)

    # Outliers
    outliers = approved[
        approved["DIAS_HASTA_APROBACION"].notna()
        & (approved["DIAS_HASTA_APROBACION"] > 365)
    ][
        [
            "CODIGO DEL DOCUMENTO",
            "NOMBRE DEL DOCUMENTO",
            "PROCESO",
            "FECHA DE CREACIÓN",
            "FECHA DE APROBACIÓN",
            "DIAS_HASTA_APROBACION",
        ]
    ].sort_values("DIAS_HASTA_APROBACION", ascending=False)

    st.markdown('<div class="section-title">Documentos con ciclo de aprobación superior a 365 días</div>', unsafe_allow_html=True)
    st.dataframe(outliers.head(50), use_container_width=True, hide_index=True)

# ============================================================
# 4. VIGENCIA
# ============================================================
with tab_vigencia:
    st.markdown('<div class="section-title">Riesgo por vencimiento documental</div>', unsafe_allow_html=True)

    v1, v2, v3 = st.columns(3)
    v1.metric("Vencidos", format_int(len(expired)))
    v2.metric("Vencen en ≤30 días", format_int((expiring_90["DIAS_PARA_VENCER"] <= 30).sum()))
    v3.metric("Vencen en 31–90 días", format_int(expiring_90["DIAS_PARA_VENCER"].between(31, 90).sum()))

    c1, c2 = st.columns(2)

    with c1:
        vig = (
            approved["ESTADO_VIGENCIA"]
            .value_counts()
            .reindex(
                ["Vencido", "0–30 días", "31–90 días", "91–180 días", "181–365 días", ">365 días"],
                fill_value=0,
            )
            .reset_index()
        )
        vig.columns = ["VIGENCIA", "DOCUMENTOS"]
        fig = px.bar(
            vig,
            x="DOCUMENTOS",
            y="VIGENCIA",
            orientation="h",
            title="Distribución de vigencia",
            text="DOCUMENTOS",
        )
        fig.update_traces(marker_color=BLUE, textposition="outside")
        base_chart(fig, 390)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        risk = (
            expired.groupby("PROCESO")
            .size()
            .reset_index(name="VENCIDOS")
            .sort_values("VENCIDOS", ascending=False)
            .head(15)
        )
        if not risk.empty:
            fig = px.bar(
                risk.sort_values("VENCIDOS"),
                x="VENCIDOS",
                y="PROCESO",
                orientation="h",
                title="Procesos con documentos vencidos",
                text="VENCIDOS",
            )
            fig.update_traces(marker_color=RED, textposition="outside")
            base_chart(fig, 390)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay documentos vencidos en el conjunto filtrado.")

    alertas = approved[
        approved["DIAS_PARA_VENCER"] <= 90
    ][
        [
            "CODIGO DEL DOCUMENTO",
            "NOMBRE DEL DOCUMENTO",
            "PROCESO",
            "FECHA DE VIGENCIA HASTA",
            "DIAS_PARA_VENCER",
            "ESTADO_VIGENCIA",
        ]
    ].sort_values("FECHA DE VIGENCIA HASTA")

    st.markdown('<div class="section-title">Matriz de documentos que requieren atención</div>', unsafe_allow_html=True)
    st.dataframe(alertas, use_container_width=True, hide_index=True)

# ============================================================
# 5. ESTRUCTURA
# ============================================================
with tab_estructura:
    st.markdown('<div class="section-title">Arquitectura documental</div>', unsafe_allow_html=True)

    c1, c2 = st.columns(2)

    with c1:
        matrix = pd.crosstab(
            active["MACROPROCESO"],
            active["TIPO DE DOCUMENTO"]
        )
        fig = px.imshow(
            matrix,
            text_auto=True,
            aspect="auto",
            title="Mapa de calor: macroproceso vs tipo de documento",
        )
        base_chart(fig, 520)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        process_counts = (
            active["PROCESO"]
            .value_counts()
            .sort_values()
            .tail(20)
            .reset_index()
        )
        process_counts.columns = ["PROCESO", "DOCUMENTOS"]
        fig = px.bar(
            process_counts,
            x="DOCUMENTOS",
            y="PROCESO",
            orientation="h",
            title="Procesos con mayor volumen documental",
            text="DOCUMENTOS",
        )
        fig.update_traces(marker_color=BLUE, textposition="outside")
        base_chart(fig, 520)
        st.plotly_chart(fig, use_container_width=True)

    # Tabla resumen
    summary = (
        active.groupby(["UNIDAD DE NEGOCIO", "MACROPROCESO"])
        .agg(
            DOCUMENTOS=("CODIGO DEL DOCUMENTO", "size"),
            APROBADOS=("ESTADO", lambda s: s.fillna("").str.upper().eq("APROBADO").sum()),
            NO_CREADOS=("ESTADO", lambda s: s.fillna("").str.upper().eq("NO CREADO").sum()),
        )
        .reset_index()
    )
    summary["FORMALIZACION_%"] = (
        summary["APROBADOS"] / summary["DOCUMENTOS"] * 100
    ).round(1)

    st.markdown('<div class="section-title">Resumen por unidad y macroproceso</div>', unsafe_allow_html=True)
    st.dataframe(
        summary.sort_values(["UNIDAD DE NEGOCIO", "MACROPROCESO"]),
        use_container_width=True,
        hide_index=True,
    )

# ============================================================
# 6. REPOSITORIO BINAPS
# ============================================================
with tab_repositorio:
    st.markdown('<div class="section-title">Conciliación entre Libro Maestro y Binaps</div>', unsafe_allow_html=True)

    if df_binaps.empty:
        st.warning("No fue posible cargar la hoja Binaps del archivo fuente.")
    else:
        master_codes = set(df["CODIGO DEL DOCUMENTO"].dropna().astype(str).str.strip())
        binaps_codes = set(df_binaps["CODIGO DEL DOCUMENTO"].dropna().astype(str).str.strip())

        visible_codes = set(df_filtered["CODIGO DEL DOCUMENTO"].dropna().astype(str).str.strip())
        visible_binaps = df_binaps[df_binaps["CODIGO DEL DOCUMENTO"].isin(visible_codes)].copy()

        matched = visible_codes & binaps_codes
        missing_in_binaps = visible_codes - binaps_codes

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Registros en Binaps", format_int(len(df_binaps)))
        r2.metric("Coinciden con el maestro", format_int(len(matched)))
        r3.metric("Maestro sin registro Binaps", format_int(len(missing_in_binaps)))
        r4.metric("Cobertura Binaps", f"{pct(len(matched), len(visible_codes)):.1f}%")

        st.caption(
            "En el archivo actual, los registros de Binaps corresponden principalmente a documentos que ya tienen gestión documental "
            "registrada. Los documentos 'No creado' quedan fuera del repositorio."
        )

        # Conciliación
        recon = df_filtered[
            [
                "CODIGO DEL DOCUMENTO",
                "NOMBRE DEL DOCUMENTO",
                "ESTADO",
                "VERSIÓN",
                "PROCESO",
                "FECHA DE APROBACIÓN",
            ]
        ].merge(
            df_binaps[
                [
                    "CODIGO DEL DOCUMENTO",
                    "NOMBRE_BINAPS",
                    "ESTADO_BINAPS",
                    "VERSION_BINAPS",
                    "FECHA_APROBACION_BINAPS",
                    "FECHA_VIGENCIA_BINAPS",
                ]
            ],
            on="CODIGO DEL DOCUMENTO",
            how="left",
        )

        recon["EN_BINAPS"] = recon["ESTADO_BINAPS"].notna()
        recon["DIFERENCIA_ESTADO"] = (
            recon["ESTADO"].fillna("").str.upper()
            != recon["ESTADO_BINAPS"].fillna("").str.upper()
        ) & recon["EN_BINAPS"]

        st.markdown('<div class="section-title">Estado de conciliación</div>', unsafe_allow_html=True)
        status_con = recon["EN_BINAPS"].map({True: "En Binaps", False: "No está en Binaps"}).value_counts().reset_index()
        status_con.columns = ["SITUACION", "DOCUMENTOS"]
        fig = px.bar(
            status_con,
            x="DOCUMENTOS",
            y="SITUACION",
            orientation="h",
            title="Presencia de documentos en el repositorio",
            text="DOCUMENTOS",
        )
        fig.update_traces(marker_color=BLUE, textposition="outside")
        base_chart(fig, 290)
        st.plotly_chart(fig, use_container_width=True)

        missing_df = recon[
            ~recon["EN_BINAPS"]
        ][
            [
                "CODIGO DEL DOCUMENTO",
                "NOMBRE DEL DOCUMENTO",
                "PROCESO",
                "ESTADO",
            ]
        ].sort_values(["PROCESO", "CODIGO DEL DOCUMENTO"])

        st.markdown('<div class="section-title">Documentos del maestro que no están en Binaps</div>', unsafe_allow_html=True)
        st.dataframe(missing_df, use_container_width=True, hide_index=True)

        name_diff = recon[
            recon["EN_BINAPS"]
            & (
                recon["NOMBRE DEL DOCUMENTO"].fillna("").str.strip().str.lower()
                != recon["NOMBRE_BINAPS"].fillna("").str.strip().str.lower()
            )
        ][
            [
                "CODIGO DEL DOCUMENTO",
                "NOMBRE DEL DOCUMENTO",
                "NOMBRE_BINAPS",
                "ESTADO",
                "ESTADO_BINAPS",
            ]
        ]

        st.markdown('<div class="section-title">Nombres que no coinciden entre Maestro y Binaps</div>', unsafe_allow_html=True)
        st.dataframe(name_diff, use_container_width=True, hide_index=True)

# ============================================================
# 7. LIBRO MAESTRO
# ============================================================
with tab_libro:
    st.markdown('<div class="section-title">Consulta operacional del Libro Maestro</div>', unsafe_allow_html=True)

    display_columns = [
        "UNIDAD DE NEGOCIO",
        "MACROPROCESO",
        "PROCESO",
        "TIPO DE DOCUMENTO",
        "CODIGO DEL DOCUMENTO",
        "NOMBRE DEL DOCUMENTO",
        "VERSIÓN",
        "ESTADO",
        "FECHA DE CREACIÓN",
        "FECHA DE VIGENCIA HASTA",
        "FECHA DE REVISIÓN",
        "FECHA DE APROBACIÓN",
        "MANEJO DEL DOCUMENTO",
        "OBSERVACIONES",
    ]

    table_df = df_filtered[display_columns].copy()

    st.dataframe(
        table_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "FECHA DE CREACIÓN": st.column_config.DateColumn("Creación", format="DD/MM/YYYY"),
            "FECHA DE VIGENCIA HASTA": st.column_config.DateColumn("Vigencia", format="DD/MM/YYYY"),
            "FECHA DE REVISIÓN": st.column_config.DateColumn("Revisión", format="DD/MM/YYYY"),
            "FECHA DE APROBACIÓN": st.column_config.DateColumn("Aprobación", format="DD/MM/YYYY"),
        },
    )

    csv = table_df.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Descargar vista filtrada en CSV",
        data=csv,
        file_name="libro_maestro_filtrado.csv",
        mime="text/csv",
    )

# ============================================================
# 8. CALIDAD DEL DATO
# ============================================================
with tab_calidad:
    st.markdown('<div class="section-title">Controles de calidad sobre el maestro</div>', unsafe_allow_html=True)

    missing = (
        active[REQUIRED_COLUMNS]
        .isna()
        .sum()
        .reset_index()
    )
    missing.columns = ["CAMPO", "VACIOS"]
    missing["% VACIOS"] = (
        missing["VACIOS"] / len(active) * 100
    ).round(1)
    missing = missing.sort_values("VACIOS", ascending=False)

    q1, q2, q3, q4 = st.columns(4)
    q1.metric("Campos críticos faltantes", format_int(int(df_filtered["DATO_CRITICO_FALTANTE"].sum())))
    q2.metric("Códigos repetidos", format_int(int(df_filtered["CODIGO_REPETIDO"].sum())))
    q3.metric("Fechas de aprobación", format_int(int(df_filtered["FECHA DE APROBACIÓN"].notna().sum())))
    q4.metric("Fechas de revisión", format_int(int(df_filtered["FECHA DE REVISIÓN"].notna().sum())))

    c1, c2 = st.columns(2)

    with c1:
        fig = px.bar(
            missing.head(12).sort_values("VACIOS"),
            x="VACIOS",
            y="CAMPO",
            orientation="h",
            title="Campos con mayor cantidad de vacíos",
            text="VACIOS",
        )
        fig.update_traces(marker_color=ORANGE, textposition="outside")
        base_chart(fig, 430)
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        duplicate_codes = (
            df_filtered[df_filtered["CODIGO_REPETIDO"]]
            .groupby("CODIGO DEL DOCUMENTO")
            .size()
            .reset_index(name="REGISTROS")
            .sort_values("REGISTROS", ascending=False)
        )

        if not duplicate_codes.empty:
            fig = px.bar(
                duplicate_codes.head(15).sort_values("REGISTROS"),
                x="REGISTROS",
                y="CODIGO DEL DOCUMENTO",
                orientation="h",
                title="Códigos repetidos que requieren revisión",
                text="REGISTROS",
            )
            fig.update_traces(marker_color=RED, textposition="outside")
            base_chart(fig, 430)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("No hay códigos repetidos en la selección actual.")

    st.markdown('<div class="section-title">Registros que requieren revisión de calidad</div>', unsafe_allow_html=True)
    quality_df = df_filtered[
        df_filtered["DATO_CRITICO_FALTANTE"] | df_filtered["CODIGO_REPETIDO"]
    ][
        [
            "CODIGO DEL DOCUMENTO",
            "NOMBRE DEL DOCUMENTO",
            "PROCESO",
            "ESTADO",
            "CODIGO_REPETIDO",
            "DATO_CRITICO_FALTANTE",
            "OBSERVACIONES",
        ]
    ]

    st.dataframe(quality_df, use_container_width=True, hide_index=True)

# ============================================================
# PIE
# ============================================================
st.divider()
st.caption(
    "Dashboard de control documental. Los cálculos se realizan sobre el Excel fuente y se recalculan al actualizar la página."
)
