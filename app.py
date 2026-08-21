import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 1. CONFIGURACIÓN DE PÁGINA
st.set_page_config(
    page_title="Libro Maestro | Electroingeniería",
    page_icon="📊",
    layout="wide"
)

# 2. ESTILOS CSS PERSONALIZADOS (PALETA CORPORATIVA)
st.markdown("""
    <style>
    .main { background-color: #F8F9FA; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    div[data-testid="stMetric"] { background-color: #FFFFFF; border-left: 5px solid #00369C; border-radius: 8px; padding: 15px; box-shadow: 0 2px 8px rgba(0,0,0,0.05); }
    div[data-testid="stMetric"] label { color: #555555 !important; font-weight: 600 !important; font-size: 0.9rem !important; }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] { color: #00369C !important; font-weight: 700 !important; }
    button[data-baseweb="tab"] { font-size: 1rem !important; font-weight: 600 !important; color: #A4A8AB !important; }
    button[data-baseweb="tab"][aria-selected="true"] { color: #00369C !important; border-bottom-color: #F6D000 !important; border-bottom-width: 3px !important; }
    h1, h2, h3 { color: #00369C; font-weight: 700; }
    </style>
""", unsafe_allow_html=True)

# 3. CARGA Y PROCESAMIENTO
@st.cache_data
def load_data():
    file_path = "M-FT-1 LIBRO MAESTRO DE DOCUMENTOS v2.xlsx"
    df = pd.read_excel(file_path, sheet_name="LIBRO MAESTRO", header=2)
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df = df.dropna(how="all")
    if 'PROCESO' in df.columns: df['PROCESO'] = df['PROCESO'].astype(str).str.strip()
    date_columns = [c for c in df.columns if 'FECHA' in c.upper()]
    for col in date_columns: df[col] = pd.to_datetime(df[col], errors='coerce')
    df['DIAS_HASTA_APROBACION'] = (df['FECHA DE APROBACIÓN'] - df['FECHA DE CREACIÓN']).dt.days if 'FECHA DE CREACIÓN' in df.columns and 'FECHA DE APROBACIÓN' in df.columns else None
    df['DIAS_HASTA_REVISION'] = (df['FECHA DE REVISIÓN'] - df['FECHA DE CREACIÓN']).dt.days if 'FECHA DE CREACIÓN' in df.columns and 'FECHA DE REVISIÓN' in df.columns else None
    return df

df_raw = load_data()

# 4. ENCABEZADO
col_logo, col_titulo = st.columns([1, 5])
with col_logo:
    if os.path.exists("Logos Electroingeniería-01.jpg"): st.image("Logos Electroingeniería-01.jpg", width=150)
with col_titulo:
    st.title("Libro Maestro de Documentos")
    st.caption("Sistema de Gestión de Calidad | Control de Documentación, Tipos y Tiempos")

st.markdown("---")

# 5. FILTROS (CON EL NUEVO FILTRO DE PROCESO)
st.sidebar.header("🔍 Filtros de Control")

# Filtro Unidad
col_unidad = 'UNIDAD DE NEGOCIO'
unidades = sorted([str(x).strip() for x in df_raw[col_unidad].dropna().unique()]) if col_unidad in df_raw.columns else []
selected_unidad = st.sidebar.multiselect("Unidad de Negocio", unidades)

# Filtro Macroproceso
col_macro = 'MACROPROCESO'
df_f_macro = df_raw[df_raw[col_unidad].isin(selected_unidad)] if selected_unidad else df_raw
macroprocesos = sorted([str(x).strip() for x in df_f_macro[col_macro].dropna().unique()]) if col_macro in df_raw.columns else []
selected_macro = st.sidebar.multiselect("Macroproceso", macroprocesos)

# NUEVO FILTRO PROCESO
col_proceso = 'PROCESO'
df_f_proceso = df_f_macro[df_f_macro[col_macro].isin(selected_macro)] if selected_macro else df_f_macro
procesos = sorted([str(x).strip() for x in df_f_proceso[col_proceso].dropna().unique()]) if col_proceso in df_raw.columns else []
selected_proceso = st.sidebar.multiselect("Proceso", procesos)

# Otros Filtros
col_estado = 'ESTADO'
selected_estado = st.sidebar.multiselect("Estado del Documento", sorted(df_raw[col_estado].dropna().unique()) if col_estado in df_raw.columns else [])
col_tipo = 'TIPO DE DOCUMENTO'
selected_tipo = st.sidebar.multiselect("Tipo de Documento", sorted(df_raw[col_tipo].dropna().unique()) if col_tipo in df_raw.columns else [])

# Aplicación de filtros
df_filtered = df_raw.copy()
if selected_unidad: df_filtered = df_filtered[df_filtered[col_unidad].isin(selected_unidad)]
if selected_macro: df_filtered = df_filtered[df_filtered[col_macro].isin(selected_macro)]
if selected_proceso: df_filtered = df_filtered[df_filtered[col_proceso].isin(selected_proceso)]
if selected_estado: df_filtered = df_filtered[df_filtered[col_estado].isin(selected_estado)]
if selected_tipo: df_filtered = df_filtered[df_filtered[col_tipo].isin(selected_tipo)]

# 6. KPIs
kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
kpi1.metric("Total Documentos", len(df_filtered))
aprob = len(df_filtered[df_filtered[col_estado].astype(str).str.strip().str.upper() == 'APROBADO']) if col_estado in df_filtered.columns else 0
pend = len(df_filtered[df_filtered[col_estado].astype(str).str.strip().str.upper() == 'NO CREADO']) if col_estado in df_filtered.columns else 0
kpi2.metric("Aprobados", aprob)
kpi3.metric("No Creados / Pendientes", pend)
avg_ap = df_filtered['DIAS_HASTA_APROBACION'].mean()
avg_re = df_filtered['DIAS_HASTA_REVISION'].mean()
kpi4.metric("Prom. Aprobación", f"{int(avg_ap)} días" if pd.notna(avg_ap) else "N/A")
kpi5.metric("Prom. Revisión", f"{int(avg_re)} días" if pd.notna(avg_re) else "N/A")

# 7. PESTAÑAS (Resto del código original intacto)
tab_volumen, tab_tiempos, tab_tabla = st.tabs(["📁 Volumen", "⏱️ Tiempos", "📋 Libro Maestro"])
with tab_volumen:
    if col_tipo in df_filtered.columns:
        st.plotly_chart(px.bar(df_filtered[col_tipo].value_counts().reset_index(), x='count', y=col_tipo, orientation='h', color_discrete_sequence=['#00369C']), use_container_width=True)
with tab_tiempos:
    if 'PROCESO' in df_filtered.columns:
        df_p = df_filtered.groupby('PROCESO')['DIAS_HASTA_APROBACION'].mean().reset_index()
        st.plotly_chart(px.bar(df_p, x='PROCESO', y='DIAS_HASTA_APROBACION', color_discrete_sequence=['#F6D000']), use_container_width=True)
with tab_tabla:
    st.dataframe(df_filtered, use_container_width=True)
