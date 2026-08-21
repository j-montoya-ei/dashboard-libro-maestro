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

# 2. ESTILOS CSS PERSONALIZADOS
st.markdown("""
    <style>
    .main { background-color: #F8F9FA; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border-left: 5px solid #00369C;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    div[data-testid="stMetric"] label { color: #555555 !important; font-weight: 600 !important; font-size: 0.9rem !important; }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] { color: #00369C !important; font-weight: 700 !important; }
    button[data-baseweb="tab"] { font-size: 1rem !important; font-weight: 600 !important; color: #A4A8AB !important; }
    button[data-baseweb="tab"][aria-selected="true"] { color: #00369C !important; border-bottom-color: #F6D000 !important; border-bottom-width: 3px !important; }
    h1, h2, h3 { color: #00369C; font-weight: 700; }
    </style>
""", unsafe_allow_html=True)

# 3. CARGA Y PROCESAMIENTO DE DATOS
@st.cache_data
def load_data():
    file_path = "M-FT-1 LIBRO MAESTRO DE DOCUMENTOS v2.xlsx"
    df = pd.read_excel(file_path, sheet_name="LIBRO MAESTRO", header=2)
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df = df.dropna(how="all")
    
    if 'PROCESO' in df.columns:
        df['PROCESO'] = df['PROCESO'].astype(str).str.strip()
    
    date_columns = [c for c in df.columns if 'FECHA' in c.upper()]
    for col in date_columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')
        
    if 'FECHA DE CREACIÓN' in df.columns and 'FECHA DE APROBACIÓN' in df.columns:
        df['DIAS_HASTA_APROBACION'] = (df['FECHA DE APROBACIÓN'] - df['FECHA DE CREACIÓN']).dt.days
    else:
        df['DIAS_HASTA_APROBACION'] = None

    if 'FECHA DE CREACIÓN' in df.columns and 'FECHA DE REVISIÓN' in df.columns:
        df['DIAS_HASTA_REVISION'] = (df['FECHA DE REVISIÓN'] - df['FECHA DE CREACIÓN']).dt.days
    else:
        df['DIAS_HASTA_REVISION'] = None
    
    return df

try:
    df_raw = load_data()

    # 4. ENCABEZADO
    col_logo, col_titulo = st.columns([1, 5])
    with col_logo:
        st.image("Logos Electroingeniería-01.jpg", width=150) if os.path.exists("Logos Electroingeniería-01.jpg") else st.warning("Logo no encontrado")
    with col_titulo:
        st.title("Libro Maestro de Documentos")
        st.caption("Sistema de Gestión de Calidad | Control de Documentación")

    st.markdown("---")

    # 5. FILTROS EN BARRA LATERAL (LÓGICA SECUENCIAL)
    st.sidebar.header("🔍 Filtros de Control")
    
    # Unidad de Negocio
    selected_unidad = st.sidebar.multiselect("Unidad de Negocio", sorted(df_raw['UNIDAD DE NEGOCIO'].dropna().unique()))
    
    # Macroproceso (Filtrado por Unidad)
    df_temp = df_raw.copy()
    if selected_unidad:
        df_temp = df_temp[df_temp['UNIDAD DE NEGOCIO'].isin(selected_unidad)]
    
    selected_macro = st.sidebar.multiselect("Macroproceso", sorted(df_temp['MACROPROCESO'].dropna().unique()))
    
    # Proceso (Filtrado por Macroproceso)
    if selected_macro:
        df_temp = df_temp[df_temp['MACROPROCESO'].isin(selected_macro)]
        
    selected_proceso = st.sidebar.multiselect("Proceso", sorted(df_temp['PROCESO'].dropna().unique()))

    # Filtros finales
    selected_estado = st.sidebar.multiselect("Estado del Documento", sorted(df_raw['ESTADO'].dropna().unique()))
    selected_tipo = st.sidebar.multiselect("Tipo de Documento", sorted(df_raw['TIPO DE DOCUMENTO'].dropna().unique()))

    # Aplicar todos los filtros al DataFrame final
    df_filtered = df_raw.copy()
    if selected_unidad: df_filtered = df_filtered[df_filtered['UNIDAD DE NEGOCIO'].isin(selected_unidad)]
    if selected_macro: df_filtered = df_filtered[df_filtered['MACROPROCESO'].isin(selected_macro)]
    if selected_proceso: df_filtered = df_filtered[df_filtered['PROCESO'].isin(selected_proceso)]
    if selected_estado: df_filtered = df_filtered[df_filtered['ESTADO'].isin(selected_estado)]
    if selected_tipo: df_filtered = df_filtered[df_filtered['TIPO DE DOCUMENTO'].isin(selected_tipo)]

    # 6. KPIS
    avg_dias_aprob = df_filtered['DIAS_HASTA_APROBACION'].mean()
    avg_dias_rev = df_filtered['DIAS_HASTA_REVISION'].mean()
    
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    kpi1.metric("Total Documentos", len(df_filtered))
    kpi2.metric("Aprobados", len(df_filtered[df_filtered['ESTADO'] == 'APROBADO']))
    kpi3.metric("Pendientes", len(df_filtered[df_filtered['ESTADO'] == 'NO CREADO']))
    kpi4.metric("Prom. Aprobación", f"{int(avg_dias_aprob)} días" if pd.notna(avg_dias_aprob) else "N/A")
    kpi5.metric("Prom. Revisión", f"{int(avg_dias_rev)} días" if pd.notna(avg_dias_rev) else "N/A")

    # 7. VISUALIZACIÓN
    tab1, tab2, tab3 = st.tabs(["📁 Volumen", "⏱️ Tiempos", "📋 Detalle"])
    
    with tab1:
        st.subheader("Documentos por Tipo")
        if not df_filtered.empty:
            st.plotly_chart(px.bar(df_filtered['TIPO DE DOCUMENTO'].value_counts().reset_index(), x='count', y='TIPO DE DOCUMENTO', orientation='h', color_discrete_sequence=['#00369C']), use_container_width=True)

    with tab2:
        st.subheader("Cuellos de Botella (Procesos)")
        if 'PROCESO' in df_filtered.columns:
            df_p = df_filtered.groupby('PROCESO')['DIAS_HASTA_APROBACION'].mean().reset_index()
            st.plotly_chart(px.bar(df_p, x='PROCESO', y='DIAS_HASTA_APROBACION', color_discrete_sequence=['#F6D000']), use_container_width=True)

    with tab3:
        st.dataframe(df_filtered, use_container_width=True)

except Exception as e:
    st.error(f"Error: {e}")
