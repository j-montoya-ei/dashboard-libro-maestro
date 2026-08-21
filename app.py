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

# 2. ESTILOS CSS CORPORATIVOS MODERNOS
st.markdown("""
    <style>
    .main { background-color: #F4F6F9; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border-left: 5px solid #00369C;
        border-radius: 10px;
        padding: 18px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.05);
    }
    div[data-testid="stMetric"] label {
        color: #6C757D !important;
        font-weight: 600 !important;
        font-size: 0.85rem !important;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #00369C !important;
        font-weight: 800 !important;
        font-size: 1.8rem !important;
    }
    button[data-baseweb="tab"] {
        font-size: 1.05rem !important;
        font-weight: 600 !important;
        color: #6C757D !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #00369C !important;
        border-bottom-color: #F6D000 !important;
        border-bottom-width: 4px !important;
    }
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
    
    for col in ['PROCESO', 'MACROPROCESO', 'UNIDAD DE NEGOCIO', 'ESTADO', 'TIPO DE DOCUMENTO']:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            
    date_columns = [c for c in df.columns if 'FECHA' in c.upper()]
    for col in date_columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')
        
    df['DIAS_HASTA_APROBACION'] = (df['FECHA DE APROBACIÓN'] - df['FECHA DE CREACIÓN']).dt.days if 'FECHA DE CREACIÓN' in df.columns and 'FECHA DE APROBACIÓN' in df.columns else None
    df['DIAS_HASTA_REVISION'] = (df['FECHA DE REVISIÓN'] - df['FECHA DE CREACIÓN']).dt.days if 'FECHA DE CREACIÓN' in df.columns and 'FECHA DE REVISIÓN' in df.columns else None
    
    return df

try:
    df_raw = load_data()

    # 4. ENCABEZADO Y LOGO
    col_logo, col_titulo = st.columns([1, 6])
    with col_logo:
        if os.path.exists("Logos Electroingeniería-01.jpg"):
            st.image("Logos Electroingeniería-01.jpg", width=140)
    with col_titulo:
        st.title("Libro Maestro de Documentos")
        st.markdown("**Sistema de Gestión de Calidad** | Control Inteligente de Tiempos, Tipos y Procesos")

    st.markdown("---")

    # 5. FILTROS EN BARRA LATERAL (EN CASCADA)
    st.sidebar.header("🔍 Filtros de Control")

    col_unidad = 'UNIDAD DE NEGOCIO'
    unidades = sorted(df_raw[col_unidad].dropna().unique()) if col_unidad in df_raw.columns else []
    selected_unidad = st.sidebar.multiselect("Unidad de Negocio", unidades)

    df_f = df_raw.copy()
    if selected_unidad:
        df_f = df_f[df_f[col_unidad].isin(selected_unidad)]

    col_macro = 'MACROPROCESO'
    macroprocesos = sorted(df_f[col_macro].dropna().unique()) if col_macro in df_f.columns else []
    selected_macro = st.sidebar.multiselect("Macroproceso", macroprocesos)

    if selected_macro:
        df_f = df_f[df_f[col_macro].isin(selected_macro)]

    col_proceso = 'PROCESO'
    procesos = sorted(df_f[col_proceso].dropna().unique()) if col_proceso in df_f.columns else []
    selected_proceso = st.sidebar.multiselect("Proceso", procesos)

    if selected_proceso:
        df_f = df_f[df_f[col_proceso].isin(selected_proceso)]

    col_estado = 'ESTADO'
    estados = sorted(df_raw[col_estado].dropna().unique()) if col_estado in df_raw.columns else []
    selected_estado = st.sidebar.multiselect("Estado del Documento", estados)

    col_tipo = 'TIPO DE DOCUMENTO'
    tipos = sorted(df_raw[col_tipo].dropna().unique()) if col_tipo in df_raw.columns else []
    selected_tipo = st.sidebar.multiselect("Tipo de Documento", tipos)

    # Aplicación definitiva de filtros
    df_filtered = df_raw.copy()
    if selected_unidad: df_filtered = df_filtered[df_filtered[col_unidad].isin(selected_unidad)]
    if selected_macro: df_filtered = df_filtered[df_filtered[col_macro].isin(selected_macro)]
    if selected_proceso: df_filtered = df_filtered[df_filtered[col_proceso].isin(selected_proceso)]
    if selected_estado: df_filtered = df_filtered[df_filtered[col_estado].isin(selected_estado)]
    if selected_tipo: df_filtered = df_filtered[df_filtered[col_tipo].isin(selected_tipo)]

    # 6. KPIS SUPERIORES
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    
    total_docs = len(df_filtered)
    aprobados = len(df_filtered[df_filtered[col_estado].str.upper() == 'APROBADO']) if col_estado in df_filtered.columns else 0
    pendientes = len(df_filtered[df_filtered[col_estado].str.upper() == 'NO CREADO']) if col_estado in df_filtered.columns else 0
    avg_aprob = df_filtered['DIAS_HASTA_APROBACION'].mean()
    avg_rev = df_filtered['DIAS_HASTA_REVISION'].mean()

    kpi1.metric("Total Documentos", total_docs)
    kpi2.metric("Aprobados", aprobados)
    kpi3.metric("Pendientes", pendientes)
    kpi4.metric("Prom. Aprobación", f"{int(avg_aprob)} días" if pd.notna(avg_aprob) else "N/A")
    kpi5.metric("Prom. Revisión", f"{int(avg_rev)} días" if pd.notna(avg_rev) else "N/A")

    st.markdown("<br>", unsafe_allow_html=True)

    # 7. ORGANIZACIÓN POR PESTAÑAS
    tab_volumen, tab_tiempos, tab_tabla = st.tabs([
        "📁 Volumen y Distribución", 
        "⏱️ Análisis de Tiempos y Cuellos de Botella", 
        "📋 Base de Datos Completa"
    ])

    with tab_volumen:
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.markdown("##### 📌 Documentos por Tipo")
            if col_tipo in df_filtered.columns and not df_filtered.empty:
                df_tipo = df_filtered[col_tipo].value_counts().reset_index()
                df_tipo.columns = ['TIPO', 'CANTIDAD']
                df_tipo = df_tipo.sort_values(by='CANTIDAD', ascending=True)
                
                fig_tipo = px.bar(
                    df_tipo, x='CANTIDAD', y='TIPO', orientation='h', text='CANTIDAD',
                    color_discrete_sequence=['#00369C']
                )
                fig_tipo.update_layout(xaxis_title="", yaxis_title="", margin=dict(t=10, b=10))
                st.plotly_chart(fig_tipo, use_container_width=True)

        with col_g2:
            st.markdown("##### 🍩 Distribución por Estado")
            if col_estado in df_filtered.columns and not df_filtered.empty:
                df_est = df_filtered[col_estado].value_counts().reset_index()
                df_est.columns = ['ESTADO', 'CANTIDAD']
                
                fig_est = px.pie(
                    df_est, values='CANTIDAD', names='ESTADO', hole=0.5,
                    color_discrete_sequence=['#00369C', '#F6D000', '#A4A8AB', '#555555']
                )
                fig_est.update_traces(textposition='inside', textinfo='percent+label')
                fig_est.update_layout(margin=dict(t=10, b=10))
                st.plotly_chart(fig_est, use_container_width=True)

    with tab_tiempos:
        st.markdown("##### 🐢 Top Procesos con Mayor Retraso en Aprobación")
        if 'PROCESO' in df_filtered.columns and 'DIAS_HASTA_APROBACION' in df_filtered.columns and not df_filtered.empty:
            df_proc = df_filtered.groupby('PROCESO')['DIAS_HASTA_APROBACION'].mean().reset_index()
            df_proc = df_proc.dropna().sort_values(by='DIAS_HASTA_APROBACION', ascending=False).head(8)
            
            fig_proc = px.bar(
                df_proc, x='DIAS_HASTA_APROBACION', y='PROCESO', orientation='h', text_auto='.0f',
                color_discrete_sequence=['#F6D000']
            )
            fig_proc.update_layout(xaxis_title="Días Promedio", yaxis_title="", margin=dict(t=10, b=10))
            st.plotly_chart(fig_proc, use_container_width=True)

    with tab_tabla:
        st.markdown("##### 📋 Registros Filtrados del Libro Maestro")
        st.dataframe(df_filtered, use_container_width=True)

except Exception as e:
    st.error(f"Error crítico en el dashboard: {e}")
