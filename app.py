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

# 2. ESTILOS CSS PERSONALIZADOS (UI/UX - PALETA CORPORATIVA)
# Azul Principal: #00369C | Amarillo Accento: #F6D000 | Gris Neutro: #A4A8AB
st.markdown("""
    <style>
    /* Estructura general y fuentes */
    .main {
        background-color: #F8F9FA;
    }
    
    /* Ocultar barra de menú por defecto de Streamlit para acabado limpio */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Tarjetas de KPI Personalizadas */
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border-left: 5px solid #00369C;
        border-radius: 8px;
        padding: 15px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.05);
    }
    
    div[data-testid="stMetric"] label {
        color: #555555 !important;
        font-weight: 600 !important;
        font-size: 0.9rem !important;
    }
    
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #00369C !important;
        font-weight: 700 !important;
    }

    /* Estilos de Pestañas (Tabs) */
    button[data-baseweb="tab"] {
        font-size: 1rem !important;
        font-weight: 600 !important;
        color: #A4A8AB !important;
    }
    
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #00369C !important;
        border-bottom-color: #F6D000 !important;
        border-bottom-width: 3px !important;
    }
    
    /* Encabezados */
    h1, h2, h3 {
        color: #00369C;
        font-weight: 700;
    }
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

    # 4. ENCABEZADO CON LOGO E IDENTIDAD DE MARCA
    col_logo, col_titulo = st.columns([1, 5])
    
    with col_logo:
        if os.path.exists("logo.png"):
            st.image("logo.png", width=140)
        else:
            st.markdown("### 🏢") # Placeholder por si el archivo aún no existe
            
    with col_titulo:
        st.title("Libro Maestro de Documentos")
        st.caption("Sistema de Gestión de Calidad | Panel Control de Tiempos y Volumen")

    st.markdown("---")

    # 5. BARRA LATERAL (FILTROS)
    st.sidebar.header("🔍 Filtros de Control")
    
    col_unidad = 'UNIDAD DE NEGOCIO'
    unidades = sorted([str(x).strip() for x in df_raw[col_unidad].dropna().unique() if str(x).strip() != '']) if col_unidad in df_raw.columns else []
    selected_unidad = st.sidebar.multiselect("Unidad de Negocio", unidades)
    
    col_macro = 'MACROPROCESO'
    macroprocesos = sorted([str(x).strip() for x in df_raw[col_macro].dropna().unique() if str(x).strip() != '']) if col_macro in df_raw.columns else []
    selected_macro = st.sidebar.multiselect("Macroproceso", macroprocesos)
    
    col_estado = 'ESTADO'
    estados = sorted([str(x).strip() for x in df_raw[col_estado].dropna().unique() if str(x).strip() != '']) if col_estado in df_raw.columns else []
    selected_estado = st.sidebar.multiselect("Estado del Documento", estados)

    col_tipo = 'TIPO DE DOCUMENTO'
    tipos = sorted([str(x).strip() for x in df_raw[col_tipo].dropna().unique() if str(x).strip() != '']) if col_tipo in df_raw.columns else []
    selected_tipo = st.sidebar.multiselect("Tipo de Documento", tipos)

    # Filtrado dinámico
    df_filtered = df_raw.copy()
    if selected_unidad and col_unidad in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_unidad].astype(str).str.strip().isin(selected_unidad)]
    if selected_macro and col_macro in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_macro].astype(str).str.strip().isin(selected_macro)]
    if selected_estado and col_estado in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_estado].astype(str).str.strip().isin(selected_estado)]
    if selected_tipo and col_tipo in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_tipo].astype(str).str.strip().isin(selected_tipo)]

    # 6. KPIS SUPERIORES
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    
    kpi1.metric("Total Documentos", len(df_filtered), help="Total de registros en el filtro actual")
    
    if col_estado in df_filtered.columns:
        aprobados = len(df_filtered[df_filtered[col_estado].astype(str).str.strip().str.upper() == 'APROBADO'])
        no_creados = len(df_filtered[df_filtered[col_estado].astype(str).str.strip().str.upper() == 'NO CREADO'])
    else:
        aprobados, no_creados = 0, 0
        
    kpi2.metric("Aprobados", aprobados)
    kpi3.metric("Pendientes / No Creados", no_creados)
    
    avg_dias_aprob = df_filtered['DIAS_HASTA_APROBACION'].mean()
    avg_dias_rev = df_filtered['DIAS_HASTA_REVISION'].mean()
    
    val_aprob_str = f"{int(avg_dias_aprob)} días" if pd.notna(avg_dias_aprob) else "N/A"
    val_rev_str = f"{int(avg_dias_rev)} días" if pd.notna(avg_dias_rev) else "N/A"
    
    kpi4.metric("Prom. Aprobación", val_aprob_str, help="Días promedio desde la Creación hasta la Aprobación")
    kpi5.metric("Prom. Revisión", val_rev_str, help="Días promedio desde la Creación hasta la Revisión")

    st.markdown("<br>", unsafe_allow_html=True)

    # 7. NAVEGACIÓN POR PESTAÑAS (UX/UI MEJORADA)
    tab_resumen, tab_tiempos, tab_tabla = st.tabs([
        "📊 Resumen Ejecutivo", 
        "⏱️ Análisis de Tiempos y Cuelgues", 
        "📋 Libro Maestro Completo"
    ])

    # --- PESTAÑA 1: RESUMEN EJECUTIVO ---
    with tab_resumen:
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.markdown("##### Distribución por Estado")
            if col_estado in df_filtered.columns:
                df_estado = df_filtered[col_estado].dropna().astype(str).str.strip().value_counts().reset_index()
                df_estado.columns = ['ESTADO', 'CANTIDAD']
                fig_estado = px.pie(
                    df_estado, values='CANTIDAD', names='ESTADO', hole=0.45,
                    color_discrete_sequence=['#00369C', '#F6D000', '#A4A8AB', '#D9D9D9']
                )
                fig_estado.update_traces(textposition='inside', textinfo='percent+label')
                fig_estado.update_layout(margin=dict(t=20, b=20, l=20, r=20))
                st.plotly_chart(fig_estado, use_container_width=True)

        with col_g2:
            st.markdown("##### Volumen por Tipo de Documento")
            if col_tipo in df_filtered.columns:
                df_tipo = df_filtered[col_tipo].dropna().astype(str).str.strip().value_counts().reset_index()
                df_tipo.columns = ['TIPO', 'CANTIDAD']
                df_tipo = df_tipo.sort_values(by='CANTIDAD', ascending=True)
                
                fig_tipo = px.bar(
                    df_tipo, y='TIPO', x='CANTIDAD', orientation='h',
                    text_auto=True, color_discrete_sequence=['#00369C']
                )
                fig_tipo.update_layout(xaxis_title="", yaxis_title="", margin=dict(t=20, b=20, l=20, r=20))
                st.plotly_chart(fig_tipo, use_container_width=True)

        st.markdown("##### Documentos por Unidad de Negocio")
        if col_unidad in df_filtered.columns:
            df_unidad = df_filtered[col_unidad].dropna().astype(str).str.strip().value_counts().reset_index()
            df_unidad.columns = ['UNIDAD', 'CANTIDAD']
            fig_unidad = px.bar(
                df_unidad, x='UNIDAD', y='CANTIDAD', text_auto=True,
                color_discrete_sequence=['#00369C']
            )
            fig_unidad.update_layout(xaxis_title="", yaxis_title="")
            st.plotly_chart(fig_unidad, use_container_width=True)

    # --- PESTAÑA 2: TIEMPOS Y CUELGUES ---
    with tab_tiempos:
        st.markdown("##### Tiempos Promedio de Ciclo (Días)")
        tiempos_data = pd.DataFrame({
            'Etapa': ['Aprobación (Creación ➔ Aprobado)', 'Revisión (Creación ➔ Revisado)'],
            'Días Promedio': [
                avg_dias_aprob if pd.notna(avg_dias_aprob) else 0, 
                avg_dias_rev if pd.notna(avg_dias_rev) else 0
            ]
        })
        fig_tiempos = px.bar(
            tiempos_data, x='Etapa', y='Días Promedio', text_auto='.0f',
            color='Etapa', color_discrete_sequence=['#00369C', '#F6D000']
        )
        fig_tiempos.update_layout(showlegend=False, yaxis_title="Días")
        st.plotly_chart(fig_tiempos, use_container_width=True)

        st.markdown("---")
        st.markdown("##### Top 10 Procesos con Mayor Tiempo de Respuesta")
        col_t1, col_t2 = st.columns(2)

        with col_t1:
            st.markdown("**🐢 Más lentos para APROBAR**")
            if 'PROCESO' in df_filtered.columns and 'DIAS_HASTA_APROBACION' in df_filtered.columns:
                df_p_aprob = df_filtered.groupby('PROCESO')['DIAS_HASTA_APROBACION'].mean().dropna().reset_index()
                df_p_aprob.columns = ['PROCESO', 'DIAS']
                df_p_aprob = df_p_aprob.sort_values(by='DIAS', ascending=False).head(10).sort_values(by='DIAS', ascending=True)
                
                fig_p_aprob = px.bar(
                    df_p_aprob, y='PROCESO', x='DIAS', orientation='h', text_auto='.0f',
                    color_discrete_sequence=['#00369C']
                )
                fig_p_aprob.update_layout(xaxis_title="Días Promedio", yaxis_title="")
                st.plotly_chart(fig_p_aprob, use_container_width=True)

        with col_t2:
            st.markdown("**🔄 Más lentos para REVISAR**")
            if 'PROCESO' in df_filtered.columns and 'DIAS_HASTA_REVISION' in df_filtered.columns:
                df_p_rev = df_filtered.groupby('PROCESO')['DIAS_HASTA_REVISION'].mean().dropna().reset_index()
                df_p_rev.columns = ['PROCESO', 'DIAS']
                df_p_rev = df_p_rev.sort_values(by='DIAS', ascending=False).head(10).sort_values(by='DIAS', ascending=True)
                
                fig_p_rev = px.bar(
                    df_p_rev, y='PROCESO', x='DIAS', orientation='h', text_auto='.0f',
                    color_discrete_sequence=['#F6D000']
                )
                fig_p_rev.update_layout(xaxis_title="Días Promedio", yaxis_title="")
                st.plotly_chart(fig_p_rev, use_container_width=True)

    # --- PESTAÑA 3: TABLA COMPLETA ---
    with tab_tabla:
        st.markdown("##### Consulta Detallada de Registros")
        df_display = df_filtered.copy()
        for col in [c for c in df_display.columns if 'FECHA' in c.upper()]:
            df_display[col] = df_display[col].dt.strftime('%Y-%m-%d').fillna('')
            
        st.dataframe(df_display, use_container_width=True)

except Exception as e:
    st.error(f"Error al procesar el dashboard: {e}")
