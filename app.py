import streamlit as st
import pandas as pd
import plotly.express as px

# Configuración de página ancha estilo Power BI
st.set_page_config(
    page_title="Dashboard - Libro Maestro de Documentos",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Dashboard - Libro Maestro de Documentos")
st.markdown("---")

@st.cache_data
def load_data():
    file_path = "M-FT-1 LIBRO MAESTRO DE DOCUMENTOS v2.xlsx"
    
    # Cargar la hoja LIBRO MAESTRO omitiendo encabezados iniciales
    df = pd.read_excel(file_path, sheet_name="LIBRO MAESTRO", header=2)
    
    # Limpieza estricta de columnas
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]
    df = df.dropna(how="all")
    
    # Normalizar columna PROCESO (quitar espacios sobrantes)
    if 'PROCESO' in df.columns:
        df['PROCESO'] = df['PROCESO'].astype(str).str.strip()
    
    # Convertir columnas de fecha a datetime para cálculos de tiempo
    date_columns = [c for c in df.columns if 'FECHA' in c.upper()]
    for col in date_columns:
        df[col] = pd.to_datetime(df[col], errors='coerce')
        
    # Cálculo de tiempos en días
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

    # --- BARRA LATERAL: FILTROS DINÁMICOS ---
    st.sidebar.header("🔍 Filtros Dinámicos")
    
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

    # --- APLICACIÓN DE FILTROS ---
    df_filtered = df_raw.copy()
    
    if selected_unidad and col_unidad in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_unidad].astype(str).str.strip().isin(selected_unidad)]
    if selected_macro and col_macro in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_macro].astype(str).str.strip().isin(selected_macro)]
    if selected_estado and col_estado in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_estado].astype(str).str.strip().isin(selected_estado)]
    if selected_tipo and col_tipo in df_filtered.columns:
        df_filtered = df_filtered[df_filtered[col_tipo].astype(str).str.strip().isin(selected_tipo)]

    # --- KPIS PRINCIPALES ---
    st.subheader("📌 Indicadores Generales")
    kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
    
    kpi1.metric("📄 Total Documentos", len(df_filtered))
    
    if col_estado in df_filtered.columns:
        aprobados = len(df_filtered[df_filtered[col_estado].astype(str).str.strip().str.upper() == 'APROBADO'])
        no_creados = len(df_filtered[df_filtered[col_estado].astype(str).str.strip().str.upper() == 'NO CREADO'])
    else:
        aprobados, no_creados = 0, 0
        
    kpi2.metric("✅ Aprobados", aprobados)
    kpi3.metric("⏳ No Creados", no_creados)
    
    avg_dias_aprob = df_filtered['DIAS_HASTA_APROBACION'].mean()
    avg_dias_rev = df_filtered['DIAS_HASTA_REVISION'].mean()
    
    val_aprob_str = f"{int(avg_dias_aprob)} días" if pd.notna(avg_dias_aprob) else "N/A"
    val_rev_str = f"{int(avg_dias_rev)} días" if pd.notna(avg_dias_rev) else "N/A"
    
    kpi4.metric("⏱️ Promedio Aprobación", val_aprob_str)
    kpi5.metric("🔄 Promedio Revisión", val_rev_str)

    st.markdown("---")

    # --- FILA 1 DE GRÁFICOS: ESTADOS Y UNIDADES ---
    g_col1, g_col2 = st.columns(2)

    with g_col1:
        st.subheader("📌 Documentos por Estado")
        if col_estado in df_filtered.columns:
            df_estado = df_filtered[col_estado].dropna().astype(str).str.strip().value_counts().reset_index()
            df_estado.columns = ['ESTADO', 'CANTIDAD']
            fig_estado = px.pie(
                df_estado, 
                values='CANTIDAD', 
                names='ESTADO', 
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Pastel
            )
            st.plotly_chart(fig_estado, use_container_width=True)

    with g_col2:
        st.subheader("🏢 Documentos por Unidad de Negocio")
        if col_unidad in df_filtered.columns:
            df_unidad = df_filtered[col_unidad].dropna().astype(str).str.strip().value_counts().reset_index()
            df_unidad.columns = ['UNIDAD DE NEGOCIO', 'CANTIDAD']
            fig_unidad = px.bar(
                df_unidad, 
                x='UNIDAD DE NEGOCIO', 
                y='CANTIDAD', 
                color='UNIDAD DE NEGOCIO',
                text='CANTIDAD'
            )
            fig_unidad.update_layout(showlegend=False)
            st.plotly_chart(fig_unidad, use_container_width=True)

    st.markdown("---")

    # --- FILA 2 DE GRÁFICOS: TIPOS Y TIEMPOS TOTALES ---
    g_col3, g_col4 = st.columns(2)

    with g_col3:
        st.subheader("📁 Documentos por Tipo")
        if col_tipo in df_filtered.columns:
            df_tipo = df_filtered[col_tipo].dropna().astype(str).str.strip().value_counts().reset_index()
            df_tipo.columns = ['TIPO DE DOCUMENTO', 'CANTIDAD']
            df_tipo = df_tipo.sort_values(by='CANTIDAD', ascending=True)
            
            fig_tipo = px.bar(
                df_tipo, 
                y='TIPO DE DOCUMENTO', 
                x='CANTIDAD', 
                orientation='h',
                color='CANTIDAD',
                color_continuous_scale='Blues',
                text='CANTIDAD'
            )
            fig_tipo.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig_tipo, use_container_width=True)

    with g_col4:
        st.subheader("⏱️ Comparativa de Tiempos de Ciclo")
        tiempos_data = pd.DataFrame({
            'Etapa': ['Aprobación (Creación ➔ Aprobado)', 'Revisión (Creación ➔ Revisado)'],
            'Días Promedio': [
                avg_dias_aprob if pd.notna(avg_dias_aprob) else 0, 
                avg_dias_rev if pd.notna(avg_dias_rev) else 0
            ]
        })
        
        fig_tiempos = px.bar(
            tiempos_data,
            x='Etapa',
            y='Días Promedio',
            color='Etapa',
            text_auto='.0f',
            color_discrete_sequence=['#2ba02b', '#ff7f0e']
        )
        fig_tiempos.update_layout(showlegend=False, yaxis_title="Días")
        st.plotly_chart(fig_tiempos, use_container_width=True)

    st.markdown("---")

    # --- FILA 3 DE GRÁFICOS (SPRINT 3: PROCESOS MÁS LENTOS) ---
    st.subheader("🐢 Análisis de Cuelgues: Top Procesos Más Lentos")
    g_col5, g_col6 = st.columns(2)

    with g_col5:
        st.markdown("##### ⏳ Top Procesos que más tardan en APROBAR (Días Promedio)")
        if 'PROCESO' in df_filtered.columns and 'DIAS_HASTA_APROBACION' in df_filtered.columns:
            df_p_aprob = df_filtered.groupby('PROCESO')['DIAS_HASTA_APROBACION'].mean().dropna().reset_index()
            df_p_aprob.columns = ['PROCESO', 'DIAS_PROMEDIO']
            df_p_aprob = df_p_aprob.sort_values(by='DIAS_PROMEDIO', ascending=False).head(10)
            df_p_aprob = df_p_aprob.sort_values(by='DIAS_PROMEDIO', ascending=True) # Para visualización horizontal ordenada
            
            fig_p_aprob = px.bar(
                df_p_aprob,
                y='PROCESO',
                x='DIAS_PROMEDIO',
                orientation='h',
                text_auto='.0f',
                color='DIAS_PROMEDIO',
                color_continuous_scale='Reds'
            )
            fig_p_aprob.update_layout(coloraxis_showscale=False, xaxis_title="Días Promedio")
            st.plotly_chart(fig_p_aprob, use_container_width=True)

    with g_col6:
        st.markdown("##### 🔄 Top Procesos que más tardan en REVISAR (Días Promedio)")
        if 'PROCESO' in df_filtered.columns and 'DIAS_HASTA_REVISION' in df_filtered.columns:
            df_p_rev = df_filtered.groupby('PROCESO')['DIAS_HASTA_REVISION'].mean().dropna().reset_index()
            df_p_rev.columns = ['PROCESO', 'DIAS_PROMEDIO']
            df_p_rev = df_p_rev.sort_values(by='DIAS_PROMEDIO', ascending=False).head(10)
            df_p_rev = df_p_rev.sort_values(by='DIAS_PROMEDIO', ascending=True)
            
            fig_p_rev = px.bar(
                df_p_rev,
                y='PROCESO',
                x='DIAS_PROMEDIO',
                orientation='h',
                text_auto='.0f',
                color='DIAS_PROMEDIO',
                color_continuous_scale='Oranges'
            )
            fig_p_rev.update_layout(coloraxis_showscale=False, xaxis_title="Días Promedio")
            st.plotly_chart(fig_p_rev, use_container_width=True)

    st.markdown("---")

    # --- TABLA INTERACTIVA ---
    st.subheader("📋 Detalle del Libro Maestro")
    
    df_display = df_filtered.copy()
    for col in [c for c in df_display.columns if 'FECHA' in c.upper()]:
        df_display[col] = df_display[col].dt.strftime('%Y-%m-%d').fillna('')
        
    st.dataframe(df_display, use_container_width=True)

except Exception as e:
    st.error(f"Error al procesar el dashboard: {e}")