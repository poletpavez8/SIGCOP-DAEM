import os
import glob
import pandas as pd
import streamlit as st

# Configuración de la página
st.set_page_config(
    page_title="Cuadro de Mando SIGCOP-DAEM",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Cuadro de Mando Presupuestario y Control Interno — DAEM San Javier")
st.caption("Sistema Integrado de Gestión, Control Interno y Seguimiento del Ciclo de Compras y Pagos")

# BOTÓN DE CARGA EN EL PANEL LATERAL
st.sidebar.header("📁 Cargar Datos")
archivo_subido = st.sidebar.file_uploader(
    "Sube una nueva planilla Excel (.xlsx, .xls) o CSV", 
    type=["xlsx", "xls", "csv"]
)

@st.cache_data
def procesar_datos(file_input):
    # Detectar si se subió un archivo desde la web o si leemos el local
    if isinstance(file_input, str):
        # Es una ruta de archivo local
        df_raw = pd.read_excel(file_input, header=None) if file_input.endswith(('.xlsx', '.xls')) else pd.read_csv(file_input, sep=None, engine='python')
    else:
        # Es un archivo subido a través del botón web
        if file_input.name.endswith(('.xlsx', '.xls')):
            df_raw = pd.read_excel(file_input, header=None)
        else:
            df_raw = pd.read_csv(file_input, sep=None, engine='python')
        
    # Detectar en qué fila están los encabezados reales
    header_row = 0
    for i, row in df_raw.iterrows():
        row_str = [str(cell).lower() for cell in row.tolist()]
        if any("fecha_solicitud" in cell or "solicitud" in cell for cell in row_str):
            header_row = i
            break
            
    # Volver a leer desde la fila detectada
    if isinstance(file_input, str):
        df = pd.read_excel(file_input, header=header_row) if file_input.endswith(('.xlsx', '.xls')) else pd.read_csv(file_input, sep=None, engine='python')
    else:
        file_input.seek(0)
        df = pd.read_excel(file_input, header=header_row) if file_input.name.endswith(('.xlsx', '.xls')) else pd.read_csv(file_input, sep=None, engine='python')

    # Eliminar columnas vacías 'Unnamed'
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', na=False)]

    # Limpieza de espacios y caracteres invisibles en nombres de columnas
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace('\ufeff', '', regex=False)
    )
    
    # Mapeo flexible para estandarizar nombres
    column_mapping = {}
    for col in df.columns:
        col_clean = col.lower().replace(" ", "_").replace("-", "_")
        if "fecha" in col_clean and "solicitud" in col_clean:
            column_mapping[col] = "Fecha_Solicitud"
        elif "monto" in col_clean:
            column_mapping[col] = "Monto_Estimado"
        elif "estado" in col_clean:
            column_mapping[col] = "Estado"
        elif "escuela" in col_clean or "establecimiento" in col_clean:
            column_mapping[col] = "Escuela"
        elif "alerta" in col_clean or "riesgo" in col_clean:
            column_mapping[col] = "Alerta_Riesgo"

    df = df.rename(columns=column_mapping)
    return df

try:
    # Lógica de carga: usa el archivo subido desde la web o el predeterminado del repositorio
    if archivo_subido is not None:
        df = procesar_datos(archivo_subido)
        st.success(f"✅ Mostrando datos cargados desde: **{archivo_subido.name}**")
    else:
        # Buscar el archivo local
        archivos_locales = glob.glob("*.xlsx") + glob.glob("*.xls") + glob.glob("*.csv")
        if archivos_locales:
            df = procesar_datos(archivos_locales[0])
            st.info(f"ℹ️ Mostrando planilla por defecto: **{archivos_locales[0]}**")
        else:
            st.warning("⚠️ No se encontró ninguna planilla por defecto. Por favor, sube un archivo Excel desde el panel izquierdo.")
            st.stop()

    # Métrica Resumen (KPIs)
    col1, col2, col3, col4 = st.columns(4)
    
    total_solicitudes = len(df)
    monto_total = df["Monto_Estimado"].sum() if "Monto_Estimado" in df.columns else 0
    
    col1.metric("Total Solicitudes", f"{total_solicitudes}")
    col2.metric("Monto Total Estimado", f"${monto_total:,.0f}".replace(",", "."))
    
    if "Estado" in df.columns:
        pendientes = df[df["Estado"].astype(str).str.contains("PENDIENTE|Facturada|Ingresada|En Revision", case=False, na=False)].shape[0]
        col3.metric("Trámites Pendientes", f"{pendientes}")
    
    if "Alerta_Riesgo" in df.columns:
        alertas = df[df["Alerta_Riesgo"].astype(str).str.contains("Alerta|Riesgo|🟡|🔴", case=False, na=False)].shape[0]
        col4.metric("Alertas de Riesgo", f"{alertas}")

    st.markdown("---")

    # Filtros de búsqueda en la barra lateral
    st.sidebar.markdown("---")
    st.sidebar.header("Filtros de Búsqueda")
    if "Escuela" in df.columns:
        escuelas = ["Todas"] + list(df["Escuela"].dropna().unique())
        escuela_sel = st.sidebar.selectbox("Filtrar por Escuela:", escuelas)
        if escuela_sel != "Todas":
            df = df[df["Escuela"] == escuela_sel]

    if "Estado" in df.columns:
        estados = ["Todos"] + list(df["Estado"].dropna().unique())
        estado_sel = st.sidebar.selectbox("Filtrar por Estado:", estados)
        if estado_sel != "Todos":
            df = df[df["Estado"] == estado_sel]

    # Tabla principal
    st.subheader("📋 Registro General de Solicitudes de Compras y Pagos")
    st.dataframe(df, use_container_width=True)

    # Gráficos
    st.markdown("---")
    st.subheader("📈 Análisis y Control de Procesos")
    col_g1, col_g2 = st.columns(2)

    with col_g1:
        if "Estado" in df.columns:
            st.write("**Distribución por Estado de Tramitación**")
            st.bar_chart(df["Estado"].value_counts())

    with col_g2:
        if "Escuela" in df.columns and "Monto_Estimado" in df.columns:
            st.write("**Monto Estimado por Establecimiento**")
            monto_escuela = df.groupby("Escuela")["Monto_Estimado"].sum()
            st.bar_chart(monto_escuela)

except Exception as e:
    st.error(f"Ocurrió un detalle al procesar el archivo: {e}")
