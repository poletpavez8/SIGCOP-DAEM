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

@st.cache_data
def cargar_datos():
    # Buscar archivos Excel o CSV en la carpeta actual
    archivos_excel = glob.glob("*.xlsx") + glob.glob("*.xls")
    archivos_csv = glob.glob("*.csv")
    
    archivo_target = None
    if archivos_excel:
        archivo_target = archivos_excel[0]
        
        # Leemos sin encabezados para detectar en qué fila están los títulos reales
        df_raw = pd.read_excel(archivo_target, header=None)
        
        header_row = 0
        for i, row in df_raw.iterrows():
            # Convertimos todas las celdas a String para evitar el error de float / NaN
            row_str = [str(cell).lower() for cell in row.tolist()]
            if any("fecha_solicitud" in cell or "solicitud" in cell for cell in row_str):
                header_row = i
                break
                
        # Leemos el Excel desde la fila correcta
        df = pd.read_excel(archivo_target, header=header_row)

    elif archivos_csv:
        archivo_target = archivos_csv[0]
        try:
            df = pd.read_csv(archivo_target, sep=None, engine='python')
        except Exception:
            df = pd.read_csv(archivo_target, sep=';')
    else:
        st.error("No se encontró ningún archivo Excel (.xlsx, .xls) o CSV (.csv) en la carpeta.")
        st.stop()

    # Eliminar columnas vacías 'Unnamed' producidas por celdas sin título en el Excel
    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', na=False)]

    # Limpieza de nombres de columnas
    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace('\ufeff', '', regex=False)
    )
    
    # Mapeo flexible para asegurar estándar
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
    return df, archivo_target

try:
    df, archivo_usado = cargar_datos()
    st.success(f"✅ Datos cargados correctamente desde: **{archivo_usado}**")

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

    # Filtros de búsqueda laterales
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
    st.error(f"Ocurrió un detalle al cargar los datos: {e}")
