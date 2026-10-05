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
    "Sube una planilla Excel (.xlsx, .xls) o CSV", 
    type=["xlsx", "xls", "csv"]
)

def extraer_datos_sep(file_input):
    """Extrae y consolida datos de planillas de control financiero tipo SEP (multihoja por escuela)."""
    try:
        xls = pd.ExcelFile(file_input)
    except Exception:
        return None

    registros = []
    hojas_omitir = ['SUBVENCION', 'COSTO REMUNERACION ', 'PROPORCION GASTOS AC', '10% AC SEP']

    for sheet in xls.sheet_names:
        if sheet.strip() in hojas_omitir:
            continue
            
        df_sheet = pd.read_excel(file_input, sheet_name=sheet, header=None)
        
        # Extraer el nombre de la Escuela
        nombre_escuela = sheet
        for i in range(min(6, len(df_sheet))):
            val = str(df_sheet.iloc[i, 1]) if len(df_sheet.columns) > 1 else ""
            if "ESCUELA:" in val.upper() or "ESCUELA :" in val.upper():
                nombre_escuela = val.replace("ESCUELA:", "").replace("Escuela:", "").strip()
                break

        # Buscar filas con transacciones
        for i in range(len(df_sheet)):
            row = df_sheet.iloc[i].tolist()
            fecha = row[1] if len(row) > 1 else None
            sol = str(row[2]) if len(row) > 2 and pd.notna(row[2]) else ""
            oc = str(row[3]) if len(row) > 3 and pd.notna(row[3]) else ""
            dp = str(row[4]) if len(row) > 4 and pd.notna(row[4]) else ""
            categoria = str(row[5]) if len(row) > 5 and pd.notna(row[5]) else ""
            detalle = str(row[6]) if len(row) > 6 and pd.notna(row[6]) else ""
            gasto = row[8] if len(row) > 8 and pd.notna(row[8]) else 0

            # Validar si es registro de gasto
            if pd.notna(fecha) and str(fecha).strip() not in ["FECHA", "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE", "nan"]:
                try:
                    monto = float(gasto) if pd.notna(gasto) else 0.0
                    if monto > 0:
                        estado = "Pagada" if dp and dp.upper() != "PENDIENTE" and dp.upper() != "NAN" else ("En Compra" if oc else "Solicitada")
                        registros.append({
                            "Cod_Solicitud": f"SOL-{sol}" if sol else "S/N",
                            "Fecha_Solicitud": str(fecha)[:10],
                            "Escuela": nombre_escuela,
                            "Orden_Compra": oc if oc else "PENDIENTE",
                            "Decreto_Pago": dp if dp else "PENDIENTE",
                            "Categoria": categoria,
                            "Detalle": detalle,
                            "Monto_Estimado": monto,
                            "Estado": estado,
                            "Alerta_Riesgo": "🟢 Normal" if estado == "Pagada" else "🟡 Riesgo Atendible"
                        })
                except ValueError:
                    continue

    if registros:
        return pd.DataFrame(registros)
    return None


def limpiar_columnas_no_deseadas(df):
    """Elimina las columnas no requeridas."""
    columnas_excluir = [
        'proveedor', 'proveedores', 'rut_proveedor', 'rut', 'rut_de_proveedor',
        'certificado_presupuestario', 'cp', 'certificado',
        'acta_recepcion_conforme', 'acta_recepcion', 'acta', 'arc',
        'factura_folio', 'factura', 'folio_factura'
    ]
    cols_a_borrar = []
    for col in df.columns:
        col_clean = str(col).lower().replace(" ", "_").replace("-", "_").strip()
        if col_clean in columnas_excluir:
            cols_a_borrar.append(col)
            
    return df.drop(columns=cols_a_borrar, errors='ignore')


def cargar_datos(file_input):
    df_sep = extraer_datos_sep(file_input)
    if df_sep is not None and not df_sep.empty:
        return limpiar_columnas_no_deseadas(df_sep)

    if isinstance(file_input, str):
        df_raw = pd.read_excel(file_input, header=None) if file_input.endswith(('.xlsx', '.xls')) else pd.read_csv(file_input, sep=None, engine='python')
    else:
        if file_input.name.endswith(('.xlsx', '.xls')):
            df_raw = pd.read_excel(file_input, header=None)
        else:
            df_raw = pd.read_csv(file_input, sep=None, engine='python')
        
    header_row = 0
    for i, row in df_raw.iterrows():
        row_str = [str(cell).lower() for cell in row.tolist()]
        if any("fecha_solicitud" in cell or "solicitud" in cell or "escuela" in cell for cell in row_str):
            header_row = i
            break
            
    if isinstance(file_input, str):
        df = pd.read_excel(file_input, header=header_row) if file_input.endswith(('.xlsx', '.xls')) else pd.read_csv(file_input, sep=None, engine='python')
    else:
        file_input.seek(0)
        if file_input.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_input, header=header_row)
        else:
            df = pd.read_csv(file_input, sep=None, engine='python')

    df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed', na=False)]
    df.columns = df.columns.astype(str).str.strip().str.replace('\ufeff', '', regex=False)
    
    column_mapping = {}
    for col in df.columns:
        col_clean = col.lower().replace(" ", "_").replace("-", "_")
        if "fecha" in col_clean and "solicitud" in col_clean:
            column_mapping[col] = "Fecha_Solicitud"
        elif "monto" in col_clean or "gasto" in col_clean:
            column_mapping[col] = "Monto_Estimado"
        elif "estado" in col_clean:
            column_mapping[col] = "Estado"
        elif "escuela" in col_clean or "establecimiento" in col_clean:
            column_mapping[col] = "Escuela"
        elif "alerta" in col_clean or "riesgo" in col_clean:
            column_mapping[col] = "Alerta_Riesgo"

    df = df.rename(columns=column_mapping)
    return limpiar_columnas_no_deseadas(df)


try:
    if archivo_subido is not None:
        df = cargar_datos(archivo_subido)
        st.success(f"✅ Datos cargados correctamente desde: **{archivo_subido.name}**")
    else:
        archivos_locales = glob.glob("*.xlsx") + glob.glob("*.xls") + glob.glob("*.csv")
        if archivos_locales:
            archivo_target = archivos_locales[0]
            df = cargar_datos(archivo_target)
            st.info(f"ℹ️ Mostrando planilla por defecto: **{archivo_target}**")
        else:
            st.warning("⚠️ No se encontró ninguna planilla. Por favor, sube un archivo Excel desde el panel izquierdo.")
            st.stop()

    # Métrica Resumen (KPIs)
    col1, col2, col3, col4 = st.columns(4)
    
    total_solicitudes = len(df)
    monto_total = df["Monto_Estimado"].sum() if "Monto_Estimado" in df.columns else 0
    
    col1.metric("Total Solicitudes / Gastos", f"{total_solicitudes}")
    col2.metric("Monto Total Ejecutado/Estimado", f"${monto_total:,.0f}".replace(",", "."))
    
    if "Estado" in df.columns:
        pendientes = df[df["Estado"].astype(str).str.contains("PENDIENTE|Facturada|Ingresada|En Revision|Solicitada|En Compra", case=False, na=False)].shape[0]
        col3.metric("Trámites Pendientes", f"{pendientes}")
    
    if "Alerta_Riesgo" in df.columns:
        alertas = df[df["Alerta_Riesgo"].astype(str).str.contains("Alerta|Riesgo|🟡|🔴", case=False, na=False)].shape[0]
        col4.metric("Alertas de Riesgo", f"{alertas}")

    st.markdown("---")

    # Filtros de búsqueda laterales
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
    st.subheader("📋 Consolidado General de Solicitudes y Gastos SEP")
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
            st.write("**Gasto Total por Escuela / Establecimiento**")
            monto_escuela = df.groupby("Escuela")["Monto_Estimado"].sum()
            st.bar_chart(monto_escuela)

except Exception as e:
    st.error(f"Ocurrió un detalle al procesar la planilla: {e}")
