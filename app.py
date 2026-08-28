from datetime import datetime
import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Compromisos y Producción", layout="wide"
)

st.title("📊 Panel de Control: Producción y Seguimiento Comercial")
st.markdown("Consolidación inteligente de reportes individuales y base.")

st.sidebar.header("📂 Gestión de Archivos")

# 1. Subir el archivo consolidado base (que contiene las 3 pestañas)
archivo_base = st.sidebar.file_uploader(
    "1. Sube tu Excel Consolidado actual (3 pestañas)",
    type=["xlsx"],
    key="base",
)

# 2. Subir los reportes individuales de los vendedores
archivos_nuevos = st.sidebar.file_uploader(
    "2. Sube los reportes individuales semanales de los vendedores",
    type=["xlsx"],
    accept_multiple_files=True,
    key="nuevos",
)

# Inicializar DataFrames para las 3 pestañas
df_plan_total = pd.DataFrame()
df_ventas_total = pd.DataFrame()
df_auditoria_total = pd.DataFrame()

# Leer el archivo consolidado base si se sube (usando header=2 para capturar las cabeceras reales)
if archivo_base is not None:
  try:
    xls_base = pd.ExcelFile(archivo_base)

    if "Planificacion_Produccion" in xls_base.sheet_names:
      df_plan_total = pd.read_excel(
          xls_base, sheet_name="Planificacion_Produccion", header=2
      ).dropna(how="all")

    if "Seguimiento_Ventas" in xls_base.sheet_names:
      df_ventas_total = pd.read_excel(
          xls_base, sheet_name="Seguimiento_Ventas", header=2
      ).dropna(how="all")

    if "Detalle_Auditoria" in xls_base.sheet_names:
      df_auditoria_total = pd.read_excel(
          xls_base, sheet_name="Detalle_Auditoria", header=2
      ).dropna(how="all")

    st.sidebar.success(
        "✅ ¡Archivo consolidado base cargado con sus cabeceras correctas!"
    )
  except Exception as e:
    st.sidebar.error(f"⚠️ Error al leer el consolidado base: {e}")

# Procesar los reportes individuales de los vendedores
if archivos_nuevos:
  for file in archivos_nuevos:
    try:
      xls = pd.ExcelFile(file)
      # Los reportes individuales tienen la cabecera real en la fila 2 (índice 2) o combinada filas 1-2.
      # Usamos header=2 que contiene: Cliente, No. Contacto, Ultimo Contacto, Cotizacion, etc.
      df_vendedor = pd.read_excel(xls, sheet_name=0, header=2).dropna(
          how="all"
      )

      # Limpiar nombres de columnas por si tienen espacios o valores NaN
      df_vendedor = df_vendedor.loc[
          :, ~df_vendedor.columns.str.contains("^Unnamed")
      ]

      # Agregar columna de trazabilidad con el nombre del archivo
      df_vendedor["Cierre_Semanal"] = file.name

      # Unir al consolidado de auditoría
      if not df_auditoria_total.empty:
        df_auditoria_total = pd.concat(
            [df_auditoria_total, df_vendedor], ignore_index=True
        ).drop_duplicates()
      else:
        df_auditoria_total = df_vendedor

      st.sidebar.success(f"✅ ¡Procesado correctamente: {file.name}!")
    except Exception as e:
      st.sidebar.error(f"❌ Error al procesar '{file.name}': {e}")

# --- SECCIÓN DE DESCARGA DEL EXCEL CONSOLIDADO FINAL ---
if (
    not df_plan_total.empty
    or not df_ventas_total.empty
    or not df_auditoria_total.empty
):
  st.sidebar.markdown("---")

  fecha_actual = datetime.now().strftime("%d-%m-%y")
  nombre_archivo = f"Planificacion_Produccion_y_Ventas_Final_{fecha_actual}.xlsx"

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    if not df_plan_total.empty:
      df_plan_total.to_excel(
          writer, sheet_name="Planificacion_Produccion", index=False
      )
    if not df_ventas_total.empty:
      df_ventas_total.to_excel(
          writer, sheet_name="Seguimiento_Ventas", index=False
      )
    if not df_auditoria_total.empty:
      df_auditoria_total.to_excel(
          writer, sheet_name="Detalle_Auditoria", index=False
      )

  excel_data = output.getvalue()

  st.sidebar.download_button(
      label="📥 Descargar Excel Consolidado Final",
      data=excel_data,
      file_name=nombre_archivo,
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )

# --- PESTAÑAS DE VISUALIZACIÓN ---
tab1, tab2, tab3 = st.tabs(
    [
        "📈 Planificación de Producción",
        "💰 Seguimiento Ventas",
        "📋 Detalle de Auditoría",
    ]
)

with tab1:
  st.subheader("Planificación y Requerimiento de Producción")
  if not df_plan_total.empty:
    st.dataframe(df_plan_total, use_container_width=True)
  else:
    st.info("Sube tu archivo consolidado base para ver esta pestaña.")

with tab2:
  st.subheader("Seguimiento de Rendimiento Comercial")
  if not df_ventas_total.empty:
    st.dataframe(df_ventas_total, use_container_width=True)
  else:
    st.info("Sube tu archivo consolidado base para ver esta pestaña.")

with tab3:
  st.subheader("Detalle General y Auditoría de Cotizaciones")
  if not df_auditoria_total.empty:
    st.dataframe(df_auditoria_total, use_container_width=True)
    st.success(f"Registros totales en auditoría: {len(df_auditoria_total)}")
  else:
    st.info(
        "Sube tu archivo consolidado o los reportes individuales de los"
        " vendedores."
    )
