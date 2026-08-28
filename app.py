from datetime import datetime
import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Compromisos y Producción", layout="wide"
)

st.title("📊 Panel de Control: Producción y Seguimiento Comercial")
st.markdown("Consolidación automática de reportes individuales de vendedores.")

st.sidebar.header("📂 Gestión de Archivos y Plantillas")

# Botón para descargar la plantilla individual oficial vacía (desde celda A1)
output_plantilla = io.BytesIO()
with pd.ExcelWriter(output_plantilla, engine="openpyxl") as writer:
  df_vacio = pd.DataFrame(
      columns=[
          "Cliente",
          "No. Contacto",
          "1o. Contacto",
          "Ultimo Contacto",
          "Cotizacion Numero",
          "Unidad",
          "Cantid",
          "Valor + IGV",
          "Status Resumido",
          "Chance Venta",
          "Mes / Año Previsto",
          "CODIGO-RESPONSABLE",
      ]
  )
  df_vacio.to_excel(writer, sheet_name="Hoja1", index=False)
plantilla_bytes = output_plantilla.getvalue()

st.sidebar.download_button(
    label="📥 Descargar Formato Individual Vacío",
    data=plantilla_bytes,
    file_name="Modelo_Control_Compromisos_Ventas.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.sidebar.markdown("---")

# 1. Subir el archivo consolidado histórico (opcional)
archivo_base = st.sidebar.file_uploader(
    "1. Sube tu Excel Consolidado actual (Opcional si es nuevo)",
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

# Inicializar DataFrames
df_plan_total = pd.DataFrame()
df_ventas_total = pd.DataFrame()
df_auditoria_total = pd.DataFrame()

# Leer el consolidado base si ya existe
if archivo_base is not None:
  try:
    xls_base = pd.ExcelFile(archivo_base)
    if "Seguimiento_Ventas" in xls_base.sheet_names:
      df_ventas_total = pd.read_excel(xls_base, sheet_name="Seguimiento_Ventas")
    if "Planificacion_Produccion" in xls_base.sheet_names:
      df_plan_total = pd.read_excel(
          xls_base, sheet_name="Planificacion_Produccion"
      )
    if "Detalle_Auditoria" in xls_base.sheet_names:
      df_auditoria_total = pd.read_excel(xls_base, sheet_name="Detalle_Auditoria")
  except Exception:
    st.sidebar.error("⚠️ El archivo consolidado base tiene un formato inválido.")

# Procesar los reportes individuales de los vendedores
if archivos_nuevos:
  for file in archivos_nuevos:
    try:
      xls = pd.ExcelFile(file)
      # Leyendo desde la fila 0 (celda A1) porque ya limpiaste el formato
      df_vendedor = pd.read_excel(xls, sheet_name=0, header=0).dropna(
          how="all"
      )

      # Validar que el archivo contenga las columnas esenciales
      columnas_texto = " ".join([str(c) for c in df_vendedor.columns])
      if "Cotizacion" not in columnas_texto and "Numero" not in columnas_texto:
        st.sidebar.error(
            f"❌ El archivo '{file.name}' no tiene la columna de Cotización en"
            " la primera fila. Verifica tu formato en A1."
        )
        continue

      # Agregar etiqueta de trazabilidad
      df_vendedor["Cierre_Semanal"] = file.name

      # Unir al consolidado general de ventas
      if not df_ventas_total.empty:
        df_ventas_total = pd.concat(
            [df_ventas_total, df_vendedor], ignore_index=True
        ).drop_duplicates()
      else:
        df_ventas_total = df_vendedor

      st.sidebar.success(f"✅ ¡Procesado correctamente: {file.name}!")
    except Exception:
      st.sidebar.error(
          f"❌ Error al procesar '{file.name}'. Revisa que comience en A1."
      )

# --- SECCIÓN DE DESCARGA DEL EXCEL CONSOLIDADO FINAL ---
if not df_ventas_total.empty or not df_plan_total.empty:
  st.sidebar.markdown("---")

  fecha_actual = datetime.now().strftime("%d-%m-%y")
  nombre_archivo = f"Planificacion_Produccion_y_Ventas_Final_{fecha_actual}.xlsx"

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    if not df_ventas_total.empty:
      df_ventas_total.to_excel(
          writer, sheet_name="Seguimiento_Ventas", index=False
      )
    if not df_plan_total.empty:
      df_plan_total.to_excel(
          writer, sheet_name="Planificacion_Produccion", index=False
      )
    else:
      pd.DataFrame(columns=["Mensaje"]).to_excel(
          writer, sheet_name="Planificacion_Produccion", index=False
      )

    if not df_auditoria_total.empty:
      df_auditoria_total.to_excel(
          writer, sheet_name="Detalle_Auditoria", index=False
      )
    else:
      pd.DataFrame(columns=["Mensaje"]).to_excel(
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
        "💰 Seguimiento Ventas Consolidado",
        "📈 Planificación Consolidada",
        "📋 Auditoría General Consolidada",
    ]
)

with tab1:
  st.subheader("Seguimiento de Ventas (Consolidado de Vendedores)")
  if not df_ventas_total.empty:
    st.dataframe(df_ventas_total, use_container_width=True)
    st.info(f"Total de registros consolidados: {len(df_ventas_total)}")
  else:
    st.info("Sube los reportes individuales de los vendedores en la barra lateral.")

with tab2:
  st.subheader("Planificación de Producción")
  if not df_plan_total.empty:
    st.dataframe(df_plan_total, use_container_width=True)
  else:
    st.info("Sin registros de planificación cargados.")

with tab3:
  st.subheader("Detalle de Auditoría")
  if not df_auditoria_total.empty:
    st.dataframe(df_auditoria_total, use_container_width=True)
  else:
    st.info("Sin registros de auditoría cargados.")
