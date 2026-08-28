from datetime import datetime
import io
import openpyxl
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Compromisos y Producción", layout="wide"
)

st.title("📊 Panel de Control: Producción y Seguimiento Comercial")
st.markdown("Consolidación inteligente preservando la plantilla y formato base.")

st.sidebar.header("📂 Gestión de Archivos")

# 1. Subir el archivo consolidado base / modelo (que contiene las 3 pestañas con su diseño original)
archivo_base = st.sidebar.file_uploader(
    "1. Sube tu Excel Consolidado o Modelo Base", type=["xlsx"], key="base"
)

# 2. Subir los reportes individuales de los vendedores
archivos_nuevos = st.sidebar.file_uploader(
    "2. Sube los reportes individuales semanales de los vendedores",
    type=["xlsx"],
    accept_multiple_files=True,
    key="nuevos",
)

# Mantener estado en session_state para que no se pierda al subir archivos individuales
if "df_auditoria" not in st.session_state:
  st.session_state["df_auditoria"] = pd.DataFrame()
if "df_plan" not in st.session_state:
  st.session_state["df_plan"] = pd.DataFrame()
if "df_ventas" not in st.session_state:
  st.session_state["df_ventas"] = pd.DataFrame()
if "wb_template" not in st.session_state:
  st.session_state["wb_template"] = None

# Leer el archivo consolidado base si se sube
if archivo_base is not None:
  try:
    # Guardar copia en memoria del workbook original para conservar estilos y formato completo
    archivo_base.seek(0)
    st.session_state["wb_template"] = openpyxl.load_workbook(archivo_base)

    archivo_base.seek(0)
    xls_base = pd.ExcelFile(archivo_base)

    if "Planificacion_Produccion" in xls_base.sheet_names:
      st.session_state["df_plan"] = pd.read_excel(
          xls_base, sheet_name="Planificacion_Produccion", header=2
      ).dropna(how="all")

    if "Seguimiento_Ventas" in xls_base.sheet_names:
      st.session_state["df_ventas"] = pd.read_excel(
          xls_base, sheet_name="Seguimiento_Ventas", header=2
      ).dropna(how="all")

    if "Detalle_Auditoria" in xls_base.sheet_names:
      st.session_state["df_auditoria"] = pd.read_excel(
          xls_base, sheet_name="Detalle_Auditoria", header=2
      ).dropna(how="all")

    st.sidebar.success("✅ ¡Plantilla base cargada y memorizada con éxito!")
  except Exception as e:
    st.sidebar.error(f"⚠️ Error al leer el archivo base: {e}")

# Procesar los reportes individuales de los vendedores
if archivos_nuevos:
  for file in archivos_nuevos:
    try:
      xls = pd.ExcelFile(file)
      df_vendedor = pd.read_excel(xls, sheet_name=0, header=2).dropna(
          how="all"
      )

      df_vendedor = df_vendedor.loc[
          :, ~df_vendedor.columns.str.contains("^Unnamed")
      ]
      df_vendedor["Cierre_Semanal"] = file.name

      if not st.session_state["df_auditoria"].empty:
        st.session_state["df_auditoria"] = pd.concat(
            [st.session_state["df_auditoria"], df_vendedor], ignore_index=True
        ).drop_duplicates()
      else:
        st.session_state["df_auditoria"] = df_vendedor

      st.sidebar.success(f"✅ ¡Vendedor procesado: {file.name}!")
    except Exception as e:
      st.sidebar.error(
          f"❌ Formato no compatible en '{file.name}'. Verifique que sea un"
          " reporte válido."
      )

# --- SECCIÓN DE DESCARGA CONSERVANDO EL FORMATO ORIGINAL ---
if (
    not st.session_state["df_plan"].empty
    or not st.session_state["df_ventas"].empty
    or not st.session_state["df_auditoria"].empty
):
  st.sidebar.markdown("---")
  fecha_actual = datetime.now().strftime("%d-%m-%y")
  nombre_archivo = f"Planificacion_Produccion_y_Ventas_Final_{fecha_actual}.xlsx"

  output = io.BytesIO()

  # Si tenemos la plantilla original cargada, la usamos para reescribir solo los datos manteniendo estilos
  if st.session_state["wb_template"] is not None:
    wb = st.session_state["wb_template"]

    # Actualizar la pestaña Detalle_Auditoria manteniendo las primeras 3 filas de títulos
    if (
        "Detalle_Auditoria" in wb.sheetnames
        and not st.session_state["df_auditoria"].empty
    ):
      ws = wb["Detalle_Auditoria"]
      # Borrar filas de datos anteriores (desde la fila 4 en adelante)
      if ws.max_row >= 4:
        ws.delete_rows(4, ws.max_row - 3)

      # Escribir los nuevos datos desde la fila 4
      for r_idx, row in enumerate(
          st.session_state["df_auditoria"].values, start=4
      ):
        for c_idx, val in enumerate(row, start=1):
          ws.cell(row=r_idx, column=c_idx, value=val)

    wb.save(output)
    excel_data = output.getvalue()
  else:
    # Respaldo plano por si no subieron plantilla base
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      if not st.session_state["df_plan"].empty:
        st.session_state["df_plan"].to_excel(
            writer, sheet_name="Planificacion_Produccion", index=False
        )
      if not st.session_state["df_ventas"].empty:
        st.session_state["df_ventas"].to_excel(
            writer, sheet_name="Seguimiento_Ventas", index=False
        )
      if not st.session_state["df_auditoria"].empty:
        st.session_state["df_auditoria"].to_excel(
            writer, sheet_name="Detalle_Auditoria", index=False
        )
    excel_data = output.getvalue()

  st.sidebar.download_button(
      label="📥 Descargar Consolidado con Formato Original",
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
  if not st.session_state["df_plan"].empty:
    st.dataframe(st.session_state["df_plan"], use_container_width=True)
  else:
    st.info("Sube tu archivo consolidado base.")

with tab2:
  st.subheader("Seguimiento de Rendimiento Comercial")
  if not st.session_state["df_ventas"].empty:
    st.dataframe(st.session_state["df_ventas"], use_container_width=True)
  else:
    st.info("Sube tu archivo consolidado base.")

with tab3:
  st.subheader("Detalle General y Auditoría de Cotizaciones")
  if not st.session_state["df_auditoria"].empty:
    st.dataframe(st.session_state["df_auditoria"], use_container_width=True)
    st.success(
        f"Registros totales en auditoría: {len(st.session_state['df_auditoria'])}"
    )
  else:
    st.info("Sube tu archivo consolidado o reportes individuales.")
