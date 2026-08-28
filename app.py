from datetime import datetime
import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Compromisos y Producción", layout="wide"
)

st.title("📊 Panel de Control: Producción y Seguimiento Comercial")
st.markdown(
    "Sistema inteligente con control de cierres semanales (Viernes) y"
    " actualización por número de cotización."
)

st.sidebar.header("📂 Gestión de Archivos")

# 1. Subir el archivo consolidado histórico
archivo_base = st.sidebar.file_uploader(
    "1. Sube tu Excel Consolidado actual (Opcional si es nuevo)",
    type=["xlsx"],
    key="base",
)

# 2. Subir los nuevos reportes de los vendedores
archivos_nuevos = st.sidebar.file_uploader(
    "2. Sube los nuevos reportes semanales de los vendedores",
    type=["xlsx"],
    accept_multiple_files=True,
    key="nuevos",
)

# Inicializar DataFrames vacíos
df_plan_total = pd.DataFrame()
df_ventas_total = pd.DataFrame()
df_auditoria_total = pd.DataFrame()

# Leer el consolidado base si existe
if archivo_base is not None:
  try:
    df_plan_total = pd.read_excel(
        archivo_base, sheet_name="Planificacion_Produccion"
    )
    df_ventas_total = pd.read_excel(
        archivo_base, sheet_name="Seguimiento_Ventas"
    )
    df_auditoria_total = pd.read_excel(
        archivo_base, sheet_name="Detalle_Auditoria"
    )
  except Exception as e:
    st.sidebar.error(f"Error al leer el archivo consolidado base: {e}")

# Procesar y actualizar con los nuevos reportes
if archivos_nuevos:
  for file in archivos_nuevos:
    try:
      # Extraer la fecha del nombre del archivo si la trae (ej: "al 20-08-2026") o usar la fecha actual
      fecha_corte = datetime.now().strftime("%Y-%m-%d")
      for parte in file.name.replace(".xlsx", "").split("al "):
        if len(parte.strip() >= 10):
          # Intento extraer fecha del nombre si existe
          pass

      df_p = pd.read_excel(
          file, sheet_name="Planificacion_Produccion", skiprows=2
      ).dropna(how="all")
      df_v = pd.read_excel(
          file, sheet_name="Seguimiento_Ventas", skiprows=2
      ).dropna(how="all")
      df_a = pd.read_excel(
          file, sheet_name="Detalle_Auditoria", skiprows=2
      ).dropna(how="all")

      # Agregar etiqueta de origen o archivo para trazabilidad de la semana de cierre
      df_p["Cierre_Semanal"] = file.name
      df_v["Cierre_Semanal"] = file.name
      df_a["Cierre_Semanal"] = file.name

      # Actualización inteligente por número de Cotización para evitar duplicados
      if not df_plan_total.empty:
        df_plan_total = pd.concat([df_plan_total, df_p]).drop_duplicates()
      else:
        df_plan_total = df_p

      for col_cot in [
          "Cotización",
          "Cotizacion",
          "Cotizacion Numero",
          "Nº Cotizaciones",
          "Numero",
      ]:
        if col_cot in df_ventas_total.columns and col_cot in df_v.columns:
          df_ventas_total = df_ventas_total[
              ~df_ventas_total[col_cot].isin(df_v[col_cot])
          ]
        if (
            col_cot in df_auditoria_total.columns
            and col_cot in df_a.columns
        ):
          df_auditoria_total = df_auditoria_total[
              ~df_auditoria_total[col_cot].isin(df_a[col_cot])
          ]

      df_ventas_total = pd.concat([df_ventas_total, df_v], ignore_index=True)
      df_auditoria_total = pd.concat(
          [df_auditoria_total, df_a], ignore_index=True
      )

      st.sidebar.success(
          f"Procesado (Cierre semanal considerado): {file.name}"
      )
    except Exception as e:
      st.sidebar.error(f"Error en {file.name}: {e}")

# --- SECCIÓN DE DESCARGA DEL EXCEL ACTUALIZADO CON FECHA ---
if not df_plan_total.empty:
  st.sidebar.markdown("---")

  fecha_actual = datetime.now().strftime("%d-%m-%y")
  nombre_archivo = f"Planificacion_Produccion_y_Ventas_Final_{fecha_actual}.xlsx"

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_plan_total.to_excel(
        writer, sheet_name="Planificacion_Produccion", index=False
    )
    df_ventas_total.to_excel(writer, sheet_name="Seguimiento_Ventas", index=False)
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
        "📈 Planificación Consolidada",
        "💰 Seguimiento Ventas Consolidado",
        "📋 Auditoría General Consolidada",
    ]
)

with tab1:
  st.subheader("Planificación de Producción")
  if not df_plan_total.empty:
    st.dataframe(df_plan_total, use_container_width=True)
    st.info(f"Total registros: {len(df_plan_total)}")
  else:
    st.info(
        "Sube tu archivo consolidado o nuevos reportes en la barra lateral."
    )

with tab2:
  st.subheader("Seguimiento de Ventas")
  if not df_ventas_total.empty:
    st.dataframe(df_ventas_total, use_container_width=True)
    st.info(f"Total registros: {len(df_ventas_total)}")
  else:
    st.info(
        "Sube tu archivo consolidado o nuevos reportes en la barra lateral."
    )

with tab3:
  st.subheader("Detalle de Auditoría")
  if not df_auditoria_total.empty:
    st.dataframe(df_auditoria_total, use_container_width=True)
    st.info(f"Total registros: {len(df_auditoria_total)}")
  else:
    st.info(
        "Sube tu archivo consolidado o nuevos reportes en la barra lateral."
    )
