from datetime import datetime
import io
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Compromisos y Producción", layout="wide"
)

st.title("📊 Panel de Control: Producción y Seguimiento Comercial")
st.markdown("Consolidación inteligente manteniendo el formato original.")

st.sidebar.header("📂 Gestión de Archivos")

# 1. Subir el archivo modelo o consolidado actual
archivo_base = st.sidebar.file_uploader(
    "1. Sube tu Excel Consolidado o Plantilla Modelo",
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

# Inicializar DataFrames y estructura base
df_ventas_total = pd.DataFrame()
columnas_modelo = None

# Leer el archivo base o modelo si se sube
if archivo_base is not None:
  try:
    xls_base = pd.ExcelFile(archivo_base)
    hoja_base = xls_base.sheet_names[0]
    df_ventas_total = pd.read_excel(xls_base, sheet_name=hoja_base, header=0)
    columnas_modelo = (
        df_ventas_total.columns.tolist()
    )  # Guardar las columnas exactas del modelo
  except Exception as e:
    st.sidebar.error(f"⚠️ Error al leer el archivo base: {e}")

# Procesar los reportes individuales de los vendedores
if archivos_nuevos:
  for file in archivos_nuevos:
    try:
      xls = pd.ExcelFile(file)
      df_vendedor = pd.read_excel(xls, sheet_name=0, header=0).dropna(
          how="all"
      )

      # Validar columnas esenciales
      columnas_texto = " ".join([str(c) for c in df_vendedor.columns])
      if "Cotizacion" not in columnas_texto and "Numero" not in columnas_texto:
        st.sidebar.error(
            f"❌ El archivo '{file.name}' no tiene la columna de Cotización en"
            " A1."
        )
        continue

      df_vendedor["Cierre_Semanal"] = file.name

      if not df_ventas_total.empty:
        df_ventas_total = pd.concat(
            [df_ventas_total, df_vendedor], ignore_index=True
        ).drop_duplicates()
      else:
        df_ventas_total = df_vendedor

      st.sidebar.success(f"✅ ¡Procesado correctamente: {file.name}!")
    except Exception as e:
      st.sidebar.error(f"❌ Error al procesar '{file.name}': {e}")

# --- SECCIÓN DE DESCARGA DEL EXCEL CONSOLIDADO FINAL ---
if not df_ventas_total.empty:
  st.sidebar.markdown("---")

  fecha_actual = datetime.now().strftime("%d-%m-%y")
  nombre_archivo = f"Control_Compromisos_Consolidado_{fecha_actual}.xlsx"

  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_ventas_total.to_excel(writer, sheet_name="Hoja1", index=False)
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
st.subheader("💰 Seguimiento de Ventas Consolidado")

if not df_ventas_total.empty:
  st.dataframe(df_ventas_total, use_container_width=True)
  # Mostrar métrica informativa de registros y columnas detectadas
  st.success(
      f"📊 Se encontraron **{len(df_ventas_total)}** registros y"
      f" **{len(df_ventas_total.columns)}** columnas cargadas"
      " correctamente."
  )
else:
  st.info(
      "📌 Sube tu plantilla modelo o archivo base arriba, y luego añade los"
      " reportes semanales de los vendedores en la opción 2 para ver el"
      " resultado consolidado."
  )
