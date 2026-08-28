from datetime import datetime
import io
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Compromisos y Producción", layout="wide"
)

st.title("📊 Panel de Control: Producción y Seguimiento Comercial")
st.markdown(
    "Consolidación automática manteniendo la estructura original del formato."
)

st.sidebar.header("📂 Gestión de Archivos")

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
df_ventas_total = pd.DataFrame()

# Leer el consolidado base si ya existe
if archivo_base is not None:
  try:
    xls_base = pd.ExcelFile(archivo_base)
    # Buscamos la hoja principal de ventas
    hoja_base = xls_base.sheet_names[0]
    df_ventas_total = pd.read_excel(xls_base, sheet_name=hoja_base, header=0)
  except Exception:
    st.sidebar.error("⚠️ El archivo consolidado base tiene un formato inválido.")

# Procesar los reportes individuales de los vendedores
if archivos_nuevos:
  for file in archivos_nuevos:
    try:
      xls = pd.ExcelFile(file)
      df_vendedor = pd.read_excel(xls, sheet_name=0, header=0).dropna(
          how="all"
      )

      # Validar que contenga columnas clave
      columnas_texto = " ".join([str(c) for c in df_vendedor.columns])
      if "Cotizacion" not in columnas_texto and "Numero" not in columnas_texto:
        st.sidebar.error(
            f"❌ El archivo '{file.name}' no tiene la estructura correcta en"
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
    except Exception:
      st.sidebar.error(
          f"❌ Error al procesar '{file.name}'. Revisa que comience en A1."
      )

# --- SECCIÓN DE DESCARGA DEL EXCEL CON LA ESTRUCTURA EXACTA ---
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
  st.info(f"Total de registros consolidados: {len(df_ventas_total)}")
else:
  st.info("Sube los reportes individuales de los vendedores en la barra lateral.")
