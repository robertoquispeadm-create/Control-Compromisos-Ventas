from datetime import datetime
import io
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Compromisos y Producción", layout="wide"
)

st.title("📊 Panel de Control: Producción y Seguimiento Comercial")
st.markdown("Consolidación inteligente con formato y alineación automática.")

st.sidebar.header("📂 Gestión de Archivos")

# 1. Subir el archivo consolidado base / modelo general (3 pestañas)
archivo_base = st.sidebar.file_uploader(
    "1. Sube tu Excel Consolidado o Modelo Base", type=["xlsx"], key="base"
)

# 2. Subir los reportes individuales semanales de los vendedores
archivos_nuevos = st.sidebar.file_uploader(
    "2. Sube los reportes individuales semanales de los vendedores",
    type=["xlsx"],
    accept_multiple_files=True,
    key="nuevos",
)

# Mantener estado en session_state
if "df_auditoria" not in st.session_state:
  st.session_state["df_auditoria"] = pd.DataFrame()
if "df_plan" not in st.session_state:
  st.session_state["df_plan"] = pd.DataFrame()
if "df_ventas" not in st.session_state:
  st.session_state["df_ventas"] = pd.DataFrame()
if "wb_template" not in st.session_state:
  st.session_state["wb_template"] = None

# Leer el archivo consolidado base subido por el usuario
if archivo_base is not None:
  try:
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

    st.sidebar.success("✅ ¡Plantilla base cargada con éxito!")
  except Exception as e:
    st.sidebar.error(f"⚠️ Error al leer el archivo base: {e}")

# --- BOTÓN AUTOMÁTICO DESDE GITHUB: DESCARGAR FORMATO DE VENDEDOR ---
nombre_archivo_modelo_github = (
    "Modelo - Control Compromisos para Ventas (FECHA).xlsx"
)
if os.path.exists(nombre_archivo_modelo_github):
  with open(nombre_archivo_modelo_github, "rb") as f:
    bytes_modelo = f.read()
  st.sidebar.download_button(
      label="📥 Descargar Formato Oficial de Vendedor",
      data=bytes_modelo,
      file_name="Formato_Oficial_Vendedor.xlsx",
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )

st.sidebar.markdown("---")

# --- VALIDACIÓN ESTRICTA Y ROBUSTA DE REPORTES INDIVIDUALES ---
if archivos_nuevos:
  for file in archivos_nuevos:
    try:
      file.seek(0)
      xls = pd.ExcelFile(file)

      df_fila1 = pd.read_excel(xls, sheet_name=0, header=1)
      df_fila2 = pd.read_excel(xls, sheet_name=0, header=2)

      columnas_totales = [str(c).strip() for c in df_fila1.columns] + [
          str(c).strip() for c in df_fila2.columns
      ]

      faltantes = []
      columnas_requeridas = ["CODIGO-RESPONSABLE", "Cotizacion", "Nombre"]

      for req in columnas_requeridas:
        encontrado = any(
            req.lower() in c.lower() for c in columnas_totales if c
        )
        if not encontrado:
          faltantes.append(req)

      if faltantes:
        st.sidebar.error(
            f"❌ **Anomalía en '{file.name}':** Le falta(n) la(s)"
            f" columna(s): **{', '.join(faltantes)}**."
        )
        continue

      df_vendedor = df_fila2.dropna(how="all")
      df_vendedor = df_vendedor.loc[
          :, ~df_vendedor.columns.str.contains("^Unnamed")
      ]
      
      # --- TRADUCCIÓN DE COLUMNAS PARA EVITAR DESCUADRES AL UNIR ---
      df_vendedor = df_vendedor.rename(columns={
          "Nombre": "Cliente",
          "Numero": "N° Cotización",
          "Cantid": "Cantidad Bruta",
          "Valor + IGV": "Importe Bruto (S/)",
          "Chance Venta": "Chance de Venta",
          "Previsto": "Mes Previsto"
      })
      
      df_vendedor["Cierre_Semanal"] = file.name

      if not st.session_state["df_auditoria"].empty:
        st.session_state["df_auditoria"] = pd.concat(
            [st.session_state["df_auditoria"], df_vendedor], ignore_index=True
        ).drop_duplicates()
      else:
        st.session_state["df_auditoria"] = df_vendedor

      st.sidebar.success(f"✅ ¡Vendedor procesado con éxito: {file.name}!")
    except Exception as e:
      st.sidebar.error(f"❌ Error al procesar '{file.name}': {e}")

# --- SECCIÓN DE DESCARGA DEL CONSOLIDADO FINAL CON FORMATO PERFECTO ---
if (
    not st.session_state["df_plan"].empty
    or not st.session_state["df_ventas"].empty
    or not st.session_state["df_auditoria"].empty
):
  st.sidebar.markdown("---")
  fecha_actual = datetime.now().strftime("%d-%m-%y")
  nombre_archivo = f"Planificacion_Produccion_y_Ventas_Final_{fecha_actual}.xlsx"

  output = io.BytesIO()
  if st.session_state["wb_template"] is not None:
    wb = st.session_state["wb_template"]
    if (
        "Detalle_Auditoria" in wb.sheetnames
        and not st.session_state["df_auditoria"].empty
    ):
      ws = wb["Detalle_Auditoria"]
      if ws.max_row >= 4:
        ws.delete_rows(4, ws.max_row - 3)
      
      # Llenar espacios vacíos para que no se exporte como 'nan'
      registros = st.session_state["df_auditoria"].fillna("").to_dict("records")
      
      for r_idx, row in enumerate(registros, start=4):
        # Mapeo usando los nombres estandarizados
        ws.cell(row=r_idx, column=1, value=row.get("Cliente", ""))
        ws.cell(row=r_idx, column=2, value=row.get("No. Contacto", ""))
        ws.cell(row=r_idx, column=3, value=row.get("1o. Contacto", ""))
        ws.cell(row=r_idx, column=4, value=row.get("Ultimo Contacto", ""))
        ws.cell(row=r_idx, column=5, value=row.get("N° Cotización", ""))
        ws.cell(row=r_idx, column=6, value=row.get("Unidad", ""))
        
        cant = row.get("Cantidad Bruta", 0)
        importe = row.get("Importe Bruto (S/)", 0)
        chance = row.get("Chance de Venta", 0)
        
        ws.cell(row=r_idx, column=7, value=cant if cant != "" else None)
        ws.cell(row=r_idx, column=8, value=importe if importe != "" else None)
        ws.cell(row=r_idx, column=9, value=chance if chance != "" else None)
        
        # Cálculo automático de los ponderados (Columna 10 y 11)
        try: val_pond = float(importe) * float(chance)
        except: val_pond = 0
        try: cant_pond = float(cant) * float(chance)
        except: cant_pond = 0
            
        ws.cell(row=r_idx, column=10, value=val_pond)
        ws.cell(row=r_idx, column=11, value=cant_pond)
        
        # Acomodar fecha y reporte en las últimas columnas
        ws.cell(row=r_idx, column=12, value=row.get("Mes Previsto", ""))
        ws.cell(row=r_idx, column=13, value=row.get("Cierre_Semanal", ""))

      # --- APLICAR CORRECCIÓN DE DISEÑO Y ALINEACIÓN EN MEMORIA ---
      data_font = Font(name="Calibri", size=10)
      border_style = Border(
          left=Side(style="thin", color="D9D9D9"),
          right=Side(style="thin", color="D9D9D9"),
          top=Side(style="thin", color="D9D9D9"),
          bottom=Side(style="thin", color="D9D9D9"),
      )

      for row_cells in ws.iter_rows(
          min_row=4, max_row=ws.max_row, min_col=1, max_col=ws.max_column
      ):
        for cell in row_cells:
          cell.font = data_font
          cell.border = border_style
          if isinstance(cell.value, (int, float)):
            cell.alignment = Alignment(horizontal="right", vertical="center")
          else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

      for col in ws.columns:
        max_length = 0
        column_letter = get_column_letter(col[0].column)
        for cell in col:
          if cell.value is not None:
            max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column_letter].width = max(max_length + 4, 14)

    wb.save(output)
    excel_data = output.getvalue()
  else:
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      if not st.session_state["df_auditoria"].empty:
        st.session_state["df_auditoria"].to_excel(
            writer, sheet_name="Detalle_Auditoria", index=False
        )
    excel_data = output.getvalue()

  st.sidebar.download_button(
      label="📥 Descargar Consolidado Final con Formato Original",
      data=excel_data,
      file_name=nombre_archivo,
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )

# --- VISUALIZACIÓN EN PESTAÑAS ---
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
