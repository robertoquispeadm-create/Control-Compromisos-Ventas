from datetime import datetime
import io
import os
import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Compromisos y Producción", layout="wide"
)

st.title("📊 Panel de Control: Seguimiento Comercial")
st.markdown("Consolidación inteligente con formato y alineación automática.")

st.sidebar.header("📂 Gestión de Archivos")

# 1. Subir el archivo consolidado base / modelo general
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

    if "Seguimiento_Ventas" in xls_base.sheet_names:
      df_temp = pd.read_excel(xls_base, sheet_name="Seguimiento_Ventas", header=0)
      df_temp.columns = [str(c).strip() for c in df_temp.columns]
      df_temp = df_temp.replace(r'^\s*$', float('nan'), regex=True).dropna(how="all")
      st.session_state["df_ventas"] = df_temp

    if "Detalle_Auditoria" in xls_base.sheet_names:
      df_temp = pd.read_excel(xls_base, sheet_name="Detalle_Auditoria", header=0)
      df_temp.columns = [str(c).strip() for c in df_temp.columns]
      df_temp = df_temp.replace(r'^\s*$', float('nan'), regex=True).dropna(how="all")
      if "Cliente" in df_temp.columns:
        df_temp = df_temp.dropna(subset=["Cliente"])
      st.session_state["df_auditoria"] = df_temp

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

# --- VALIDACIÓN Y LIMPIEZA ESTRICTA DE REPORTES INDIVIDUALES ---
if archivos_nuevos:
  for file in archivos_nuevos:
    try:
      file.seek(0)
      df_vendedor = pd.read_excel(file, sheet_name=0, header=0)
      df_vendedor.columns = [str(c).strip() for c in df_vendedor.columns]

      faltantes = []
      columnas_requeridas = ["Cliente", "N° Cotización", "CODIGO-RESPONSABLE"]

      for req in columnas_requeridas:
        encontrado = any(
            req.lower() in str(c).lower() for c in df_vendedor.columns if pd.notna(c)
        )
        if not encontrado:
          faltantes.append(req)

      if faltantes:
        st.sidebar.error(
            f"❌ **Anomalía en '{file.name}':** Le falta(n) la(s)"
            f" columna(s): **{', '.join(faltantes)}**."
        )
        continue

      # --- LIMPIEZA PROFUNDA DE FILAS VACÍAS O ESPACIOS EN BLANCO ---
      df_vendedor = df_vendedor.replace(r'^\s*$', float('nan'), regex=True)
      df_vendedor = df_vendedor.dropna(how="all")
      if "Cliente" in df_vendedor.columns:
        df_vendedor = df_vendedor.dropna(subset=["Cliente"])

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

      st.sidebar.success(f"✅ ¡Vendedor procesado con éxito: {file.name}!")
    except Exception as e:
      st.sidebar.error(f"❌ Error al procesar '{file.name}': {e}")

# --- SECCIÓN DE DESCARGA DEL CONSOLIDADO FINAL ---
if (
    not st.session_state["df_ventas"].empty
    or not st.session_state["df_auditoria"].empty
):
  st.sidebar.markdown("---")
  fecha_actual = datetime.now().strftime("%d-%m-%y")
  nombre_archivo = f"Control_Compromisos_Ventas_Final_{fecha_actual}.xlsx"

  output = io.BytesIO()
  if st.session_state["wb_template"] is not None:
    wb = st.session_state["wb_template"]
    if (
        "Detalle_Auditoria" in wb.sheetnames
        and not st.session_state["df_auditoria"].empty
    ):
      ws = wb["Detalle_Auditoria"]
      if ws.max_row >= 2:
        ws.delete_rows(2, ws.max_row - 1)
      
      df_final = st.session_state["df_auditoria"].replace(r'^\s*$', float('nan'), regex=True)
      df_final = df_final.dropna(how="all")
      if "Cliente" in df_final.columns:
        df_final = df_final.dropna(subset=["Cliente"])
      
      registros = df_final.fillna("").to_dict("records")
      
      for r_idx, row in enumerate(registros, start=2):
        ws.cell(row=r_idx, column=1, value=row.get("Cliente", ""))
        ws.cell(row=r_idx, column=2, value=row.get("No. Contacto", ""))
        ws.cell(row=r_idx, column=3, value=row.get("1o. Contacto", ""))
        ws.cell(row=r_idx, column=4, value=row.get("Ultimo Contacto", ""))
        ws.cell(row=r_idx, column=5, value=row.get("N° Cotización", ""))
        ws.cell(row=r_idx, column=6, value=row.get("Unidad", ""))
        
        cant = row.get("Cantidad", row.get("Cantidad Bruta", 0))
        importe = row.get("Valor + IGV", row.get("Importe Bruto (S/)", 0))
        chance = row.get("Chance de Venta", 0)
        
        ws.cell(row=r_idx, column=7, value=cant if cant != "" else None)
        ws.cell(row=r_idx, column=8, value=importe if importe != "" else None)
        ws.cell(row=r_idx, column=9, value=chance if chance != "" else None)
        
        try: val_pond = float(importe) * float(chance)
        except: val_pond = 0
        try: cant_pond = float(cant) * float(chance)
        except: cant_pond = 0
            
        ws.cell(row=r_idx, column=10, value=val_pond)
        ws.cell(row=r_idx, column=11, value=cant_pond)
        
        ws.cell(row=r_idx, column=12, value=row.get("Mes Previsto", ""))
        ws.cell(row=r_idx, column=13, value=row.get("Cierre_Semanal", ""))

      data_font = Font(name="Calibri", size=10)
      border_style = Border(
          left=Side(style="thin", color="D9D9D9"),
          right=Side(style="thin", color="D9D9D9"),
          top=Side(style="thin", color="D9D9D9"),
          bottom=Side(style="thin", color="D9D9D9"),
      )

      for row_cells in ws.iter_rows(
          min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column
      ):
        for cell in row_cells:
          cell.font = data_font
          cell.border = border_style
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
      label="📥 Descargar Consolidado Final sin Filas Vacías",
      data=excel_data,
      file_name=nombre_archivo,
      mime=(
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      ),
  )

# --- VISUALIZACIÓN EN PESTAÑAS ---
tab1, tab2, tab3 = st.tabs(
    [
        "📦 Resumen por Mes y Unidad",
        "💰 Seguimiento Ventas",
        "📋 Detalle de Auditoría",
    ]
)

with tab1:
  st.subheader("Resumen Proyectado: Cantidades y Valores por Unidad (Chance > 0%)")
  if not st.session_state["df_auditoria"].empty:
    df_aud = st.session_state["df_auditoria"].copy()
    
    col_cantidad = "Cantidad" if "Cantidad" in df_aud.columns else "Cantidad Bruta"
    col_importe = "Valor + IGV" if "Valor + IGV" in df_aud.columns else "Importe Bruto (S/)"
    
    df_aud["Cantidad_Num"] = pd.to_numeric(df_aud[col_cantidad], errors="coerce").fillna(0) if col_cantidad in df_aud.columns else 0
    df_aud["Importe_Num"] = pd.to_numeric(df_aud[col_importe], errors="coerce").fillna(0) if col_importe in df_aud.columns else 0
    df_aud["Chance Num"] = pd.to_numeric(df_aud["Chance de Venta"], errors="coerce").fillna(0) if "Chance de Venta" in df_aud.columns else 0
    
    # --- FILTRAR CHANCE DE VENTA > 0% ---
    df_aud = df_aud[df_aud["Chance Num"] > 0]
    
    if not df_aud.empty and "Mes Previsto" in df_aud.columns and "Unidad" in df_aud.columns:
      def formatear_mes_yyyy_mm(val):
        if pd.isna(val):
          return "Sin Especificar"
        try:
          dt = pd.to_datetime(val)
          return dt.strftime("%Y-%m")
        except:
          s = str(val).strip()
          return s[:7] if len(s) >= 7 else s

      df_aud["Mes Formateado"] = df_aud["Mes Previsto"].apply(formatear_mes_yyyy_mm)
      
      meses_disponibles = sorted([str(m) for m in df_aud["Mes Formateado"].unique()])
      mes_actual_default = datetime.now().strftime("%Y-%m")
      defaults = [mes_actual_default] if mes_actual_default in meses_disponibles else meses_disponibles
      
      st.markdown("### 🎛️ Filtros Interactivos")
      meses_seleccionados = st.multiselect(
          "Selecciona el Mes Previsto (YYYY-MM):",
          options=meses_disponibles,
          default=defaults if defaults else meses_disponibles
      )
      
      if meses_seleccionados:
        df_filtrado = df_aud[df_aud["Mes Formateado"].isin(meses_seleccionados)]
      else:
        df_filtrado = df_aud.copy()

      df_resumen = df_filtrado.groupby(["Mes Formateado", "Unidad"], dropna=False).agg(
          Total_Cantidad=("Cantidad_Num", "sum"),
          Total_Valor_IGV=("Importe_Num", "sum"),
          Total_Cotizaciones=("N° Cotización", "count") if "N° Cotización" in df_filtrado.columns else ("Cliente", "count")
      ).reset_index()
      
      df_resumen_display = df_resumen.copy()
      df_resumen_display["Total_Cantidad"] = df_resumen_display["Total_Cantidad"].apply(lambda x: f"{x:,.2f}")
      df_resumen_display["Total_Valor_IGV"] = df_resumen_display["Total_Valor_IGV"].apply(lambda x: f"S/ {x:,.2f}")
      df_resumen_display["Total_Cotizaciones"] = df_resumen_display["Total_Cotizaciones"].astype(str)
      
      df_styled = df_resumen_display.style.set_properties(**{'text-align': 'left'})
      st.dataframe(df_styled, use_container_width=True)
      
      total_general_cant = df_resumen["Total_Cantidad"].sum()
      total_general_val = df_resumen["Total_Valor_IGV"].sum()
      
      col1, col2 = st.columns(2)
      with col1:
        st.metric(label="📦 Cantidad Total Filtrada (Chance > 0%)", value=f"{total_general_cant:,.2f}")
      with col2:
        st.metric(label="💰 Valor Total con IGV (S/) Filtrado", value=f"S/ {total_general_val:,.2f}")
    else:
      st.warning("No hay registros con Chance de Venta mayor a 0% o faltan columnas requeridas.")
  else:
    st.info("Sube tus reportes individuales o archivo base para visualizar el resumen.")

with tab2:
  st.subheader("Seguimiento de Rendimiento Comercial")
  if not st.session_state["df_ventas"].empty:
    df_ventas_styled = st.session_state["df_ventas"].style.set_properties(**{'text-align': 'left'})
    st.dataframe(df_ventas_styled, use_container_width=True)
  else:
    st.info("Sube tu archivo consolidado base.")

with tab3:
  st.subheader("Detalle General y Auditoría de Cotizaciones")
  if not st.session_state["df_auditoria"].empty:
    df_auditoria_styled = st.session_state["df_auditoria"].style.set_properties(**{'text-align': 'left'})
    st.dataframe(df_auditoria_styled, use_container_width=True)
    st.success(
        f"Registros totales en auditoría: {len(st.session_state['df_auditoria'])}"
    )
  else:
    st.info("Sube tu archivo consolidado o reportes individuales.")
