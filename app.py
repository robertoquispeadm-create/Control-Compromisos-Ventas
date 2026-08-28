from datetime import datetime
import io
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Compromisos y Producción", layout="wide"
)

st.title("📊 Panel de Control: Seguimiento Comercial")
st.markdown("Consolidación inteligente con preservación de historial comercial.")

st.sidebar.header("📂 Gestión de Archivos")

# 1. Subir el archivo consolidado base
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

if "df_auditoria" not in st.session_state:
  st.session_state["df_auditoria"] = pd.DataFrame()
if "df_ventas" not in st.session_state:
  st.session_state["df_ventas"] = pd.DataFrame()
if "wb_template" not in st.session_state:
  st.session_state["wb_template"] = None

# Cargar plantilla base
if archivo_base is not None:
  try:
    archivo_base.seek(0)
    st.session_state["wb_template"] = openpyxl.load_workbook(archivo_base)

    archivo_base.seek(0)
    xls_base = pd.ExcelFile(archivo_base)

    if "Seguimiento_Ventas" in xls_base.sheet_names:
      df_temp = pd.read_excel(xls_base, sheet_name="Seguimiento_Ventas", header=0)
      df_temp.columns = [str(c).strip() for c in df_temp.columns]
      st.session_state["df_ventas"] = df_temp.dropna(how="all")

    if "Detalle_Auditoria" in xls_base.sheet_names:
      df_temp = pd.read_excel(xls_base, sheet_name="Detalle_Auditoria", header=0)
      df_temp.columns = [str(c).strip() for c in df_temp.columns]
      df_temp = df_temp.dropna(subset=["Cliente"]) if "Cliente" in df_temp.columns else df_temp
      st.session_state["df_auditoria"] = df_temp

    st.sidebar.success("✅ ¡Plantilla base cargada con éxito!")
  except Exception as e:
    st.sidebar.error(f"⚠️ Error al leer el archivo base: {e}")

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

# --- PROCESAMIENTO, VALIDACIÓN Y GESTIÓN DE HISTORIAL ---
if archivos_nuevos:
  lista_nuevos_dfs = []
  
  columnas_obligatorias = [
      "Cliente", "No. Contacto", "1o. Contacto", "Ultimo Contacto", 
      "N° Cotización", "Unidad", "Cantidad", "Valor + IGV", 
      "Chance de Venta", "Mes Previsto"
  ]

  for file in archivos_nuevos:
    try:
      file.seek(0)
      df_v = pd.read_excel(file, sheet_name=0, header=0)
      df_v.columns = [str(c).strip() for c in df_v.columns]

      # 1. Validación estricta de columnas
      faltantes = [col for col in columnas_obligatorias if col not in df_v.columns]
      tiene_responsable = "CODIGO-RESPONSABLE" in df_v.columns or "RESPONSABLE" in df_v.columns
      if not tiene_responsable:
        faltantes.append("CODIGO-RESPONSABLE (o RESPONSABLE)")

      if faltantes:
        st.sidebar.error(
            f"❌ **Rechazado ('{file.name}')**: Le falta(n) la(s) columna(s): **{', '.join(faltantes)}**."
        )
        continue

      df_v = df_v.dropna(subset=["Cliente"])

      df_normalizado = pd.DataFrame()
      df_normalizado["Cliente"] = df_v["Cliente"]
      df_normalizado["No. Contacto"] = df_v["No. Contacto"]
      df_normalizado["1o. Contacto"] = df_v["1o. Contacto"]
      df_normalizado["Ultimo Contacto"] = df_v["Ultimo Contacto"]
      df_normalizado["N° Cotización"] = df_v["N° Cotización"]
      df_normalizado["Unidad"] = df_v["Unidad"]
      df_normalizado["Cantidad"] = pd.to_numeric(df_v["Cantidad"], errors="coerce").fillna(0)
      df_normalizado["Valor + IGV"] = pd.to_numeric(df_v["Valor + IGV"], errors="coerce").fillna(0)
      df_normalizado["Chance de Venta"] = pd.to_numeric(df_v["Chance de Venta"], errors="coerce").fillna(0)
      
      # Cálculos automáticos de ponderados
      df_normalizado["Valor Ponderado (S/)"] = df_normalizado["Valor + IGV"] * df_normalizado["Chance de Venta"]
      df_normalizado["Cantidad Ponderada Prod."] = df_normalizado["Cantidad"] * df_normalizado["Chance de Venta"]
      
      df_normalizado["Mes Previsto"] = df_v["Mes Previsto"]
      
      resp_col = "CODIGO-RESPONSABLE" if "CODIGO-RESPONSABLE" in df_v.columns else "RESPONSABLE"
      df_normalizado["RESPONSABLE"] = df_v[resp_col]

      lista_nuevos_dfs.append(df_normalizado)
      st.sidebar.success(f"✅ ¡Vendedor validado y procesado: {file.name}!")
    except Exception as e:
      st.sidebar.error(f"❌ Error al procesar '{file.name}': {e}")

  if lista_nuevos_dfs:
    df_agregado = pd.concat(lista_nuevos_dfs, ignore_index=True)
    
    if not st.session_state["df_auditoria"].empty:
      df_combinado = pd.concat([st.session_state["df_auditoria"], df_agregado], ignore_index=True)
    else:
      df_combinado = df_agregado

    # --- REGLA DE DUPLICADOS EXACTOS (Misma Cotización + Misma Fecha Contacto + Mismo Chance) ---
    # Si estas tres columnas son idénticas, se considera duplicado redundante y se elimina (keep="last").
    # Si cambia la fecha o el chance, se conservan ambos registros para formar parte del historial.
    df_combinado = df_combinado.drop_duplicates(
        subset=["N° Cotización", "Ultimo Contacto", "Chance de Venta"], 
        keep="last"
    )
    
    st.session_state["df_auditoria"] = df_combinado

# --- GENERACIÓN DEL EXCEL FINAL ---
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
      
      df_final = st.session_state["df_auditoria"].dropna(subset=["Cliente"])
      registros = df_final.to_dict("records")
      
      # Escritura exacta fila por fila (13 columnas)
      for r_idx, row in enumerate(registros, start=2):
        ws.cell(row=r_idx, column=1, value=row.get("Cliente", ""))
        ws.cell(row=r_idx, column=2, value=row.get("No. Contacto", ""))
        ws.cell(row=r_idx, column=3, value=row.get("1o. Contacto", ""))
        ws.cell(row=r_idx, column=4, value=row.get("Ultimo Contacto", ""))
        ws.cell(row=r_idx, column=5, value=row.get("N° Cotización", ""))
        ws.cell(row=r_idx, column=6, value=row.get("Unidad", ""))
        
        cant = row.get("Cantidad", 0)
        importe = row.get("Valor + IGV", 0)
        chance = row.get("Chance de Venta", 0)
        
        ws.cell(row=r_idx, column=7, value=cant)
        ws.cell(row=r_idx, column=8, value=importe)
        ws.cell(row=r_idx, column=9, value=chance)
        
        try: val_pond = float(importe) * float(chance)
        except: val_pond = 0
        try: cant_pond = float(cant) * float(chance)
        except: cant_pond = 0
            
        ws.cell(row=r_idx, column=10, value=val_pond)
        ws.cell(row=r_idx, column=11, value=cant_pond)
        
        ws.cell(row=r_idx, column=12, value=row.get("Mes Previsto", ""))
        ws.cell(row=r_idx, column=13, value=row.get("RESPONSABLE", ""))

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
      label="📥 Descargar Consolidado Final Alineado",
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
    
    df_aud["Cantidad_Num"] = pd.to_numeric(df_aud["Cantidad"], errors="coerce").fillna(0)
    df_aud["Importe_Num"] = pd.to_numeric(df_aud["Valor + IGV"], errors="coerce").fillna(0)
    df_aud["Chance Num"] = pd.to_numeric(df_aud["Chance de Venta"], errors="coerce").fillna(0)
    
    # --- FILTRAR CHANCE DE VENTA > 0% PARA EL RESUMEN ---
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
          Total_Cotizaciones=("N° Cotización", "count")
      ).reset_index()
      
      df_resumen_display = df_resumen.copy()
      df_resumen_display["Total_Cantidad"] = df_resumen_display["Total_Cantidad"].apply(lambda x: f"{x:,.2f}")
      df_resumen_display["Total_Valor_IGV"] = df_resumen_display["Total_Valor_IGV"].apply(lambda x: f"S/ {x:,.2f}")
      df_resumen_display["Total_Cotizaciones"] = df_resumen_display["Total_Cotizaciones"].astype(str)
      
      st.dataframe(df_resumen_display.style.set_properties(**{'text-align': 'left'}), use_container_width=True)
      
      total_general_cant = df_resumen["Total_Cantidad"].sum()
      total_general_val = df_resumen["Total_Valor_IGV"].sum()
      
      col1, col2 = st.columns(2)
      with col1:
        st.metric(label="📦 Cantidad Total Filtrada (Chance > 0%)", value=f"{total_general_cant:,.2f}")
      with col2:
        st.metric(label="💰 Valor Total con IGV (S/) Filtrado", value=f"S/ {total_general_val:,.2f}")
    else:
      st.warning("No hay registros con Chance de Venta mayor a 0%.")
  else:
    st.info("Sube tus archivos para visualizar el resumen.")

with tab2:
  st.subheader("Seguimiento de Rendimiento Comercial")
  if not st.session_state["df_ventas"].empty:
    st.dataframe(st.session_state["df_ventas"].style.set_properties(**{'text-align': 'left'}), use_container_width=True)
  else:
    st.info("Sube tu archivo consolidado base.")

with tab3:
  st.subheader("Detalle General y Auditoría de Cotizaciones")
  if not st.session_state["df_auditoria"].empty:
    st.dataframe(st.session_state["df_auditoria"].style.set_properties(**{'text-align': 'left'}), use_container_width=True)
    st.success(f"Registros totales en el historial de auditoría: {len(st.session_state['df_auditoria'])}")
  else:
    st.info("Sube tus archivos.")
