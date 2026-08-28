import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import datetime, timedelta
import os

st.set_page_config(page_title="Control de Compromisos y Producción", layout="wide")

st.title("📊 Panel de Control: Producción y Seguimiento Comercial")
st.markdown("Sistema interconectado para la lectura de reportes semanales, cálculo de requerimiento de producción y auditoría de ventas.")

# --- CONFIGURACIÓN DE CARPETA (LOCAL O COMPARTIDA) ---
# Puedes cambiar esta ruta a la carpeta sincronizada de OneDrive en tu PC (ej. C:/Users/TuUsuario/OneDrive/CarpetaCompartida/)
DEFAULT_PATH = "./" 

st.sidebar.header("📂 Origen de Datos")
folder_path = st.sidebar.text_input("Ruta de la carpeta compartida o local:", value=DEFAULT_PATH)

# Botón para buscar archivos automáticamente en la carpeta
if st.sidebar.button("🔄 Cargar y Actualizar desde Carpeta"):
    try:
        files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xlsx') and 'Control Compromisos' in f]
        
        if not files:
            # Si no encuentra por nombre específico, toma todos los xlsx de cotizaciones
            files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xlsx') and not f.startswith('Reporte_')]
        
        all_data = []
        for file_path in files:
            file_name = os.path.basename(file_path)
            match = re.search(r'al (\d{2}-\d{2}-\d{4})', file_name)
            if match:
                date_str = match.group(1)
                file_date = datetime.strptime(date_str, "%d-%m-%Y")
                days_to_friday = 4 - file_date.weekday()
                friday_date = file_date + timedelta(days=days_to_friday)
            else:
                file_date = datetime.now()
                friday_date = file_date
                
            df = pd.read_excel(file_path, header=None)
            header_row_idx = None
            for i, row in df.head(15).iterrows():
                row_str = " ".join(str(x).lower() for x in row.tolist())
                if "nombre" in row_str and "cantid" in row_str:
                    header_row_idx = i
                    break
                    
            if header_row_idx is not None:
                headers = df.iloc[header_row_idx].tolist()
                clean_headers = [str(h).strip() if not pd.isna(h) else f"Unnamed_{j}" for j, h in enumerate(headers)]
                data = df.iloc[header_row_idx+1:].copy()
                data.columns = clean_headers
                
                col_mapping = {}
                for col in clean_headers:
                    col_lower = col.lower()
                    if "nombre" in col_lower: col_mapping[col] = "Cliente"
                    elif "cantid" in col_lower: col_mapping[col] = "Cantidad"
                    elif "valor" in col_lower or "igv" in col_lower: col_mapping[col] = "Importe"
                    elif "unidad" in col_lower: col_mapping[col] = "Unidad"
                    elif "chance" in col_lower or "porcentual" in col_lower: col_mapping[col] = "Chance_Venta"
                    elif "numero" in col_lower or "cotizacion" in col_mapping if 'col_mapping' in locals() else col.lower() == 'cotizacion': col_mapping[col] = "Cotizacion"
                    elif "cotizacion" in col_lower or "numero" in col_lower: col_mapping[col] = "Cotizacion"
                
                data = data.rename(columns=col_mapping)
                if 'Cliente' in data.columns:
                    data = data.dropna(subset=['Cliente'])
                    data = data[data['Cliente'].astype(str).str.len() > 2]
                    data['Archivo_Origen'] = file_name
                    data['Fecha_Reporte'] = file_date
                    data['Viernes_Semana'] = friday_date
                    all_data.append(data)

        if all_data:
            df_final = pd.concat(all_data, ignore_index=True)
            df_final['Cantidad'] = pd.to_numeric(df_final['Cantidad'], errors='coerce').fillna(0)
            df_final['Importe'] = pd.to_numeric(df_final['Importe'], errors='coerce').fillna(0)
            df_final['Chance_Venta'] = pd.to_numeric(df_final['Chance_Venta'], errors='coerce').fillna(0)
            df_final['Valor_Ponderado'] = df_final['Importe'] * df_final['Chance_Venta']
            df_final['Cantidad_Ponderada_Prod'] = df_final['Cantidad'] * df_final['Chance_Venta']

            st.session_state['df_final'] = df_final
            st.sidebar.success(f"¡Se procesaron {len(files)} archivos correctamente!")
        else:
            st.sidebar.error("No se encontraron registros válidos en los archivos de la ruta.")
    except Exception as e:
        st.sidebar.error(f"Error al leer la carpeta: {e}")

# --- PANTALLA PRINCIPAL SI EXISTEN DATOS ---
if 'df_final' in st.session_state:
    df_final = st.session_state['df_final']

    tab1, tab2, tab3 = st.tabs(["🏭 Planificación de Producción", "📈 Seguimiento Comercial", "📋 Auditoría Detallada"])

    with tab1:
        st.subheader("Requerimiento Realista de Producción (Bolsas)")
        df_final['Mes_Previsto_Str'] = pd.to_datetime(df_final['Previsto'], errors='coerce').dt.strftime('%Y-%m')
        
        prod_summary = df_final[df_final['Unidad'] == 'Bolsas'].groupby('Mes_Previsto_Str').agg(
            Cotizaciones=('Cotizacion', 'count'),
            Cantidad_Bruta=('Cantidad', 'sum'),
            Cantidad_Realista_Producir=('Cantidad_Ponderada_Prod', 'sum')
        ).reset_index()
        
        prod_summary['Ratio_Conversion'] = prod_summary['Cantidad_Realista_Producir'] / prod_summary['Cantidad_Bruta']
        
        st.dataframe(prod_summary.style.format({
            'Cantidad_Bruta': '{:,.0f}',
            'Cantidad_Realista_Producir': '{:,.0f}',
            'Ratio_Conversion': '{:.1%}'
        }), use_container_width=True)

    with tab2:
        st.subheader("Rendimiento del Pipeline de Ventas por Semana")
        sales_summary = df_final.groupby('Viernes_Semana').agg(
            Cotizaciones=('Cotizacion', 'count'),
            Pipeline_Bruto=('Importe', 'sum'),
            Pipeline_Realista=('Valor_Ponderado', 'sum'),
            Chance_Promedio=('Chance_Venta', 'mean')
        ).reset_index()
        
        sales_summary['Efectividad'] = sales_summary['Pipeline_Realista'] / sales_summary['Pipeline_Bruto']
        
        st.dataframe(sales_summary.style.format({
            'Pipeline_Bruto': 'S/ {:,.2f}',
            'Pipeline_Realista': 'S/ {:,.2f}',
            'Chance_Promedio': '{:.1%}',
            'Efectividad': '{:.1%}'
        }), use_container_width=True)

    with tab3:
        st.subheader("Base Consolidada de Cotizaciones")
        st.dataframe(df_final[['Cliente', 'No. Contacto', 'Ultimo Contacto', 'Cotizacion', 'Unidad', 'Cantidad', 'Importe', 'Chance_Venta', 'Valor_Ponderado', 'Previsto']], use_container_width=True)

    # Botón para exportar el consolidado directamente a la carpeta compartida o descarga
    st.sidebar.markdown("---")
    output_filename = "Reporte_Consolidado_Actualizado.xlsx"
    if st.sidebar.button("💾 Guardar Consolidado Actualizado"):
        df_final.to_excel(os.path.join(folder_path, output_filename), index=False)
        st.sidebar.success(f"Archivo guardado con éxito como: {output_filename}")

else:
    st.info("👈 Ingresa la ruta de tu carpeta compartida en la barra lateral y presiona **Cargar y Actualizar** para visualizar el panel interactivo.")
