import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Compromisos y Producción", layout="wide"
)

st.title("📊 Panel de Control: Producción y Seguimiento Comercial")
st.markdown(
    "Sistema en la nube para la lectura de reportes semanales, visualización de"
    " bases de datos y control de responsables."
)

st.sidebar.header("📂 Carga de Archivos")
uploaded_file = st.sidebar.file_uploader(
    "Sube tu archivo Excel (.xlsx)", type=["xlsx"]
)

if uploaded_file is not None:
  try:
    # Leer las 3 pestañas principales del excel respetando la estructura
    df_plan = pd.read_excel(
        uploaded_file, sheet_name="Planificacion_Produccion", skiprows=2
    )
    df_ventas = pd.read_excel(
        uploaded_file, sheet_name="Seguimiento_Ventas", skiprows=2
    )
    df_auditoria = pd.read_excel(
        uploaded_file, sheet_name="Detalle_Auditoria", skiprows=2
    )

    # Limpiar filas completamente vacías
    df_plan = df_plan.dropna(how="all")
    df_ventas = df_ventas.dropna(how="all")
    df_auditoria = df_auditoria.dropna(how="all")

    st.success("¡Archivo cargado y procesado exitosamente desde la nube!")

    # Pestañas de visualización en la app web
    tab1, tab2, tab3 = st.tabs(
        [
            "📈 Planificación de Producción",
            "💰 Seguimiento de Ventas",
            "📋 Detalle de Auditoría",
        ]
    )

    with tab1:
      st.subheader("Planificación y Requerimiento de Producción")
      st.dataframe(df_plan, use_container_width=True)

    with tab2:
      st.subheader("Seguimiento de Rendimiento Comercial")
      st.dataframe(df_ventas, use_container_width=True)

    with tab3:
      st.subheader("Detalle General de Cotizaciones y Auditoría")
      st.dataframe(df_auditoria, use_container_width=True)

  except Exception as e:
    st.error(
        f"Ocurrió un error al procesar el archivo Excel. Detalle técnico: {e}"
    )
else:
  st.info(
      "👈 Por favor, sube tu archivo Excel en la barra lateral para comenzar a"
      " visualizar el panel con sus respectivas columnas de responsables."
  )
