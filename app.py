# --- VALIDACIÓN ESTRICTA Y ROBUSTA DE REPORTES INDIVIDUALES ---
if archivos_nuevos:
  for file in archivos_nuevos:
    try:
      file.seek(0)
      xls = pd.ExcelFile(file)

      # Leer la fila 1 y fila 2 para capturar todas las cabeceras reales del formato
      df_fila1 = pd.read_excel(xls, sheet_name=0, header=1)
      df_fila2 = pd.read_excel(xls, sheet_name=0, header=2)

      # Unir todos los nombres de columnas de ambas filas en una sola lista de texto limpio
      columnas_totales = [str(c).strip() for c in df_fila1.columns] + [
          str(c).strip() for c in df_fila2.columns
      ]

      # Buscar si están las columnas clave en todo el conjunto de cabeceras
      faltantes = []
      columnas_requeridas = ["CODIGO-RESPONSABLE", "Cotizacion", "Nombre"]

      for req in columnas_requeridas:
        # Verificamos si alguna columna contiene el texto requerido
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

      # Procesar el dataframe usando header=2 para los datos correctos
      df_vendedor = df_fila2.dropna(how="all")
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
