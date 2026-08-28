import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def corregir_detalle_auditoria(filepath):
  wb = openpyxl.load_workbook(filepath)

  if "Detalle_Auditoria" in wb.sheetnames:
    ws = wb["Detalle_Auditoria"]

    # Estilos profesionales para la auditoría
    header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    border_style = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # Recorrer filas y celdas para asegurar formato unificado
    for row in ws.iter_rows(
        min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
    ):
      for cell in row:
        if cell.row == 1:
          cell.fill = header_fill
          cell.font = header_font
          cell.alignment = Alignment(
              horizontal="center", vertical="center", wrap_text=True
          )
        else:
          cell.font = data_font
          cell.border = border_style
          # Alineación condicional: números a la derecha, textos/fechas a la izquierda
          if isinstance(cell.value, (int, float)):
            cell.alignment = Alignment(horizontal="right", vertical="center")
          else:
            cell.alignment = Alignment(horizontal="left", vertical="center")

    # Auto-ajuste preciso del ancho de columnas para evitar columnas solapadas o truncadas
    for col in ws.columns:
      max_length = 0
      column_letter = get_column_letter(col[0].column)
      for cell in col:
        if cell.value is not None:
          max_length = max(max_length, len(str(cell.value)))
      ws.column_dimensions[column_letter].width = max(max_length + 5, 14)

    wb.save(filepath)


# Ejecutar la función sobre tu archivo generado
corregir_detalle_auditoria("Planificacion_Produccion_y_Ventas_Final_28-08-26.xlsx")
