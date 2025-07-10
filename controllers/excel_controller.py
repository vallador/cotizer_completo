# controllers/excel_controller.py
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
import os
from datetime import datetime


class ExcelController:
    """
    Controlador responsable de generar los archivos de cotización en formato Excel.
    """

    def __init__(self, cotizacion_controller, aiu_manager):
        """
        Inicializa el controlador de Excel.

        Args:
            cotizacion_controller: Instancia del controlador principal.
            aiu_manager: Instancia del gestor de valores AIU.
        """
        self.cotizacion_controller = cotizacion_controller
        self.aiu_manager = aiu_manager

    def generate_excel(self, activities, tipo_persona, administracion, imprevistos, utilidad, iva_utilidad):
        """
        Genera un archivo Excel de cotización con capítulos organizados y coloreados.
        Este es ahora un MÉTODO de la clase.
        """
        # Crear un nuevo archivo Excel
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Cotización"

        # --- Estilos (sin cambios) ---
        fill_color = PatternFill(start_color='008000', end_color='008000', fill_type='solid')
        thick_border = Side(border_style="thick", color="000000")
        thin_border = Side(border_style="thin", color="000000")
        chapter_font = Font(bold=True, color="FFFFFF")
        chapter_fill = PatternFill(start_color='008000', end_color='008000', fill_type='solid')
        chapter_alignment = Alignment(horizontal="center", vertical="center")
        chapter_activity_colors = {
            1: "C8E6C9", 2: "BBDEFB", 3: "DCEDC8", 4: "FFE0B2", 5: "E1BEE7",
            6: "FFF9C4", 7: "B2EBF2", 8: "F8BBD9", 9: "F0F4C3", 10: "D1C4E9"
        }

        # --- Anchos de columna (sin cambios) ---
        sheet.column_dimensions['A'].width = 5
        sheet.column_dimensions['B'].width = 50
        sheet.column_dimensions['C'].width = 10
        sheet.column_dimensions['D'].width = 10
        sheet.column_dimensions['E'].width = 15
        sheet.column_dimensions['F'].width = 15

        # --- Encabezados (sin cambios) ---
        headers = ["Item", "Descripción", "Cantidad", "Unidad", "Precio Unitario", "Total"]
        for col_num, header in enumerate(headers, start=1):
            col_letter = get_column_letter(col_num)
            cell = sheet[f"{col_letter}1"]
            cell.value = header
            cell.font = Font(bold=True, color="FFFFFF")  # Letra blanca
            cell.alignment = Alignment(horizontal="center")
            cell.fill = fill_color  # Fondo verde

        # --- Organizar actividades por capítulos (sin cambios) ---
        capitulos = {}
        actividades_sin_capitulo = []
        for activity in activities:
            chapter_id = activity.get('chapter_id')
            chapter_name = activity.get('chapter_name')
            if chapter_id and chapter_name:
                if chapter_id not in capitulos:
                    capitulos[chapter_id] = {'nombre': chapter_name, 'actividades': []}
                capitulos[chapter_id]['actividades'].append(activity)
            else:
                actividades_sin_capitulo.append(activity)

        # --- Insertar actividades en las filas (sin cambios) ---
        row_num = 2
        item_num = 1

        # Función auxiliar para insertar filas y aplicar formato
        def _insert_activity_row(activity, fill=None):
            nonlocal row_num, item_num
            sheet[f"A{row_num}"].value = item_num
            sheet[f"B{row_num}"].value = activity['descripcion']
            sheet[f"C{row_num}"].value = activity['cantidad']
            sheet[f"D{row_num}"].value = activity['unidad']
            sheet[f"E{row_num}"].value = activity['valor_unitario']
            sheet[f"F{row_num}"].value = f"=C{row_num}*E{row_num}"

            for col in range(1, 7):
                cell = sheet.cell(row=row_num, column=col)
                if fill:
                    cell.fill = fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            sheet[f"B{row_num}"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            sheet[f"E{row_num}"].number_format = '"$"#,##0.00'
            sheet[f"F{row_num}"].number_format = '"$"#,##0.00'
            row_num += 1
            item_num += 1

        # Insertar actividades sin capítulo
        for activity in actividades_sin_capitulo:
            _insert_activity_row(activity)

        # Insertar actividades por capítulos
        for chapter_id, capitulo_data in sorted(capitulos.items()):
            sheet.merge_cells(f"A{row_num}:F{row_num}")
            cell = sheet[f"A{row_num}"]
            cell.value = capitulo_data['nombre'].upper()
            cell.font = chapter_font
            cell.alignment = chapter_alignment
            cell.fill = chapter_fill
            row_num += 1

            activity_fill = PatternFill(start_color=chapter_activity_colors.get(chapter_id, "FFFFFF"),
                                        fill_type='solid')
            for activity in capitulo_data['actividades']:
                _insert_activity_row(activity, fill=activity_fill)

        # --- Calcular totales (sin cambios, pero con formato mejorado) ---
        total_row = row_num
        subtotal_cell = f"F{total_row}"
        sheet.merge_cells(f"A{total_row}:E{total_row}")
        sheet[f"A{total_row}"].value = "TOTAL COSTOS DIRECTOS DE OBRA"
        sheet[f"A{total_row}"].font = Font(bold=True)
        sheet[f"A{total_row}"].alignment = Alignment(horizontal="right")
        sheet[subtotal_cell].value = f"=SUM(F2:F{total_row - 1})"
        sheet[subtotal_cell].number_format = '"$"#,##0.00'
        sheet[subtotal_cell].font = Font(bold=True)
        row_num += 1

        if tipo_persona == "jurídica":
            sheet.merge_cells(f"A{row_num}:E{row_num}");
            sheet[f"A{row_num}"].value = f"ADMINISTRACIÓN ({administracion}%)";
            sheet[f"F{row_num}"].value = f"={subtotal_cell}*{administracion / 100}";
            sheet[f"F{row_num}"].number_format = '"$"#,##0.00';
            sheet[f"A{row_num}"].alignment = Alignment(horizontal="right");
            row_num += 1
            sheet.merge_cells(f"A{row_num}:E{row_num}");
            sheet[f"A{row_num}"].value = f"IMPREVISTOS ({imprevistos}%)";
            sheet[f"F{row_num}"].value = f"={subtotal_cell}*{imprevistos / 100}";
            sheet[f"F{row_num}"].number_format = '"$"#,##0.00';
            sheet[f"A{row_num}"].alignment = Alignment(horizontal="right");
            row_num += 1
            util_cell = f"F{row_num}";
            sheet.merge_cells(f"A{row_num}:E{row_num}");
            sheet[f"A{row_num}"].value = f"UTILIDAD ({utilidad}%)";
            sheet[f"F{row_num}"].value = f"={subtotal_cell}*{utilidad / 100}";
            sheet[f"F{row_num}"].number_format = '"$"#,##0.00';
            sheet[f"A{row_num}"].alignment = Alignment(horizontal="right");
            row_num += 1
            sheet.merge_cells(f"A{row_num}:E{row_num}");
            sheet[f"A{row_num}"].value = f"IVA SOBRE UTILIDAD ({iva_utilidad}%)";
            sheet[f"F{row_num}"].value = f"={util_cell}*{iva_utilidad / 100}";
            sheet[f"F{row_num}"].number_format = '"$"#,##0.00';
            sheet[f"A{row_num}"].alignment = Alignment(horizontal="right");
            row_num += 1
        else:  # Persona Natural
            sheet.merge_cells(f"A{row_num}:E{row_num}");
            sheet[f"A{row_num}"].value = f"IVA ({iva_utilidad}%)";
            sheet[f"F{row_num}"].value = f"={subtotal_cell}*{iva_utilidad / 100}";
            sheet[f"F{row_num}"].number_format = '"$"#,##0.00';
            sheet[f"A{row_num}"].alignment = Alignment(horizontal="right");
            row_num += 1

        sheet.merge_cells(f"A{row_num}:E{row_num}")
        sheet[f"A{row_num}"].value = "VALOR TOTAL COTIZACIÓN"
        sheet[f"A{row_num}"].font = Font(bold=True)
        sheet[f"A{row_num}"].alignment = Alignment(horizontal="right")
        sheet[f"F{row_num}"].value = f"=SUM(F{total_row}:F{row_num - 1})"
        sheet[f"F{row_num}"].number_format = '"$"#,##0.00'
        sheet[f"F{row_num}"].font = Font(bold=True)

        # --- Bordes (sin cambios) ---
        for row in sheet.iter_rows(min_row=1, max_row=row_num, min_col=1, max_col=6):
            for cell in row:
                cell.border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)

        # --- Guardar archivo (sin cambios) ---
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        # Guardar en una carpeta 'output' para mantener el orden
        output_dir = os.path.join(os.getcwd(), "output")
        os.makedirs(output_dir, exist_ok=True)
        excel_path = os.path.join(output_dir, f"cotizacion_{timestamp}.xlsx")

        workbook.save(excel_path)
        workbook.close()

        print(f"Archivo Excel guardado en: {excel_path}")
        return excel_path
