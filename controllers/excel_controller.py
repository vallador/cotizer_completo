# En controllers/excel_controller.py

import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
from openpyxl.utils import get_column_letter
from datetime import datetime
import os


class ExcelController:
    def __init__(self, cotizacion_controller, aiu_manager):
        self.cotizacion_controller = cotizacion_controller
        self.aiu_manager = aiu_manager

    def aplicar_bordes_totales(self, sheet, start_row, end_row, color_hex):
        """Aplica bordes exteriores gruesos y bordes interiores delgados a un bloque de celdas."""
        thin = Side(border_style="thin", color="000000")
        medium = Side(border_style="medium", color="000000")

        fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

        for row in range(start_row, end_row + 1):
            for col in range(1, 7):  # Columnas A (1) a F (6)
                cell = sheet.cell(row=row, column=col)
                cell.fill = fill

                # Determinar bordes según posición
                left = medium if col == 1 else thin
                right = medium if col == 6 else thin
                top = medium if row == start_row else thin
                bottom = medium if row == end_row else thin

                cell.border = Border(left=left, right=right, top=top, bottom=bottom)

    def aplicar_bordes_totales_con_marco_grueso(self, sheet, start_row, end_row, start_col=1, end_col=6,
                                                color_hex="D9EAD3"):
        """
        Aplica bordes delgados internos y un marco grueso alrededor de un bloque rectangular.
        - start_row, end_row: filas del bloque.
        - start_col, end_col: columnas del bloque (por defecto A:F → 1:6).
        """
        thin = Side(border_style="thin", color="000000")
        medium = Side(border_style="medium", color="000000")
        fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row=row, column=col)
                cell.fill = fill

                # Determinar bordes por posición (marco grueso)
                left = medium if col == start_col else thin
                right = medium if col == end_col else thin
                top = medium if row == start_row else thin
                bottom = medium if row == end_row else thin

                cell.border = Border(left=left, right=right, top=top, bottom=bottom)
    def generate_excel(self,activities, items, tipo_persona, administracion, imprevistos, utilidad, iva_utilidad):
        """
        Genera un archivo Excel de cotización profesional, manejando capítulos,
        formato de celdas y lógica de AIU/IVA.
        """
        # 1. --- CONFIGURACIÓN INICIAL DEL LIBRO Y LA HOJA ---
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Cotización"

        # 2. --- DEFINICIÓN DE ESTILOS REUTILIZABLES ---
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="008080", end_color="008080", fill_type="solid")
        chapter_font = Font(bold=True, color="FFFFFF")
        chapter_fill = PatternFill(start_color="004D40", end_color="004D40", fill_type="solid")
        total_font = Font(bold=True)

        ### SOLUCIÓN AL PROBLEMA 2 y 3: DEFINIR ALINEACIONES ###
        # Alineación para la descripción: centrada verticalmente y con ajuste de texto
        description_alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        # Alineación para el resto de columnas: centrada en ambas direcciones
        center_alignment = Alignment(horizontal="center", vertical="center")

        # 3. --- CONFIGURACIÓN DE LA HOJA (ANCHOS Y ENCABEZADOS) ---
        sheet.column_dimensions['A'].width = 5
        sheet.column_dimensions['B'].width = 60
        sheet.column_dimensions['C'].width = 12
        sheet.column_dimensions['D'].width = 15
        sheet.column_dimensions['E'].width = 20
        sheet.column_dimensions['F'].width = 20

        headers = ["Item", "Descripción", "Cantidad", "Unidad", "Precio Unitario", "Total"]
        for col_num, header_title in enumerate(headers, 1):
            cell = sheet.cell(row=1, column=col_num)
            cell.value = header_title
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_alignment

        # 4. --- PROCESAMIENTO DE LOS ÍTEMS (CAPÍTULOS Y ACTIVIDADES) ---
        row_num = 2
        item_counter = 1

        for item in items:
            if item['type'] == 'chapter':
                sheet.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=6)
                cell = sheet.cell(row=row_num, column=1)
                cell.value = item['name'].upper()
                cell.font = chapter_font
                cell.fill = chapter_fill
                cell.alignment = center_alignment
                row_num += 1


            # Reemplaza tu bloque 'elif item['type'] == 'activity':' con este:

            elif item['type'] == 'activity':

                # --- LÓGICA PARA INSERTAR FILA DE ACTIVIDAD (VERSIÓN ROBUSTA) ---

                # 1. Asignar valores a las celdas. Para las celdas numéricas,

                #    asegurarse de que el tipo de dato sea float/int, no string.

                total_actividad = float(item['cantidad']) * float(item['valor_unitario'])

                sheet.cell(row=row_num, column=1).value = item_counter

                sheet.cell(row=row_num, column=2).value = item['descripcion']

                sheet.cell(row=row_num, column=3).value = float(item['cantidad'])

                sheet.cell(row=row_num, column=4).value = item['unidad']

                sheet.cell(row=row_num, column=5).value = float(item['valor_unitario'])

                sheet.cell(row=row_num, column=6).value = total_actividad

                # 2. Aplicar alineaciones a todas las celdas de la fila.

                #    Esto es seguro y no causa errores de tipo.

                sheet.cell(row=row_num, column=1).alignment = center_alignment

                sheet.cell(row=row_num, column=2).alignment = description_alignment

                sheet.cell(row=row_num, column=3).alignment = center_alignment

                sheet.cell(row=row_num, column=4).alignment = center_alignment

                sheet.cell(row=row_num, column=5).alignment = center_alignment

                sheet.cell(row=row_num, column=6).alignment = center_alignment

                sheet.cell(row=row_num, column=5).number_format = '"$"#,##0.00'

                sheet.cell(row=row_num, column=6).number_format = '"$"#,##0.00'

                # 3. Incrementar contadores para la siguiente fila.

                item_counter += 1

                row_num += 1

        # 5. --- CÁLCULO DE SUBTOTAL, TOTALES Y AIU ---
        subtotal_row_num = row_num
        subtotal_cell_address = f"F{subtotal_row_num}"

        # Fila de Subtotal
        sheet.merge_cells(f"A{subtotal_row_num}:E{subtotal_row_num}")
        sheet[f"A{subtotal_row_num}"].value = "SUBTOTAL COSTOS DIRECTOS"
        sheet[f"A{subtotal_row_num}"].font = total_font
        sheet[f"A{subtotal_row_num}"].alignment = Alignment(horizontal="right")
        sheet[subtotal_cell_address].value = f"=SUM(F2:F{subtotal_row_num - 1})"
        sheet[subtotal_cell_address].number_format = '"$"#,##0.00'
        sheet[subtotal_cell_address].font = total_font
        row_num += 1

        ### SOLUCIÓN AL PROBLEMA 1: LÓGICA PARA PERSONA JURÍDICA VS NATURAL ###
        if tipo_persona.lower() == "jurídica":
            # Lógica completa para AIU
            # Aplicar estilos a las filas de totales

            admin_row_num = row_num
            impr_row_num = row_num + 1
            util_row_num = row_num + 2
            iva_row_num = row_num + 3
            total_row_num = row_num + 4
            # Aplicar estilos y bordes a las filas de totales
            for i in range(admin_row_num, total_row_num + 1):
                sheet[f"A{i}"].alignment = Alignment(horizontal="right")
                sheet[f"F{i}"].number_format = '"$"#,##0.00'
                sheet[f"A{i}"].font = total_font
                sheet[f"F{i}"].font = total_font
            # Aplicar color y bordes gruesos por fuera
            self.aplicar_bordes_totales(sheet, subtotal_row_num, total_row_num, color_hex="D9EAD3")

            # Administración
            sheet.merge_cells(f"A{admin_row_num}:E{admin_row_num}");
            sheet[f"A{admin_row_num}"].value = f"ADMINISTRACIÓN ({administracion}%)"
            sheet[f"F{admin_row_num}"].value = f"={subtotal_cell_address}*({administracion}/100)"

            # Imprevistos
            sheet.merge_cells(f"A{impr_row_num}:E{impr_row_num}");
            sheet[f"A{impr_row_num}"].value = f"IMPREVISTOS ({imprevistos}%)"
            sheet[f"F{impr_row_num}"].value = f"={subtotal_cell_address}*({imprevistos}/100)"

            # Utilidad
            sheet.merge_cells(f"A{util_row_num}:E{util_row_num}");
            sheet[f"A{util_row_num}"].value = f"UTILIDAD ({utilidad}%)"
            sheet[f"F{util_row_num}"].value = f"={subtotal_cell_address}*({utilidad}/100)"

            # IVA sobre Utilidad
            sheet.merge_cells(f"A{iva_row_num}:E{iva_row_num}");
            sheet[f"A{iva_row_num}"].value = f"IVA SOBRE UTILIDAD ({iva_utilidad}%)"
            sheet[f"F{iva_row_num}"].value = f"=F{util_row_num}*({iva_utilidad}/100)"

            # Fila de Total Final
            sheet.merge_cells(f"A{total_row_num}:E{total_row_num}");
            sheet[f"A{total_row_num}"].value = "VALOR TOTAL COTIZACIÓN"
            sheet[f"F{total_row_num}"].value = f"=SUM(F{subtotal_row_num}:F{iva_row_num})"

            # Aplicar estilos a las filas de totales
            for i in range(admin_row_num, total_row_num + 1):
                sheet[f"A{i}"].alignment = Alignment(horizontal="right")
                sheet[f"F{i}"].number_format = '"$"#,##0.00'
                if i == total_row_num:
                    sheet[f"A{i}"].font = total_font
                    sheet[f"F{i}"].font = total_font
                    sheet[f"A{i}"].fill = header_fill
                    sheet[f"F{i}"].fill = header_fill
            self.aplicar_bordes_totales_con_marco_grueso(sheet, subtotal_row_num, total_row_num, color_hex="D9EAD3")

        else:  # Persona Natural
            # Lógica simple con IVA sobre el subtotal
            iva_row_num = row_num
            total_row_num = row_num + 1

            # IVA
            sheet.merge_cells(f"A{iva_row_num}:E{iva_row_num}");
            sheet[f"A{iva_row_num}"].value = f"IVA ({iva_utilidad}%)"
            sheet[f"F{iva_row_num}"].value = f"={subtotal_cell_address}*({iva_utilidad}/100)"
            sheet[f"A{iva_row_num}"].alignment = Alignment(horizontal="right")
            sheet[f"F{iva_row_num}"].number_format = '"$"#,##0.00'

            # Fila de Total Final
            sheet.merge_cells(f"A{total_row_num}:E{total_row_num}");
            sheet[f"A{total_row_num}"].value = "VALOR TOTAL COTIZACIÓN"
            sheet[f"F{total_row_num}"].value = f"=SUM(F{subtotal_row_num}, F{iva_row_num})"
            sheet[f"A{total_row_num}"].font = total_font;
            sheet[f"F{total_row_num}"].font = total_font
            sheet[f"A{total_row_num}"].alignment = Alignment(horizontal="right")
            sheet[f"F{total_row_num}"].number_format = '"$"#,##0.00'
            sheet[f"A{total_row_num}"].fill = header_fill
            sheet[f"F{total_row_num}"].fill = header_fill

            self.aplicar_bordes_totales_con_marco_grueso(sheet, subtotal_row_num, total_row_num, color_hex="D9EAD3")

        # 6. --- GUARDAR EL ARCHIVO ---
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        export_dir = "exports"
        if not os.path.exists(export_dir): os.makedirs(export_dir)
        excel_path = os.path.join(export_dir, f"cotizacion_{timestamp}.xlsx")

        try:
            workbook.save(excel_path)
            print(f"Archivo Excel guardado en: {excel_path}")
            return excel_path
        except Exception as e:
            print(f"Error al guardar el archivo Excel: {e}")
            return None
