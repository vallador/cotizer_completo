# controllers/excel_controller.py
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, numbers
import os
from datetime import datetime


class ExcelController:
    def __init__(self, cotizacion_controller):
        self.cotizacion_controller = cotizacion_controller

    def generate_excel(self, cotizacion_id=None, cotizacion=None, tipo_persona=None, activities=None,
                       administracion=None, imprevistos=None, utilidad=None, iva_utilidad=None):
        """
        Genera un archivo Excel de cotización.
        Puede recibir un ID de cotización, un objeto cotización, o datos individuales.
        """
        # Si se proporciona un ID, obtener la cotización
        if cotizacion_id and not cotizacion:
            cotizacion = self.cotizacion_controller.obtener_cotizacion(cotizacion_id)

        # Crear un nuevo archivo Excel
        excel_path = os.path.join(os.getcwd(), "cotizacion.xlsx")
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Cotización"

        # Estilos
        fill_color = PatternFill(start_color='008000', end_color='008000', fill_type='solid')
        thick_border = Side(border_style="thick", color="000000")  # Borde grueso
        thin_border = Side(border_style="thin", color="000000")  # Borde delgado

        # Estilo para encabezados de capítulo
        chapter_font = Font(bold=True, color="FFFFFF")
        chapter_fill = PatternFill(start_color='008000', end_color='008000', fill_type='solid')
        chapter_alignment = Alignment(horizontal="center", vertical="center")

        # Ajustar anchos de columna
        sheet.column_dimensions['A'].width = 5  # Ajusta el ancho de la columna A
        sheet.column_dimensions['B'].width = 50  # Ajusta el ancho de la columna B

        # Establecer encabezados de la tabla
        headers = ["Item", "Descripción", "Cantidad", "Unidad", "Precio Unitario", "Total"]
        for col_num, header in enumerate(headers, start=1):
            col_letter = get_column_letter(col_num)
            cell = sheet[f"{col_letter}1"]
            cell.value = header
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center")
            cell.fill = fill_color

        # Determinar los valores de AIU y las actividades
        if cotizacion:
            # Usar datos de la cotización
            tipo_persona = cotizacion.cliente.tipo
            activities = []
            for act in cotizacion.actividades:
                activities.append({
                    'descripcion': act.descripcion,
                    'cantidad': act.cantidad,
                    'unidad': act.unidad,
                    'valor_unitario': act.valor_unitario,
                    'total': act.valor_total,
                    'capitulo_id': act.categoria,
                    'capitulo_nombre': act.categoria
                })
            administracion = cotizacion.administracion
            imprevistos = cotizacion.imprevistos
            utilidad = cotizacion.utilidad
            iva_utilidad = cotizacion.iva_utilidad
        else:
            # Convertir valores a float si son strings
            if isinstance(administracion, str):
                administracion = float(administracion.replace(',', '.'))
            if isinstance(imprevistos, str):
                imprevistos = float(imprevistos.replace(',', '.'))
            if isinstance(utilidad, str):
                utilidad = float(utilidad.replace(',', '.'))
            if isinstance(iva_utilidad, str):
                iva_utilidad = float(iva_utilidad.replace(',', '.'))

        # Organizar actividades por capítulos
        capitulos = {}
        actividades_sin_capitulo = []

        for activity in activities:
            capitulo_id = activity.get('capitulo_id')
            capitulo_nombre = activity.get('capitulo_nombre')

            if capitulo_id and capitulo_nombre:
                if capitulo_id not in capitulos:
                    capitulos[capitulo_id] = {
                        'nombre': capitulo_nombre,
                        'actividades': []
                    }
                capitulos[capitulo_id]['actividades'].append(activity)
            else:
                actividades_sin_capitulo.append(activity)

        # Insertar las actividades en las filas, agrupadas por capítulos
        row_num = 2
        item_num = 1

        # Primero insertar actividades sin capítulo
        for activity in actividades_sin_capitulo:
            sheet[f"A{row_num}"].value = item_num
            sheet[f"B{row_num}"].value = activity['descripcion']
            sheet[f"C{row_num}"].value = activity['cantidad']
            sheet[f"D{row_num}"].value = activity['unidad']
            sheet[f"E{row_num}"].value = activity['valor_unitario']
            # Insertar una fórmula para calcular el total de la actividad
            sheet[f"F{row_num}"].value = f"=C{row_num}*E{row_num}"

            # Formato contable sin decimales para valores monetarios
            sheet[f"E{row_num}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
            sheet[f"F{row_num}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1

            row_num += 1
            item_num += 1

        # Luego insertar actividades por capítulos
        for capitulo_id, capitulo_data in capitulos.items():
            # Insertar encabezado de capítulo
            sheet.merge_cells(f"A{row_num}:F{row_num}")
            cell = sheet[f"A{row_num}"]
            cell.value = capitulo_data['nombre'].upper()
            cell.font = chapter_font
            cell.alignment = chapter_alignment
            cell.fill = chapter_fill

            row_num += 1

            # Insertar actividades del capítulo
            for activity in capitulo_data['actividades']:
                sheet[f"A{row_num}"].value = item_num
                sheet[f"B{row_num}"].value = activity['descripcion']
                sheet[f"C{row_num}"].value = activity['cantidad']
                sheet[f"D{row_num}"].value = activity['unidad']
                sheet[f"E{row_num}"].value = activity['valor_unitario']
                # Insertar una fórmula para calcular el total de la actividad
                sheet[f"F{row_num}"].value = f"=C{row_num}*E{row_num}"

                # Formato contable sin decimales para valores monetarios
                sheet[f"E{row_num}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
                sheet[f"F{row_num}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1

                row_num += 1
                item_num += 1

        # Calcular totales
        total_row = row_num

        # Ajustar el texto y la alineación en la columna B
        for row in sheet.iter_rows(min_col=2, max_col=2, min_row=2, max_row=sheet.max_row):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="center")

        # Ajustar alineación central en columnas específicas
        for row in sheet.iter_rows(min_col=1, max_col=6, min_row=2, max_row=total_row - 1):
            for cell in row:
                if cell.column in [1, 3, 4, 5, 6]:  # Columna A (1), C (3), D (4), E (5), F (6)
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        # Fila de subtotal
        sheet[f"A{total_row}"].fill = fill_color
        subtotal_cell = f"F{total_row}"
        sheet.merge_cells(f"A{total_row}:E{total_row}")
        sheet[f"A{total_row}"].font = Font(bold=True)
        sheet[f"A{total_row}"].alignment = Alignment(horizontal="right")
        sheet[f"A{total_row}"].value = "TOTAL COSTOS DIRECTOS DE OBRA"
        sheet[f"F{total_row}"].value = f"=SUM(F2:F{total_row - 1})"  # Fórmula de subtotal
        sheet[f"F{total_row}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1  # Formato contable sin decimales

        # Aplicar AIU si el cliente es jurídico
        if tipo_persona == "jurídica":
            # Fórmulas para Administración, Imprevistos y Utilidad
            sheet.merge_cells(f"A{total_row + 1}:E{total_row + 1}")
            sheet[f"A{total_row + 1}"].value = f"ADMINISTRACIÓN {administracion}%"
            sheet[f"A{total_row + 1}"].alignment = Alignment(horizontal="right")
            sheet[f"F{total_row + 1}"].value = f"={administracion / 100}*{subtotal_cell}"
            sheet[f"F{total_row + 1}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1

            sheet.merge_cells(f"A{total_row + 2}:E{total_row + 2}")
            sheet[f"A{total_row + 2}"].value = f"IMPREVISTOS {imprevistos}%"
            sheet[f"F{total_row + 2}"].value = f"={imprevistos / 100}*{subtotal_cell}"
            sheet[f"F{total_row + 2}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
            sheet[f"A{total_row + 2}"].alignment = Alignment(horizontal="right")

            util_cell = f"F{total_row + 3}"
            sheet.merge_cells(f"A{total_row + 3}:E{total_row + 3}")
            sheet[f"A{total_row + 3}"].value = f"UTILIDAD {utilidad}%"
            sheet[f"F{total_row + 3}"].value = f"={utilidad / 100}*{subtotal_cell}"
            sheet[f"F{total_row + 3}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
            sheet[f"A{total_row + 3}"].alignment = Alignment(horizontal="right")

            sheet.merge_cells(f"A{total_row + 4}:E{total_row + 4}")
            sheet[f"A{total_row + 4}"].value = f"IVA SOBRE UTILIDAD {iva_utilidad}%"
            sheet[f"F{total_row + 4}"].value = f"={iva_utilidad / 100}*{util_cell}"
            sheet[f"F{total_row + 4}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
            sheet[f"A{total_row + 4}"].alignment = Alignment(horizontal="right")

            sheet[f"A{total_row + 5}"].fill = fill_color
            sheet.merge_cells(f"A{total_row + 5}:E{total_row + 5}")
            sheet[f"A{total_row + 5}"].value = f"VALOR TOTAL COTIZACIÓN A TODO COSTO"
            sheet[f"A{total_row + 5}"].font = Font(bold=True)
            sheet[f"A{total_row + 5}"].alignment = Alignment(horizontal="right")
            sheet[f"F{total_row + 5}"].value = f"=SUM(F{total_row}:F{total_row + 4})"  # Fórmula TOTAL
            sheet[f"F{total_row + 5}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
        else:
            # Para clientes naturales, solo mostrar el total con IVA
            sheet.merge_cells(f"A{total_row + 1}:E{total_row + 1}")
            sheet[f"A{total_row + 1}"].value = f"IVA {iva_utilidad}%"
            sheet[f"F{total_row + 1}"].value = f"={iva_utilidad / 100}*{subtotal_cell}"
            sheet[f"F{total_row + 1}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1
            sheet[f"A{total_row + 1}"].alignment = Alignment(horizontal="right")

            sheet[f"A{total_row + 2}"].fill = fill_color
            sheet.merge_cells(f"A{total_row + 2}:E{total_row + 2}")
            sheet[f"A{total_row + 2}"].value = f"VALOR TOTAL COTIZACIÓN"
            sheet[f"A{total_row + 2}"].font = Font(bold=True)
            sheet[f"A{total_row + 2}"].alignment = Alignment(horizontal="right")
            sheet[f"F{total_row + 2}"].value = f"=SUM(F{total_row}:F{total_row + 1})"  # Fórmula TOTAL
            sheet[f"F{total_row + 2}"].number_format = numbers.FORMAT_NUMBER_COMMA_SEPARATED1

        # Ajustar automáticamente el ancho de las columnas
        for col_num, _ in enumerate(headers, start=1):
            col_letter = get_column_letter(col_num)
            sheet.column_dimensions[col_letter].auto_size = True

        max_row = sheet.max_row
        max_col = sheet.max_column

        # Aplicar bordes a todas las celdas
        for row in sheet.iter_rows(min_row=1, max_row=max_row, min_col=1, max_col=max_col):
            for cell in row:
                # Aplicar bordes delgados en todas las celdas internas
                cell.border = Border(
                    left=thin_border, right=thin_border, top=thin_border, bottom=thin_border
                )

        # Aplicar bordes gruesos alrededor de los datos (bordes externos)
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell = sheet.cell(row=row, column=col)
                if row == 1:
                    cell.border = Border(top=thick_border, left=cell.border.left, right=cell.border.right,
                                         bottom=cell.border.bottom)

                # Borde inferior grueso para la última fila
                if row == max_row:
                    cell.border = Border(bottom=thick_border, left=cell.border.left, right=cell.border.right,
                                         top=cell.border.top)

                # Borde izquierdo grueso para la primera columna
                if col == 1:
                    cell.border = Border(left=thick_border, right=cell.border.right, top=cell.border.top,
                                         bottom=cell.border.bottom)

                # Borde derecho grueso para la última columna
                if col == max_col:
                    cell.border = Border(right=thick_border, left=cell.border.left, top=cell.border.top,
                                         bottom=cell.border.bottom)

        # Generar nombre de archivo con fecha y hora
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if cotizacion:
            excel_path = os.path.join(os.getcwd(), f"cotizacion_{cotizacion.numero}_{timestamp}.xlsx")
        else:
            excel_path = os.path.join(os.getcwd(), f"cotizacion_{timestamp}.xlsx")

        # Guardar el archivo Excel
        workbook.save(excel_path)
        workbook.close()

        print(f"Archivo Excel guardado en: {excel_path}")
        return excel_path
