import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
import os
import re
from datetime import datetime
import calendar
from PIL import Image
import io
import tempfile
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
import matplotlib.pyplot as plt
import numpy as np
from openpyxl.utils import get_column_letter


class WordController:
    def __init__(self, cotizacion_controller):
        self.cotizacion_controller = cotizacion_controller

        # DEBUG: Mostrar directorio actual
        print(f"DEBUG: Directorio actual: {os.getcwd()}")

        self.templates_dir = os.path.join(os.getcwd(), 'controllers', 'data', 'templates')
        print(f"DEBUG: templates_dir calculado: {self.templates_dir}")

        # Verificar si existe con la ruta calculada
        if not os.path.exists(self.templates_dir):
            print(f"DEBUG: No existe {self.templates_dir}, probando rutas alternativas...")

            # Probar diferentes rutas posibles
            rutas_alternativas = [
                os.path.join(os.getcwd(), 'data', 'templates'),
                os.path.join(os.getcwd(), 'controllers', 'data', 'templates'),
                os.path.join(os.path.dirname(__file__), '..', 'data', 'templates'),
                os.path.join(os.path.dirname(__file__), 'data', 'templates')
            ]

            for ruta in rutas_alternativas:
                ruta_abs = os.path.abspath(ruta)
                print(f"DEBUG: Probando ruta: {ruta_abs}")
                if os.path.exists(ruta_abs):
                    self.templates_dir = ruta_abs
                    print(f"DEBUG: ¡ENCONTRADO! Usando templates_dir: {self.templates_dir}")
                    break
            else:
                print(f"DEBUG: No se encontró el directorio de templates en ninguna ruta")
                # Crear directorio de plantillas si no existe
                os.makedirs(self.templates_dir)
                print(f"DEBUG: Directorio creado: {self.templates_dir}")
        else:
            print(f"DEBUG: templates_dir existe: {self.templates_dir}")

        # Mostrar contenido del directorio si existe
        if os.path.exists(self.templates_dir):
            contenido = os.listdir(self.templates_dir)
            print(f"DEBUG: Contenido de templates_dir: {contenido}")

        # Nombres de plantillas predeterminadas
        self.template_juridica_larga = os.path.join(self.templates_dir, 'plantilla_juridica_larga.docx')
        self.template_juridica_corta = os.path.join(self.templates_dir, 'plantilla_juridica_corta.docx')
        self.template_natural_corta = os.path.join(self.templates_dir, 'plantilla_natural_corta.docx')

        print(f"DEBUG: template_juridica_larga = {self.template_juridica_larga}")
        print(f"DEBUG: template_juridica_corta = {self.template_juridica_corta}")
        print(f"DEBUG: template_natural_corta = {self.template_natural_corta}")

        # Verificar existencia de cada plantilla
        print(f"DEBUG: ¿Existe template_juridica_larga? {os.path.exists(self.template_juridica_larga)}")
        print(f"DEBUG: ¿Existe template_juridica_corta? {os.path.exists(self.template_juridica_corta)}")
        print(f"DEBUG: ¿Existe template_natural_corta? {os.path.exists(self.template_natural_corta)}")

    def generate_natural_cotizacion(self, config, excel_path):
        """Genera una cotización para una persona natural usando plantilla."""
        try:
            # Verificar que existe la plantilla
            if not os.path.exists(self.template_natural_corta):
                raise FileNotFoundError(f"No se encontró la plantilla: {self.template_natural_corta}")

            # Cargar la plantilla
            doc = Document(self.template_natural_corta)

            # Preparar datos para reemplazo
            replace_data = self._prepare_replace_data_natural(config)

            # Reemplazar marcadores en todo el documento
            self._replace_placeholders_in_document(doc, replace_data)

            # Insertar tabla de Excel si existe
            self._insert_excel_table_in_document(doc, excel_path)

            # Guardar documento
            output_path = self._save_document(doc, "Natural", config.get("lugar", "Cotizacion"))

            return output_path

        except Exception as e:
            print(f"Error al generar cotización natural: {e}")
            raise

    def generate_juridica_cotizacion(self, config, excel_path):
        """Genera una cotización para una persona jurídica usando plantilla."""
        try:
            # Determinar qué plantilla usar
            template_path = self.template_juridica_larga if config.get("formato",
                                                                       "largo") == "largo" else self.template_juridica_corta

            # Verificar que existe la plantilla
            if not os.path.exists(template_path):
                raise FileNotFoundError(f"No se encontró la plantilla: {template_path}")

            # Cargar la plantilla
            doc = Document(template_path)

            # Preparar datos para reemplazo
            replace_data = self._prepare_replace_data_juridica(config)

            # Reemplazar marcadores en todo el documento
            self._replace_placeholders_in_document(doc, replace_data)

            # Insertar tabla de Excel si existe
            self._insert_excel_table_in_document(doc, excel_path)

            # Guardar documento
            output_path = self._save_document(doc, "Juridica", config.get("lugar", "Cotizacion"))

            return output_path

        except Exception as e:
            print(f"Error al generar cotización jurídica: {e}")
            raise

    def _prepare_replace_data_natural(self, config):
        """Prepara los datos para reemplazar en plantilla de persona natural."""
        fecha_actual = datetime.now()
        meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
                 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
        fecha_formateada = f"{fecha_actual.day} de {meses[fecha_actual.month - 1]} de {fecha_actual.year}"

        # Preparar forma de pago
        forma_pago = self._format_forma_pago(config)

        return {
            'fecha': config.get('fecha', fecha_formateada),
            'titulo': config.get('titulo', 'COTIZACIÓN DE SERVICIOS'),
            'lugar': config.get('lugar', 'No especificado'),
            'concepto': config.get('concepto', 'Servicios de construcción y mantenimiento'),
            'validez': str(config.get('validez', '30')),
            'cuadrillas': config.get('cuadrillas', 'una'),
            'operarios_num': str(config.get('operarios_num', '2')),
            'operarios_letra': config.get('operarios_letra', 'dos'),
            'plazo_dias': str(config.get('plazo_dias', '15')),
            'plazo_tipo': config.get('plazo_tipo', 'hábiles'),
            'forma_pago': forma_pago
        }

    def _prepare_replace_data_juridica(self, config):
        """Prepara los datos para reemplazar en plantilla de persona jurídica."""
        fecha_actual = datetime.now()
        meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
                 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
        fecha_formateada = f"{fecha_actual.day} de {meses[fecha_actual.month - 1]} de {fecha_actual.year}"

        # Preparar forma de pago
        forma_pago = self._format_forma_pago(config)

        return {
            'fecha': config.get('fecha', fecha_formateada),
            'titulo': config.get('titulo', 'PROPUESTA TÉCNICO ECONÓMICA'),
            'lugar': config.get('lugar', 'No especificado'),
            'concepto': config.get('concepto', 'Servicios de construcción y mantenimiento'),
            'validez': str(config.get('validez', '30')),
            'cuadrillas': config.get('cuadrillas', 'una'),
            'operarios_num': str(config.get('operarios_num', '2')),
            'operarios_letra': config.get('operarios_letra', 'dos'),
            'plazo_dias': str(config.get('plazo_dias', '15')),
            'plazo_tipo': config.get('plazo_tipo', 'hábiles'),
            'forma_pago': forma_pago,
            'director_obra': config.get('director_obra', 'A definir'),
            'residente_obra': config.get('residente_obra', 'A definir'),
            'tecnologo_sgsst': config.get('tecnologo_sgsst', 'A definir')
        }

    def _format_forma_pago(self, config):
        """Formatea la forma de pago según la configuración."""
        if config.get('pago_contraentrega', False):
            return "Contraentrega"
        elif config.get('pago_personalizado'):
            return config['pago_personalizado']
        elif config.get('pago_porcentajes', False):
            anticipo = config.get('anticipo', 30)
            avance = config.get('avance', 40)
            avance_requerido = config.get('avance_requerido', 50)
            final = config.get('final', 30)
            return f"{anticipo}% de anticipo, {avance}% al {avance_requerido}% de avance y {final}% al finalizar."
        else:
            return "Contraentrega"

    def _replace_placeholders_in_document(self, doc, replace_data):
        """Reemplaza todos los marcadores en el documento."""
        # Reemplazar en párrafos
        for paragraph in doc.paragraphs:
            self._replace_in_paragraph(paragraph, replace_data)

        # Reemplazar en tablas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, replace_data)

        # Reemplazar en headers y footers
        for section in doc.sections:
            header = section.header
            footer = section.footer
            for paragraph in header.paragraphs:
                self._replace_in_paragraph(paragraph, replace_data)
            for paragraph in footer.paragraphs:
                self._replace_in_paragraph(paragraph, replace_data)

    def _replace_in_paragraph(self, paragraph, replace_data):
        """Reemplaza marcadores en un párrafo específico."""
        for key, value in replace_data.items():
            marker = f"{{{{{key}}}}}"
            if marker in paragraph.text:
                self._replace_text_in_paragraph(paragraph, marker, str(value))

    def _insert_excel_table_in_document(self, doc, excel_path):
        """Busca el marcador de tabla y lo reemplaza con la imagen de Excel."""
        if not excel_path or not os.path.exists(excel_path):
            return

        # Buscar el marcador {{tabla_cotizacion}} en todo el documento
        for paragraph in doc.paragraphs:
            if '{{tabla_cotizacion}}' in paragraph.text:
                # Limpiar el párrafo
                paragraph.clear()

                # Generar imagen de la tabla
                table_image_path = self._capture_excel_table(excel_path)
                if table_image_path and os.path.exists(table_image_path):
                    try:
                        # Insertar la imagen
                        run = paragraph.add_run()
                        run.add_picture(table_image_path, width=Inches(6.5))

                        # Limpiar archivo temporal
                        os.remove(table_image_path)
                    except Exception as e:
                        print(f"Error al insertar imagen de tabla: {e}")
                        paragraph.add_run("[Error al generar la tabla de cotización]")
                else:
                    paragraph.add_run("[No se pudo generar la tabla de cotización]")
                break

        # También buscar en tablas
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        if '{{tabla_cotizacion}}' in paragraph.text:
                            paragraph.clear()
                            table_image_path = self._capture_excel_table(excel_path)
                            if table_image_path and os.path.exists(table_image_path):
                                try:
                                    run = paragraph.add_run()
                                    run.add_picture(table_image_path, width=Inches(6.0))
                                    os.remove(table_image_path)
                                except Exception as e:
                                    print(f"Error al insertar imagen de tabla en celda: {e}")
                                    paragraph.add_run("[Error al generar la tabla de cotización]")
                            return

    def _save_document(self, doc, tipo, lugar):
        """Guarda el documento y retorna la ruta."""
        output_dir = os.path.join(os.getcwd(), "output")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        filename = f"Cotizacion_{tipo}_{lugar.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        output_path = os.path.join(output_dir, filename)

        doc.save(output_path)
        return output_path

    # Mantener el resto de métodos originales que funcionan bien
    def generate_word_document(self, cotizacion_id, excel_path, datos_adicionales, formato='auto'):
        """
        Genera un documento Word a partir de una cotización y datos adicionales.

        Args:
            cotizacion_id: ID de la cotización
            excel_path: Ruta al archivo Excel de la cotización
            datos_adicionales: Diccionario con datos adicionales para el documento
            formato: 'largo', 'corto' o 'auto' (selecciona automáticamente según tipo de cliente)

        Returns:
            Ruta al documento Word generado
        """
        # Obtener la cotización
        cotizacion = self.cotizacion_controller.obtener_cotizacion(cotizacion_id)
        if not cotizacion:
            raise ValueError(f"No se encontró la cotización con ID {cotizacion_id}")

        # Determinar la plantilla a utilizar
        template_path = self._select_template(cotizacion, formato)

        # Capturar la tabla de Excel como imagen
        table_image_path = self._capture_excel_table(excel_path)

        # Preparar los datos para reemplazo
        replace_data = self._prepare_replace_data(cotizacion, datos_adicionales)

        # Generar el documento
        output_path = self._generate_document(template_path, replace_data, table_image_path)

        return output_path

    def _select_template(self, cotizacion, formato):
        """
        Selecciona la plantilla adecuada según el tipo de cliente y formato.

        Args:
            cotizacion: Objeto Cotizacion
            formato: 'largo', 'corto' o 'auto'

        Returns:
            Ruta a la plantilla seleccionada
        """
        tipo_cliente = cotizacion.cliente.tipo.lower()

        if formato == 'auto':
            # Selección automática según tipo de cliente
            if tipo_cliente == 'jurídica':
                return self.template_juridica_larga
            else:
                return self.template_natural_corta
        elif formato == 'largo':
            if tipo_cliente == 'jurídica':
                return self.template_juridica_larga
            else:
                # Para clientes naturales no hay formato largo, usar corto
                return self.template_natural_corta
        elif formato == 'corto':
            if tipo_cliente == 'jurídica':
                return self.template_juridica_corta
            else:
                return self.template_natural_corta
        else:
            # Formato no reconocido, usar selección automática
            return self._select_template(cotizacion, 'auto')

    def _capture_excel_table(self, excel_path):
        """
        Captura la tabla de Excel como imagen.

        Args:
            excel_path: Ruta al archivo Excel

        Returns:
            Ruta a la imagen temporal de la tabla
        """
        try:
            # Cargar el libro de Excel
            wb = load_workbook(excel_path)
            sheet = wb.active

            # Determinar el rango de la tabla (asumiendo que comienza en A1)
            max_row = sheet.max_row
            max_col = sheet.max_column

            # Crear una figura de matplotlib para renderizar la tabla
            fig, ax = plt.subplots(figsize=(12, max_row * 0.4))
            ax.axis('tight')
            ax.axis('off')

            # Extraer datos de la tabla
            data = []
            for row in range(1, max_row + 1):
                row_data = []
                for col in range(1, max_col + 1):
                    cell_value = sheet.cell(row=row, column=col).value
                    row_data.append(str(cell_value) if cell_value is not None else "")
                data.append(row_data)

            # Crear la tabla
            table = ax.table(cellText=data,
                             colLabels=None,
                             cellLoc='center',
                             loc='center',
                             colWidths=[0.1, 0.4, 0.1, 0.1, 0.15, 0.15])

            # Estilo de la tabla
            table.auto_set_font_size(False)
            table.set_fontsize(10)
            table.scale(1, 1.5)

            # Estilo para la primera fila (encabezados)
            for (row, col), cell in table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontweight='bold')
                    cell.set_facecolor('#008000')  # Verde
                    cell.set_text_props(color='white')

            # Guardar la imagen en un archivo temporal
            temp_image_path = os.path.join(tempfile.gettempdir(),
                                           f"tabla_cotizacion_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
            plt.savefig(temp_image_path, bbox_inches='tight', dpi=300)
            plt.close()

            return temp_image_path

        except Exception as e:
            print(f"Error al capturar la tabla de Excel: {e}")
            # En caso de error, devolver None y manejar en el método que llama
            return None

    def _prepare_replace_data(self, cotizacion, datos_adicionales):
        """
        Prepara los datos para reemplazar en la plantilla.

        Args:
            cotizacion: Objeto Cotizacion
            datos_adicionales: Diccionario con datos adicionales

        Returns:
            Diccionario con los datos para reemplazo
        """
        # Obtener fecha actual en formato "DD de Mes de AAAA"
        fecha_actual = datetime.now()
        meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
                 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre']
        fecha_formateada = f"{fecha_actual.day} de {meses[fecha_actual.month - 1]} de {fecha_actual.year}"

        # Preparar forma de pago
        forma_pago = datos_adicionales.get('forma_pago', 'contraentrega')
        texto_forma_pago = ""

        if forma_pago == 'contraentrega':
            texto_forma_pago = "Contraentrega"
        else:
            # Formato: X% al inicio, Y% al Z% de avance, W% al finalizar
            inicio = datos_adicionales.get('porcentaje_inicio', 0)
            avance = datos_adicionales.get('porcentaje_avance', 0)
            avance_requerido = datos_adicionales.get('avance_requerido', 50)
            final = datos_adicionales.get('porcentaje_final', 0)

            texto_forma_pago = f"{inicio}% al inicio de obra, {avance}% al {avance_requerido}% de avance y {final}% al finalizar la intervención"

        # Crear diccionario de reemplazo
        replace_data = {
            'fecha': fecha_formateada,
            'cliente': cotizacion.cliente.nombre,
            'nit': cotizacion.cliente.nit or 'N/A',
            'direccion': cotizacion.cliente.direccion or 'N/A',
            'referencia': datos_adicionales.get('referencia', 'Servicios de construcción y mantenimiento'),
            'validez': datos_adicionales.get('validez', '30'),
            'cuadrillas': datos_adicionales.get('cuadrillas', 'una'),
            'operarios': datos_adicionales.get('operarios', '2'),
            'plazo': datos_adicionales.get('plazo', '15'),
            'forma_pago': texto_forma_pago,
            'porcentaje_inicio': datos_adicionales.get('porcentaje_inicio', '30'),
            'porcentaje_avance': datos_adicionales.get('porcentaje_avance', '40'),
            'avance_requerido': datos_adicionales.get('avance_requerido', '50'),
            'porcentaje_final': datos_adicionales.get('porcentaje_final', '30')
        }

        return replace_data

    def _generate_document(self, template_path, replace_data, table_image_path):
        """
        Genera el documento Word reemplazando los marcadores y añadiendo la imagen.

        Args:
            template_path: Ruta a la plantilla
            replace_data: Diccionario con datos para reemplazo
            table_image_path: Ruta a la imagen de la tabla

        Returns:
            Ruta al documento generado
        """
        try:
            # Cargar la plantilla
            doc = Document(template_path)

            # Reemplazar marcadores en párrafos
            for paragraph in doc.paragraphs:
                for key, value in replace_data.items():
                    marker = f"{{{{{key}}}}}"
                    if marker in paragraph.text:
                        # Reemplazar manteniendo el formato
                        self._replace_text_in_paragraph(paragraph, marker, str(value))

                # Reemplazar marcador de tabla con imagen
                if '{{tabla_cotizacion}}' in paragraph.text:
                    # Eliminar el texto del marcador
                    paragraph.text = ""

                    # Insertar la imagen si existe
                    if table_image_path and os.path.exists(table_image_path):
                        run = paragraph.add_run()
                        run.add_picture(table_image_path, width=Inches(6))

            # Reemplazar marcadores en tablas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for key, value in replace_data.items():
                                marker = f"{{{{{key}}}}}"
                                if marker in paragraph.text:
                                    self._replace_text_in_paragraph(paragraph, marker, str(value))

            # Generar nombre de archivo para el documento
            output_dir = os.path.join(os.getcwd(), 'output')
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            output_filename = f"Cotizacion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            output_path = os.path.join(output_dir, output_filename)

            # Guardar el documento
            doc.save(output_path)

            # Eliminar la imagen temporal si existe
            if table_image_path and os.path.exists(table_image_path):
                try:
                    os.remove(table_image_path)
                except:
                    pass

            return output_path

        except Exception as e:
            print(f"Error al generar el documento Word: {e}")
            raise

    def _replace_text_in_paragraph(self, paragraph, marker, text):
        """
        Reemplaza un marcador en un párrafo manteniendo el formato.

        Args:
            paragraph: Objeto párrafo de python-docx
            marker: Marcador a reemplazar
            text: Texto de reemplazo
        """
        if marker in paragraph.text:
            # Iterar sobre los runs para mantener el formato
            for run in paragraph.runs:
                if marker in run.text:
                    run.text = run.text.replace(marker, text)