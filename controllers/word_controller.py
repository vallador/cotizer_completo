import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
import os
import re
from datetime import datetime
import calendar
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

        base_path = os.path.dirname(os.path.abspath(__file__))
        self.templates_dir = os.path.join(base_path, 'data', 'templates')

        # Nombres de plantillas que deben existir en el directorio
        self.template_juridica_larga = os.path.join(self.templates_dir, 'plantilla_juridica_larga.docx')
        self.template_juridica_corta = os.path.join(self.templates_dir, 'plantilla_juridica_corta.docx')
        self.template_natural_corta = os.path.join(self.templates_dir, 'plantilla_natural_corta.docx')

        # Verificar que existen las plantillas
        self._verify_templates_exist()

    def _verify_templates_exist(self):
        """Verifica que las plantillas existan"""
        templates = {
            'Jurídica Larga': self.template_juridica_larga,
            'Jurídica Corta': self.template_juridica_corta,
            'Natural Corta': self.template_natural_corta
        }

        missing_templates = []
        for name, path in templates.items():
            if not os.path.exists(path):
                missing_templates.append(f"  - {name}: {path}")

        if missing_templates:
            print("⚠️  ADVERTENCIA: Las siguientes plantillas no se encontraron:")
            for template in missing_templates:
                print(template)
            print(f"\nAsegúrate de que estén en: {self.templates_dir}")

    def generate_word_document(self, cotizacion_id, excel_path, datos_adicionales, formato='auto',
                               template_personalizada=None):
        """
        Genera un documento Word usando las plantillas existentes.

        Args:
            cotizacion_id: ID de la cotización
            excel_path: Ruta al archivo Excel de la cotización
            datos_adicionales: Diccionario con datos adicionales para el documento
            formato: 'largo', 'corto' o 'auto' (selecciona automáticamente según tipo de cliente)
            template_personalizada: Ruta a una plantilla personalizada (opcional)

        Returns:
            Ruta al documento Word generado
        """
        # Obtener la cotización
        cotizacion = self.cotizacion_controller.obtener_cotizacion(cotizacion_id)
        if not cotizacion:
            raise ValueError(f"No se encontró la cotización con ID {cotizacion_id}")

        # Determinar la plantilla a utilizar
        if template_personalizada and os.path.exists(template_personalizada):
            template_path = template_personalizada
        else:
            template_path = self._select_template(cotizacion, formato)

        # Verificar que la plantilla existe
        if not os.path.exists(template_path):
            raise FileNotFoundError(f"No se encontró la plantilla: {template_path}")


        # Preparar los datos para reemplazo
        replace_data = self._prepare_replace_data(cotizacion, datos_adicionales)

        # Generar el documento
        output_path = self._generate_document(template_path, replace_data)

        return output_path

    def _select_template(self, cotizacion, formato):
        """
        Selecciona la plantilla adecuada según el tipo de cliente y formato.
        """
        tipo_cliente = cotizacion.cliente.tipo.lower()

        if formato == 'auto':
            if tipo_cliente == 'jurídica':
                return self.template_juridica_larga
            else:
                return self.template_natural_corta
        elif formato == 'largo':
            if tipo_cliente == 'jurídica':
                return self.template_juridica_larga
            else:
                return self.template_natural_corta
        elif formato == 'corto':
            if tipo_cliente == 'jurídica':
                return self.template_juridica_corta
            else:
                return self.template_natural_corta
        else:
            return self._select_template(cotizacion, 'auto')

    def debug_template_markers(self, template_path=None):
        """
        Ver qué marcadores están en una plantilla.
        """
        results = {}

        if template_path:
            if os.path.exists(template_path):
                results[os.path.basename(template_path)] = self._extract_markers_from_template(template_path)
            else:
                results['ERROR'] = f"No se encontró la plantilla: {template_path}"
        else:
            # Analizar todas las plantillas
            templates = {
                'plantilla_juridica_larga.docx': self.template_juridica_larga,
                'plantilla_juridica_corta.docx': self.template_juridica_corta,
                'plantilla_natural_corta.docx': self.template_natural_corta
            }

            for name, path in templates.items():
                if os.path.exists(path):
                    results[name] = self._extract_markers_from_template(path)
                else:
                    results[name] = ['PLANTILLA NO ENCONTRADA']

        return results

    def _extract_markers_from_template(self, template_path):
        """Extrae todos los marcadores {{}} de una plantilla"""
        try:
            doc = Document(template_path)
            markers_found = set()

            # Buscar en párrafos
            for paragraph in doc.paragraphs:
                matches = re.findall(r'\{\{([^}]+)\}\}', paragraph.text)
                markers_found.update(matches)

            # Buscar en tablas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            matches = re.findall(r'\{\{([^}]+)\}\}', paragraph.text)
                            markers_found.update(matches)

            # Buscar en encabezados y pies de página
            for section in doc.sections:
                if section.header:
                    for paragraph in section.header.paragraphs:
                        matches = re.findall(r'\{\{([^}]+)\}\}', paragraph.text)
                        markers_found.update(matches)

                if section.footer:
                    for paragraph in section.footer.paragraphs:
                        matches = re.findall(r'\{\{([^}]+)\}\}', paragraph.text)
                        markers_found.update(matches)

            return sorted(list(markers_found))

        except Exception as e:
            return [f"ERROR: {str(e)}"]

    def list_available_markers(self):
        """Lista todos los marcadores disponibles"""
        sample_data = {
            'fecha': 'Fecha formateada (ej: 16 de septiembre de 2025)',
            'cliente': 'Nombre del cliente',
            'nit': 'NIT del cliente',
            'direccion': 'Dirección del cliente',
            'referencia': 'Referencia del proyecto',
            'validez': 'Días de validez de la oferta',
            'cuadrillas': 'Número de cuadrillas',
            'operarios': 'Número de operarios',
            'plazo': 'Días de plazo de ejecución',
            'forma_pago': 'Forma de pago (se genera automáticamente según configuración)',
            'tabla_cotizacion': 'Se reemplaza por imagen de la tabla de Excel'
        }

        print("📋 MARCADORES DISPONIBLES PARA USAR EN LAS PLANTILLAS:")
        print("=" * 60)
        for marker, description in sample_data.items():
            print(f"{{{{ {marker:<18} }}}} → {description}")
        print("=" * 60)

        return sample_data


    def _prepare_replace_data(self, cotizacion, datos_adicionales):
        """Prepara los datos para reemplazar en la plantilla."""
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
            inicio = datos_adicionales.get('porcentaje_inicio', 0)
            avance = datos_adicionales.get('porcentaje_avance', 0)
            avance_requerido = datos_adicionales.get('avance_requerido', 50)
            final = datos_adicionales.get('porcentaje_final', 0)

            texto_forma_pago = f"{inicio}% al inicio de obra, {avance}% al {avance_requerido}% de avance y {final}% al finalizar la intervención"

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
            'forma_pago': texto_forma_pago
        }

        return replace_data

    def _generate_document(self, template_path, replace_data):
        """
        Genera el documento Word reemplazando los marcadores.
        """
        try:
            # Cargar la plantilla
            doc = Document(template_path)

            # Reemplazar marcadores en párrafos
            for paragraph in doc.paragraphs:
                for key, value in replace_data.items():
                    marker = f"{{{{{key}}}}}"
                    if marker in paragraph.text:
                        self._replace_text_in_paragraph(paragraph, marker, str(value))


            # Reemplazar marcadores en tablas
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for key, value in replace_data.items():
                                marker = f"{{{{{key}}}}}"
                                if marker in paragraph.text:
                                    self._replace_text_in_paragraph(paragraph, marker, str(value))

            # Reemplazar marcadores en encabezados y pies de página
            for section in doc.sections:
                if section.header:
                    for paragraph in section.header.paragraphs:
                        for key, value in replace_data.items():
                            marker = f"{{{{{key}}}}}"
                            if marker in paragraph.text:
                                self._replace_text_in_paragraph(paragraph, marker, str(value))

                if section.footer:
                    for paragraph in section.footer.paragraphs:
                        for key, value in replace_data.items():
                            marker = f"{{{{{key}}}}}"
                            if marker in paragraph.text:
                                self._replace_text_in_paragraph(paragraph, marker, str(value))

            # Generar archivo de salida
            output_dir = os.path.join(os.getcwd(), 'output')
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            output_filename = f"Cotizacion_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
            output_path = os.path.join(output_dir, output_filename)

            doc.save(output_path)

            return output_path

        except Exception as e:
            print(f"Error al generar el documento Word: {e}")
            raise

    def _replace_text_in_paragraph(self, paragraph, marker, text):
        """Reemplaza el marcador manteniendo el texto estático y el formato de la plantilla."""
        if marker in paragraph.text:
            # Buscamos el marcador en todo el texto del párrafo
            # Si el marcador está repartido en varios 'runs', este bucle lo unifica
            for run in paragraph.runs:
                if marker in run.text:
                    # Caso simple: el marcador está en un solo bloque
                    run.text = run.text.replace(marker, str(text))
                else:
                    # Caso complejo: el marcador está dividido.
                    # Forzamos el reemplazo en el contenido del párrafo
                    # pero python-docx intentará mantener el estilo base del párrafo.
                    full_text = paragraph.text.replace(marker, str(text))
                    paragraph.text = full_text
                    break