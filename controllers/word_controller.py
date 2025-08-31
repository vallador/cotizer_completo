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
        self.templates_dir = os.path.join(os.getcwd(), 'data', 'templates')
        
        # Crear directorio de plantillas si no existe
        if not os.path.exists(self.templates_dir):
            os.makedirs(self.templates_dir)
            
        # Nombres de plantillas predeterminadas
        self.template_juridica_larga = os.path.join(self.templates_dir, 'plantilla_juridica_larga.docx')
        self.template_juridica_corta = os.path.join(self.templates_dir, 'plantilla_juridica_corta.docx')
        self.template_natural_corta = os.path.join(self.templates_dir, 'plantilla_natural_corta.docx')
        
        # Crear plantillas predeterminadas si no existen
        self._create_default_templates_if_not_exist()
    
    def _create_default_templates_if_not_exist(self):
        """Crea las plantillas predeterminadas si no existen"""
        # Plantilla jurídica larga
        if not os.path.exists(self.template_juridica_larga):
            self._create_juridica_larga_template()
            
        # Plantilla jurídica corta
        if not os.path.exists(self.template_juridica_corta):
            self._create_juridica_corta_template()
            
        # Plantilla natural corta
        if not os.path.exists(self.template_natural_corta):
            self._create_natural_corta_template()
    
    def _create_juridica_larga_template(self):
        """Crea la plantilla larga para empresas jurídicas"""
        doc = Document()
        
        # Configuración de página
        sections = doc.sections
        for section in sections:
            section.page_width = Inches(8.5)
            section.page_height = Inches(11)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
        
        # Título
        title = doc.add_heading('COTIZACIÓN DE SERVICIOS', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Fecha
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run('Fecha: ').bold = True
        p.add_run('{{fecha}}')
        
        # Información del cliente
        doc.add_heading('INFORMACIÓN DEL CLIENTE', level=1)
        p = doc.add_paragraph()
        p.add_run('Cliente: ').bold = True
        p.add_run('{{cliente}}')
        p = doc.add_paragraph()
        p.add_run('NIT: ').bold = True
        p.add_run('{{nit}}')
        p = doc.add_paragraph()
        p.add_run('Dirección: ').bold = True
        p.add_run('{{direccion}}')
        
        # Referencia
        doc.add_heading('REFERENCIA', level=1)
        doc.add_paragraph('{{referencia}}')
        
        # Tabla de cotización
        doc.add_heading('DETALLE DE COTIZACIÓN', level=1)
        doc.add_paragraph('{{tabla_cotizacion}}')
        
        # Condiciones comerciales
        doc.add_heading('CONDICIONES COMERCIALES', level=1)
        
        # Validez
        p = doc.add_paragraph()
        p.add_run('Validez de la Oferta: ').bold = True
        p.add_run('{{validez}} días calendario a partir de la fecha de emisión.')
        
        # Personal
        p = doc.add_paragraph()
        p.add_run('Personal de Obra: ').bold = True
        p.add_run('{{cuadrillas}} cuadrilla(s) conformada(s) por {{operarios}} operario(s).')
        
        # Plazo
        p = doc.add_paragraph()
        p.add_run('Plazo de Ejecución: ').bold = True
        p.add_run('{{plazo}} días hábiles.')
        
        # Forma de pago
        p = doc.add_paragraph()
        p.add_run('Forma de Pago: ').bold = True
        p.add_run('{{forma_pago}}')
        
        # Notas adicionales
        doc.add_heading('NOTAS ADICIONALES', level=1)
        doc.add_paragraph('1. Los precios incluyen IVA y todos los impuestos aplicables.')
        doc.add_paragraph('2. No incluye trámites ni permisos ante entidades municipales.')
        doc.add_paragraph('3. Los materiales serán de primera calidad y cumplen con las normas técnicas vigentes.')
        doc.add_paragraph('4. Cualquier trabajo adicional no contemplado en esta cotización será objeto de una nueva propuesta.')
        
        # Firma
        doc.add_paragraph('\n\n')
        p = doc.add_paragraph('_______________________________')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = doc.add_paragraph('Firma y Sello')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Guardar plantilla
        doc.save(self.template_juridica_larga)
    
    def _create_juridica_corta_template(self):
        """Crea la plantilla corta para empresas jurídicas"""
        doc = Document()

        # Título
        title = doc.add_heading('COTIZACIÓN DE SERVICIOS', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Fecha e información básica
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run('Fecha: ').bold = True
        p.add_run('{{fecha}}')

        # Información del cliente (resumida)
        p = doc.add_paragraph()
        p.add_run('Cliente: ').bold = True
        p.add_run('{{cliente}} - NIT: {{nit}}')

        # Referencia
        p = doc.add_paragraph()
        p.add_run('Referencia: ').bold = True
        p.add_run('{{referencia}}')

        # Tabla de cotización
        doc.add_heading('DETALLE DE COTIZACIÓN', level=1)
        doc.add_paragraph('{{tabla_cotizacion}}')

        # Condiciones resumidas
        doc.add_heading('CONDICIONES', level=1)

        # Lista de condiciones
        p = doc.add_paragraph('• Validez: ')
        p.add_run('{{validez}} días')

        p = doc.add_paragraph('• Personal: ')
        p.add_run('{{cuadrillas}} cuadrilla(s) - {{operarios}} operario(s)')

        p = doc.add_paragraph('• Plazo: ')
        p.add_run('{{plazo}} días hábiles')

        p = doc.add_paragraph('• Forma de Pago: ')
        p.add_run('{{forma_pago}}')

        # Firma
        doc.add_paragraph('\n')
        p = doc.add_paragraph('_______________________________')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = doc.add_paragraph('Firma y Sello')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Guardar plantilla
        doc.save(self.template_juridica_corta)
    
    def _create_natural_corta_template(self):
        """Crea la plantilla corta para personas naturales"""
        doc = Document()
        
        # Título
        title = doc.add_heading('COTIZACIÓN DE SERVICIOS', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Fecha e información básica
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.add_run('Fecha: ').bold = True
        p.add_run('{{fecha}}')
        
        # Información del cliente (solo nombre)
        p = doc.add_paragraph()
        p.add_run('Cliente: ').bold = True
        p.add_run('{{cliente}}')
        
        # Referencia
        p = doc.add_paragraph()
        p.add_run('Referencia: ').bold = True
        p.add_run('{{referencia}}')
        
        # Tabla de cotización
        doc.add_heading('DETALLE DE COTIZACIÓN', level=1)
        doc.add_paragraph('{{tabla_cotizacion}}')
        
        # Condiciones resumidas
        doc.add_heading('CONDICIONES', level=1)
        
        # Lista de condiciones
        p = doc.add_paragraph('• Validez: ')
        p.add_run('{{validez}} días')
        
        p = doc.add_paragraph('• Personal: ')
        p.add_run('{{cuadrillas}} cuadrilla(s) - {{operarios}} operario(s)')
        
        p = doc.add_paragraph('• Plazo: ')
        p.add_run('{{plazo}} días hábiles')
        
        p = doc.add_paragraph('• Forma de Pago: ')
        p.add_run('{{forma_pago}}')
        
        # Notas adicionales
        doc.add_paragraph('Nota: Los precios incluyen IVA y todos los impuestos aplicables.')
        
        # Firma
        doc.add_paragraph('\n')
        p = doc.add_paragraph('_______________________________')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p = doc.add_paragraph('Firma')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Guardar plantilla
        doc.save(self.template_natural_corta)
    
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
            temp_image_path = os.path.join(tempfile.gettempdir(), f"tabla_cotizacion_{datetime.now().strftime('%Y%m%d%H%M%S')}.png")
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
    
    def generate_natural_cotizacion(self, config, excel_path):
        """Genera una cotización para una persona natural."""
        doc = Document()

        # Estilos (opcional, pero recomendado para consistencia)
        # ... (puedes definir estilos aquí)

        # --- Contenido del documento ---
        doc.add_heading(config["titulo"], level=1)

        p = doc.add_paragraph()
        p.add_run("Fecha: ").bold = True
        p.add_run(config["fecha"])

        p = doc.add_paragraph()
        p.add_run("Lugar de intervención: ").bold = True
        p.add_run(config["lugar"])

        doc.add_heading("1. Concepto General del Trabajo", level=2)
        doc.add_paragraph(config["concepto"])

        doc.add_heading("2. Condiciones Comerciales", level=2)

        p = doc.add_paragraph()
        p.add_run("Validez de la oferta: ").bold = True
        p.add_run(f"{config["validez"]} días")

        p = doc.add_paragraph()
        p.add_run("Personal de obra: ").bold = True
        p.add_run(f"{config["cuadrillas"]} cuadrilla(s) con {config["operarios_num"]} ({config["operarios_letra"]}) operarios.")

        p = doc.add_paragraph()
        p.add_run("Plazo de ejecución: ").bold = True
        p.add_run(f"{config["plazo_dias"]} días {config["plazo_tipo"]}")

        p = doc.add_paragraph()
        p.add_run("Forma de pago: ").bold = True
        if config["pago_personalizado"]:
            p.add_run(config["pago_personalizado"])
        elif config["pago_contraentrega"]:
            p.add_run("Contraentrega")
        elif config["pago_porcentajes"]:
            pago_str = f"{config["anticipo"]}% de anticipo, {config["avance"]}% al {config["avance_requerido"]}% de avance y {config["final"]}% al finalizar."
            p.add_run(pago_str)

        doc.add_heading("3. Detalle de la Cotización", level=2)
        # Aquí puedes insertar la tabla de Excel como imagen
        if excel_path and os.path.exists(excel_path):
            try:
                table_image_path = self._capture_excel_table(excel_path)
                if table_image_path:
                    doc.add_picture(table_image_path, width=Inches(6.0))
                    os.remove(table_image_path) # Limpiar imagen temporal
            except Exception as e:
                print(f"Error al insertar la tabla de Excel: {e}")
                doc.add_paragraph("[Error al generar la tabla de cotización]")
        else:
            doc.add_paragraph("[No se encontró el archivo de cotización en Excel]")

        doc.add_heading("4. Notas Adicionales", level=2)
        if config["incluir_iva"]:
            doc.add_paragraph("Los precios presentados en esta cotización incluyen IVA.")
        if config["incluir_materiales"]:
            doc.add_paragraph("Los materiales a utilizar serán de primera calidad y cumplirán con las normas técnicas vigentes.")

        # --- Guardar el documento ---
        output_dir = os.path.join(os.getcwd(), "output")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        filename = f"Cotizacion_Natural_{config["lugar"].replace(" ", "_")}_{datetime.now().strftime("%Y%m%d")}.docx"
        output_path = os.path.join(output_dir, filename)

        doc.save(output_path)
        return output_path

    def generate_juridica_cotizacion(self, config, excel_path):
        """Genera una cotización para una persona jurídica."""
        doc = Document()

        # --- Contenido del documento ---
        # 1. Portadas (si se seleccionan)
        if config["secciones_incluir"].get("portadas"):
            self._add_portadas(doc, config)

        # 2. Contenido de Separadores (si se seleccionan)
        if config["secciones_incluir"].get("contenido_separadores"):
            self._add_contenido_separadores(doc, config)

        # 3. Carta de Presentación (si se selecciona)
        if config["secciones_incluir"].get("carta_presentacion"):
            self._add_carta_presentacion(doc, config)

        # 4. Páginas Estándar (Pólizas, Personal, Director de Obra) (si se seleccionan)
        if config["secciones_incluir"].get("paginas_estandar"):
            self._add_paginas_estandar(doc, config)

        # 5. Cuadro de Experiencia (si se selecciona)
        if config["secciones_incluir"].get("cuadro_experiencia"):
            self._add_cuadro_experiencia(doc, config)

        # 6. Certificados de Trabajos Similares (si se selecciona)
        if config["secciones_incluir"].get("certificados_trabajos"):
            self._add_certificados_trabajos(doc, config)

        # 7. Información de Seguridad de Alturas (si se selecciona)
        if config["secciones_incluir"].get("seguridad_alturas"):
            self._add_seguridad_alturas(doc, config)

        # 8. Resumen del Programa de Prevención y Protección contra Caídas en Alturas (si se selecciona)
        if config["secciones_incluir"].get("programa_prevencion"):
            self._add_programa_prevencion(doc, config)

        # 9. Estándares Mínimos de SGSST Certificado Ministerio (si se selecciona)
        if config["secciones_incluir"].get("sgsst_certificado"):
            self._add_sgsst_certificado(doc, config)

        # 10. Presupuesto de Obra y Programación de Obra (si se selecciona)
        if config["secciones_incluir"].get("presupuesto_programacion"):
            self._add_presupuesto_programacion(doc, config, excel_path)

        # 11. Documentación Legal y Financiera (si se selecciona)
        if config["secciones_incluir"].get("documentacion_legal"):
            self._add_documentacion_legal(doc, config)

        # 12. Anexos (si se seleccionan)
        if config["secciones_incluir"].get("anexos"):
            self._add_anexos(doc, config)

        # --- Combinar archivos preexistentes ---
        if config["archivos_combinar"]:
            self._combine_external_files(doc, config["archivos_combinar"])

        # --- Guardar el documento ---
        output_dir = os.path.join(os.getcwd(), "output")
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        filename = f"Cotizacion_Juridica_{config["lugar"].replace(" ", "_")}_{datetime.now().strftime("%Y%m%d")}.docx"
        output_path = os.path.join(output_dir, filename)

        doc.save(output_path)
        return output_path

    def _add_portadas(self, doc, config):
        doc.add_page_break()
        doc.add_heading("PORTADA DE LA COTIZACIÓN", level=0).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"\n\n\nProyecto: {config["titulo"]}").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Lugar: {config["lugar"]}").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Fecha: {config["fecha"]}").alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_page_break()

    def _add_contenido_separadores(self, doc, config):
        doc.add_page_break()
        doc.add_heading("CONTENIDO DE SEPARADORES", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Aquí iría el contenido de los separadores. La paginación se manejará al final.")
        doc.add_page_break()

    def _add_carta_presentacion(self, doc, config):
        doc.add_page_break()
        doc.add_heading("CARTA DE PRESENTACIÓN", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"Bogotá D.C., {config["fecha"]}")
        doc.add_paragraph("\nSeñores")
        doc.add_paragraph(config["lugar"])
        doc.add_paragraph("Ciudad")
        doc.add_paragraph("\nAsunto: Presentación de Propuesta para " + config["titulo"])
        doc.add_paragraph("\nCordial Saludo,")
        doc.add_paragraph("Por medio de la presente, nos permitimos presentar a su consideración la propuesta para la ejecución de los trabajos de " + config["concepto"] + ", en las instalaciones de " + config["lugar"] + ", de acuerdo a los requerimientos solicitados.")
        doc.add_paragraph("\nEsperamos que la presente propuesta sea de su agrado y quedamos a su disposición para cualquier aclaración.")
        doc.add_paragraph("\nAtentamente,")
        doc.add_paragraph("\n_______________________________")
        doc.add_paragraph("Firma y Sello de la Empresa")
        doc.add_page_break()

    def _add_paginas_estandar(self, doc, config):
        doc.add_page_break()
        doc.add_heading("PÁGINAS ESTÁNDAR", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Aquí se incluirían las páginas estándar con información de pólizas, personal y director de obra.")

        doc.add_heading("Pólizas", level=2)
        for poliza_key, checked in config["polizas_incluir"].items():
            if checked:
                doc.add_paragraph(f"- {poliza_key.replace("_", " ").title()}: [Texto de la póliza correspondiente]")

        doc.add_heading("Personal de Obra", level=2)
        doc.add_paragraph(f"Cantidad de Personal de Obra: {config["cuadrillas"]} cuadrilla(s) con {config["operarios_num"]} ({config["operarios_letra"]}) operarios.")

        doc.add_heading("Personal Técnico", level=2)
        if config["director_obra"]:
            doc.add_paragraph(f"Director de Obra: {config["director_obra"]}")
        if config["residente_obra"]:
            doc.add_paragraph(f"Residente de Obra: {config["residente_obra"]}")
        if config["tecnologo_sgsst"]:
            doc.add_paragraph(f"Tecnólogo SGSST: {config["tecnologo_sgsst"]}")
        doc.add_page_break()

    def _add_cuadro_experiencia(self, doc, config):
        doc.add_page_break()
        doc.add_heading("CUADRO DE EXPERIENCIA", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Aquí se insertaría el cuadro de experiencia con material fotográfico.")
        doc.add_page_break()

    def _add_certificados_trabajos(self, doc, config):
        doc.add_page_break()
        doc.add_heading("CERTIFICADOS DE TRABAJOS SIMILARES", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Aquí se adjuntarían los certificados de trabajos similares.")
        doc.add_page_break()

    def _add_seguridad_alturas(self, doc, config):
        doc.add_page_break()
        doc.add_heading("INFORMACIÓN DE SEGURIDAD DE ALTURAS", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Aquí se incluiría la información de seguridad de alturas.")
        doc.add_page_break()

    def _add_programa_prevencion(self, doc, config):
        doc.add_page_break()
        doc.add_heading("RESUMEN DEL PROGRAMA DE PREVENCIÓN Y PROTECCIÓN CONTRA CAÍDAS EN ALTURAS", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Aquí se incluiría el resumen del programa de prevención y protección contra caídas en alturas.")
        doc.add_page_break()

    def _add_sgsst_certificado(self, doc, config):
        doc.add_page_break()
        doc.add_heading("ESTÁNDARES MÍNIMOS DE SGSST CERTIFICADO MINISTERIO", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Aquí se incluirían los estándares mínimos de SGSST certificado por el ministerio.")
        doc.add_page_break()

    def _add_presupuesto_programacion(self, doc, config, excel_path):
        doc.add_page_break()
        doc.add_heading("PRESUPUESTO DE OBRA Y PROGRAMACIÓN DE OBRA", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Aquí se insertaría el presupuesto de obra y la programación de obra.")
        # Reutilizar la lógica de la tabla de Excel
        if excel_path and os.path.exists(excel_path):
            try:
                table_image_path = self._capture_excel_table(excel_path)
                if table_image_path:
                    doc.add_picture(table_image_path, width=Inches(6.0))
                    os.remove(table_image_path) # Limpiar imagen temporal
            except Exception as e:
                print(f"Error al insertar la tabla de Excel en cotización jurídica: {e}")
                doc.add_paragraph("[Error al generar la tabla de cotización]")
        else:
            doc.add_paragraph("[No se encontró el archivo de cotización en Excel]")
        doc.add_page_break()

    def _add_documentacion_legal(self, doc, config):
        doc.add_page_break()
        doc.add_heading("DOCUMENTACIÓN LEGAL Y FINANCIERA", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Aquí se incluiría la documentación legal y financiera.")
        doc.add_page_break()

    def _add_anexos(self, doc, config):
        doc.add_page_break()
        doc.add_heading("ANEXOS", level=1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("Aquí se incluirían los anexos.")
        doc.add_page_break()

    def _combine_external_files(self, main_doc, file_paths):
        """Combina archivos externos (DOCX y PDF) en el documento principal."""
        for file_path in file_paths:
            if file_path.lower().endswith(".docx"):
                try:
                    sub_doc = Document(file_path)
                    for element in sub_doc.element.body:
                        main_doc.element.body.append(element)
                    main_doc.add_page_break() # Añadir salto de página entre documentos
                except Exception as e:
                    print(f"Error al combinar DOCX {file_path}: {e}")
                    main_doc.add_paragraph(f"[Error al insertar archivo DOCX: {os.path.basename(file_path)}]")
            elif file_path.lower().endswith(".pdf"):
                # Para PDF, se necesitaría una librería externa como PyPDF2 para leer y luego convertir a imagen o texto
                # Esto es más complejo y requeriría un renderizado o extracción de texto.
                # Por ahora, solo se añade una nota.
                main_doc.add_paragraph(f"[Archivo PDF '{os.path.basename(file_path)}' no se puede insertar directamente. Se requiere procesamiento adicional.]")
                main_doc.add_page_break()
            else:
                main_doc.add_paragraph(f"[Tipo de archivo no soportado para combinación: {os.path.basename(file_path)}]")
                main_doc.add_page_break()

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


