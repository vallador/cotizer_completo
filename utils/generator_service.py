import os
from datetime import datetime
from controllers.excel_controller import ExcelController
from controllers.word_controller import WordController
from utils.excel_to_word import ExcelToWordAutomation


class GeneratorService:
    def __init__(self):
        # Inicializamos los controladores
        self.excel_ctrl = ExcelController()
        self.word_ctrl = WordController()
        self.automation = ExcelToWordAutomation()

    def generate_quotation(self, cotizacion_data):
        """
        Método principal que orquesta la generación según el tipo de persona.
        """
        tipo_persona = cotizacion_data.get('tipo_persona', 'natural').lower()

        # 1. PASO COMÚN: Generar el Excel (Base de datos y cálculos)
        # Se asume que generate_excel devuelve la ruta del archivo creado
        excel_path = self.excel_ctrl.generate_excel(
            items=cotizacion_data['items'],
            activities=cotizacion_data['activities'],
            tipo_persona=tipo_persona,
            administracion=cotizacion_data.get('administracion', 0),
            imprevistos=cotizacion_data.get('imprevistos', 0),
            utilidad=cotizacion_data.get('utilidad', 0),
            iva_utilidad=cotizacion_data.get('iva_utilidad', 0),
            nombre_cliente=cotizacion_data.get('nombre_cliente', "Cliente")
        )

        if not excel_path:
            return False, "Error al generar el archivo Excel base."

        # 2. BIFURCACIÓN DE LÓGICA (Natural vs Jurídica)
        if tipo_persona == "juridica":
            return self._handle_juridica_flow(excel_path, cotizacion_data)
        else:
            return self._handle_natural_flow(excel_path, cotizacion_data)

    def _handle_juridica_flow(self, excel_path, data):
        """
        Flujo complejo: Usa win32com para pegar tabla con formatos en Plantilla Base.
        """
        print("Iniciando flujo de Persona Jurídica (Excel -> Word Automation)...")

        template_path = "templates/plantilla_base.docx"
        pdf_output = f"outputs/Cotizacion_Juridica_{datetime.now().strftime('%Y%m%d')}.pdf"

        # Usamos el método de tu clase ExcelToWordAutomation que maneja win32com
        success, message = self.automation.insertar_tabla_y_convertir_pdf(
            excel_path=excel_path,
            word_path=template_path,
            pdf_output_path=pdf_output
        )

        return success, f"PDF Jurídico generado: {pdf_output}" if success else message

    def _handle_natural_flow(self, excel_path, data):
        """
        Flujo simple: Usa python-docx para reemplazo de marcadores.
        """
        print("Iniciando flujo de Persona Natural (Marcadores)...")

        template_path = "templates/persona_natural_corta.docx"

        try:
            # Usamos el WordController para reemplazar {{nombre}}, {{fecha}}, etc.
            word_path = self.word_ctrl.generate_word_document(
                cotizacion_id=data.get('id'),
                excel_path=excel_path,
                datos_adicionales=data,
                template_personalizada=template_path
            )
            return True, f"Word Natural generado: {word_path}"
        except Exception as e:
            return False, f"Error en flujo Natural: {str(e)}"