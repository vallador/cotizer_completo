from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFormLayout, QCheckBox, QMessageBox, QFileDialog, QGroupBox, QComboBox, QSpinBox, QRadioButton, QButtonGroup
from PyQt5.QtCore import Qt
import os
from controllers.word_controller import WordController

class WordConfigDialog(QDialog):
    def __init__(self, cotizacion_controller, cotizacion_id, excel_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configuración de Documento Word")
        self.setMinimumWidth(600)
        
        self.cotizacion_controller = cotizacion_controller
        self.cotizacion_id = cotizacion_id
        self.excel_path = excel_path
        self.word_controller = WordController(cotizacion_controller)
        
        # Obtener la cotización
        self.cotizacion = cotizacion_controller.obtener_cotizacion(cotizacion_id)
        if not self.cotizacion:
            QMessageBox.critical(self, "Error", f"No se encontró la cotización con ID {cotizacion_id}")
            self.reject()
            return
        
        # Crear layout principal
        main_layout = QVBoxLayout(self)
        
        # Grupo de formato de documento
        format_group = QGroupBox("Formato del Documento")
        format_layout = QVBoxLayout()
        
        # Opciones de formato
        self.formato_largo_radio = QRadioButton("Formato Largo (Detallado)")
        self.formato_corto_radio = QRadioButton("Formato Corto (Resumido)")
        
        # Grupo de botones para formato
        formato_group = QButtonGroup(self)
        formato_group.addButton(self.formato_largo_radio)
        formato_group.addButton(self.formato_corto_radio)
        
        # Seleccionar formato predeterminado según tipo de cliente
        if self.cotizacion.cliente.tipo.lower() == 'jurídica':
            self.formato_largo_radio.setChecked(True)
        else:
            self.formato_corto_radio.setChecked(True)
            # Deshabilitar formato largo para clientes naturales
            self.formato_largo_radio.setEnabled(False)
        
        format_layout.addWidget(self.formato_largo_radio)
        format_layout.addWidget(self.formato_corto_radio)
        format_group.setLayout(format_layout)
        main_layout.addWidget(format_group)
        
        # Grupo de información del trabajo
        info_group = QGroupBox("Información del Trabajo")
        info_layout = QFormLayout()
        
        self.referencia_input = QLineEdit("Servicios de construcción y mantenimiento")
        self.validez_spin = QSpinBox()
        self.validez_spin.setRange(1, 180)
        self.validez_spin.setValue(30)
        
        info_layout.addRow("Referencia del Trabajo:", self.referencia_input)
        info_layout.addRow("Validez de la Oferta (días):", self.validez_spin)
        
        info_group.setLayout(info_layout)
        main_layout.addWidget(info_group)
        
        # Grupo de personal y plazos
        personal_group = QGroupBox("Personal y Plazos")
        personal_layout = QFormLayout()
        
        self.cuadrillas_combo = QComboBox()
        self.cuadrillas_combo.addItems(["una", "dos", "tres"])
        
        self.operarios_spin = QSpinBox()
        self.operarios_spin.setRange(1, 20)
        self.operarios_spin.setValue(2)
        
        self.plazo_spin = QSpinBox()
        self.plazo_spin.setRange(1, 365)
        self.plazo_spin.setValue(15)
        
        personal_layout.addRow("Número de Cuadrillas:", self.cuadrillas_combo)
        personal_layout.addRow("Número de Operarios:", self.operarios_spin)
        personal_layout.addRow("Plazo de Ejecución (días hábiles):", self.plazo_spin)
        
        personal_group.setLayout(personal_layout)
        main_layout.addWidget(personal_group)
        
        # Grupo de forma de pago
        pago_group = QGroupBox("Forma de Pago")
        pago_layout = QVBoxLayout()
        
        # Opciones de pago
        self.pago_contraentrega_radio = QRadioButton("Contraentrega")
        self.pago_porcentajes_radio = QRadioButton("Porcentajes")
        
        # Grupo de botones para forma de pago
        pago_button_group = QButtonGroup(self)
        pago_button_group.addButton(self.pago_contraentrega_radio)
        pago_button_group.addButton(self.pago_porcentajes_radio)
        
        self.pago_contraentrega_radio.setChecked(True)
        
        pago_layout.addWidget(self.pago_contraentrega_radio)
        pago_layout.addWidget(self.pago_porcentajes_radio)
        
        # Grupo de porcentajes (inicialmente oculto)
        self.porcentajes_group = QGroupBox("Desglose de Porcentajes")
        porcentajes_layout = QFormLayout()
        
        self.porcentaje_inicio_spin = QSpinBox()
        self.porcentaje_inicio_spin.setRange(0, 100)
        self.porcentaje_inicio_spin.setValue(30)
        
        self.porcentaje_avance_spin = QSpinBox()
        self.porcentaje_avance_spin.setRange(0, 100)
        self.porcentaje_avance_spin.setValue(40)
        
        self.avance_requerido_spin = QSpinBox()
        self.avance_requerido_spin.setRange(1, 99)
        self.avance_requerido_spin.setValue(50)
        
        self.porcentaje_final_spin = QSpinBox()
        self.porcentaje_final_spin.setRange(0, 100)
        self.porcentaje_final_spin.setValue(30)
        
        porcentajes_layout.addRow("Porcentaje al Inicio (%):", self.porcentaje_inicio_spin)
        porcentajes_layout.addRow("Porcentaje al Avance (%):", self.porcentaje_avance_spin)
        porcentajes_layout.addRow("Avance Requerido (%):", self.avance_requerido_spin)
        porcentajes_layout.addRow("Porcentaje al Finalizar (%):", self.porcentaje_final_spin)
        
        self.porcentajes_group.setLayout(porcentajes_layout)
        self.porcentajes_group.setVisible(False)
        
        # Conectar cambio de forma de pago
        self.pago_contraentrega_radio.toggled.connect(self.toggle_porcentajes)
        
        pago_layout.addWidget(self.porcentajes_group)
        pago_group.setLayout(pago_layout)
        main_layout.addWidget(pago_group)
        
        # Botones de acción
        buttons_layout = QHBoxLayout()
        
        self.generate_btn = QPushButton("Generar Documento")
        self.generate_btn.clicked.connect(self.generate_document)
        
        self.cancel_btn = QPushButton("Cancelar")
        self.cancel_btn.clicked.connect(self.reject)
        
        buttons_layout.addWidget(self.generate_btn)
        buttons_layout.addWidget(self.cancel_btn)
        
        main_layout.addLayout(buttons_layout)
    
    def toggle_porcentajes(self, checked):
        """Muestra u oculta el grupo de porcentajes según la selección"""
        self.porcentajes_group.setVisible(not checked)
    
    def generate_document(self):
        """Genera el documento Word con la configuración actual"""
        try:
            # Determinar formato
            formato = 'largo' if self.formato_largo_radio.isChecked() else 'corto'
            
            # Determinar forma de pago
            forma_pago = 'contraentrega' if self.pago_contraentrega_radio.isChecked() else 'porcentajes'
            
            # Preparar datos adicionales
            datos_adicionales = {
                'referencia': self.referencia_input.text(),
                'validez': self.validez_spin.value(),
                'cuadrillas': self.cuadrillas_combo.currentText(),
                'operarios': self.operarios_spin.value(),
                'plazo': self.plazo_spin.value(),
                'forma_pago': forma_pago
            }
            
            # Agregar porcentajes si aplica
            if forma_pago == 'porcentajes':
                datos_adicionales.update({
                    'porcentaje_inicio': self.porcentaje_inicio_spin.value(),
                    'porcentaje_avance': self.porcentaje_avance_spin.value(),
                    'avance_requerido': self.avance_requerido_spin.value(),
                    'porcentaje_final': self.porcentaje_final_spin.value()
                })
                
                # Validar que los porcentajes sumen 100%
                total = (datos_adicionales['porcentaje_inicio'] + 
                         datos_adicionales['porcentaje_avance'] + 
                         datos_adicionales['porcentaje_final'])
                
                if total != 100:
                    QMessageBox.warning(self, "Advertencia", 
                                       f"Los porcentajes suman {total}%. Se recomienda que sumen 100%.")
                    reply = QMessageBox.question(self, "Continuar", 
                                               "¿Desea continuar de todos modos?",
                                               QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                    if reply == QMessageBox.No:
                        return
            
            # Generar el documento
            word_path = self.word_controller.generate_word_document(
                self.cotizacion_id,
                self.excel_path,
                datos_adicionales,
                formato
            )
            
            if word_path and os.path.exists(word_path):
                QMessageBox.information(self, "Éxito", f"Documento generado correctamente en:\n{word_path}")
                
                # Preguntar si desea enviar por correo
                reply = QMessageBox.question(self, "Enviar por Correo", 
                                           "¿Desea enviar el documento por correo electrónico?",
                                           QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                
                if reply == QMessageBox.Yes:
                    # Importar aquí para evitar importación circular
                    from views.email_dialog import SendEmailDialog
                    
                    # Abrir diálogo de envío de correo
                    email_dialog = SendEmailDialog([word_path, self.excel_path], self)
                    email_dialog.exec_()
                
                self.accept()
            else:
                QMessageBox.critical(self, "Error", "No se pudo generar el documento Word.")
                
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al generar el documento: {str(e)}")
