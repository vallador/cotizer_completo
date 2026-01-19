import sys
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QLabel, QLineEdit, QTextEdit, QComboBox, QCheckBox,
                             QPushButton, QGroupBox, QScrollArea, QWidget,
                             QDateEdit, QSpinBox, QFileDialog, QListWidget,
                             QMessageBox, QTabWidget, QGridLayout)
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QFont, QColor
import os


class ImprovedWordConfigDialog(QDialog):
    def __init__(self, parent=None, client_type="natural", client_data=None, precarga_data=None):
        super().__init__(parent)
        self.client_type = client_type
        self.client_data = client_data or {}
        self.precarga_data = precarga_data or {}
        self.selected_files = []

        self.setWindowTitle("Configuración de Documento Word - Cotización")
        self.setModal(True)
        self.resize(800, 600)

        self.setup_ui()
        self.precargar_datos()
        self.load_default_values()
        print("Tipo de cliente:", self.client_type)

    def setup_ui(self):
        """Configura la interfaz de usuario"""
        main_layout = QVBoxLayout()
        self.setLayout(main_layout)

        # Título
        title_label = QLabel(f"Configuración para Cliente {self.client_type.title()}")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # Crear pestañas
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # Pestaña de información básica
        self.create_basic_info_tab()

        # Pestaña de condiciones comerciales
        self.create_commercial_conditions_tab()

        # Pestaña específica según tipo de cliente
        if self.client_type == 'juridica':
            self.create_juridica_specific_tab()
        else:
            self.create_natural_specific_tab()

        # Botones
        buttons_layout = QHBoxLayout()

        self.cancel_button = QPushButton("Cancelar")
        self.cancel_button.clicked.connect(self.reject)
        buttons_layout.addWidget(self.cancel_button)

        buttons_layout.addStretch()

        self.generate_button = QPushButton("Generar Documento")
        self.generate_button.clicked.connect(self.validate_and_accept)
        self.generate_button.setDefault(True)
        buttons_layout.addWidget(self.generate_button)

        main_layout.addLayout(buttons_layout)

    def create_basic_info_tab(self):
        """Crea la pestaña de información básica"""
        tab = QWidget()
        layout = QFormLayout()

        # Fecha actual
        self.fecha_edit = QDateEdit(QDate.currentDate())
        self.fecha_edit.setCalendarPopup(True)
        layout.addRow("Fecha del documento:", self.fecha_edit)

        # Referencia del trabajo
        self.referencia_edit = QLineEdit()
        self.referencia_edit.setPlaceholderText("Ej: Mantenimiento de fachadas - Conjunto Residencial")
        layout.addRow("Referencia del trabajo:*", self.referencia_edit)

        # Título del trabajo
        self.titulo_edit = QLineEdit()
        self.titulo_edit.setPlaceholderText("Ej: COTIZACIÓN PARA MANTENIMIENTO DE FACHADAS")
        layout.addRow("Título del trabajo:*", self.titulo_edit)

        # Lugar de intervención
        self.lugar_edit = QLineEdit()
        self.lugar_edit.setPlaceholderText("Ej: Conjunto Residencial Las Flores")
        layout.addRow("Lugar de intervención:*", self.lugar_edit)

        # Concepto general del trabajo
        self.concepto_edit = QTextEdit()
        self.concepto_edit.setMaximumHeight(80)
        self.concepto_edit.setPlaceholderText("Descripción general del trabajo a realizar...")
        layout.addRow("Concepto general:*", self.concepto_edit)

        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "Información Básica")

    def precargar_datos(self):
        """Precarga los datos del cliente en los campos correspondientes"""
        if self.precarga_data:
            # Precargar referencia si está disponible
            if 'referencia' in self.precarga_data and hasattr(self, 'referencia_edit'):
                self.referencia_edit.setText(self.precarga_data['referencia'])

            # Precargar título si está disponible
            if 'titulo' in self.precarga_data and hasattr(self, 'titulo_edit'):
                self.titulo_edit.setText(self.precarga_data['titulo'])

            # Precargar lugar si está disponible
            if 'lugar' in self.precarga_data and hasattr(self, 'lugar_edit'):
                self.lugar_edit.setText(self.precarga_data['lugar'])

            # Precargar concepto si está disponible
            if 'concepto' in self.precarga_data and hasattr(self, 'concepto_edit'):
                self.concepto_edit.setText(self.precarga_data['concepto'])

    def create_commercial_conditions_tab(self):
        """Crea la pestaña de condiciones comerciales"""
        tab = QWidget()
        layout = QFormLayout()

        # Validez de la oferta
        self.validez_edit = QSpinBox()
        self.validez_edit.setRange(1, 365)
        self.validez_edit.setValue(30)
        self.validez_edit.setSuffix(" días")
        layout.addRow("Validez de la oferta:", self.validez_edit)

        # Personal de obra
        personal_group = QGroupBox("Personal de Obra")
        personal_layout = QGridLayout()

        self.cuadrillas_edit = QSpinBox()
        self.cuadrillas_edit.setRange(1, 10)
        self.cuadrillas_edit.setValue(1)
        personal_layout.addWidget(QLabel("Número de cuadrillas:*"), 0, 0)
        personal_layout.addWidget(self.cuadrillas_edit, 0, 1)

        self.operarios_num_edit = QSpinBox()
        self.operarios_num_edit.setRange(1, 20)
        self.operarios_num_edit.setValue(2)
        personal_layout.addWidget(QLabel("Operarios por cuadrilla:*"), 1, 0)
        personal_layout.addWidget(self.operarios_num_edit, 1, 1)

        self.operarios_letra_edit = QLineEdit()
        self.operarios_letra_edit.setPlaceholderText("Ej: dos")
        personal_layout.addWidget(QLabel("Operarios (en letra):"), 2, 0)
        personal_layout.addWidget(self.operarios_letra_edit, 2, 1)

        personal_group.setLayout(personal_layout)
        layout.addRow(personal_group)

        # Plazo de ejecución
        plazo_group = QGroupBox("Plazo de Ejecución")
        plazo_layout = QGridLayout()

        self.plazo_dias_edit = QSpinBox()
        self.plazo_dias_edit.setRange(1, 365)
        self.plazo_dias_edit.setValue(15)
        plazo_layout.addWidget(QLabel("Días:*"), 0, 0)
        plazo_layout.addWidget(self.plazo_dias_edit, 0, 1)

        self.plazo_tipo_combo = QComboBox()
        self.plazo_tipo_combo.addItems(["hábiles", "calendario"])
        plazo_layout.addWidget(QLabel("Tipo:"), 0, 2)
        plazo_layout.addWidget(self.plazo_tipo_combo, 0, 3)

        plazo_group.setLayout(plazo_layout)
        layout.addRow(plazo_group)

        # Forma de pago
        pago_group = QGroupBox("Forma de Pago")
        pago_layout = QVBoxLayout()

        self.pago_contraentrega_radio = QCheckBox("Contraentrega")
        self.pago_contraentrega_radio.toggled.connect(self.toggle_payment_details)
        pago_layout.addWidget(self.pago_contraentrega_radio)

        self.pago_porcentajes_radio = QCheckBox("Por porcentajes")
        self.pago_porcentajes_radio.toggled.connect(self.toggle_payment_details)
        pago_layout.addWidget(self.pago_porcentajes_radio)

        # Detalles de pago por porcentajes
        self.pago_details_widget = QWidget()
        pago_details_layout = QGridLayout()

        self.anticipo_edit = QSpinBox()
        self.anticipo_edit.setRange(0, 100)
        self.anticipo_edit.setSuffix("%")
        pago_details_layout.addWidget(QLabel("Anticipo:"), 0, 0)
        pago_details_layout.addWidget(self.anticipo_edit, 0, 1)

        self.avance_edit = QSpinBox()
        self.avance_edit.setRange(0, 100)
        self.avance_edit.setSuffix("%")
        pago_details_layout.addWidget(QLabel("Al % de avance:"), 1, 0)
        pago_details_layout.addWidget(self.avance_edit, 1, 1)

        self.avance_requerido_edit = QSpinBox()
        self.avance_requerido_edit.setRange(1, 100)
        self.avance_requerido_edit.setValue(50)
        self.avance_requerido_edit.setSuffix("%")
        pago_details_layout.addWidget(QLabel("Cuando se alcance:"), 1, 2)
        pago_details_layout.addWidget(self.avance_requerido_edit, 1, 3)

        self.final_edit = QSpinBox()
        self.final_edit.setRange(0, 100)
        self.final_edit.setSuffix("%")
        pago_details_layout.addWidget(QLabel("Al finalizar:"), 2, 0)
        pago_details_layout.addWidget(self.final_edit, 2, 1)

        self.pago_details_widget.setLayout(pago_details_layout)
        self.pago_details_widget.setEnabled(False)
        pago_layout.addWidget(self.pago_details_widget)

        # Campo de texto libre para forma de pago personalizada
        self.pago_personalizado_edit = QTextEdit()
        self.pago_personalizado_edit.setMaximumHeight(60)
        self.pago_personalizado_edit.setPlaceholderText("Escriba aquí una forma de pago personalizada...")
        pago_layout.addWidget(QLabel("Forma de pago personalizada:"))
        pago_layout.addWidget(self.pago_personalizado_edit)

        pago_group.setLayout(pago_layout)
        layout.addRow(pago_group)

        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "Condiciones Comerciales")

    def create_natural_specific_tab(self):
        """Crea la pestaña específica para persona natural"""
        tab = QWidget()
        layout = QVBoxLayout()

        info_label = QLabel("Configuración específica para Persona Natural")
        info_label.setStyleSheet("font-weight: bold; color: #2E8B57;")
        layout.addWidget(info_label)

        # Formato simple
        format_group = QGroupBox("Formato del Documento")
        format_layout = QVBoxLayout()

        self.formato_simple_check = QCheckBox("Usar formato simplificado (recomendado)")
        self.formato_simple_check.setChecked(True)
        format_layout.addWidget(self.formato_simple_check)

        info_text = QLabel(
            "El formato simplificado excluye información empresarial compleja y se enfoca en los aspectos esenciales de la cotización.")
        info_text.setWordWrap(True)
        info_text.setStyleSheet("color: #666; font-style: italic;")
        format_layout.addWidget(info_text)

        format_group.setLayout(format_layout)
        layout.addWidget(format_group)

        # Información adicional
        adicional_group = QGroupBox("Información Adicional")
        adicional_layout = QFormLayout()

        self.incluir_iva_check = QCheckBox("Los precios incluyen IVA")
        self.incluir_iva_check.setChecked(True)
        adicional_layout.addRow(self.incluir_iva_check)

        self.incluir_materiales_check = QCheckBox("Incluir nota sobre materiales")
        self.incluir_materiales_check.setChecked(True)
        adicional_layout.addRow(self.incluir_materiales_check)

        adicional_group.setLayout(adicional_layout)
        layout.addWidget(adicional_group)

        layout.addStretch()
        tab.setLayout(layout)
        self.tab_widget.addTab(tab, "Configuración Específica")

    def update_files_list(self):
        """Actualiza la lista visual de archivos seleccionados"""
        try:
            # Si tienes un QListWidget para mostrar los archivos
            if hasattr(self, 'files_list_widget'):
                self.files_list_widget.clear()
                for file_path in self.selected_files:
                    # Mostrar solo el nombre del archivo, no la ruta completa
                    file_name = file_path.split('/')[-1].split('\\')[-1]
                    self.files_list_widget.addItem(f"{file_name} ({file_path})")

            # Si tienes un QLabel para mostrar el contador
            if hasattr(self, 'files_count_label'):
                count = len(self.selected_files)
                self.files_count_label.setText(f"Archivos seleccionados: {count}")

        except Exception as e:
            print(f"Error actualizando lista de archivos: {str(e)}")

    def create_juridica_specific_tab(self):
        """Crea la pestaña específica para persona jurídica"""
        tab = QWidget()
        scroll = QScrollArea()
        scroll_widget = QWidget()
        layout = QVBoxLayout()

        info_label = QLabel("Configuración específica para Cliente Jurídico")
        info_label.setStyleSheet("font-weight: bold; color: #1E90FF;")
        layout.addWidget(info_label)

        # Tipo de cotización jurídica
        tipo_group = QGroupBox("Tipo de Cotización")
        tipo_layout = QVBoxLayout()

        self.cotizacion_completa_radio = QCheckBox("Cotización completa (con todos los anexos)")
        self.cotizacion_completa_radio.setChecked(True)
        self.cotizacion_completa_radio.toggled.connect(self.toggle_juridica_options)
        tipo_layout.addWidget(self.cotizacion_completa_radio)

        self.cotizacion_basica_radio = QCheckBox("Cotización básica (sin anexos)")
        self.cotizacion_basica_radio.toggled.connect(self.toggle_juridica_options)
        tipo_layout.addWidget(self.cotizacion_basica_radio)

        tipo_group.setLayout(tipo_layout)
        layout.addWidget(tipo_group)

        # Secciones de la cotización completa
        self.secciones_widget = QWidget()
        secciones_layout = QVBoxLayout()

        secciones_group = QGroupBox("Secciones a Incluir")
        secciones_form = QVBoxLayout()

        # Lista de secciones disponibles (Clave, Nombre, Archivo/Tipo)
        self.secciones_checks = {}
        # NOTA: 'carta_presentacion' Eliminada según solicitud
        self.secciones_definitions = [
            ("portadas", "Portadas"),
            ("contenido_separadores", "Contenido (Separadores)"),
            ("propuesta_tecnica", "Propuesta Técnica (Word)"),
            ("presupuesto_programacion", "Presupuesto y Programación (Generado)"),
            ("paginas_estandar", "Páginas estándar"),
            ("cuadro_experiencia", "Cuadro de experiencia"),
            ("certificados_trabajos", "Certificados"),
            ("seguridad_alturas", "Seguridad en alturas"),
            ("programa_prevencion", "Programa prevención caídas"),
            ("sgsst_certificado", "Estándares SGSST"),
            ("documentacion_legal", "Documentación legal"),
            ("anexos", "Anexos")
        ]

        # Checkboxes (lado izquierdo)
        for key, label in self.secciones_definitions:
            check = QCheckBox(label)
            # Por defecto todas marcadas excepto anexos opcionales si se requiere
            check.setChecked(True)  
            check.stateChanged.connect(self.update_order_list_from_checks)
            self.secciones_checks[key] = check
            secciones_form.addWidget(check)

        secciones_group.setLayout(secciones_form)
        
        # --- NUEVO: Lista de Orden (Drag & Drop) ---
        ordering_group = QGroupBox("Orden del PDF Final (Arrastrar para reordenar)")
        ordering_layout = QVBoxLayout()

        self.order_list = QListWidget()
        self.order_list.setDragDropMode(QListWidget.InternalMove)
        self.order_list.setSelectionMode(QListWidget.SingleSelection)
        ordering_layout.addWidget(self.order_list)
        
        ordering_group.setLayout(ordering_layout)

        # Layout horizontal para Checks + Lista
        h_layout = QHBoxLayout()
        h_layout.addWidget(secciones_group)
        h_layout.addWidget(ordering_group)
        secciones_layout.addLayout(h_layout)

        # Archivos a combinar (Botones abajo)
        archivos_group = QGroupBox("Archivos Externos Adicionales")
        archivos_layout = QHBoxLayout()

        self.seleccionar_archivos_btn = QPushButton("Agregar Archivo Externo")
        self.seleccionar_archivos_btn.clicked.connect(self.add_external_file)
        archivos_layout.addWidget(self.seleccionar_archivos_btn)

        self.limpiar_archivos_btn = QPushButton("Remover Item Seleccionado")
        self.limpiar_archivos_btn.clicked.connect(self.remove_selected_item)
        archivos_layout.addWidget(self.limpiar_archivos_btn)
        
        archivos_group.setLayout(archivos_layout)
        secciones_layout.addWidget(archivos_group)

        # Configuración de pólizas
        polizas_group = QGroupBox("Pólizas a Incluir")
        polizas_layout = QVBoxLayout()

        self.polizas_checks = {}
        polizas = [
            ("manejo_anticipo", "Manejo del anticipo"),
            ("cumplimiento_contrato", "Cumplimiento del contrato"),
            ("calidad_servicio", "Calidad del servicio"),
            ("pago_salarios", "Pago de salarios"),
            ("prestaciones_sociales", "Prestaciones sociales e indemnizaciones"),
            ("responsabilidad_civil", "Seguro de responsabilidad civil extracontractual"),
            ("calidad_funcionamiento", "Calidad y correcto funcionamiento")
        ]

        for key, label in polizas:
            check = QCheckBox(label)
            self.polizas_checks[key] = check
            polizas_layout.addWidget(check)

        polizas_group.setLayout(polizas_layout)
        secciones_layout.addWidget(polizas_group)

        # Información del personal técnico
        personal_group = QGroupBox("Personal Técnico")
        personal_form = QFormLayout()

        self.director_obra_edit = QLineEdit()
        self.director_obra_edit.setPlaceholderText("Ej: 25% de disposición en obra")
        personal_form.addRow("Director de obra:", self.director_obra_edit)

        self.residente_obra_edit = QLineEdit()
        self.residente_obra_edit.setPlaceholderText("Nombre del residente de obra")
        personal_form.addRow("Residente de obra:", self.residente_obra_edit)

        self.tecnologo_sgsst_edit = QLineEdit()
        self.tecnologo_sgsst_edit.setPlaceholderText("Nombre del tecnólogo SGSST")
        personal_form.addRow("Tecnólogo SGSST:", self.tecnologo_sgsst_edit)

        personal_group.setLayout(personal_form)
        secciones_layout.addWidget(personal_group)

        # Finalizar Layouts
        self.secciones_widget.setLayout(secciones_layout)
        layout.addWidget(self.secciones_widget)

        layout.addStretch()
        scroll_widget.setLayout(layout)
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)

        tab_layout = QVBoxLayout()
        tab_layout.addWidget(scroll)
        tab.setLayout(tab_layout)

        self.tab_widget.addTab(tab, "Configuración Específica")

        # Inicializar lista
        self.update_order_list_from_checks()

    def update_order_list_from_checks(self):
        """Sincroniza la lista de orden con los checkboxes marcados, manteniendo orden existente."""
        current_items = {}
        for i in range(self.order_list.count()):
            item = self.order_list.item(i)
            key = item.data(Qt.UserRole)
            current_items[key] = item # Guardar referencia

        # Solo procesamos las secciones estándar (no archivos externos)
        # Definimos 'external_file' como prefijo para los externos
        
        # 1. Agregar nuevos marcados
        for key, label in self.secciones_definitions:
            if self.secciones_checks[key].isChecked():
                if key not in current_items:
                    # Agregar al final
                    from PyQt5.QtWidgets import QListWidgetItem
                    item = QListWidgetItem(label)
                    item.setData(Qt.UserRole, key)
                    # Colorear diferente la parte generada
                    if key == "presupuesto_programacion":
                        item.setBackground(QColor("#e6f3ff")) # Azul claro
                    elif key == "contenido_separadores":
                        item.setBackground(QColor("#fff0e6")) # Naranja claro
                    elif key == "propuesta_tecnica":
                        item.setBackground(QColor("#e6ffe6")) # Verde muy claro
                    self.order_list.addItem(item)
            else:
                # 2. Remover desmarcados
                if key in current_items:
                    row = self.order_list.row(current_items[key])
                    self.order_list.takeItem(row)

    def toggle_payment_details(self):
        """Activa/desactiva los detalles de pago por porcentajes"""
        if self.pago_porcentajes_radio.isChecked():
            self.pago_details_widget.setEnabled(True)
            self.pago_contraentrega_radio.setChecked(False)
        elif self.pago_contraentrega_radio.isChecked():
            self.pago_details_widget.setEnabled(False)
            self.pago_porcentajes_radio.setChecked(False)

    def toggle_juridica_options(self):
        """Activa/desactiva las opciones específicas de cotización jurídica"""
        if self.cotizacion_completa_radio.isChecked():
            self.secciones_widget.setEnabled(True)
            self.cotizacion_basica_radio.setChecked(False)
        elif self.cotizacion_basica_radio.isChecked():
            self.secciones_widget.setEnabled(False)
            self.cotizacion_completa_radio.setChecked(False)

    def add_external_file(self):
        """Agrega un archivo externo a la lista de orden."""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Seleccionar archivos", "", "PDF Files (*.pdf);;All Files (*)"
        )
        if files:
            from PyQt5.QtWidgets import QListWidgetItem
            for f in files:
                name = os.path.basename(f)
                item = QListWidgetItem(f"[EXT] {name}")
                item.setData(Qt.UserRole, f"external::{f}") # Guardar ruta completa
                item.setToolTip(f)
                item.setBackground(QColor("#f0fff0")) # Verde claro
                self.order_list.addItem(item)

    def remove_selected_item(self):
        """Remueve el item seleccionado de la lista (y desmarca checkbox si aplica)."""
        row = self.order_list.currentRow()
        if row >= 0:
            item = self.order_list.item(row)
            key = item.data(Qt.UserRole)
            
            # Si es una sección estándar, desmarcar checkbox
            if not key.startswith("external::") and key in self.secciones_checks:
                self.secciones_checks[key].setChecked(False)
            else:
                # Si es externo, solo borrar
                self.order_list.takeItem(row)

    def get_order_list(self):
        """Retorna la lista ordenada de claves/rutas."""
        order = []
        for i in range(self.order_list.count()):
            item = self.order_list.item(i)
            order.append(item.data(Qt.UserRole))
        return order

    def load_default_values(self):
        """Carga valores por defecto"""
        # Valores por defecto para operarios en letra
        self.operarios_letra_edit.setText("dos")

        # Configurar pago contraentrega por defecto
        self.pago_contraentrega_radio.setChecked(True)

        # Para jurídica, configurar cotización completa por defecto
        if self.client_type == 'juridica':
            self.cotizacion_completa_radio.setChecked(True)

    def validate_and_accept(self):
        """Valida los campos y acepta el diálogo"""
        errors = []

        # Validar campos obligatorios
        if not self.referencia_edit.text().strip():
            errors.append("La referencia del trabajo es obligatoria")

        if not self.titulo_edit.text().strip():
            errors.append("El título del trabajo es obligatorio")

        if not self.lugar_edit.text().strip():
            errors.append("El lugar de intervención es obligatorio")

        if not self.concepto_edit.toPlainText().strip():
            errors.append("El concepto general del trabajo es obligatorio")

        # Validar forma de pago
        if not self.pago_contraentrega_radio.isChecked() and not self.pago_porcentajes_radio.isChecked():
            if not self.pago_personalizado_edit.toPlainText().strip():
                errors.append("Debe seleccionar una forma de pago o escribir una personalizada")

        # Validar porcentajes si está seleccionado pago por porcentajes
        if self.pago_porcentajes_radio.isChecked():
            total = self.anticipo_edit.value() + self.avance_edit.value() + self.final_edit.value()
            if total != 100:
                errors.append(f"Los porcentajes de pago deben sumar 100% (actual: {total}%)")

        # Mostrar errores si los hay
        if errors:
            QMessageBox.warning(self, "Campos requeridos", "\n".join(errors))
            return

        # Si todo está bien, aceptar el diálogo
        self.accept()

    def get_config(self):
        """Obtiene la configuración del diálogo"""
        config = {
            # Información básica
            'fecha': self.fecha_edit.date().toString("dd/MM/yyyy"),
            'referencia': self.referencia_edit.text().strip(),
            'titulo': self.titulo_edit.text().strip(),
            'lugar': self.lugar_edit.text().strip(),
            'concepto': self.concepto_edit.toPlainText().strip(),

            # Condiciones comerciales
            'validez': self.validez_edit.value(),
            'cuadrillas': self.cuadrillas_edit.value(),
            'operarios_num': self.operarios_num_edit.value(),
            'operarios_letra': self.operarios_letra_edit.text().strip() or "dos",
            'plazo_dias': self.plazo_dias_edit.value(),
            'plazo_tipo': self.plazo_tipo_combo.currentText(),

            # Forma de pago
            'pago_contraentrega': self.pago_contraentrega_radio.isChecked(),
            'pago_porcentajes': self.pago_porcentajes_radio.isChecked(),
            'anticipo': self.anticipo_edit.value(),
            'avance': self.avance_edit.value(),
            'avance_requerido': self.avance_requerido_edit.value(),
            'final': self.final_edit.value(),
            'pago_personalizado': self.pago_personalizado_edit.toPlainText().strip(),

            # Archivos seleccionados
            'archivos_combinar': self.selected_files.copy(),

            # Configuración específica del tipo de cliente
            'client_type': self.client_type
        }

        # Configuración específica para persona natural
        if self.client_type == 'natural':
            config.update({
                'formato_simple': self.formato_simple_check.isChecked(),
                'incluir_iva': self.incluir_iva_check.isChecked(),
                'incluir_materiales': self.incluir_materiales_check.isChecked()
            })

            # Configuración específica para persona jurídica
        elif self.client_type == 'juridica':
            config.update({
                'cotizacion_completa': self.cotizacion_completa_radio.isChecked(),
                'cotizacion_basica': self.cotizacion_basica_radio.isChecked(),
                'secciones_incluir': {key: check.isChecked() for key, check in self.secciones_checks.items()},
                'section_order': self.get_order_list(), # NUEVO: Orden definido por el usuario
                'polizas_incluir': {key: check.isChecked() for key, check in self.polizas_checks.items()},
                'director_obra': self.director_obra_edit.text().strip(),
                'residente_obra': self.residente_obra_edit.text().strip(),
                'tecnologo_sgsst': self.tecnologo_sgsst_edit.text().strip()
            })

        return config


