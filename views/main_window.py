from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QComboBox, QTableWidget, \
    QTableWidgetItem, QMessageBox, QWidget, QDoubleSpinBox, QHeaderView, QTextEdit, QFileDialog, QSplitter, QScrollArea, QGridLayout
from controllers.cotizacion_controller import CotizacionController
from models.cliente import Cliente
from views.data_management_window import DataManagementWindow
from views.email_dialog import SendEmailDialog
from views.word_dialog import WordConfigDialog
from views.cotizacion_file_dialog import CotizacionFileDialog
from PyQt5.QtCore import Qt, pyqtSlot
from controllers.excel_controller import ExcelController
from controllers.word_controller import WordController
from PyQt5.QtGui import QPalette, QColor
import os
from datetime import datetime

class EditableTableWidgetItem(QTableWidgetItem):
    def __init__(self, value, editable=True):
        if isinstance(value, (int, float)):
            super().__init__(str(value))
        else:
            super().__init__(value)
        
        if not editable:
            self.setFlags(self.flags() & ~Qt.ItemIsEditable)

class MainWindow(QMainWindow):
    def __init__(self, controller=None):
        super().__init__()
        
        # Inicializar controlador
        self.controller = controller if controller else CotizacionController()
        self.aiu_manager = self.controller.aiu_manager
        
        # Configurar ventana
        self.setWindowTitle("Sistema de Cotizaciones")
        self.setGeometry(100, 100, 1200, 800)
        
        # Crear widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principal
        main_layout = QVBoxLayout(central_widget)
        
        # Sección de cliente
        client_group = QWidget()
        client_layout = QVBoxLayout(client_group)
        
        client_title = QLabel("Información del Cliente")
        client_title.setStyleSheet("font-size: 16px; font-weight: bold;")
        client_layout.addWidget(client_title)
        
        # Combo para seleccionar cliente existente
        client_combo_layout = QHBoxLayout()
        client_combo_layout.setContentsMargins(0, 0, 0, 0)  # Eliminar márgenes
        client_combo_layout.setAlignment(Qt.AlignLeft)  # Alinear a la izquierda
        
        client_label = QLabel("Cliente:")
        client_label.setFixedWidth(60)  # Ancho fijo para alineación
        client_combo_layout.addWidget(client_label)
        
        self.client_combo = QComboBox()
        self.client_combo.setMinimumWidth(300)  # Ancho mínimo para mostrar nombres completos
        self.client_combo.currentIndexChanged.connect(self.actualizar_datos_cliente)
        client_combo_layout.addWidget(self.client_combo)
        
        manage_data_btn = QPushButton("Gestionar Datos")
        client_combo_layout.addWidget(manage_data_btn)
        manage_data_btn.clicked.connect(self.open_data_management)
        
        # Añadir espacio flexible al final para empujar todo a la izquierda
        client_combo_layout.addStretch(1)
        
        client_layout.addLayout(client_combo_layout)
        
        # Formulario de cliente usando GridLayout para mejor alineación
        client_form_layout = QGridLayout()
        client_form_layout.setContentsMargins(0, 0, 0, 0)  # Eliminar márgenes
        
        # Tipo (primera fila, primera columna)
        tipo_label = QLabel("Tipo:")
        tipo_label.setFixedWidth(60)  # Ancho fijo para alineación
        client_form_layout.addWidget(tipo_label, 0, 0)
        
        self.tipo_combo = QComboBox()
        self.tipo_combo.addItems(["Natural", "Jurídica"])
        self.tipo_combo.setMinimumWidth(150)
        client_form_layout.addWidget(self.tipo_combo, 0, 1)
        
        # NIT/CC (primera fila, tercera columna)
        nit_label = QLabel("NIT/CC:")
        nit_label.setFixedWidth(60)  # Ancho fijo para alineación
        client_form_layout.addWidget(nit_label, 0, 2)
        
        self.nit_input = QLineEdit()
        self.nit_input.setMinimumWidth(150)
        client_form_layout.addWidget(self.nit_input, 0, 3)
        
        # Nombre (segunda fila, primera columna)
        nombre_label = QLabel("Nombre:")
        nombre_label.setFixedWidth(60)  # Ancho fijo para alineación
        client_form_layout.addWidget(nombre_label, 1, 0)
        
        self.nombre_input = QLineEdit()
        self.nombre_input.setMinimumWidth(150)
        client_form_layout.addWidget(self.nombre_input, 1, 1)
        
        # Teléfono (segunda fila, tercera columna)
        telefono_label = QLabel("Teléfono:")
        telefono_label.setFixedWidth(60)  # Ancho fijo para alineación
        client_form_layout.addWidget(telefono_label, 1, 2)
        
        self.telefono_input = QLineEdit()
        self.telefono_input.setMinimumWidth(150)
        client_form_layout.addWidget(self.telefono_input, 1, 3)
        
        # Dirección (tercera fila, primera columna)
        direccion_label = QLabel("Dirección:")
        direccion_label.setFixedWidth(60)  # Ancho fijo para alineación
        client_form_layout.addWidget(direccion_label, 2, 0)
        
        self.direccion_input = QLineEdit()
        self.direccion_input.setMinimumWidth(150)
        client_form_layout.addWidget(self.direccion_input, 2, 1)
        
        # Email (tercera fila, tercera columna)
        email_label = QLabel("Email:")
        email_label.setFixedWidth(60)  # Ancho fijo para alineación
        client_form_layout.addWidget(email_label, 2, 2)
        
        self.email_input = QLineEdit()
        self.email_input.setMinimumWidth(150)
        client_form_layout.addWidget(self.email_input, 2, 3)
        
        # Añadir espacio flexible al final de cada fila
        client_form_layout.setColumnStretch(4, 1)
        
        client_layout.addLayout(client_form_layout)
        
        # Botón para guardar cliente
        save_client_btn = QPushButton("Guardar Cliente")
        save_client_btn.clicked.connect(self.guardar_cliente)
        save_client_btn.setMaximumWidth(150)  # Limitar ancho del botón
        
        # Contenedor para el botón alineado a la izquierda
        save_btn_layout = QHBoxLayout()
        save_btn_layout.setContentsMargins(0, 0, 0, 0)  # Eliminar márgenes
        save_btn_layout.addWidget(save_client_btn)
        save_btn_layout.addStretch(1)  # Espacio flexible para empujar el botón a la izquierda
        
        client_layout.addLayout(save_btn_layout)
        
        main_layout.addWidget(client_group)
        
        # Contenedor principal para actividades y panel de selección
        activities_container = QSplitter(Qt.Horizontal)
        
        # Sección de actividades (lado izquierdo)
        activities_group = QWidget()
        activities_layout = QVBoxLayout(activities_group)
        
        activities_title = QLabel("Actividades")
        activities_title.setStyleSheet("font-size: 16px; font-weight: bold;")
        activities_layout.addWidget(activities_title)
        
        # Tabla de actividades
        self.activities_table = QTableWidget(0, 6)
        self.activities_table.setHorizontalHeaderLabels(["Descripción", "Cantidad", "Unidad", "Valor Unitario", "Total", ""])
        self.activities_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.activities_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.activities_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.activities_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.activities_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.activities_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        self.activities_table.itemChanged.connect(self.on_item_changed)
        activities_layout.addWidget(self.activities_table)
        
        # Sección de totales
        totals_layout = QHBoxLayout()
        
        totals_layout.addWidget(QLabel("Subtotal:"))
        self.subtotal_label = QLabel("0.00")
        totals_layout.addWidget(self.subtotal_label)
        
        totals_layout.addWidget(QLabel("IVA:"))
        self.iva_label = QLabel("0.00")
        totals_layout.addWidget(self.iva_label)
        
        totals_layout.addWidget(QLabel("Total:"))
        self.total_label = QLabel("0.00")
        totals_layout.addWidget(self.total_label)
        activities_layout.addLayout(totals_layout)
        
        # Agregar el grupo de actividades al contenedor principal
        activities_container.addWidget(activities_group)
        
        # Panel de selección de actividades (lado derecho)
        selection_panel = QWidget()
        selection_layout = QVBoxLayout(selection_panel)
        
        selection_title = QLabel("Selección de Actividades")
        selection_title.setStyleSheet("font-size: 16px; font-weight: bold;")
        selection_layout.addWidget(selection_title)
        
        # Filtros y búsqueda
        filter_layout = QVBoxLayout()
        
        # Búsqueda
        search_layout = QHBoxLayout()
        search_layout.addWidget(QLabel("Buscar:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Buscar actividades...")
        self.search_input.textChanged.connect(self.filter_activities)
        search_layout.addWidget(self.search_input)
        filter_layout.addLayout(search_layout)
        
        # Filtro por categoría
        category_layout = QHBoxLayout()
        category_layout.addWidget(QLabel("Categoría:"))
        self.category_combo = QComboBox()
        self.category_combo.currentIndexChanged.connect(self.filter_activities)
        category_layout.addWidget(self.category_combo)
        filter_layout.addLayout(category_layout)
        
        selection_layout.addLayout(filter_layout)
        
        # Selección de actividad predefinida
        predefined_layout = QVBoxLayout()
        
        # Combo de actividades con descripción completa
        activity_combo_layout = QVBoxLayout()
        activity_combo_layout.addWidget(QLabel("Actividades:"))
        self.activity_combo = QComboBox()
        self.activity_combo.setMinimumWidth(350)  # Ancho mínimo para mostrar más texto
        self.activity_combo.currentIndexChanged.connect(self.update_related_activities)
        activity_combo_layout.addWidget(self.activity_combo)
        
        # Mostrar descripción completa en un área de texto
        activity_combo_layout.addWidget(QLabel("Descripción completa:"))
        self.activity_description = QTextEdit()
        self.activity_description.setReadOnly(True)
        self.activity_description.setMinimumHeight(80)
        activity_combo_layout.addWidget(self.activity_description)
        
        predefined_layout.addLayout(activity_combo_layout)
        
        # Combo de actividades relacionadas
        related_activities_layout = QVBoxLayout()
        related_activities_layout.addWidget(QLabel("Actividades Relacionadas:"))
        self.related_activities_combo = QComboBox()
        self.related_activities_combo.setMinimumWidth(350)  # Ancho mínimo para mostrar más texto
        related_activities_layout.addWidget(self.related_activities_combo)
        
        # Mostrar descripción completa de la actividad relacionada
        related_activities_layout.addWidget(QLabel("Descripción completa:"))
        self.related_activity_description = QTextEdit()
        self.related_activity_description.setReadOnly(True)
        self.related_activity_description.setMinimumHeight(80)
        related_activities_layout.addWidget(self.related_activity_description)
        
        predefined_layout.addLayout(related_activities_layout)
        
        # Cantidad para actividad predefinida
        pred_quantity_layout = QHBoxLayout()
        pred_quantity_layout.addWidget(QLabel("Cantidad:"))
        self.pred_quantity_spinbox = QDoubleSpinBox()
        self.pred_quantity_spinbox.setMinimum(0.01)
        self.pred_quantity_spinbox.setMaximum(9999.99)
        self.pred_quantity_spinbox.setValue(1.0)
        pred_quantity_layout.addWidget(self.pred_quantity_spinbox)
        predefined_layout.addLayout(pred_quantity_layout)
        
        # Botones para agregar actividades predefinidas
        add_buttons_layout = QVBoxLayout()
        add_buttons_layout.addWidget(QLabel("Agregar:"))
        buttons_layout = QHBoxLayout()
        
        add_activity_btn = QPushButton("Agregar Actividad")
        add_activity_btn.clicked.connect(self.add_selected_activity)
        buttons_layout.addWidget(add_activity_btn)
        
        add_related_btn = QPushButton("Agregar Relacionada")
        add_related_btn.clicked.connect(self.add_related_activity)
        buttons_layout.addWidget(add_related_btn)
        
        add_buttons_layout.addLayout(buttons_layout)
        predefined_layout.addLayout(add_buttons_layout)
        
        selection_layout.addLayout(predefined_layout)
        
        # Sección de entrada manual
        manual_entry_layout = QVBoxLayout()
        manual_entry_title = QLabel("Entrada Manual")
        manual_entry_title.setStyleSheet("font-size: 14px; font-weight: bold;")
        manual_entry_layout.addWidget(manual_entry_title)
        
        # Descripción
        description_layout = QVBoxLayout()
        description_layout.addWidget(QLabel("Descripción:"))
        self.description_input = QLineEdit()
        description_layout.addWidget(self.description_input)
        manual_entry_layout.addLayout(description_layout)
        
        # Cantidad, Unidad y Valor en una fila
        quantity_unit_value_layout = QHBoxLayout()
        
        # Cantidad
        quantity_layout = QVBoxLayout()
        quantity_layout.addWidget(QLabel("Cantidad:"))
        self.quantity_spinbox = QDoubleSpinBox()
        self.quantity_spinbox.setMinimum(0.01)
        self.quantity_spinbox.setMaximum(9999.99)
        self.quantity_spinbox.setValue(1.0)
        quantity_layout.addWidget(self.quantity_spinbox)
        quantity_unit_value_layout.addLayout(quantity_layout)
        
        # Unidad
        unit_layout = QVBoxLayout()
        unit_layout.addWidget(QLabel("Unidad:"))
        self.unit_input = QLineEdit()
        unit_layout.addWidget(self.unit_input)
        quantity_unit_value_layout.addLayout(unit_layout)
        
        # Valor unitario
        value_layout = QVBoxLayout()
        value_layout.addWidget(QLabel("Valor Unitario:"))
        self.value_spinbox = QDoubleSpinBox()
        self.value_spinbox.setMinimum(0.01)
        self.value_spinbox.setMaximum(9999999.99)
        self.value_spinbox.setValue(1000.0)
        value_layout.addWidget(self.value_spinbox)
        quantity_unit_value_layout.addLayout(value_layout)
        
        manual_entry_layout.addLayout(quantity_unit_value_layout)
        
        # Botón para agregar actividad manual
        add_manual_btn = QPushButton("Agregar Actividad Manual")
        add_manual_btn.clicked.connect(self.add_manual_activity)
        manual_entry_layout.addWidget(add_manual_btn)
        
        selection_layout.addLayout(manual_entry_layout)
        
        # Agregar espacio flexible al final para que los elementos se alineen en la parte superior
        selection_layout.addStretch()
        
        # Agregar el panel de selección al contenedor principal
        activities_container.addWidget(selection_panel)
        
        # Establecer proporciones iniciales del splitter (70% izquierda, 30% derecha)
        activities_container.setSizes([700, 300])
        
        main_layout.addWidget(activities_container)
        
        # Botones de acción
        actions_layout = QHBoxLayout()
        
        # Botones para gestión de archivos de cotización
        save_file_btn = QPushButton("Guardar como Archivo")
        save_file_btn.clicked.connect(self.save_as_file)
        actions_layout.addWidget(save_file_btn)
        
        open_file_btn = QPushButton("Abrir Archivo")
        open_file_btn.clicked.connect(self.open_file_dialog)
        actions_layout.addWidget(open_file_btn)
        
        generate_excel_btn = QPushButton("Generar Excel")
        generate_excel_btn.clicked.connect(self.generate_excel)
        actions_layout.addWidget(generate_excel_btn)
        
        generate_word_btn = QPushButton("Generar Word")
        generate_word_btn.clicked.connect(self.generate_word)
        actions_layout.addWidget(generate_word_btn)
        
        send_email_btn = QPushButton("Enviar por Correo")
        send_email_btn.clicked.connect(self.send_email)
        actions_layout.addWidget(send_email_btn)
        
        clear_btn = QPushButton("Limpiar")
        clear_btn.clicked.connect(self.clear_form)
        actions_layout.addWidget(clear_btn)
        
        # Botón para cambiar tema
        theme_btn = QPushButton("Modo Oscuro")
        theme_btn.clicked.connect(self.toggle_theme)
        actions_layout.addWidget(theme_btn)
        
        main_layout.addLayout(actions_layout)
        
        # Inicializar tema claro por defecto
        self.dark_mode = False
        self.theme_btn = theme_btn
        
        # Cargar datos iniciales
        self.load_initial_data()
        
    def load_initial_data(self):
        """Carga los datos iniciales en la interfaz"""
        try:
            # Cargar clientes
            self.refresh_client_combo()
            
            # Cargar categorías
            self.refresh_category_combo()
            
            # Cargar actividades
            self.refresh_activity_combo()
        except Exception as e:
            print(f"Error al cargar datos iniciales: {str(e)}")
        
    def toggle_theme(self):
        """Cambia entre tema claro y oscuro"""
        if not self.dark_mode:
            # Aplicar tema oscuro
            self.dark_mode = True
            self.theme_btn.setText("Modo Claro")
            self.apply_dark_theme()
        else:
            # Aplicar tema claro
            self.dark_mode = False
            self.theme_btn.setText("Modo Oscuro")
            self.apply_light_theme()
    
    def apply_dark_theme(self):
        """Aplica el tema oscuro a la aplicación"""
        palette = QPalette()
        
        # Colores para el tema oscuro
        palette.setColor(QPalette.Window, QColor(53, 53, 53))
        palette.setColor(QPalette.WindowText, Qt.white)
        palette.setColor(QPalette.Base, QColor(25, 25, 25))
        palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        palette.setColor(QPalette.ToolTipBase, Qt.white)
        palette.setColor(QPalette.ToolTipText, Qt.white)
        palette.setColor(QPalette.Text, Qt.white)
        palette.setColor(QPalette.Button, QColor(53, 53, 53))
        palette.setColor(QPalette.ButtonText, Qt.white)
        palette.setColor(QPalette.BrightText, Qt.red)
        palette.setColor(QPalette.Link, QColor(42, 130, 218))
        palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
        palette.setColor(QPalette.HighlightedText, Qt.black)
        
        # Aplicar la paleta
        self.setPalette(palette)
        
        # Aplicar estilos adicionales para mejorar la visibilidad
        self.setStyleSheet("""
            QPushButton { 
                color: white; 
                background-color: #444444;
                border: 1px solid #666666;
                padding: 5px;
                border-radius: 3px;
            }
            QPushButton:hover { 
                background-color: #555555; 
            }
            QPushButton:pressed { 
                background-color: #333333; 
            }
            QComboBox {
                color: white;
                background-color: #444444;
                selection-background-color: #666666;
                border: 1px solid #666666;
                padding: 5px;
                border-radius: 3px;
            }
            QComboBox QAbstractItemView {
                color: white;
                background-color: #444444;
                selection-background-color: #666666;
            }
            QLineEdit, QTextEdit, QDoubleSpinBox {
                color: white;
                background-color: #333333;
                border: 1px solid #666666;
                padding: 5px;
                border-radius: 3px;
            }
            QTableWidget {
                color: white;
                background-color: #333333;
                gridline-color: #666666;
                border: 1px solid #666666;
            }
            QHeaderView::section {
                color: white;
                background-color: #444444;
                border: 1px solid #666666;
            }
            QTableWidget::item:selected {
                color: black;
                background-color: #42a2d6;
            }
            QLabel {
                color: white;
            }
        """)
    
    def apply_light_theme(self):
        """Aplica el tema claro a la aplicación"""
        # Restaurar la paleta por defecto
        self.setPalette(self.style().standardPalette())
        
        # Limpiar estilos personalizados
        self.setStyleSheet("")
    
    def refresh_client_combo(self):
        """Actualiza el combo de clientes"""
        try:
            # Guardar el cliente seleccionado actualmente
            current_client_id = self.client_combo.currentData()
            
            # Limpiar el combo
            self.client_combo.clear()
            
            # Agregar los clientes
            clients = self.controller.get_all_clients()
            for client in clients:
                self.client_combo.addItem(client['nombre'], client['id'])
            
            # Restaurar la selección anterior si existe
            if current_client_id:
                index = self.client_combo.findData(current_client_id)
                if index >= 0:
                    self.client_combo.setCurrentIndex(index)
        except Exception as e:
            print(f"Error al actualizar combo de clientes: {str(e)}")
    
    def actualizar_datos_cliente(self):
        """Actualiza los campos del formulario con los datos del cliente seleccionado"""
        try:
            client_id = self.client_combo.currentData()
            if client_id:
                client = self.controller.database_manager.get_client_by_id(client_id)
                if client:
                    # Actualizar campos
                    self.tipo_combo.setCurrentText(client['tipo'].capitalize())
                    self.nombre_input.setText(client['nombre'])
                    self.direccion_input.setText(client['direccion'] or "")
                    self.nit_input.setText(client['nit'] or "")
                    self.telefono_input.setText(client['telefono'] or "")
                    self.email_input.setText(client['email'] or "")
        except Exception as e:
            print(f"Error al actualizar datos del cliente: {str(e)}")
    
    def guardar_cliente(self):
        """Guarda los datos del cliente en la base de datos"""
        try:
            # Obtener datos del formulario
            client_data = {
                'tipo': self.tipo_combo.currentText().lower(),
                'nombre': self.nombre_input.text(),
                'direccion': self.direccion_input.text(),
                'nit': self.nit_input.text(),
                'telefono': self.telefono_input.text(),
                'email': self.email_input.text()
            }
            
            # Validar datos
            if not client_data['nombre']:
                QMessageBox.warning(self, "Error", "El nombre del cliente es obligatorio.")
                return
            
            # Verificar si es una actualización o un nuevo cliente
            client_id = self.client_combo.currentData()
            
            if client_id:
                # Actualizar cliente existente
                self.controller.database_manager.update_client(client_id, client_data)
                QMessageBox.information(self, "Éxito", "Cliente actualizado correctamente.")
            else:
                # Agregar nuevo cliente
                client_id = self.controller.add_client(client_data)
                QMessageBox.information(self, "Éxito", "Cliente agregado correctamente.")
            
            # Actualizar combo de clientes
            self.refresh_client_combo()
            
            # Seleccionar el cliente recién guardado
            index = self.client_combo.findData(client_id)
            if index >= 0:
                self.client_combo.setCurrentIndex(index)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al guardar cliente: {str(e)}")
    
    def refresh_category_combo(self):
        """Actualiza el combo de categorías"""
        try:
            # Guardar la categoría seleccionada actualmente
            current_category_index = self.category_combo.currentIndex()
            
            # Limpiar el combo
            self.category_combo.clear()
            
            # Agregar opción "Todas"
            self.category_combo.addItem("Todas", None)
            
            # Agregar las categorías
            categories = self.controller.get_all_categories()
            for category in categories:
                self.category_combo.addItem(category['nombre'], category['id'])
            
            # Restaurar la selección anterior si existe
            if current_category_index >= 0 and current_category_index < self.category_combo.count():
                self.category_combo.setCurrentIndex(current_category_index)
        except Exception as e:
            print(f"Error al actualizar combo de categorías: {str(e)}")
    
    def refresh_activity_combo(self):
        """Actualiza el combo de actividades"""
        try:
            # Guardar la actividad seleccionada actualmente
            current_activity_id = self.activity_combo.currentData()
            
            # Limpiar el combo
            self.activity_combo.clear()
            
            # Agregar las actividades
            activities = self.controller.get_all_activities()
            for activity in activities:
                # Mostrar descripción completa en el combo
                self.activity_combo.addItem(activity['descripcion'], activity['id'])
            
            # Restaurar la selección anterior si existe
            if current_activity_id:
                index = self.activity_combo.findData(current_activity_id)
                if index >= 0:
                    self.activity_combo.setCurrentIndex(index)
                    
            # Actualizar la descripción completa
            self.update_activity_description()
        except Exception as e:
            print(f"Error al actualizar combo de actividades: {str(e)}")
    
    def update_activity_description(self):
        """Actualiza el área de texto con la descripción completa de la actividad seleccionada"""
        try:
            activity_id = self.activity_combo.currentData()
            if activity_id:
                activity = self.controller.database_manager.get_activity_by_id(activity_id)
                if activity:
                    # Mostrar descripción completa
                    description = f"Descripción: {activity['descripcion']}\n"
                    description += f"Unidad: {activity['unidad']}\n"
                    description += f"Valor unitario: {activity['valor_unitario']}\n"
                    if 'categoria_nombre' in activity and activity['categoria_nombre']:
                        description += f"Categoría: {activity['categoria_nombre']}"
                    
                    self.activity_description.setText(description)
                else:
                    self.activity_description.clear()
            else:
                self.activity_description.clear()
        except Exception as e:
            print(f"Error al actualizar descripción de actividad: {str(e)}")
    
    def update_related_activity_description(self):
        """Actualiza el área de texto con la descripción completa de la actividad relacionada seleccionada"""
        try:
            activity_id = self.related_activities_combo.currentData()
            if activity_id:
                activity = self.controller.database_manager.get_activity_by_id(activity_id)
                if activity:
                    # Mostrar descripción completa
                    description = f"Descripción: {activity['descripcion']}\n"
                    description += f"Unidad: {activity['unidad']}\n"
                    description += f"Valor unitario: {activity['valor_unitario']}\n"
                    if 'categoria_nombre' in activity and activity['categoria_nombre']:
                        description += f"Categoría: {activity['categoria_nombre']}"
                    
                    self.related_activity_description.setText(description)
                else:
                    self.related_activity_description.clear()
            else:
                self.related_activity_description.clear()
        except Exception as e:
            print(f"Error al actualizar descripción de actividad relacionada: {str(e)}")
    
    def update_related_activities(self):
        """Actualiza el combo de actividades relacionadas"""
        try:
            # Guardar la actividad relacionada seleccionada actualmente
            current_related_activity_id = self.related_activities_combo.currentData()
            
            # Limpiar el combo
            self.related_activities_combo.clear()
            
            # Obtener la actividad seleccionada
            activity_id = self.activity_combo.currentData()
            
            # Actualizar la descripción de la actividad principal
            self.update_activity_description()
            
            if activity_id:
                # Obtener actividades relacionadas
                related_activities = self.controller.get_related_activities(activity_id)
                
                # Agregar las actividades relacionadas
                for activity in related_activities:
                    self.related_activities_combo.addItem(activity['descripcion'], activity['id'])
                
                # Restaurar la selección anterior si existe
                if current_related_activity_id:
                    index = self.related_activities_combo.findData(current_related_activity_id)
                    if index >= 0:
                        self.related_activities_combo.setCurrentIndex(index)
                        
                # Actualizar la descripción de la actividad relacionada
                self.update_related_activity_description()
        except Exception as e:
            print(f"Error al actualizar actividades relacionadas: {str(e)}")
    
    def filter_activities(self):
        """Filtra las actividades según el texto de búsqueda y la categoría seleccionada"""
        try:
            # Obtener texto de búsqueda
            search_text = self.search_input.text()
            
            # Obtener categoría seleccionada
            category_id = self.category_combo.currentData()
            
            # Buscar actividades
            activities = self.controller.search_activities(search_text, category_id)
            
            # Guardar la actividad seleccionada actualmente
            current_activity_id = self.activity_combo.currentData()
            
            # Limpiar el combo
            self.activity_combo.clear()
            
            # Agregar las actividades filtradas
            for activity in activities:
                self.activity_combo.addItem(activity['descripcion'], activity['id'])
            
            # Restaurar la selección anterior si existe
            if current_activity_id:
                index = self.activity_combo.findData(current_activity_id)
                if index >= 0:
                    self.activity_combo.setCurrentIndex(index)
                    
            # Actualizar la descripción de la actividad
            self.update_activity_description()
            
            # Actualizar actividades relacionadas
            self.update_related_activities()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al filtrar actividades: {str(e)}")
    
    def add_selected_activity(self):
        """Agrega la actividad seleccionada a la tabla"""
        try:
            # Obtener la actividad seleccionada
            activity_id = self.activity_combo.currentData()
            
            if not activity_id:
                QMessageBox.warning(self, "Error", "Debe seleccionar una actividad.")
                return
            
            # Obtener datos de la actividad
            activity = self.controller.database_manager.get_activity_by_id(activity_id)
            
            if not activity:
                QMessageBox.warning(self, "Error", "La actividad seleccionada no existe.")
                return
            
            # Obtener cantidad
            cantidad = self.pred_quantity_spinbox.value()
            
            # Agregar a la tabla
            self.add_activity_to_table(activity['descripcion'], cantidad, activity['unidad'], activity['valor_unitario'])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al agregar actividad: {str(e)}")
    
    def add_related_activity(self):
        """Agrega la actividad relacionada seleccionada a la tabla"""
        try:
            # Obtener la actividad relacionada seleccionada
            activity_id = self.related_activities_combo.currentData()
            
            if not activity_id:
                QMessageBox.warning(self, "Error", "Debe seleccionar una actividad relacionada.")
                return
            
            # Obtener datos de la actividad
            activity = self.controller.database_manager.get_activity_by_id(activity_id)
            
            if not activity:
                QMessageBox.warning(self, "Error", "La actividad seleccionada no existe.")
                return
            
            # Obtener cantidad
            cantidad = self.pred_quantity_spinbox.value()
            
            # Agregar a la tabla
            self.add_activity_to_table(activity['descripcion'], cantidad, activity['unidad'], activity['valor_unitario'])
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al agregar actividad relacionada: {str(e)}")
    
    def add_manual_activity(self):
        """Agrega una actividad manual a la tabla"""
        try:
            # Obtener datos del formulario
            descripcion = self.description_input.text()
            cantidad = self.quantity_spinbox.value()
            unidad = self.unit_input.text()
            valor_unitario = self.value_spinbox.value()
            
            # Validar datos
            if not descripcion:
                QMessageBox.warning(self, "Error", "La descripción es obligatoria.")
                return
            
            if not unidad:
                QMessageBox.warning(self, "Error", "La unidad es obligatoria.")
                return
            
            # Agregar a la tabla
            self.add_activity_to_table(descripcion, cantidad, unidad, valor_unitario)
            
            # Limpiar campos
            self.description_input.clear()
            self.quantity_spinbox.setValue(1.0)
            self.unit_input.clear()
            self.value_spinbox.setValue(1000.0)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al agregar actividad manual: {str(e)}")
    
    def add_activity_to_table(self, descripcion, cantidad, unidad, valor_unitario):
        """Agrega una actividad a la tabla"""
        try:
            # Calcular total
            total = cantidad * valor_unitario
            
            # Agregar fila
            row = self.activities_table.rowCount()
            self.activities_table.insertRow(row)
            
            # Descripción
            self.activities_table.setItem(row, 0, EditableTableWidgetItem(descripcion))
            
            # Cantidad
            self.activities_table.setItem(row, 1, EditableTableWidgetItem(cantidad))
            
            # Unidad
            self.activities_table.setItem(row, 2, EditableTableWidgetItem(unidad))
            
            # Valor unitario
            self.activities_table.setItem(row, 3, EditableTableWidgetItem(valor_unitario))
            
            # Total
            self.activities_table.setItem(row, 4, EditableTableWidgetItem(total, False))
            
            # Botón eliminar
            delete_btn = QPushButton("Eliminar")
            delete_btn.clicked.connect(lambda _, r=row: self.delete_activity(r))
            self.activities_table.setCellWidget(row, 5, delete_btn)
            
            # Actualizar totales
            self.update_totals()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al agregar actividad a la tabla: {str(e)}")
    
    def delete_activity(self, row):
        """Elimina una actividad de la tabla"""
        try:
            self.activities_table.removeRow(row)
            self.update_totals()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al eliminar actividad: {str(e)}")
    
    def on_item_changed(self, item):
        """Se ejecuta cuando cambia un item de la tabla"""
        try:
            # Solo procesar cambios en columnas editables (0-3)
            if item.column() > 3:
                return
            
            row = item.row()
            
            # Obtener valores
            try:
                cantidad = float(self.activities_table.item(row, 1).text())
                valor_unitario = float(self.activities_table.item(row, 3).text())
                total = cantidad * valor_unitario
                
                # Actualizar total
                self.activities_table.setItem(row, 4, EditableTableWidgetItem(total, False))
                
                # Actualizar totales generales
                self.update_totals()
            except (ValueError, AttributeError):
                # Ignorar errores de conversión durante la edición
                pass
        except Exception as e:
            print(f"Error en on_item_changed: {str(e)}")
    
    def update_totals(self):
        """Actualiza los totales de la cotización"""
        try:
            # Calcular subtotal
            subtotal = 0
            for row in range(self.activities_table.rowCount()):
                try:
                    total = float(self.activities_table.item(row, 4).text())
                    subtotal += total
                except (ValueError, AttributeError):
                    pass
            
            # Obtener valores AIU
            aiu_values = self.aiu_manager.get_aiu_values()
            
            # Calcular IVA (sobre utilidad para jurídicas, sobre subtotal para naturales)
            iva = 0
            if self.tipo_combo.currentText().lower() == "jurídica":
                utilidad = subtotal * (aiu_values['utilidad'] / 100)
                iva = utilidad * (aiu_values['iva_sobre_utilidad'] / 100)
            else:
                iva = subtotal * (aiu_values['iva_sobre_utilidad'] / 100)
            
            # Calcular total
            total = subtotal + iva
            
            # Actualizar etiquetas
            self.subtotal_label.setText(f"{subtotal:.2f}")
            self.iva_label.setText(f"{iva:.2f}")
            self.total_label.setText(f"{total:.2f}")
        except Exception as e:
            print(f"Error al actualizar totales: {str(e)}")
    
    def generate_excel(self, show_message=True):
        """Genera un archivo Excel con la cotización actual"""
        try:
            # Verificar que haya actividades
            if self.activities_table.rowCount() == 0:
                QMessageBox.warning(self, "Error", "Debe agregar al menos una actividad a la cotización.")
                return None
            
            # Obtener actividades de la tabla
            activities = []
            for row in range(self.activities_table.rowCount()):
                activity = {
                    'descripcion': self.activities_table.item(row, 0).text(),
                    'cantidad': float(self.activities_table.item(row, 1).text()),
                    'unidad': self.activities_table.item(row, 2).text(),
                    'valor_unitario': float(self.activities_table.item(row, 3).text()),
                    'total': float(self.activities_table.item(row, 4).text())
                }
                activities.append(activity)
            
            # Obtener valores AIU
            aiu_values = self.aiu_manager.get_aiu_values()
            
            # Crear controlador de Excel
            excel_controller = ExcelController(self.controller)
            
            # Generar Excel
            excel_path = excel_controller.generate_excel(
                activities=activities,
                tipo_persona=self.tipo_combo.currentText().lower(),
                administracion=aiu_values['administracion'],
                imprevistos=aiu_values['imprevistos'],
                utilidad=aiu_values['utilidad'],
                iva_utilidad=aiu_values['iva_sobre_utilidad']
            )
            
            if show_message:
                QMessageBox.information(self, "Éxito", f"Archivo Excel generado correctamente en:\n{excel_path}")
            
            return excel_path
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al generar el Excel: {str(e)}")
            return None
    
    def generate_word(self):
        """Genera un documento Word con la cotización actual"""
        try:
            # Verificar que haya actividades
            if self.activities_table.rowCount() == 0:
                QMessageBox.warning(self, "Error", "Debe agregar al menos una actividad a la cotización.")
                return
            
            # Verificar que haya un cliente seleccionado
            if self.client_combo.currentIndex() == -1:
                QMessageBox.warning(self, "Error", "Debe seleccionar un cliente.")
                return
            
            # Generar Excel primero (necesario para la imagen)
            excel_path = self.generate_excel(show_message=False)
            if not excel_path:
                return
            
            # Obtener datos del cliente
            client_id = self.client_combo.currentData()
            client = self.controller.database_manager.get_client_by_id(client_id)
            
            # Abrir diálogo para configurar documento Word
            dialog = WordConfigDialog(self, client['tipo'])
            if dialog.exec_():
                config = dialog.get_config()
                
                # Crear controlador de Word
                word_controller = WordController()
                
                # Generar Word
                word_path = word_controller.generate_word(
                    excel_path=excel_path,
                    client_data=client,
                    config=config
                )
                
                QMessageBox.information(self, "Éxito", f"Documento Word generado correctamente en:\n{word_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al generar el Word: {str(e)}")
    
    def clear_form(self):
        """Limpia el formulario"""
        try:
            # Limpiar tabla de actividades
            self.activities_table.setRowCount(0)
            
            # Limpiar totales
            self.subtotal_label.setText("0.00")
            self.iva_label.setText("0.00")
            self.total_label.setText("0.00")
            
            # Limpiar campos de búsqueda
            self.search_input.clear()
            self.category_combo.setCurrentIndex(0)
            
            # Limpiar campos de actividad manual
            self.description_input.clear()
            self.quantity_spinbox.setValue(1.0)
            self.unit_input.clear()
            self.value_spinbox.setValue(1000.0)
            
            # Actualizar combos
            self.refresh_activity_combo()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al limpiar formulario: {str(e)}")
    
    def save_as_file(self):
        """Guarda la cotización actual como un archivo"""
        try:
            # Verificar que haya actividades
            if self.activities_table.rowCount() == 0:
                QMessageBox.warning(self, "Error", "Debe agregar al menos una actividad a la cotización.")
                return
            
            # Verificar que haya un cliente seleccionado
            if self.client_combo.currentIndex() == -1:
                QMessageBox.warning(self, "Error", "Debe seleccionar un cliente.")
                return
            
            # Obtener datos del cliente
            client_id = self.client_combo.currentData()
            client_data = self.controller.database_manager.get_client_by_id(client_id)
            
            # Obtener actividades de la tabla
            activities = []
            for row in range(self.activities_table.rowCount()):
                activity = {
                    'descripcion': self.activities_table.item(row, 0).text(),
                    'cantidad': float(self.activities_table.item(row, 1).text()),
                    'unidad': self.activities_table.item(row, 2).text(),
                    'valor_unitario': float(self.activities_table.item(row, 3).text()),
                    'total': float(self.activities_table.item(row, 4).text()),
                    'id': row + 1  # Asignar un ID único basado en el índice
                }
                activities.append(activity)
            
            # Obtener valores AIU
            aiu_values = self.aiu_manager.get_aiu_values()
            
            # Calcular totales
            subtotal = sum(activity['total'] for activity in activities)
            administracion = subtotal * (aiu_values['administracion'] / 100)
            imprevistos = subtotal * (aiu_values['imprevistos'] / 100)
            utilidad = subtotal * (aiu_values['utilidad'] / 100)
            iva_utilidad = utilidad * (aiu_values['iva_sobre_utilidad'] / 100)
            total = subtotal + administracion + imprevistos + utilidad + iva_utilidad
            
            # Crear datos de la cotización
            cotizacion_data = {
                'fecha': datetime.now().strftime('%Y-%m-%d'),
                'cliente': client_data,
                'actividades': activities,
                'subtotal': subtotal,
                'administracion': administracion,
                'imprevistos': imprevistos,
                'utilidad': utilidad,
                'iva_utilidad': iva_utilidad,
                'total': total,
                'aiu_values': aiu_values
            }
            
            # Abrir diálogo para guardar archivo
            file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Cotización", "", "Archivos de Cotización (*.cotiz);;Todos los archivos (*)")
            
            if file_path:
                # Asegurar que tenga la extensión correcta
                if not file_path.endswith('.cotiz'):
                    file_path += '.cotiz'
                
                # Guardar archivo
                self.controller.file_manager.guardar_cotizacion(cotizacion_data, file_path)
                QMessageBox.information(self, "Éxito", f"Cotización guardada correctamente en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al guardar archivo: {str(e)}")
    
    def open_file_dialog(self):
        """Abre el diálogo para seleccionar un archivo de cotización"""
        try:
            dialog = CotizacionFileDialog(self.controller, self)
            if dialog.exec_():
                cotizacion_data = dialog.get_selected_cotizacion()
                if cotizacion_data:
                    self.load_cotizacion_from_file(cotizacion_data)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al abrir archivo: {str(e)}")
    
    def load_cotizacion_from_file(self, cotizacion_data):
        """Carga una cotización desde un archivo"""
        try:
            # Limpiar formulario
            self.clear_form()
            
            # Cargar datos del cliente
            client_data = cotizacion_data['cliente']
            
            # Buscar cliente en la base de datos o crearlo si no existe
            clients = self.controller.get_all_clients()
            client_exists = False
            client_id = None
            
            for client in clients:
                if client['nombre'] == client_data['nombre'] and client['tipo'] == client_data['tipo']:
                    client_exists = True
                    client_id = client['id']
                    break
            
            if not client_exists:
                # Crear nuevo cliente
                client_id = self.controller.add_client(client_data)
                self.refresh_client_combo()
            
            # Seleccionar cliente en el combo
            index = self.client_combo.findData(client_id)
            if index >= 0:
                self.client_combo.setCurrentIndex(index)
            
            # Cargar actividades
            for activity in cotizacion_data['actividades']:
                row = self.activities_table.rowCount()
                self.activities_table.insertRow(row)
                
                # Descripción
                self.activities_table.setItem(row, 0, EditableTableWidgetItem(activity['descripcion']))
                
                # Cantidad
                self.activities_table.setItem(row, 1, EditableTableWidgetItem(activity['cantidad']))
                
                # Unidad
                self.activities_table.setItem(row, 2, EditableTableWidgetItem(activity['unidad']))
                
                # Valor unitario
                self.activities_table.setItem(row, 3, EditableTableWidgetItem(activity['valor_unitario']))
                
                # Total
                total = activity['cantidad'] * activity['valor_unitario']
                self.activities_table.setItem(row, 4, EditableTableWidgetItem(total, False))
                
                # Botón eliminar
                delete_btn = QPushButton("Eliminar")
                delete_btn.clicked.connect(lambda _, r=row: self.delete_activity(r))
                self.activities_table.setCellWidget(row, 5, delete_btn)
            
            # Actualizar totales
            self.update_totals()
            
            QMessageBox.information(self, "Éxito", "Cotización cargada correctamente.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al cargar cotización: {str(e)}")
    
    def send_email(self):
        """Envía la cotización por correo electrónico"""
        try:
            # Verificar que haya actividades
            if self.activities_table.rowCount() == 0:
                QMessageBox.warning(self, "Error", "Debe agregar al menos una actividad a la cotización.")
                return
            
            # Verificar que haya un cliente seleccionado
            if self.client_combo.currentIndex() == -1:
                QMessageBox.warning(self, "Error", "Debe seleccionar un cliente.")
                return
            
            # Generar Excel
            excel_path = self.generate_excel(show_message=False)
            if not excel_path:
                return
            
            # Abrir diálogo para enviar correo
            dialog = SendEmailDialog(self, excel_path)
            dialog.exec_()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al enviar correo: {str(e)}")
    
    def open_data_management(self):
        """Abre la ventana de gestión de datos"""
        try:
            self.data_management_window = DataManagementWindow(self.controller, self)
            self.data_management_window.closed.connect(self.on_data_management_closed)
            self.data_management_window.show()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al abrir la ventana de gestión de datos: {str(e)}")
    
    @pyqtSlot()
    def on_data_management_closed(self):
        """Se ejecuta cuando se cierra la ventana de gestión de datos"""
        # Actualizar combos
        self.refresh_client_combo()
        self.refresh_category_combo()
        self.refresh_activity_combo()
