from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QComboBox, QTableWidget, \
    QTableWidgetItem, QMessageBox, QWidget, QDoubleSpinBox, QHeaderView, QTextEdit, QFileDialog
from controllers.cotizacion_controller import CotizacionController
from models.cliente import Cliente
from views.data_management_window import DataManagementWindow
from views.email_dialog import SendEmailDialog
from views.word_dialog import WordConfigDialog
from PyQt5.QtCore import Qt, pyqtSlot
from controllers.excel_controller import ExcelController
from controllers.word_controller import WordController
from PyQt5.QtGui import QPalette, QColor
import os

class EditableTableWidgetItem(QTableWidgetItem):
    def __init__(self, value, editable=True):
        super().__init__(str(value))
        self.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        if editable:
            self.setFlags(self.flags() | Qt.ItemIsEditable)
            
class MainWindow(QMainWindow):
    def __init__(self, controller=None):
        super().__init__()
        self.setWindowTitle("Sistema de Cotizaciones")
        self.setGeometry(100, 100, 1000, 800)

        # Inicializa el controlador
        self.controller = controller if controller else CotizacionController()
        self.xls = ExcelController(self.controller)

        # Widget central y layout principal
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Botón para abrir la ventana de gestión de datos
        self.open_data_management_btn = QPushButton("Gestionar Datos")
        self.open_data_management_btn.clicked.connect(self.open_data_management)
        main_layout.addWidget(self.open_data_management_btn)

        # Sección de cliente
        client_layout = QHBoxLayout()
        client_layout.addWidget(QLabel("Cliente:"))
        self.client_combo = QComboBox()
        self.load_clientes()
        self.client_combo.currentIndexChanged.connect(self.actualizar_datos_cliente)
        client_layout.addWidget(self.client_combo)

        # Campos de entrada para nuevo cliente
        self.nombre_input = QLineEdit()
        self.nombre_input.setPlaceholderText("Nombre del Cliente")
        client_layout.addWidget(self.nombre_input)

        self.tipo_combo = QComboBox()
        self.tipo_combo.addItems(['natural', 'jurídica'])
        client_layout.addWidget(self.tipo_combo)

        self.direccion_input = QLineEdit()
        self.direccion_input.setPlaceholderText("Dirección del Cliente")
        client_layout.addWidget(self.direccion_input)

        self.nit_input = QLineEdit()
        self.nit_input.setPlaceholderText("NIT (opcional)")
        client_layout.addWidget(self.nit_input)

        self.telefono_input = QLineEdit()
        self.telefono_input.setPlaceholderText("Teléfono (opcional)")
        client_layout.addWidget(self.telefono_input)

        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("email")
        client_layout.addWidget(self.email_input)

        self.guardar_cliente_btn = QPushButton("Guardar Cliente")
        self.guardar_cliente_btn.clicked.connect(self.guardar_cliente)
        client_layout.addWidget(self.guardar_cliente_btn)

        main_layout.addLayout(client_layout)
        
        # Seccion AIU
        # Valores AIU configurables
        aiu_values = self.controller.aiu_manager.obtener_aiu()
        self.admin_input = QDoubleSpinBox()
        self.admin_input.setValue(aiu_values['administracion'])
        self.imprevistos_input = QDoubleSpinBox()
        self.imprevistos_input.setValue(aiu_values['imprevistos'])
        self.utilidad_input = QDoubleSpinBox()
        self.utilidad_input.setValue(aiu_values['utilidad'])
        self.iva_input = QDoubleSpinBox()
        self.iva_input.setValue(aiu_values['iva_sobre_utilidad'])

        # Agregar campos al layout
        aiu_layout = QHBoxLayout()
        aiu_layout.addWidget(QLabel("Administración:"))
        aiu_layout.addWidget(self.admin_input)
        aiu_layout.addWidget(QLabel("Imprevistos:"))
        aiu_layout.addWidget(self.imprevistos_input)
        aiu_layout.addWidget(QLabel("Utilidad:"))
        aiu_layout.addWidget(self.utilidad_input)
        aiu_layout.addWidget(QLabel("IVA:"))
        aiu_layout.addWidget(self.iva_input)

        main_layout.addLayout(aiu_layout)

        # Sección de actividades (modificada)
        activities_layout = QVBoxLayout()
        activities_layout.addWidget(QLabel("Actividades:"))

        # Layout para selección de actividad y cantidad
        activity_selection_layout = QHBoxLayout()

        # Combo para categorías
        self.category_combo = QComboBox()
        self.category_combo.addItem("Todas las categorías", None)
        self.load_categories()
        self.category_combo.currentIndexChanged.connect(self.filter_activities)
        activity_selection_layout.addWidget(QLabel("Categoría:"), 1)
        activity_selection_layout.addWidget(self.category_combo, 2)

        # Combo para actividades
        self.activity_combo = QComboBox()
        self.load_activities()
        activity_selection_layout.addWidget(self.activity_combo, 3)

        # Campo de búsqueda
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Buscar actividad...")
        self.search_input.textChanged.connect(self.filter_activities)
        activity_selection_layout.addWidget(self.search_input, 2)

        self.quantity_spinbox = QDoubleSpinBox()
        self.quantity_spinbox.setRange(0.01, 1000000)
        self.quantity_spinbox.setValue(1)
        self.quantity_spinbox.setDecimals(2)
        activity_selection_layout.addWidget(QLabel("Cantidad:"), 1)
        activity_selection_layout.addWidget(self.quantity_spinbox, 1)

        add_activity_btn = QPushButton("Agregar Actividad")
        add_activity_btn.clicked.connect(self.add_selected_activity)
        activity_selection_layout.addWidget(add_activity_btn, 1)

        activities_layout.addLayout(activity_selection_layout)

        # Sección para actividades relacionadas
        related_activities_layout = QHBoxLayout()
        related_activities_layout.addWidget(QLabel("Actividades Relacionadas:"))
        self.related_activities_combo = QComboBox()
        related_activities_layout.addWidget(self.related_activities_combo, 3)
        
        add_related_btn = QPushButton("Agregar Relacionada")
        add_related_btn.clicked.connect(self.add_related_activity)
        related_activities_layout.addWidget(add_related_btn, 1)
        
        activities_layout.addLayout(related_activities_layout)

        # Tabla de actividades
        self.activities_table = QTableWidget(0, 6)
        self.activities_table.setHorizontalHeaderLabels(
            ["Descripción", "Cantidad", "Unidad", "Precio Unitario", "Total", ""])
        self.activities_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.activities_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.EditKeyPressed)
        self.activities_table.itemChanged.connect(self.on_item_changed)
        activities_layout.addWidget(self.activities_table)

        main_layout.addLayout(activities_layout)

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
        main_layout.addLayout(totals_layout)

        # Botones de acción
        actions_layout = QHBoxLayout()
        generate_excel_btn = QPushButton("Generar Excel")
        generate_excel_btn.clicked.connect(self.generate_excel)
        actions_layout.addWidget(generate_excel_btn)
        
        generate_word_btn = QPushButton("Generar Word")
        generate_word_btn.clicked.connect(self.generate_word)
        actions_layout.addWidget(generate_word_btn)
        
        save_quotation_btn = QPushButton("Guardar Cotización")
        save_quotation_btn.clicked.connect(self.save_quotation)
        actions_layout.addWidget(save_quotation_btn)
        
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
        
    def toggle_theme(self):
        """Cambia entre tema claro y oscuro"""
        self.dark_mode = not self.dark_mode
        
        if self.dark_mode:
            # Aplicar tema oscuro
            self.theme_btn.setText("Modo Claro")
            self.apply_dark_theme()
        else:
            # Aplicar tema claro
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
        
        self.setPalette(palette)
    
    def apply_light_theme(self):
        """Aplica el tema claro a la aplicación"""
        self.setPalette(self.style().standardPalette())
        
    def generate_word(self):
        """Genera un documento Word a partir de la cotización actual"""
        # Verificar que haya actividades
        if self.activities_table.rowCount() == 0:
            QMessageBox.warning(self, "Error", "Debe agregar al menos una actividad a la cotización.")
            return
            
        # Verificar que haya un cliente seleccionado
        if self.client_combo.currentIndex() == -1 and not self.nombre_input.text():
            QMessageBox.warning(self, "Error", "Debe seleccionar o crear un cliente.")
            return
            
        # Primero guardar la cotización si no está guardada
        cotizacion_id = None
        excel_path = None
        
        # Generar Excel primero (necesario para la tabla)
        try:
            excel_path = self.generate_excel(show_message=False)
            if not excel_path:
                return
                
            # Si hay una cotización guardada, usarla
            if hasattr(self, 'last_quotation_id') and self.last_quotation_id:
                cotizacion_id = self.last_quotation_id
            else:
                # Guardar la cotización
                result = self.save_quotation(show_message=False)
                if result and 'id' in result:
                    cotizacion_id = result['id']
                    self.last_quotation_id = cotizacion_id
                else:
                    QMessageBox.warning(self, "Error", "No se pudo guardar la cotización.")
                    return
                    
            # Abrir diálogo de configuración Word
            word_dialog = WordConfigDialog(self.controller, cotizacion_id, excel_path, self)
            word_dialog.exec_()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al generar el documento Word: {str(e)}")
            
    def send_email(self):
        """Abre el diálogo para enviar cotización por correo"""
        # Verificar si hay archivos para enviar
        attachments = []
        
        # Si hay un Excel generado, añadirlo
        if hasattr(self, 'last_excel_path') and self.last_excel_path and os.path.exists(self.last_excel_path):
            attachments.append(self.last_excel_path)
            
        # Si hay un Word generado, añadirlo
        if hasattr(self, 'last_word_path') and self.last_word_path and os.path.exists(self.last_word_path):
            attachments.append(self.last_word_path)
            
        # Si no hay archivos, preguntar si desea generar Excel
        if not attachments:
            reply = QMessageBox.question(self, "Sin Archivos", 
                                       "No hay archivos para enviar. ¿Desea generar un Excel primero?",
                                       QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                excel_path = self.generate_excel(show_message=False)
                if excel_path:
                    attachments.append(excel_path)
                else:
                    return
            else:
                return
                
        # Abrir diálogo de envío de correo
        email_dialog = SendEmailDialog(attachments, self)
        email_dialog.exec_()

    def generate_excel(self, show_message=True):
        # Verificar que haya actividades
        if self.activities_table.rowCount() == 0:
            QMessageBox.warning(self, "Error", "Debe agregar al menos una actividad a la cotización.")
            return
            
        # Verificar que haya un cliente seleccionado
        if self.client_combo.currentIndex() == -1 and not self.nombre_input.text():
            QMessageBox.warning(self, "Error", "Debe seleccionar o crear un cliente.")
            return
            
        # Obtener el tipo de persona
        tipo_persona = self.tipo_combo.currentText()
        
        # Obtener valores AIU
        administracion = self.admin_input.value()
        imprevistos = self.imprevistos_input.value()
        utilidad = self.utilidad_input.value()
        iva_utilidad = self.iva_input.value()

        # Crear una lista para almacenar las actividades seleccionadas
        actividades = []

        # Recorrer las filas de la tabla
        for row in range(self.activities_table.rowCount()):
            descripcion = self.activities_table.item(row, 0).text()
            cantidad = float(self.activities_table.item(row, 1).text())
            unidad = self.activities_table.item(row, 2).text()
            valor_unitario = float(self.activities_table.item(row, 3).text())
            total = float(self.activities_table.item(row, 4).text())
            
            # Obtener el ID de la actividad si está disponible
            activity_id = None
            if self.activities_table.item(row, 0).data(Qt.UserRole):
                activity_id = self.activities_table.item(row, 0).data(Qt.UserRole)

            # Agregar los datos de la actividad a la lista
            actividades.append({
                'id': activity_id,
                'descripcion': descripcion,
                'cantidad': cantidad,
                'unidad': unidad,
                'valor_unitario': valor_unitario,
                'total': total
            })
            
        # Llamar al controlador para generar el Excel
        excel_path = self.xls.generate_excel(
            tipo_persona=tipo_persona, 
            activities=actividades, 
            administracion=administracion, 
            imprevistos=imprevistos, 
            utilidad=utilidad, 
            iva_utilidad=iva_utilidad
        )
        
        if show_message:
            QMessageBox.information(self, "Éxito", f"Excel creado correctamente en:\n{excel_path}")
            
        # Guardar la ruta del Excel para uso posterior
        self.last_excel_path = excel_path
        return excel_path
        
    def save_quotation(self, show_message=True):
        # Verificar que haya actividades
        if self.activities_table.rowCount() == 0:
            QMessageBox.warning(self, "Error", "Debe agregar al menos una actividad a la cotización.")
            return
            
        # Verificar que haya un cliente seleccionado o creado
        client_id = None
        if self.client_combo.currentIndex() != -1:
            clients = self.controller.get_all_clients()
            client_id = clients[self.client_combo.currentIndex()]['id']
        elif self.nombre_input.text():
            # Crear nuevo cliente
            client_data = {
                'nombre': self.nombre_input.text(),
                'tipo': self.tipo_combo.currentText(),
                'direccion': self.direccion_input.text(),
                'nit': self.nit_input.text() if self.tipo_combo.currentText() == 'jurídica' else None,
                'telefono': self.telefono_input.text(),
                'email': self.email_input.text()
            }
            client_id = self.controller.add_client(client_data)
            
            # Actualizar la lista de clientes y seleccionar el nuevo
            self.load_clientes()
            clients = self.controller.get_all_clients()
            for i, client in enumerate(clients):
                if client['id'] == client_id:
                    self.client_combo.setCurrentIndex(i)
                    break
        else:
            QMessageBox.warning(self, "Error", "Debe seleccionar o crear un cliente.")
            return
            
        # Crear lista de actividades
        activities = []
        for row in range(self.activities_table.rowCount()):
            descripcion = self.activities_table.item(row, 0).text()
            cantidad = float(self.activities_table.item(row, 1).text())
            unidad = self.activities_table.item(row, 2).text()
            valor_unitario = float(self.activities_table.item(row, 3).text())
            total = float(self.activities_table.item(row, 4).text())
            
            # Obtener el ID de la actividad si está disponible
            activity_id = None
            if self.activities_table.item(row, 0).data(Qt.UserRole):
                activity_id = self.activities_table.item(row, 0).data(Qt.UserRole)
            
            # Si no hay ID, buscar la actividad por descripción
            if activity_id is None:
                all_activities = self.controller.get_all_activities()
                for act in all_activities:
                    if act['descripcion'] == descripcion:
                        activity_id = act['id']
                        break
            
            # Si aún no hay ID, es una actividad personalizada, agregarla a la BD
            if activity_id is None:
                activity_data = {
                    'descripcion': descripcion,
                    'unidad': unidad,
                    'valor_unitario': valor_unitario
                }
                activity_id = self.controller.add_activity(activity_data)
                
            activities.append({
                'actividad_id': activity_id,
                'descripcion': descripcion,
                'cantidad': cantidad,
                'unidad': unidad,
                'valor_unitario': valor_unitario,
                'total': total
            })
            
        # Obtener valores AIU
        aiu_values = {
            'administracion': self.admin_input.value(),
            'imprevistos': self.imprevistos_input.value(),
            'utilidad': self.utilidad_input.value(),
            'iva_sobre_utilidad': self.iva_input.value()
        }
        
        # Guardar la cotización
        try:
            result = self.controller.generate_quotation(client_id, activities, aiu_values)
            
            if show_message:
                QMessageBox.information(self, "Éxito", 
                    f"Cotización guardada correctamente.\nNúmero: {result['numero']}\nTotal: {result['total']:.2f}")
            
            # Guardar el ID para uso posterior
            self.last_quotation_id = result['id']
            
            if show_message:
                # Preguntar si desea generar Excel
                reply = QMessageBox.question(self, "Generar Excel", 
                    "¿Desea generar el archivo Excel de esta cotización?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes)
                if reply == QMessageBox.Yes:
                    self.generate_excel()
            
            return result
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al guardar la cotización: {str(e)}")
            return None
    
    def guardar_cliente(self):
        """Guarda un nuevo cliente en la base de datos"""
        # Verificar que se haya ingresado un nombre
        if not self.nombre_input.text():
            QMessageBox.warning(self, "Error", "Debe ingresar un nombre para el cliente.")
            return
            
        # Crear diccionario con datos del cliente
        client_data = {
            'nombre': self.nombre_input.text(),
            'tipo': self.tipo_combo.currentText(),
            'direccion': self.direccion_input.text(),
            'nit': self.nit_input.text() if self.tipo_combo.currentText() == 'jurídica' else None,
            'telefono': self.telefono_input.text(),
            'email': self.email_input.text()
        }
        
        # Guardar cliente
        try:
            client_id = self.controller.add_client(client_data)
            QMessageBox.information(self, "Éxito", "Cliente guardado correctamente.")
            
            # Recargar lista de clientes
            self.load_clientes()
            
            # Seleccionar el cliente recién creado
            clients = self.controller.get_all_clients()
            for i, client in enumerate(clients):
                if client['id'] == client_id:
                    self.client_combo.setCurrentIndex(i)
                    # Actualizar los campos con los datos del cliente seleccionado
                    self.actualizar_campos_cliente(client)
                    break
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al guardar el cliente: {str(e)}")
    
    def load_clientes(self):
        """Carga la lista de clientes en el combo box"""
        self.client_combo.clear()
        clients = self.controller.get_all_clients()
        for client in clients:
            self.client_combo.addItem(f"{client['nombre']} ({client['tipo']})")
    
    def load_activities(self):
        """Carga la lista de actividades en el combo box"""
        self.activity_combo.clear()
        activities = self.controller.get_all_activities()
        for activity in activities:
            item_text = f"{activity['descripcion']} ({activity['unidad']} - ${activity['valor_unitario']})"
            self.activity_combo.addItem(item_text)
            # Guardar el ID de la actividad como dato de usuario
            self.activity_combo.setItemData(self.activity_combo.count() - 1, activity['id'], Qt.UserRole)
    
    def load_categories(self):
        """Carga la lista de categorías en el combo box"""
        self.category_combo.clear()
        self.category_combo.addItem("Todas las categorías", None)
        categories = self.controller.get_all_categories()
        for category in categories:
            self.category_combo.addItem(category['nombre'], category['id'])
    
    def filter_activities(self):
        """Filtra las actividades según el texto de búsqueda y la categoría seleccionada"""
        search_text = self.search_input.text()
        category_id = self.category_combo.currentData()
        
        self.activity_combo.clear()
        activities = self.controller.search_activities(search_text, category_id)
        
        for activity in activities:
            item_text = f"{activity['descripcion']} ({activity['unidad']} - ${activity['valor_unitario']})"
            self.activity_combo.addItem(item_text)
            # Guardar el ID de la actividad como dato de usuario
            self.activity_combo.setItemData(self.activity_combo.count() - 1, activity['id'], Qt.UserRole)
        
        # Actualizar actividades relacionadas si hay una actividad seleccionada
        self.update_related_activities()
    
    def update_related_activities(self):
        """Actualiza la lista de actividades relacionadas con la actividad seleccionada"""
        self.related_activities_combo.clear()
        
        if self.activity_combo.currentIndex() == -1:
            return
        
        activity_id = self.activity_combo.currentData(Qt.UserRole)
        if activity_id is None:
            return
        
        related_activities = self.controller.get_related_activities(activity_id)
        
        for activity in related_activities:
            item_text = f"{activity['descripcion']} ({activity['unidad']} - ${activity['valor_unitario']})"
            self.related_activities_combo.addItem(item_text)
            # Guardar el ID de la actividad como dato de usuario
            self.related_activities_combo.setItemData(self.related_activities_combo.count() - 1, activity['id'], Qt.UserRole)
    
    def actualizar_datos_cliente(self, index):
        """Actualiza los campos de cliente al seleccionar uno de la lista"""
        if index == -1:
            return
            
        clients = self.controller.get_all_clients()
        if index < len(clients):
            client = clients[index]
            self.actualizar_campos_cliente(client)
    
    def actualizar_campos_cliente(self, client):
        """Actualiza los campos de la interfaz con los datos del cliente"""
        self.tipo_combo.setCurrentText(client['tipo'])
        self.nombre_input.setText(client['nombre'])
        self.direccion_input.setText(client['direccion'] if client['direccion'] else '')
        self.nit_input.setText(client['nit'] if client['nit'] else '')
        self.telefono_input.setText(client['telefono'] if client['telefono'] else '')
        self.email_input.setText(client['email'] if client['email'] else '')
    
    def add_selected_activity(self):
        """Agrega la actividad seleccionada a la tabla"""
        index = self.activity_combo.currentIndex()
        if index == -1:
            return
            
        # Obtener datos de la actividad seleccionada
        activity_id = self.activity_combo.currentData(Qt.UserRole)
        activity_text = self.activity_combo.currentText()
        
        # Extraer descripción, unidad y valor unitario del texto
        parts = activity_text.split('(')
        descripcion = parts[0].strip()
        
        # Extraer unidad y valor unitario
        unit_price_part = parts[1].replace(')', '').strip()
        unit_price_parts = unit_price_part.split('-')
        unidad = unit_price_parts[0].strip()
        valor_unitario = float(unit_price_parts[1].strip().replace('$', ''))
        
        # Obtener cantidad
        cantidad = self.quantity_spinbox.value()
        
        # Calcular total
        total = cantidad * valor_unitario
        
        # Agregar a la tabla
        self.add_activity_to_table(
            descripcion,
            cantidad,
            unidad,
            valor_unitario,
            total,
            activity_id
        )
        
        # Actualizar totales
        self.update_totals()
        
        # Actualizar actividades relacionadas
        self.update_related_activities()
    
    def add_related_activity(self):
        """Agrega la actividad relacionada seleccionada a la tabla"""
        index = self.related_activities_combo.currentIndex()
        if index == -1:
            return
            
        # Obtener datos de la actividad seleccionada
        activity_id = self.related_activities_combo.currentData(Qt.UserRole)
        activity_text = self.related_activities_combo.currentText()
        
        # Extraer descripción, unidad y valor unitario del texto
        parts = activity_text.split('(')
        descripcion = parts[0].strip()
        
        # Extraer unidad y valor unitario
        unit_price_part = parts[1].replace(')', '').strip()
        unit_price_parts = unit_price_part.split('-')
        unidad = unit_price_parts[0].strip()
        valor_unitario = float(unit_price_parts[1].strip().replace('$', ''))
        
        # Obtener cantidad
        cantidad = self.quantity_spinbox.value()
        
        # Calcular total
        total = cantidad * valor_unitario
        
        # Agregar a la tabla
        self.add_activity_to_table(
            descripcion,
            cantidad,
            unidad,
            valor_unitario,
            total,
            activity_id
        )
        
        # Actualizar totales
        self.update_totals()
    
    def add_activity_to_table(self, descripcion, cantidad, unidad, valor_unitario, total, activity_id=None):
        """Agrega una actividad a la tabla"""
        row = self.activities_table.rowCount()
        self.activities_table.insertRow(row)
        
        # Crear item para descripción con el ID de la actividad como dato de usuario
        desc_item = EditableTableWidgetItem(descripcion, False)
        if activity_id is not None:
            desc_item.setData(Qt.UserRole, activity_id)
        
        # Agregar datos a la tabla
        self.activities_table.setItem(row, 0, desc_item)
        self.activities_table.setItem(row, 1, EditableTableWidgetItem(cantidad, True))
        self.activities_table.setItem(row, 2, EditableTableWidgetItem(unidad, False))
        self.activities_table.setItem(row, 3, EditableTableWidgetItem(valor_unitario, True))
        self.activities_table.setItem(row, 4, EditableTableWidgetItem(total, False))
        
        # Agregar botón para eliminar
        delete_btn = QPushButton("X")
        delete_btn.clicked.connect(lambda: self.delete_activity(row))
        self.activities_table.setCellWidget(row, 5, delete_btn)
    
    def delete_activity(self, row):
        """Elimina una actividad de la tabla"""
        self.activities_table.removeRow(row)
        self.update_totals()
    
    def on_item_changed(self, item):
        """Actualiza los cálculos cuando cambia un valor en la tabla"""
        if item.column() in [1, 3]:  # Cantidad o Valor Unitario
            row = item.row()
            try:
                cantidad = float(self.activities_table.item(row, 1).text())
                valor_unitario = float(self.activities_table.item(row, 3).text())
                total = cantidad * valor_unitario
                
                # Actualizar total
                self.activities_table.setItem(row, 4, EditableTableWidgetItem(total, False))
                
                # Actualizar totales generales
                self.update_totals()
            except (ValueError, AttributeError):
                pass
    
    def update_totals(self):
        """Actualiza los totales de la cotización"""
        subtotal = 0
        
        # Calcular subtotal
        for row in range(self.activities_table.rowCount()):
            try:
                total = float(self.activities_table.item(row, 4).text())
                subtotal += total
            except (ValueError, AttributeError):
                pass
        
        # Calcular IVA (19%)
        iva = subtotal * 0.19
        
        # Calcular total
        total = subtotal + iva
        
        # Actualizar etiquetas
        self.subtotal_label.setText(f"{subtotal:.2f}")
        self.iva_label.setText(f"{iva:.2f}")
        self.total_label.setText(f"{total:.2f}")
    
    def clear_form(self):
        """Limpia el formulario"""
        # Limpiar tabla de actividades
        self.activities_table.setRowCount(0)
        
        # Limpiar totales
        self.subtotal_label.setText("0.00")
        self.iva_label.setText("0.00")
        self.total_label.setText("0.00")
        
        # Limpiar cliente
        self.client_combo.setCurrentIndex(-1)
        self.nombre_input.clear()
        self.direccion_input.clear()
        self.nit_input.clear()
        self.telefono_input.clear()
        self.email_input.clear()
        
        # Limpiar variables de estado
        if hasattr(self, 'last_quotation_id'):
            delattr(self, 'last_quotation_id')
        if hasattr(self, 'last_excel_path'):
            delattr(self, 'last_excel_path')
        if hasattr(self, 'last_word_path'):
            delattr(self, 'last_word_path')
    
    def open_data_management(self):
        """Abre la ventana de gestión de datos"""
        try:
            data_window = DataManagementWindow(self.controller, self)
            data_window.exec_()
            
            # Recargar datos
            self.load_clientes()
            self.load_activities()
            self.load_categories()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al abrir la ventana de gestión de datos: {str(e)}")
