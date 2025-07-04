from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget, QPushButton, QLineEdit, QFormLayout, QListWidget, QListWidgetItem, QMessageBox, QDoubleSpinBox, QComboBox, QLabel, QGroupBox, QSplitter, QTableWidget, QTableWidgetItem, QHeaderView, QCheckBox, QRadioButton, QButtonGroup
from PyQt5.QtCore import Qt, pyqtSignal, pyqtSlot
from utils.filter_manager import FilterManager

class DataManagementWindow(QMainWindow):
    # Señal para notificar cuando se cierra la ventana
    closed = pyqtSignal()
    
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.filter_manager = FilterManager(controller.database_manager)
        self.setWindowTitle("Gestión de Datos")
        self.setGeometry(150, 150, 700, 600)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Crear pestañas para clientes, actividades y productos
        self.tab_widget = QTabWidget()
        self.tab_widget.addTab(self.setup_client_tab(), "Clientes")
        self.tab_widget.addTab(self.setup_activity_tab(), "Actividades")
        self.tab_widget.addTab(self.setup_product_tab(), "Productos")
        self.tab_widget.addTab(self.setup_category_tab(), "Categorías")
        self.tab_widget.addTab(self.setup_relations_tab(), "Relaciones")

        main_layout.addWidget(self.tab_widget)
        
        # Inicializar todos los combos y listas
        self.refresh_all_data()
        
    def refresh_all_data(self):
        """Refresca todos los datos en la interfaz"""
        try:
            # Refrescar listas
            self.refresh_client_list()
            self.refresh_activity_list()
            self.refresh_product_list()
            self.refresh_category_list()
            
            # Refrescar combos
            self.refresh_category_combo(self.activity_category_combo)
            self.refresh_category_combo(self.activity_category_select)
            self.refresh_category_combo(self.product_category_combo)
            self.refresh_category_combo(self.product_category_select)
            
            # Refrescar combos de actividades
            self.refresh_activity_combo(self.main_activity_combo)
            self.refresh_activity_combo(self.related_activity_combo)
            self.refresh_activity_combo(self.activity_for_product_combo)
            
            # Refrescar combo de productos
            self.refresh_product_combo(self.product_combo)
        except Exception as e:
            print(f"Error al refrescar datos: {str(e)}")
        
    def closeEvent(self, event):
        # Emitir señal cuando se cierra la ventana
        self.closed.emit()
        super().closeEvent(event)

    def setup_client_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Formulario para agregar cliente
        form_layout = QFormLayout()
        self.client_name_input = QLineEdit()
        self.client_type_combo = QComboBox()
        self.client_type_combo.addItems(['Natural', 'Jurídica'])
        self.client_address_input = QLineEdit()
        self.client_nit_input = QLineEdit()
        self.client_phone_input = QLineEdit()
        self.client_email_input = QLineEdit()
        
        form_layout.addRow("Nombre:", self.client_name_input)
        form_layout.addRow("Tipo:", self.client_type_combo)
        form_layout.addRow("Dirección:", self.client_address_input)
        form_layout.addRow("NIT (para jurídica):", self.client_nit_input)
        form_layout.addRow("Teléfono:", self.client_phone_input)
        form_layout.addRow("Email:", self.client_email_input)

        # Botones para agregar y eliminar clientes
        buttons_layout = QHBoxLayout()
        add_client_btn = QPushButton("Agregar Cliente")
        add_client_btn.clicked.connect(self.add_client)
        delete_client_btn = QPushButton("Eliminar Cliente")
        delete_client_btn.clicked.connect(self.delete_client)
        buttons_layout.addWidget(add_client_btn)
        buttons_layout.addWidget(delete_client_btn)

        # Lista de clientes existentes
        self.client_list = QListWidget()
        self.client_list.itemClicked.connect(self.client_selected)

        layout.addLayout(form_layout)
        layout.addLayout(buttons_layout)
        layout.addWidget(QLabel("Lista de Clientes:"))
        layout.addWidget(self.client_list)

        return tab

    def setup_activity_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Sección de búsqueda y filtrado
        search_group = QGroupBox("Búsqueda y Filtrado")
        search_layout = QHBoxLayout()
        
        self.activity_search_input = QLineEdit()
        self.activity_search_input.setPlaceholderText("Buscar actividades...")
        self.activity_search_input.textChanged.connect(self.filter_activities)
        
        self.activity_category_combo = QComboBox()
        self.activity_category_combo.addItem("Todas las categorías", None)
        self.activity_category_combo.currentIndexChanged.connect(self.filter_activities)
        
        search_layout.addWidget(QLabel("Buscar:"))
        search_layout.addWidget(self.activity_search_input, 3)
        search_layout.addWidget(QLabel("Categoría:"))
        search_layout.addWidget(self.activity_category_combo, 2)
        
        search_group.setLayout(search_layout)
        layout.addWidget(search_group)

        # Formulario para agregar actividad
        form_group = QGroupBox("Detalles de Actividad")
        form_layout = QFormLayout()
        
        self.activity_description_input = QLineEdit()
        self.activity_unit_input = QLineEdit()
        self.activity_price_input = QDoubleSpinBox()
        self.activity_price_input.setRange(0, 1000000000)
        self.activity_price_input.setDecimals(2)
        
        self.activity_category_select = QComboBox()
        self.activity_category_select.addItem("Sin categoría", None)
        
        form_layout.addRow("Descripción:", self.activity_description_input)
        form_layout.addRow("Unidad:", self.activity_unit_input)
        form_layout.addRow("Valor Unitario:", self.activity_price_input)
        form_layout.addRow("Categoría:", self.activity_category_select)
        
        form_group.setLayout(form_layout)
        layout.addWidget(form_group)

        # Botones para agregar y eliminar actividades
        buttons_layout = QHBoxLayout()
        add_activity_btn = QPushButton("Agregar Actividad")
        add_activity_btn.clicked.connect(self.add_activity)
        update_activity_btn = QPushButton("Actualizar Actividad")
        update_activity_btn.clicked.connect(self.update_activity)
        delete_activity_btn = QPushButton("Eliminar Actividad")
        delete_activity_btn.clicked.connect(self.delete_activity)
        clear_activity_btn = QPushButton("Limpiar Campos")
        clear_activity_btn.clicked.connect(self.clear_activity_form)
        
        buttons_layout.addWidget(add_activity_btn)
        buttons_layout.addWidget(update_activity_btn)
        buttons_layout.addWidget(delete_activity_btn)
        buttons_layout.addWidget(clear_activity_btn)
        
        layout.addLayout(buttons_layout)

        # Lista de actividades existentes
        self.activity_list = QListWidget()
        self.activity_list.setWordWrap(True)
        self.activity_list.setMaximumWidth(700)
        self.activity_list.itemClicked.connect(self.activity_selected)

        layout.addWidget(QLabel("Lista de Actividades:"))
        layout.addWidget(self.activity_list)

        return tab

    def setup_product_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Sección de búsqueda y filtrado
        search_group = QGroupBox("Búsqueda y Filtrado")
        search_layout = QHBoxLayout()
        
        self.product_search_input = QLineEdit()
        self.product_search_input.setPlaceholderText("Buscar productos...")
        self.product_search_input.textChanged.connect(self.filter_products)
        
        self.product_category_combo = QComboBox()
        self.product_category_combo.addItem("Todas las categorías", None)
        self.product_category_combo.currentIndexChanged.connect(self.filter_products)
        
        search_layout.addWidget(QLabel("Buscar:"))
        search_layout.addWidget(self.product_search_input, 3)
        search_layout.addWidget(QLabel("Categoría:"))
        search_layout.addWidget(self.product_category_combo, 2)
        
        search_group.setLayout(search_layout)
        layout.addWidget(search_group)
        
        # Formulario para agregar producto
        form_group = QGroupBox("Detalles de Producto")
        form_layout = QFormLayout()
        
        self.product_name_input = QLineEdit()
        self.product_description_input = QLineEdit()
        self.product_unit_input = QLineEdit()
        self.product_price_input = QDoubleSpinBox()
        self.product_price_input.setRange(0, 1000000000)
        self.product_price_input.setDecimals(2)
        
        self.product_category_select = QComboBox()
        self.product_category_select.addItem("Sin categoría", None)
        
        form_layout.addRow("Nombre:", self.product_name_input)
        form_layout.addRow("Descripción:", self.product_description_input)
        form_layout.addRow("Unidad:", self.product_unit_input)
        form_layout.addRow("Precio Unitario:", self.product_price_input)
        form_layout.addRow("Categoría:", self.product_category_select)
        
        form_group.setLayout(form_layout)
        layout.addWidget(form_group)
        
        # Botones para agregar y eliminar productos
        buttons_layout = QHBoxLayout()
        add_product_btn = QPushButton("Agregar Producto")
        add_product_btn.clicked.connect(self.add_product)
        update_product_btn = QPushButton("Actualizar Producto")
        update_product_btn.clicked.connect(self.update_product)
        delete_product_btn = QPushButton("Eliminar Producto")
        delete_product_btn.clicked.connect(self.delete_product)
        clear_product_btn = QPushButton("Limpiar Campos")
        clear_product_btn.clicked.connect(self.clear_product_form)
        
        buttons_layout.addWidget(add_product_btn)
        buttons_layout.addWidget(update_product_btn)
        buttons_layout.addWidget(delete_product_btn)
        buttons_layout.addWidget(clear_product_btn)
        
        layout.addLayout(buttons_layout)
        
        # Lista de productos existentes
        self.product_list = QListWidget()
        self.product_list.itemClicked.connect(self.product_selected)
        
        layout.addWidget(QLabel("Lista de Productos:"))
        layout.addWidget(self.product_list)
        
        return tab
    
    def setup_category_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Formulario para agregar categoría
        form_group = QGroupBox("Detalles de Categoría")
        form_layout = QFormLayout()
        
        self.category_name_input = QLineEdit()
        self.category_description_input = QLineEdit()
        
        form_layout.addRow("Nombre:", self.category_name_input)
        form_layout.addRow("Descripción:", self.category_description_input)
        
        form_group.setLayout(form_layout)
        layout.addWidget(form_group)
        
        # Botones para agregar y eliminar categorías
        buttons_layout = QHBoxLayout()
        add_category_btn = QPushButton("Agregar Categoría")
        add_category_btn.clicked.connect(self.add_category)
        update_category_btn = QPushButton("Actualizar Categoría")
        update_category_btn.clicked.connect(self.update_category)
        delete_category_btn = QPushButton("Eliminar Categoría")
        delete_category_btn.clicked.connect(self.delete_category)
        clear_category_btn = QPushButton("Limpiar Campos")
        clear_category_btn.clicked.connect(self.clear_category_form)
        
        buttons_layout.addWidget(add_category_btn)
        buttons_layout.addWidget(update_category_btn)
        buttons_layout.addWidget(delete_category_btn)
        buttons_layout.addWidget(clear_category_btn)
        
        layout.addLayout(buttons_layout)
        
        # Lista de categorías existentes
        self.category_list = QListWidget()
        self.category_list.itemClicked.connect(self.category_selected)
        
        layout.addWidget(QLabel("Lista de Categorías:"))
        layout.addWidget(self.category_list)
        
        return tab
    
    def setup_relations_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)
        
        # Sección para relacionar actividades
        activities_group = QGroupBox("Relaciones entre Actividades")
        activities_layout = QVBoxLayout()
        
        # Selección de actividad principal
        main_activity_layout = QHBoxLayout()
        main_activity_layout.addWidget(QLabel("Actividad Principal:"))
        self.main_activity_combo = QComboBox()
        self.main_activity_combo.currentIndexChanged.connect(self.main_activity_changed)
        main_activity_layout.addWidget(self.main_activity_combo)
        
        activities_layout.addLayout(main_activity_layout)
        
        # Lista de actividades relacionadas
        self.related_activities_list = QListWidget()
        activities_layout.addWidget(QLabel("Actividades Relacionadas:"))
        activities_layout.addWidget(self.related_activities_list)
        
        # Agregar relación
        add_relation_layout = QHBoxLayout()
        add_relation_layout.addWidget(QLabel("Agregar Relación:"))
        self.related_activity_combo = QComboBox()
        add_relation_layout.addWidget(self.related_activity_combo)
        add_relation_btn = QPushButton("Agregar")
        add_relation_btn.clicked.connect(self.add_activity_relation)
        add_relation_layout.addWidget(add_relation_btn)
        
        activities_layout.addLayout(add_relation_layout)
        
        # Eliminar relación
        remove_relation_btn = QPushButton("Eliminar Relación Seleccionada")
        remove_relation_btn.clicked.connect(self.remove_activity_relation)
        activities_layout.addWidget(remove_relation_btn)
        
        activities_group.setLayout(activities_layout)
        layout.addWidget(activities_group)
        
        # Sección para relacionar actividades y productos
        products_group = QGroupBox("Relaciones entre Actividades y Productos")
        products_layout = QVBoxLayout()
        
        # Selección de actividad
        activity_product_layout = QHBoxLayout()
        activity_product_layout.addWidget(QLabel("Actividad:"))
        self.activity_for_product_combo = QComboBox()
        self.activity_for_product_combo.currentIndexChanged.connect(self.activity_for_product_changed)
        activity_product_layout.addWidget(self.activity_for_product_combo)
        
        products_layout.addLayout(activity_product_layout)
        
        # Lista de productos relacionados
        self.related_products_list = QListWidget()
        products_layout.addWidget(QLabel("Productos Relacionados:"))
        products_layout.addWidget(self.related_products_list)
        
        # Agregar relación con producto
        add_product_relation_layout = QHBoxLayout()
        add_product_relation_layout.addWidget(QLabel("Agregar Producto:"))
        self.product_combo = QComboBox()
        add_product_relation_layout.addWidget(self.product_combo)
        
        add_product_relation_layout.addWidget(QLabel("Cantidad:"))
        self.product_quantity_input = QDoubleSpinBox()
        self.product_quantity_input.setRange(0.01, 1000)
        self.product_quantity_input.setValue(1)
        add_product_relation_layout.addWidget(self.product_quantity_input)
        
        add_product_relation_btn = QPushButton("Agregar")
        add_product_relation_btn.clicked.connect(self.add_product_relation)
        add_product_relation_layout.addWidget(add_product_relation_btn)
        
        products_layout.addLayout(add_product_relation_layout)
        
        # Eliminar relación con producto
        remove_product_relation_btn = QPushButton("Eliminar Producto Seleccionado")
        remove_product_relation_btn.clicked.connect(self.remove_product_relation)
        products_layout.addWidget(remove_product_relation_btn)
        
        products_group.setLayout(products_layout)
        layout.addWidget(products_group)
        
        return tab

    def add_client(self):
        name = self.client_name_input.text()
        tipo = self.client_type_combo.currentText()
        direccion = self.client_address_input.text()
        nit = self.client_nit_input.text()
        telefono = self.client_phone_input.text()
        email = self.client_email_input.text()
        
        if not name:
            QMessageBox.warning(self, "Error", "El nombre del cliente es obligatorio.")
            return
        
        client_data = {
            'nombre': name,
            'tipo': tipo,
            'direccion': direccion,
            'nit': nit,
            'telefono': telefono,
            'email': email
        }
        
        try:
            self.controller.add_client(client_data)
            QMessageBox.information(self, "Éxito", "Cliente agregado correctamente.")
            self.refresh_client_list()
            self.clear_client_form()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al agregar cliente: {str(e)}")
    
    def delete_client(self):
        selected_items = self.client_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Debe seleccionar un cliente para eliminar.")
            return
        
        client_id = selected_items[0].data(Qt.UserRole)
        
        reply = QMessageBox.question(self, "Confirmar", "¿Está seguro de eliminar este cliente?",
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                self.controller.database_manager.delete_client(client_id)
                QMessageBox.information(self, "Éxito", "Cliente eliminado correctamente.")
                self.refresh_client_list()
                self.clear_client_form()
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al eliminar cliente: {str(e)}")
    
    def client_selected(self, item):
        client_id = item.data(Qt.UserRole)
        client_data = self.controller.database_manager.get_client_by_id(client_id)
        
        self.client_name_input.setText(client_data['nombre'])
        self.client_type_combo.setCurrentText(client_data['tipo'])
        self.client_address_input.setText(client_data['direccion'] if client_data['direccion'] else '')
        self.client_nit_input.setText(client_data['nit'] if client_data['nit'] else '')
        self.client_phone_input.setText(client_data['telefono'] if client_data['telefono'] else '')
        self.client_email_input.setText(client_data['email'] if client_data['email'] else '')
    
    def clear_client_form(self):
        self.client_name_input.clear()
        self.client_type_combo.setCurrentIndex(0)
        self.client_address_input.clear()
        self.client_nit_input.clear()
        self.client_phone_input.clear()
        self.client_email_input.clear()
    
    def refresh_client_list(self):
        try:
            self.client_list.clear()
            clients = self.controller.get_all_clients()
            for client in clients:
                item = QListWidgetItem(f"{client['nombre']} ({client['tipo']})")
                item.setData(Qt.UserRole, client['id'])
                self.client_list.addItem(item)
        except Exception as e:
            print(f"Error al refrescar lista de clientes: {str(e)}")
    
    def add_activity(self):
        description = self.activity_description_input.text()
        unit = self.activity_unit_input.text()
        price = self.activity_price_input.value()
        category_id = self.activity_category_select.currentData()
        
        if not description or not unit:
            QMessageBox.warning(self, "Error", "La descripción y la unidad son obligatorias.")
            return
        
        activity_data = {
            'descripcion': description,
            'unidad': unit,
            'valor_unitario': price,
            'categoria_id': category_id
        }
        
        try:
            self.controller.add_activity(activity_data)
            QMessageBox.information(self, "Éxito", "Actividad agregada correctamente.")
            self.refresh_activity_list()
            self.clear_activity_form()
            
            # Actualizar combos de actividades en la pestaña de relaciones
            self.refresh_activity_combo(self.main_activity_combo)
            self.refresh_activity_combo(self.related_activity_combo)
            self.refresh_activity_combo(self.activity_for_product_combo)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al agregar actividad: {str(e)}")
    
    def update_activity(self):
        selected_items = self.activity_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Debe seleccionar una actividad para actualizar.")
            return
        
        activity_id = selected_items[0].data(Qt.UserRole)
        description = self.activity_description_input.text()
        unit = self.activity_unit_input.text()
        price = self.activity_price_input.value()
        category_id = self.activity_category_select.currentData()
        
        if not description or not unit:
            QMessageBox.warning(self, "Error", "La descripción y la unidad son obligatorias.")
            return
        
        activity_data = {
            'descripcion': description,
            'unidad': unit,
            'valor_unitario': price,
            'categoria_id': category_id
        }
        
        try:
            self.controller.update_activity(activity_id, activity_data)
            QMessageBox.information(self, "Éxito", "Actividad actualizada correctamente.")
            self.refresh_activity_list()
            
            # Actualizar combos de actividades en la pestaña de relaciones
            self.refresh_activity_combo(self.main_activity_combo)
            self.refresh_activity_combo(self.related_activity_combo)
            self.refresh_activity_combo(self.activity_for_product_combo)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al actualizar actividad: {str(e)}")
    
    def delete_activity(self):
        selected_items = self.activity_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Debe seleccionar una actividad para eliminar.")
            return
        
        activity_id = selected_items[0].data(Qt.UserRole)
        
        reply = QMessageBox.question(self, "Confirmar", "¿Está seguro de eliminar esta actividad?",
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                self.controller.delete_activity(activity_id)
                QMessageBox.information(self, "Éxito", "Actividad eliminada correctamente.")
                self.refresh_activity_list()
                self.clear_activity_form()
                
                # Actualizar combos de actividades en la pestaña de relaciones
                self.refresh_activity_combo(self.main_activity_combo)
                self.refresh_activity_combo(self.related_activity_combo)
                self.refresh_activity_combo(self.activity_for_product_combo)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al eliminar actividad: {str(e)}")
    
    def activity_selected(self, item):
        activity_id = item.data(Qt.UserRole)
        activity_data = self.controller.database_manager.get_activity_by_id(activity_id)
        
        self.activity_description_input.setText(activity_data['descripcion'])
        self.activity_unit_input.setText(activity_data['unidad'])
        self.activity_price_input.setValue(activity_data['valor_unitario'])
        
        # Seleccionar categoría
        category_id = activity_data.get('categoria_id')
        index = self.activity_category_select.findData(category_id)
        self.activity_category_select.setCurrentIndex(index if index >= 0 else 0)
    
    def clear_activity_form(self):
        self.activity_description_input.clear()
        self.activity_unit_input.clear()
        self.activity_price_input.setValue(0)
        self.activity_category_select.setCurrentIndex(0)

    def refresh_activity_list(self):
        try:
            self.activity_list.clear()
            activities = self.controller.get_all_activities()
            for activity in activities:
                # Formato en múltiples líneas
                text = f"{activity['descripcion']}\n   📏 {activity['unidad']} | 💰 ${activity['valor_unitario']:,.2f}"

                item = QListWidgetItem(text)
                item.setData(Qt.UserRole, activity['id'])
                self.activity_list.addItem(item)
        except Exception as e:
            print(f"Error al refrescar lista de actividades: {str(e)}")
    
    def filter_activities(self):
        try:
            search_text = self.activity_search_input.text()
            category_id = self.activity_category_combo.currentData()
            
            self.activity_list.clear()
            activities = self.controller.search_activities(search_text, category_id)
            
            for activity in activities:
                text = f"{activity['descripcion']}\n   📏 {activity['unidad']} | 💰 ${activity['valor_unitario']:,.2f}"
                item = QListWidgetItem(text)
                item.setData(Qt.UserRole, activity['id'])
                self.activity_list.addItem(item)
        except Exception as e:
            print(f"Error al filtrar actividades: {str(e)}")
    
    def add_product(self):
        name = self.product_name_input.text()
        description = self.product_description_input.text()
        unit = self.product_unit_input.text()
        price = self.product_price_input.value()
        category_id = self.product_category_select.currentData()
        
        if not name or not unit:
            QMessageBox.warning(self, "Error", "El nombre y la unidad son obligatorios.")
            return
        
        product_data = {
            'nombre': name,
            'descripcion': description,
            'unidad': unit,
            'precio_unitario': price,
            'categoria_id': category_id
        }
        
        try:
            self.controller.database_manager.add_product(product_data)
            QMessageBox.information(self, "Éxito", "Producto agregado correctamente.")
            self.refresh_product_list()
            self.clear_product_form()
            
            # Actualizar combo de productos en la pestaña de relaciones
            self.refresh_product_combo(self.product_combo)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al agregar producto: {str(e)}")
    
    def update_product(self):
        selected_items = self.product_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Debe seleccionar un producto para actualizar.")
            return
        
        product_id = selected_items[0].data(Qt.UserRole)
        name = self.product_name_input.text()
        description = self.product_description_input.text()
        unit = self.product_unit_input.text()
        price = self.product_price_input.value()
        category_id = self.product_category_select.currentData()
        
        if not name or not unit:
            QMessageBox.warning(self, "Error", "El nombre y la unidad son obligatorios.")
            return
        
        product_data = {
            'nombre': name,
            'descripcion': description,
            'unidad': unit,
            'precio_unitario': price,
            'categoria_id': category_id
        }
        
        try:
            self.controller.database_manager.update_product(product_id, product_data)
            QMessageBox.information(self, "Éxito", "Producto actualizado correctamente.")
            self.refresh_product_list()
            
            # Actualizar combo de productos en la pestaña de relaciones
            self.refresh_product_combo(self.product_combo)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al actualizar producto: {str(e)}")
    
    def delete_product(self):
        selected_items = self.product_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Debe seleccionar un producto para eliminar.")
            return
        
        product_id = selected_items[0].data(Qt.UserRole)
        
        reply = QMessageBox.question(self, "Confirmar", "¿Está seguro de eliminar este producto?",
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                self.controller.database_manager.delete_product(product_id)
                QMessageBox.information(self, "Éxito", "Producto eliminado correctamente.")
                self.refresh_product_list()
                self.clear_product_form()
                
                # Actualizar combo de productos en la pestaña de relaciones
                self.refresh_product_combo(self.product_combo)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al eliminar producto: {str(e)}")
    
    def product_selected(self, item):
        product_id = item.data(Qt.UserRole)
        product_data = self.controller.database_manager.get_product_by_id(product_id)
        
        self.product_name_input.setText(product_data['nombre'])
        self.product_description_input.setText(product_data.get('descripcion', ''))
        self.product_unit_input.setText(product_data['unidad'])
        self.product_price_input.setValue(product_data['precio_unitario'])
        
        # Seleccionar categoría
        category_id = product_data.get('categoria_id')
        index = self.product_category_select.findData(category_id)
        self.product_category_select.setCurrentIndex(index if index >= 0 else 0)
    
    def clear_product_form(self):
        self.product_name_input.clear()
        self.product_description_input.clear()
        self.product_unit_input.clear()
        self.product_price_input.setValue(0)
        self.product_category_select.setCurrentIndex(0)
    
    def refresh_product_list(self):
        try:
            self.product_list.clear()
            products = self.controller.get_all_products()
            for product in products:
                item = QListWidgetItem(f"{product['nombre']} ({product['unidad']} - ${product['precio_unitario']})")
                item.setData(Qt.UserRole, product['id'])
                self.product_list.addItem(item)
        except Exception as e:
            print(f"Error al refrescar lista de productos: {str(e)}")
    
    def filter_products(self):
        try:
            search_text = self.product_search_input.text()
            category_id = self.product_category_combo.currentData()
            
            self.product_list.clear()
            products = self.filter_manager.search_products(search_text, category_id)
            
            for product in products:
                item = QListWidgetItem(f"{product['nombre']} ({product['unidad']} - ${product['precio_unitario']})")
                item.setData(Qt.UserRole, product['id'])
                self.product_list.addItem(item)
        except Exception as e:
            print(f"Error al filtrar productos: {str(e)}")
    
    def add_category(self):
        name = self.category_name_input.text()
        description = self.category_description_input.text()
        
        if not name:
            QMessageBox.warning(self, "Error", "El nombre de la categoría es obligatorio.")
            return
        
        category_data = {
            'nombre': name,
            'descripcion': description
        }
        
        try:
            self.controller.database_manager.add_category(category_data)
            QMessageBox.information(self, "Éxito", "Categoría agregada correctamente.")
            self.refresh_category_list()
            self.clear_category_form()
            
            # Actualizar combos de categorías
            self.refresh_category_combo(self.activity_category_combo)
            self.refresh_category_combo(self.activity_category_select)
            self.refresh_category_combo(self.product_category_combo)
            self.refresh_category_combo(self.product_category_select)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al agregar categoría: {str(e)}")
    
    def update_category(self):
        selected_items = self.category_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Debe seleccionar una categoría para actualizar.")
            return
        
        category_id = selected_items[0].data(Qt.UserRole)
        name = self.category_name_input.text()
        description = self.category_description_input.text()
        
        if not name:
            QMessageBox.warning(self, "Error", "El nombre de la categoría es obligatorio.")
            return
        
        category_data = {
            'nombre': name,
            'descripcion': description
        }
        
        try:
            self.controller.database_manager.update_category(category_id, category_data)
            QMessageBox.information(self, "Éxito", "Categoría actualizada correctamente.")
            self.refresh_category_list()
            
            # Actualizar combos de categorías
            self.refresh_category_combo(self.activity_category_combo)
            self.refresh_category_combo(self.activity_category_select)
            self.refresh_category_combo(self.product_category_combo)
            self.refresh_category_combo(self.product_category_select)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al actualizar categoría: {str(e)}")
    
    def delete_category(self):
        selected_items = self.category_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Error", "Debe seleccionar una categoría para eliminar.")
            return
        
        category_id = selected_items[0].data(Qt.UserRole)
        
        reply = QMessageBox.question(self, "Confirmar", "¿Está seguro de eliminar esta categoría?",
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            try:
                self.controller.database_manager.delete_category(category_id)
                QMessageBox.information(self, "Éxito", "Categoría eliminada correctamente.")
                self.refresh_category_list()
                self.clear_category_form()
                
                # Actualizar combos de categorías
                self.refresh_category_combo(self.activity_category_combo)
                self.refresh_category_combo(self.activity_category_select)
                self.refresh_category_combo(self.product_category_combo)
                self.refresh_category_combo(self.product_category_select)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al eliminar categoría: {str(e)}")
    
    def category_selected(self, item):
        category_id = item.data(Qt.UserRole)
        category_data = self.controller.database_manager.get_category_by_id(category_id)
        
        self.category_name_input.setText(category_data['nombre'])
        self.category_description_input.setText(category_data.get('descripcion', ''))
    
    def clear_category_form(self):
        self.category_name_input.clear()
        self.category_description_input.clear()
    
    def refresh_category_list(self):
        try:
            self.category_list.clear()
            categories = self.controller.get_all_categories()
            for category in categories:
                item = QListWidgetItem(category['nombre'])
                item.setData(Qt.UserRole, category['id'])
                self.category_list.addItem(item)
        except Exception as e:
            print(f"Error al refrescar lista de categorías: {str(e)}")
    
    def refresh_category_combo(self, combo):
        try:
            # Guardar el ítem seleccionado actualmente
            current_data = combo.currentData()
            
            # Limpiar el combo
            combo.clear()
            
            # Agregar el ítem "Sin categoría" o "Todas las categorías"
            if combo == self.activity_category_combo or combo == self.product_category_combo:
                combo.addItem("Todas las categorías", None)
            else:
                combo.addItem("Sin categoría", None)
            
            # Agregar las categorías
            categories = self.controller.get_all_categories()
            for category in categories:
                combo.addItem(category['nombre'], category['id'])
            
            # Restaurar el ítem seleccionado si existe
            if current_data is not None:
                index = combo.findData(current_data)
                if index >= 0:
                    combo.setCurrentIndex(index)
        except Exception as e:
            print(f"Error al refrescar combo de categorías: {str(e)}")
    
    def refresh_activity_combo(self, combo):
        try:
            # Guardar el ítem seleccionado actualmente
            current_data = combo.currentData()
            
            # Limpiar el combo
            combo.clear()
            
            # Agregar las actividades
            activities = self.controller.get_all_activities()
            for activity in activities:
                description = activity['descripcion']
                if len(description) > 60:
                    description = description[:37] + "..."

                text = f"{description} ({activity['unidad']} - ${activity['valor_unitario']:,.2f})"
                combo.addItem(text, activity['id'])
            
            # Restaurar el ítem seleccionado si existe
            if current_data is not None:
                index = combo.findData(current_data)
                if index >= 0:
                    combo.setCurrentIndex(index)
        except Exception as e:
            print(f"Error al refrescar combo de actividades: {str(e)}")
    
    def refresh_product_combo(self, combo):
        try:
            # Guardar el ítem seleccionado actualmente
            current_data = combo.currentData()
            
            # Limpiar el combo
            combo.clear()
            
            # Agregar los productos
            products = self.controller.get_all_products()
            for product in products:
                combo.addItem(f"{product['nombre']} ({product['unidad']} - ${product['precio_unitario']})", product['id'])
            
            # Restaurar el ítem seleccionado si existe
            if current_data is not None:
                index = combo.findData(current_data)
                if index >= 0:
                    combo.setCurrentIndex(index)
        except Exception as e:
            print(f"Error al refrescar combo de productos: {str(e)}")
    
    def main_activity_changed(self, index):
        try:
            if index < 0:
                return
                
            activity_id = self.main_activity_combo.itemData(index)
            self.refresh_related_activities_list(activity_id)
        except Exception as e:
            print(f"Error al cambiar actividad principal: {str(e)}")
    
    def refresh_related_activities_list(self, activity_id):
        try:
            self.related_activities_list.clear()
            related_activities = self.controller.get_related_activities(activity_id)
            
            for activity in related_activities:
                text = f"{activity['descripcion']}\n   📏 {activity['unidad']} | 💰 ${activity['valor_unitario']:,.2f}"
                item = QListWidgetItem(text)
                item.setData(Qt.UserRole, activity['id'])
                item.setData(Qt.UserRole + 1, activity.get('relation_id'))
                self.related_activities_list.addItem(item)
        except Exception as e:
            print(f"Error al refrescar lista de actividades relacionadas: {str(e)}")
    
    def add_activity_relation(self):
        try:
            main_index = self.main_activity_combo.currentIndex()
            related_index = self.related_activity_combo.currentIndex()
            
            if main_index < 0 or related_index < 0:
                QMessageBox.warning(self, "Error", "Debe seleccionar ambas actividades.")
                return
                
            main_activity_id = self.main_activity_combo.itemData(main_index)
            related_activity_id = self.related_activity_combo.itemData(related_index)
            
            if main_activity_id == related_activity_id:
                QMessageBox.warning(self, "Error", "No puede relacionar una actividad consigo misma.")
                return
                
            # Verificar si ya existe la relación
            related_activities = self.controller.get_related_activities(main_activity_id)
            for activity in related_activities:
                if activity['id'] == related_activity_id:
                    QMessageBox.warning(self, "Error", "Esta relación ya existe.")
                    return
            
            # Agregar la relación
            self.controller.database_manager.add_related_activity(main_activity_id, related_activity_id)
            QMessageBox.information(self, "Éxito", "Relación agregada correctamente.")
            
            # Actualizar la lista
            self.refresh_related_activities_list(main_activity_id)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al agregar relación: {str(e)}")
    
    def remove_activity_relation(self):
        try:
            selected_items = self.related_activities_list.selectedItems()
            if not selected_items:
                QMessageBox.warning(self, "Error", "Debe seleccionar una relación para eliminar.")
                return
                
            relation_id = selected_items[0].data(Qt.UserRole + 1)
            
            reply = QMessageBox.question(self, "Confirmar", "¿Está seguro de eliminar esta relación?",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.controller.database_manager.delete_related_activity(relation_id)
                QMessageBox.information(self, "Éxito", "Relación eliminada correctamente.")
                
                # Actualizar la lista
                main_index = self.main_activity_combo.currentIndex()
                if main_index >= 0:
                    main_activity_id = self.main_activity_combo.itemData(main_index)
                    self.refresh_related_activities_list(main_activity_id)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al eliminar relación: {str(e)}")
    
    def activity_for_product_changed(self, index):
        try:
            if index < 0:
                return
                
            activity_id = self.activity_for_product_combo.itemData(index)
            self.refresh_related_products_list(activity_id)
        except Exception as e:
            print(f"Error al cambiar actividad para productos: {str(e)}")
    
    def refresh_related_products_list(self, activity_id):
        try:
            self.related_products_list.clear()
            related_products = self.controller.database_manager.get_products_by_activity(activity_id)
            
            for product in related_products:
                item = QListWidgetItem(f"{product['nombre']} ({product['unidad']} - ${product['precio_unitario']}) - Cantidad: {product['cantidad']}")
                item.setData(Qt.UserRole, product['id'])
                item.setData(Qt.UserRole + 1, product.get('relation_id'))
                self.related_products_list.addItem(item)
        except Exception as e:
            print(f"Error al refrescar lista de productos relacionados: {str(e)}")
    
    def add_product_relation(self):
        try:
            activity_index = self.activity_for_product_combo.currentIndex()
            product_index = self.product_combo.currentIndex()
            
            if activity_index < 0 or product_index < 0:
                QMessageBox.warning(self, "Error", "Debe seleccionar actividad y producto.")
                return
                
            activity_id = self.activity_for_product_combo.itemData(activity_index)
            product_id = self.product_combo.itemData(product_index)
            quantity = self.product_quantity_input.value()
            
            # Verificar si ya existe la relación
            related_products = self.controller.database_manager.get_products_by_activity(activity_id)
            for product in related_products:
                if product['id'] == product_id:
                    QMessageBox.warning(self, "Error", "Esta relación ya existe.")
                    return
            
            # Agregar la relación
            self.controller.database_manager.add_activity_product(activity_id, product_id, quantity)
            QMessageBox.information(self, "Éxito", "Relación agregada correctamente.")
            
            # Actualizar la lista
            self.refresh_related_products_list(activity_id)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al agregar relación: {str(e)}")
    
    def remove_product_relation(self):
        try:
            selected_items = self.related_products_list.selectedItems()
            if not selected_items:
                QMessageBox.warning(self, "Error", "Debe seleccionar una relación para eliminar.")
                return
                
            relation_id = selected_items[0].data(Qt.UserRole + 1)
            
            reply = QMessageBox.question(self, "Confirmar", "¿Está seguro de eliminar esta relación?",
                                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.controller.database_manager.delete_activity_product(relation_id)
                QMessageBox.information(self, "Éxito", "Relación eliminada correctamente.")
                
                # Actualizar la lista
                activity_index = self.activity_for_product_combo.currentIndex()
                if activity_index >= 0:
                    activity_id = self.activity_for_product_combo.itemData(activity_index)
                    self.refresh_related_products_list(activity_id)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al eliminar relación: {str(e)}")
