from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QComboBox, QTableWidget, \
    QTableWidgetItem, QMessageBox, QWidget, QDoubleSpinBox, QHeaderView, QTextEdit
from controllers.cotizacion_controller import CotizacionController
from models.cliente import Cliente
from views.data_management_window import DataManagementWindow
from PyQt5.QtCore import Qt, pyqtSlot
from controllers.excel_controller import ExcelController

class EditableTableWidgetItem(QTableWidgetItem):
    def __init__(self, value, editable=True):
        super().__init__(str(value))
        self.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        if editable:
            self.setFlags(self.flags() | Qt.ItemIsEditable)
            
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Cotizaciones")
        self.setGeometry(100, 100, 1000, 800)

        # Inicializa el controlador
        self.controller = CotizacionController()
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

        self.activity_combo = QComboBox()
        self.load_activities()
        activity_selection_layout.addWidget(self.activity_combo, 3)

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
        
        save_quotation_btn = QPushButton("Guardar Cotización")
        save_quotation_btn.clicked.connect(self.save_quotation)
        actions_layout.addWidget(save_quotation_btn)
        
        clear_btn = QPushButton("Limpiar")
        clear_btn.clicked.connect(self.clear_form)
        actions_layout.addWidget(clear_btn)
        
        main_layout.addLayout(actions_layout)

    def generate_excel(self):
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

            # Agregar los datos de la actividad a la lista
            actividades.append({
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
        
        QMessageBox.information(self, "Éxito", f"Excel creado correctamente en:\n{excel_path}")
        
    def save_quotation(self):
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

            activities.append({
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
            QMessageBox.information(self, "Éxito", 
                f"Cotización guardada correctamente.\nNúmero: {result['numero']}\nTotal: {result['total']:.2f}")
            
            # Preguntar si desea generar Excel
            reply = QMessageBox.question(self, "Generar Excel", 
                "¿Desea generar el archivo Excel de esta cotización?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
                
            if reply == QMessageBox.Yes:
                excel_path = self.xls.generate_excel(cotizacion_id=result['id'])
                QMessageBox.information(self, "Excel Generado", 
                    f"Excel creado correctamente en:\n{excel_path}")
                    
            # Limpiar el formulario
            self.clear_form()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al guardar la cotización: {str(e)}")
        
    def clear_form(self):
        """Limpia el formulario para una nueva cotización"""
        # Limpiar tabla de actividades
        self.activities_table.setRowCount(0)
        
        # Limpiar campos de cliente
        self.nombre_input.clear()
        self.direccion_input.clear()
        self.nit_input.clear()
        self.telefono_input.clear()
        self.email_input.clear()
        
        # Resetear combos
        self.client_combo.setCurrentIndex(-1)
        self.tipo_combo.setCurrentIndex(0)
        
        # Resetear totales
        self.update_totals()
        
    def load_clientes(self):
        """Carga los clientes en el combo box"""
        clientes = self.controller.get_all_clients()
        self.client_combo.clear()
        for cliente in clientes:
            self.client_combo.addItem(cliente['nombre'])
            
    def load_activities(self):
        """Carga las actividades en el combo box"""
        activities = self.controller.get_all_activities()
        self.activity_combo.clear()
        for activity in activities:
            self.activity_combo.addItem(activity['descripcion'], activity)

    def delete_activity(self, row):
        """Elimina una actividad de la tabla"""
        self.activities_table.removeRow(row)
        self.update_totals()

    def on_item_changed(self, item):
        """Maneja el cambio de valores en la tabla de actividades"""
        if item.column() in [1, 3]:  # Cantidad o Precio Unitario
            row = item.row()
            try:
                cantidad_item = self.activities_table.item(row, 1)
                precio_unitario_item = self.activities_table.item(row, 3)

                if cantidad_item is None or precio_unitario_item is None:
                    print(f"Error: Item no encontrado en la fila {row}")
                    return

                cantidad = float(cantidad_item.text())
                precio_unitario = float(precio_unitario_item.text())

                total = cantidad * precio_unitario
                total_item = self.activities_table.item(row, 4)
                if total_item is None:
                    total_item = QTableWidgetItem()
                    self.activities_table.setItem(row, 4, total_item)
                total_item.setText(f"{total:.2f}")

                self.update_totals()
            except ValueError as e:
                print(f"Error de valor en la fila {row}: {e}")
                QMessageBox.warning(self, "Error", "Por favor, ingrese valores numéricos válidos.")
                item.setText("1.00")  # Valor por defecto en caso de error
            except Exception as e:
                print(f"Error inesperado en la fila {row}: {e}")
                QMessageBox.warning(self, "Error", f"Ocurrió un error inesperado: {e}")

    def update_totals(self):
        """Actualiza los totales mostrados en la interfaz"""
        subtotal = 0
        for row in range(self.activities_table.rowCount()):
            total_item = self.activities_table.item(row, 4)
            if total_item:
                try:
                    subtotal += float(total_item.text())
                except (ValueError, TypeError):
                    pass

        # Aplicar IVA según el tipo de cliente
        tipo_cliente = self.tipo_combo.currentText()
        iva_rate = self.iva_input.value() / 100
        
        if tipo_cliente == 'jurídica':
            # Para clientes jurídicos, el IVA se aplica solo a la utilidad
            utilidad = subtotal * (self.utilidad_input.value() / 100)
            iva = utilidad * iva_rate
        else:
            # Para clientes naturales, el IVA se aplica al subtotal
            iva = subtotal * iva_rate
            
        total = subtotal + iva
        
        # Si es jurídica, añadir AIU al total
        if tipo_cliente == 'jurídica':
            admin = subtotal * (self.admin_input.value() / 100)
            imprevistos = subtotal * (self.imprevistos_input.value() / 100)
            utilidad = subtotal * (self.utilidad_input.value() / 100)
            total = subtotal + admin + imprevistos + utilidad + iva

        self.subtotal_label.setText(f"{subtotal:.2f}")
        self.iva_label.setText(f"{iva:.2f}")
        self.total_label.setText(f"{total:.2f}")

    def guardar_cliente(self):
        """Guarda un nuevo cliente en la base de datos"""
        nombre = self.nombre_input.text()
        tipo = self.tipo_combo.currentText()
        direccion = self.direccion_input.text()
        nit = self.nit_input.text() if self.tipo_combo.currentText() == 'jurídica' else None
        telefono = self.telefono_input.text()
        email = self.email_input.text()

        if nombre and direccion:
            cliente = Cliente(nombre=nombre, tipo=tipo, direccion=direccion, nit=nit, telefono=telefono, email=email)
            self.controller.add_client(cliente.__dict__)
            QMessageBox.information(self, "Éxito", "Cliente guardado correctamente.")
            self.limpiar_campos_cliente()
            # Recargar la lista de clientes
            self.load_clientes()
        else:
            QMessageBox.warning(self, "Error", "Por favor, completa todos los campos requeridos.")

    def limpiar_campos_cliente(self):
        """Limpia los campos de entrada de cliente"""
        self.nombre_input.clear()
        self.direccion_input.clear()
        self.nit_input.clear()
        self.telefono_input.clear()
        self.email_input.clear()
        self.tipo_combo.setCurrentIndex(0)

    @pyqtSlot()
    def open_data_management(self):
        """Abre la ventana de gestión de datos"""
        self.data_management_window = DataManagementWindow(self.controller)
        self.data_management_window.show()
        # Conectar señal de cierre para actualizar datos
        self.data_management_window.closed.connect(self.refresh_data)

    def refresh_data(self):
        """Actualiza los datos después de cambios en la ventana de gestión"""
        self.load_clientes()
        self.load_activities()

    def add_selected_activity(self):
        """Añade la actividad seleccionada a la tabla"""
        if self.activity_combo.currentData():
            activity = self.activity_combo.currentData()
            cantidad = self.quantity_spinbox.value()
            
            row_position = self.activities_table.rowCount()
            self.activities_table.insertRow(row_position)

            # Añadir datos a la tabla
            self.activities_table.setItem(row_position, 0, QTableWidgetItem(activity['descripcion']))
            self.activities_table.setItem(row_position, 1, QTableWidgetItem(str(cantidad)))
            self.activities_table.setItem(row_position, 2, QTableWidgetItem(activity['unidad']))
            self.activities_table.setItem(row_position, 3, QTableWidgetItem(str(activity['valor_unitario'])))

            # Calcular el total
            total = float(cantidad) * float(activity['valor_unitario'])
            self.activities_table.setItem(row_position, 4, QTableWidgetItem(f"{total:.2f}"))
            
            # Añadir botón de eliminar
            delete_btn = QPushButton("X")
            delete_btn.setStyleSheet("background-color: red; color: white;")
            delete_btn.clicked.connect(lambda: self.delete_activity(row_position))
            self.activities_table.setCellWidget(row_position, 5, delete_btn)

            self.update_totals()

    def actualizar_datos_cliente(self):
        """Actualiza los campos con los datos del cliente seleccionado"""
        index = self.client_combo.currentIndex()
        if index >= 0:
            clients = self.controller.get_all_clients()
            if index < len(clients):
                cliente_seleccionado = clients[index]
                
                # Llenar los campos de entrada con los datos del cliente
                self.nombre_input.setText(cliente_seleccionado['nombre'])
                self.direccion_input.setText(cliente_seleccionado.get('direccion', ''))
                self.nit_input.setText(cliente_seleccionado.get('nit', ''))
                self.telefono_input.setText(cliente_seleccionado.get('telefono', ''))
                self.email_input.setText(cliente_seleccionado.get('email', ''))
                
                # Establecer el tipo de cliente
                tipo = cliente_seleccionado.get('tipo', 'natural')
                index = self.tipo_combo.findText(tipo)
                if index >= 0:
                    self.tipo_combo.setCurrentIndex(index)
