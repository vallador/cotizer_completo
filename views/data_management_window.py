from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QTabWidget, QWidget, QPushButton, QLineEdit, QFormLayout, QListWidget, QMessageBox
from PyQt5.QtCore import pyqtSlot
class DataManagementWindow(QMainWindow):
    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.setWindowTitle("Gestión de Datos")
        self.setGeometry(150, 150, 600, 400)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Crear pestañas para clientes, actividades y productos
        self.tab_widget = QTabWidget()
        self.tab_widget.addTab(self.setup_client_tab(), "Clientes")
        self.tab_widget.addTab(self.setup_activity_tab(), "Actividades")
        self.tab_widget.addTab(self.setup_product_tab(), "Productos")

        main_layout.addWidget(self.tab_widget)

    def setup_client_tab(self):
        tab = QWidget()
        layout = QVBoxLayout(tab)

        # Formulario para agregar cliente
        form_layout = QFormLayout()
        self.client_name_input = QLineEdit()
        self.client_type_input = QLineEdit()
        form_layout.addRow("Nombre:", self.client_name_input)
        form_layout.addRow("Tipo:", self.client_type_input)

        add_client_btn = QPushButton("Agregar Cliente")
        add_client_btn.clicked.connect(self.add_client)

        # Lista de clientes existentes
        self.client_list = QListWidget()
        self.refresh_client_list()

        layout.addLayout(form_layout)
        layout.addWidget(add_client_btn)
        layout.addWidget(self.client_list)

        return tab

    def setup_activity_tab(self):
        # Similar a setup_client_tab, pero para actividades
        pass

    def setup_product_tab(self):
        # Similar a setup_client_tab, pero para productos
        pass

    def add_client(self):
        name = self.client_name_input.text()
        client_type = self.client_type_input.text()
        if name and client_type:
            self.controller.add_client({"nombre": name, "tipo": client_type})
            self.refresh_client_list()
            self.client_name_input.clear()
            self.client_type_input.clear()
            QMessageBox.information(self, "Éxito", "Cliente agregado correctamente.")
        else:
            QMessageBox.warning(self, "Error", "Por favor, complete todos los campos.")

    def refresh_client_list(self):
        self.client_list.clear()
        clients = self.controller.get_all_clients()
        for client in clients:
            self.client_list.addItem(f"{client['nombre']} ({client['tipo']})")