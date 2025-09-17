import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QPushButton,
    QListWidget, QLineEdit, QMessageBox, QAbstractItemView
)


class TextManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestor de Textos con Arrastrar y Soltar")
        self.setGeometry(100, 100, 400, 300)

        # Widget central
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Layout principal vertical
        self.layout = QVBoxLayout()
        central_widget.setLayout(self.layout)

        # Lista de textos con soporte para arrastrar
        self.list_widget = QListWidget()
        self.list_widget.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.layout.addWidget(self.list_widget)

        # Campo de texto y botón para agregar
        input_layout = QHBoxLayout()
        self.text_input = QLineEdit()
        self.text_input.setPlaceholderText("Escribe aquí tu texto...")
        btn_add = QPushButton("Agregar")
        btn_add.clicked.connect(self.agregar_texto)
        input_layout.addWidget(self.text_input)
        input_layout.addWidget(btn_add)
        self.layout.addLayout(input_layout)

    def agregar_texto(self):
        texto = self.text_input.text().strip()
        if texto:
            self.list_widget.addItem(texto)
            self.text_input.clear()
        else:
            QMessageBox.warning(self, "Advertencia", "El texto está vacío.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    ventana = TextManager()
    ventana.show()
    sys.exit(app.exec())
