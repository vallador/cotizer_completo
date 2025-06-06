from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QListWidget, QListWidgetItem, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt
from utils.cotizacion_file_manager import CotizacionFileManager
import os

class CotizacionFileDialog(QDialog):
    """
    Diálogo para gestionar archivos de cotización (guardar, cargar, editar)
    """
    
    def __init__(self, controller, parent=None):
        super().__init__(parent)
        self.controller = controller
        self.file_manager = CotizacionFileManager()
        self.selected_cotizacion = None
        
        self.setWindowTitle("Gestión de Archivos de Cotización")
        self.setGeometry(200, 200, 700, 500)
        
        self.setup_ui()
        self.load_cotizaciones()
        
    def setup_ui(self):
        """Configura la interfaz de usuario del diálogo"""
        layout = QVBoxLayout(self)
        
        # Título
        title_label = QLabel("Gestión de Archivos de Cotización")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title_label)
        
        # Lista de cotizaciones
        self.cotizaciones_list = QListWidget()
        self.cotizaciones_list.setAlternatingRowColors(True)
        self.cotizaciones_list.itemDoubleClicked.connect(self.open_cotizacion)
        self.cotizaciones_list.itemClicked.connect(self.update_info_label)
        layout.addWidget(self.cotizaciones_list)
        
        # Información de la cotización seleccionada
        self.info_label = QLabel("Seleccione una cotización para ver detalles")
        layout.addWidget(self.info_label)
        
        # Botones de acción
        buttons_layout = QHBoxLayout()
        
        self.open_btn = QPushButton("Abrir")
        self.open_btn.clicked.connect(self.open_cotizacion)
        buttons_layout.addWidget(self.open_btn)
        
        self.save_btn = QPushButton("Guardar Actual")
        self.save_btn.clicked.connect(self.save_current_cotizacion)
        buttons_layout.addWidget(self.save_btn)
        
        self.import_btn = QPushButton("Importar")
        self.import_btn.clicked.connect(self.import_cotizacion)
        buttons_layout.addWidget(self.import_btn)
        
        self.export_btn = QPushButton("Exportar")
        self.export_btn.clicked.connect(self.export_cotizacion)
        buttons_layout.addWidget(self.export_btn)
        
        self.delete_btn = QPushButton("Eliminar")
        self.delete_btn.clicked.connect(self.delete_cotizacion)
        buttons_layout.addWidget(self.delete_btn)
        
        self.close_btn = QPushButton("Cerrar")
        self.close_btn.clicked.connect(self.close)
        buttons_layout.addWidget(self.close_btn)
        
        layout.addLayout(buttons_layout)
        
    def load_cotizaciones(self):
        """Carga la lista de cotizaciones disponibles"""
        self.cotizaciones_list.clear()
        cotizaciones = self.file_manager.listar_cotizaciones()
        
        for cotizacion in cotizaciones:
            # Crear texto para el item
            if 'error' in cotizacion:
                text = f"Error en archivo: {cotizacion['filename']}"
            else:
                text = f"#{cotizacion['numero']} - {cotizacion['cliente']} - Total: ${cotizacion.get('total', 0):.2f}"
            
            # Crear item y guardar ruta como dato
            item = QListWidgetItem(text)
            item.setData(Qt.UserRole, cotizacion['filepath'])
            
            # Si hay error, marcar en rojo
            if 'error' in cotizacion:
                item.setForeground(Qt.red)
            
            self.cotizaciones_list.addItem(item)
    
    def update_info_label(self):
        """Actualiza la etiqueta de información con los detalles de la cotización seleccionada"""
        selected_items = self.cotizaciones_list.selectedItems()
        if not selected_items:
            self.info_label.setText("Seleccione una cotización para ver detalles")
            self.selected_cotizacion = None
            return
            
        filepath = selected_items[0].data(Qt.UserRole)
        try:
            cotizacion = self.file_manager.cargar_cotizacion(filepath)
            self.selected_cotizacion = cotizacion  # Guardar la cotización seleccionada
            
            cliente = cotizacion.get('cliente', {}).get('nombre', 'Desconocido')
            fecha = cotizacion.get('fecha', 'Desconocida')
            numero = cotizacion.get('numero', '000')
            total = cotizacion.get('total', 0)
            actividades = len(cotizacion.get('actividades', []))
            
            info_text = f"Cotización #{numero} - Cliente: {cliente} - Fecha: {fecha} - Total: ${total:.2f} - Actividades: {actividades}"
            self.info_label.setText(info_text)
        except Exception as e:
            self.info_label.setText(f"Error al cargar información: {str(e)}")
            self.selected_cotizacion = None
    
    def get_selected_cotizacion(self):
        """
        Devuelve la cotización seleccionada actualmente.
        
        Returns:
            dict: Datos de la cotización seleccionada o None si no hay selección
        """
        selected_items = self.cotizaciones_list.selectedItems()
        if not selected_items:
            return None
            
        filepath = selected_items[0].data(Qt.UserRole)
        try:
            return self.file_manager.cargar_cotizacion(filepath)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al obtener la cotización seleccionada: {str(e)}")
            return None
    
    def open_cotizacion(self):
        """Abre la cotización seleccionada para edición"""
        selected_items = self.cotizaciones_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Advertencia", "Debe seleccionar una cotización para abrir.")
            return
            
        filepath = selected_items[0].data(Qt.UserRole)
        try:
            cotizacion = self.file_manager.cargar_cotizacion(filepath)
            self.selected_cotizacion = cotizacion  # Guardar la cotización seleccionada
            
            # Enviar la cotización al controlador para cargarla en la interfaz
            self.controller.load_cotizacion_from_file(cotizacion)
            
            QMessageBox.information(self, "Éxito", "Cotización cargada correctamente.")
            self.accept()  # Cerrar el diálogo
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al abrir la cotización: {str(e)}")
    
    def save_current_cotizacion(self):
        """Guarda la cotización actual como archivo"""
        try:
            # Obtener datos de la cotización actual
            cotizacion_data = self.controller.get_current_cotizacion_data()
            
            if not cotizacion_data:
                QMessageBox.warning(self, "Advertencia", "No hay una cotización activa para guardar.")
                return
                
            # Guardar como archivo
            filepath = self.file_manager.guardar_cotizacion(cotizacion_data)
            
            QMessageBox.information(self, "Éxito", f"Cotización guardada correctamente en:\n{filepath}")
            
            # Actualizar la lista
            self.load_cotizaciones()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al guardar la cotización: {str(e)}")
    
    def import_cotizacion(self):
        """Importa una cotización desde un archivo JSON externo"""
        filepath, _ = QFileDialog.getOpenFileName(
            self, "Importar Cotización", "", "Archivos JSON (*.json);;Archivos de Cotización (*.cotiz);;Todos los archivos (*)"
        )
        
        if not filepath:
            return
            
        try:
            # Cargar el archivo
            cotizacion = self.file_manager.cargar_cotizacion(filepath)
            self.selected_cotizacion = cotizacion  # Guardar la cotización seleccionada
            
            # Copiar al directorio de cotizaciones
            new_filepath = self.file_manager.guardar_cotizacion(cotizacion)
            
            QMessageBox.information(self, "Éxito", f"Cotización importada correctamente como:\n{os.path.basename(new_filepath)}")
            
            # Actualizar la lista
            self.load_cotizaciones()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al importar la cotización: {str(e)}")
    
    def export_cotizacion(self):
        """Exporta la cotización seleccionada a un archivo JSON externo"""
        selected_items = self.cotizaciones_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Advertencia", "Debe seleccionar una cotización para exportar.")
            return
            
        filepath = selected_items[0].data(Qt.UserRole)
        
        # Solicitar ubicación para guardar
        save_path, _ = QFileDialog.getSaveFileName(
            self, "Exportar Cotización", os.path.basename(filepath), "Archivos JSON (*.json);;Archivos de Cotización (*.cotiz)"
        )
        
        if not save_path:
            return
            
        try:
            # Cargar la cotización
            cotizacion = self.file_manager.cargar_cotizacion(filepath)
            
            # Guardar en la nueva ubicación
            with open(save_path, 'w', encoding='utf-8') as f:
                import json
                json.dump(cotizacion, f, ensure_ascii=False, indent=4)
            
            QMessageBox.information(self, "Éxito", f"Cotización exportada correctamente a:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al exportar la cotización: {str(e)}")
    
    def delete_cotizacion(self):
        """Elimina la cotización seleccionada"""
        selected_items = self.cotizaciones_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "Advertencia", "Debe seleccionar una cotización para eliminar.")
            return
            
        filepath = selected_items[0].data(Qt.UserRole)
        
        # Confirmar eliminación
        reply = QMessageBox.question(
            self, "Confirmar Eliminación", 
            "¿Está seguro de que desea eliminar esta cotización?\nEsta acción no se puede deshacer.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                os.remove(filepath)
                QMessageBox.information(self, "Éxito", "Cotización eliminada correctamente.")
                self.load_cotizaciones()
                self.selected_cotizacion = None  # Limpiar la cotización seleccionada
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al eliminar la cotización: {str(e)}")
