import os
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox
from views.main_window import MainWindow
from controllers.cotizacion_controller import CotizacionController

def main():
    """Función principal que inicia la aplicación"""
    # Verificar si existe el directorio data
    if not os.path.exists('data'):
        os.makedirs('data')
    
    # Verificar si existe la base de datos
    if not os.path.exists('data/cotizaciones.db'):
        # Si no existe, ejecutar el script de inicialización
        try:
            from reset_db import reset_database_auto
            reset_database_auto()
            print("Base de datos inicializada correctamente.")
        except Exception as e:
            print(f"Error al inicializar la base de datos: {e}")
            return
    
    # Crear la aplicación
    app = QApplication(sys.argv)
    
    # Crear el controlador
    controller = CotizacionController()
    
    # Crear la ventana principal
    main_window = MainWindow(controller)
    main_window.show()
    

    
    # Ejecutar la aplicación
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
