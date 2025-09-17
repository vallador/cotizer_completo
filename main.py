
import os
import sys

# Importaciones de PyQt5
from PyQt5.QtWidgets import QApplication

# Importaciones de los componentes de la aplicación
from utils.database_manager import DatabaseManager
from utils.aiu_manager import AIUManager
from controllers.cotizacion_controller import CotizacionController
from controllers.excel_controller import ExcelController
from views.main_window import MainWindow


def initialize_database_if_needed(db_path: str):
    """
    Verifica si el archivo de la base de datos existe. Si no, ejecuta el
    script de inicialización para crearla y poblarla.
    """
    if not os.path.exists(db_path):
        print(f"Base de datos no encontrada en '{db_path}'. Inicializando...")
        try:
            # Importación local para evitar que sea una dependencia global
            from reset_db import reset_database_auto
            reset_database_auto()
            print("Base de datos inicializada correctamente.")
        except ImportError:
            print("Error: No se pudo encontrar el script 'reset_db.py'.")
            sys.exit(1)
        except Exception as e:
            print(f"Error fatal al inicializar la base de datos: {e}")


            # Salir si la base de datos no se puede crear, ya que la app no puede funcionar.
            sys.exit(1)


def main():
    """
    Función principal que configura e inicia la aplicación de cotizaciones.
    """
    # --- 1. Configuración de Rutas y Verificación de la Base de Datos ---
    # Construir una ruta robusta al archivo de la base de datos
    base_dir = os.path.dirname(os.path.abspath(__file__))
    db_path = os.path.join(base_dir, 'data', 'cotizaciones.db')

    # Asegurarse de que la base de datos exista antes de continuar
    initialize_database_if_needed(db_path)

    # --- 2. Creación e Inyección de Dependencias (Patrón de Inversión de Control) ---

    # Nivel 1: Managers (componentes de bajo nivel, sin dependencias de controladores)
    print("Inicializando gestores...")
    db_manager = DatabaseManager(db_path=db_path)
    aiu_manager = AIUManager(database_manager=db_manager)

    # Nivel 2: Controladores (componentes de lógica de negocio, dependen de managers)
    print("Inicializando controladores...")
    cotizacion_controller = CotizacionController(database_manager=db_manager)
    excel_controller = ExcelController(cotizacion_controller=cotizacion_controller, aiu_manager=aiu_manager)

    # --- 3. Creación de la Interfaz Gráfica (GUI) ---

    # La vista (MainWindow) depende de los controladores para funcionar
    print("Creando la interfaz gráfica...")
    app = QApplication(sys.argv)

    # Se inyectan los controladores en la ventana principal
    main_window = MainWindow(
        cotizacion_controller=cotizacion_controller,
        excel_controller=excel_controller
    )
    main_window.show()
    print("Aplicación iniciada. Mostrando ventana principal.")

    # --- 4. Ejecución y Cierre Limpio de la Aplicación ---

    # app.exec_() inicia el bucle de eventos de la aplicación
    exit_code = app.exec_()

    # Cerrar recursos antes de salir del script
    print("Cerrando la aplicación...")
    db_manager.close()

    sys.exit(exit_code)


if __name__ == "__main__":
    main()