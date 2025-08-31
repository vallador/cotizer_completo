# controllers/cotizacion_controller.py

# Importa las clases, no las uses para crear instancias aquí
from utils.database_manager import DatabaseManager
from utils.aiu_manager import AIUManager
from utils.filter_manager import FilterManager
from utils.cotizacion_file_manager import CotizacionFileManager


class CotizacionController:
    def __init__(self, database_manager: DatabaseManager):
        """
        El controlador principal. Recibe una instancia ya creada del
        DatabaseManager y crea sus propios sub-componentes a partir de ella.
        """
        # --- 1. Guardar la dependencia principal ---
        self.database_manager = database_manager

        # --- 2. Crear e inicializar los managers que dependen del database_manager ---
        # Se le pasa la instancia de database_manager a los managers que la necesiten.
        self.aiu_manager = AIUManager(self.database_manager)
        self.filter_manager = FilterManager(self.database_manager)

        # --- 3. Inicializar managers independientes ---
        self.file_manager = CotizacionFileManager()

    # --- Métodos que delegan en el DatabaseManager ---
    # El controlador actúa como un intermediario (fachada)

    def add_client(self, client_data):
        """Agrega un cliente a la base de datos"""
        return self.database_manager.add_client(**client_data)

    def get_all_clients(self):
        """Obtiene todos los clientes de la base de datos"""
        return self.database_manager.get_all_clients()

    def add_chapter(self, chapter_data):
        """Agrega un capítulo a la base de datos"""
        # Pasa el diccionario completo (nombre y descripcion) al manager
        return self.database_manager.add_chapter(**chapter_data)

    def update_chapter(self, chapter_id, chapter_data):
        """Actualiza un capítulo en la base de datos"""
        return self.database_manager.update_chapter(chapter_id, **chapter_data)

    def get_all_chapters(self):
        """Obtiene todos los capítulos de la base de datos"""
        return self.database_manager.get_all_chapters()

    def delete_chapter(self, chapter_id):
        """Pasa la solicitud de eliminación al DatabaseManager."""
        print(f"DEBUG (Controller): Pasando solicitud para eliminar capítulo ID: {chapter_id}")
        return self.database_manager.delete_chapter(chapter_id)

    def add_activity(self, activity_data):
        """Agrega una actividad a la base de datos"""
        return self.database_manager.add_activity(**activity_data)
        return self.database_manager.add_activity(**activity_data)



    def update_activity(self, activity_id, activity_data):
        """Actualiza una actividad en la base de datos"""
        return self.database_manager.update_activity(activity_id, **activity_data)

    def delete_activity(self, activity_id):
        """Elimina una actividad de la base de datos"""
        return self.database_manager.delete_activity(activity_id)

    def get_all_activities(self):
        """Obtiene todas las actividades de la base de datos"""
        return self.database_manager.get_all_activities()

    def get_all_categories(self):
        """Obtiene todas las categorías de la base de datos"""
        return self.database_manager.get_all_categories()

    def get_all_products(self):
        return self.database_manager.get_all_products()

    def search_products(self, search_text, category_id=None):
        return self.filter_manager.search_products(search_text, category_id)

    # --- Métodos que delegan en otros managers ---

    def search_activities(self, search_text, category_id=None):
        """Busca actividades por texto y categoría"""
        return self.filter_manager.search_activities(search_text, category_id)

    def get_related_activities(self, activity_id):
        """
        Obtiene las actividades relacionadas con una actividad por su ID.
        Este método ahora delega correctamente en el database_manager.
        """
        # ¡CORRECCIÓN! La lógica se mueve al DatabaseManager, el controlador solo llama.
        return self.database_manager.get_related_activities(activity_id)

    def save_cotizacion_to_file(self, cotizacion_data, filepath):
        """Guarda una cotización como archivo"""
        return self.file_manager.guardar_cotizacion(cotizacion_data, filepath)

    def load_cotizacion_from_file_path(self, filepath):
        """Carga una cotización desde un archivo"""
        return self.file_manager.cargar_cotizacion(filepath)

    def list_cotizacion_files(self):
        """Lista todos los archivos de cotización disponibles"""
        return self.file_manager.listar_cotizaciones()

    # --- Métodos con lógica propia del controlador ---

    def generate_quotation(self, client_id, activities, aiu_values):
        """Genera una cotización (lógica de negocio)"""
        # (Este método parece no usarse, pero lo dejamos por si acaso)
        client = self.database_manager.get_client_by_id(client_id)
        subtotal = sum(activity['total'] for activity in activities)
        administracion = subtotal * (aiu_values['administracion'] / 100)
        imprevistos = subtotal * (aiu_values['imprevistos'] / 100)
        utilidad = subtotal * (aiu_values['utilidad'] / 100)
        iva_utilidad = utilidad * (aiu_values['iva_sobre_utilidad'] / 100)
        total = subtotal + administracion + imprevistos + utilidad + iva_utilidad

        quotation_data = {
            'cliente_id': client_id,
            'subtotal': subtotal,
            'administracion': administracion,
            'imprevistos': imprevistos,
            'utilidad': utilidad,
            'iva_utilidad': iva_utilidad,
            'total': total,
            'actividades': activities,
            'cliente': client
        }
        return quotation_data
