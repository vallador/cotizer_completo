import json
from utils.cotizacion_file_manager import CotizacionFileManager
import os

class CotizacionController:
    def __init__(self, database_manager=None, aiu_manager=None):
        from utils.database_manager import DatabaseManager
        from utils.aiu_manager import AIUManager
        from utils.filter_manager import FilterManager
        
        # Definir la ruta de la base de datos
        db_path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'data', 'cotizaciones.db')
        
        self.database_manager = database_manager if database_manager else DatabaseManager(db_path)
        self.database_manager.connect()
        
        self.aiu_manager = aiu_manager if aiu_manager else AIUManager()
        self.file_manager = CotizacionFileManager()
        self.filter_manager = FilterManager(self.database_manager)
        
    def add_client(self, client_data):
        """Agrega un cliente a la base de datos"""
        return self.database_manager.add_client(client_data)
        
    def get_all_clients(self):
        """Obtiene todos los clientes de la base de datos"""
        return self.database_manager.get_all_clients()
        
    def add_activity(self, activity_data):
        """Agrega una actividad a la base de datos"""
        return self.database_manager.add_activity(activity_data)
        
    def update_activity(self, activity_id, activity_data):
        """Actualiza una actividad en la base de datos"""
        return self.database_manager.update_activity(activity_id, activity_data)
        
    def delete_activity(self, activity_id):
        """Elimina una actividad de la base de datos"""
        return self.database_manager.delete_activity(activity_id)
        
    def get_all_activities(self):
        """Obtiene todas las actividades de la base de datos"""
        return self.database_manager.get_all_activities()
        
    def get_all_products(self):
        """Obtiene todos los productos de la base de datos"""
        return self.database_manager.get_all_products()
        
    def get_all_categories(self):
        """Obtiene todas las categorías de la base de datos"""
        return self.database_manager.get_all_categories()
        
    def search_activities(self, search_text, category_id=None):
        """Busca actividades por texto y categoría"""
        return self.filter_manager.search_activities(search_text, category_id)
        
    def get_related_activities(self, activity_id):
        """Obtiene las actividades relacionadas con una actividad"""
        return self.database_manager.get_related_activities(activity_id)
        
    def generate_quotation(self, client_id, activities, aiu_values):
        """Genera una cotización y la guarda en la base de datos"""
        # Obtener datos del cliente
        client = self.database_manager.get_client_by_id(client_id)
        
        # Asegurar que cada actividad tenga un ID válido
        for i, activity in enumerate(activities):
            if 'id' not in activity or activity['id'] is None:
                # Asignar un ID temporal basado en el índice
                activity['id'] = i + 1
        
        # Calcular subtotal
        subtotal = sum(activity['total'] for activity in activities)
        
        # Calcular valores AIU
        administracion = subtotal * (aiu_values['administracion'] / 100)
        imprevistos = subtotal * (aiu_values['imprevistos'] / 100)
        utilidad = subtotal * (aiu_values['utilidad'] / 100)
        iva_utilidad = utilidad * (aiu_values['iva_sobre_utilidad'] / 100)
        
        # Calcular total
        total = subtotal + administracion + imprevistos + utilidad + iva_utilidad
        
        # Crear datos de la cotización
        quotation_data = {
            'cliente_id': client_id,
            'subtotal': subtotal,
            'administracion': administracion,
            'imprevistos': imprevistos,
            'utilidad': utilidad,
            'iva_utilidad': iva_utilidad,
            'total': total,
            'actividades': activities
        }
        
        # Guardar en la base de datos
        result = self.database_manager.save_quotation(quotation_data)
        
        # Agregar datos del cliente al resultado para uso posterior
        result['cliente'] = client
        
        return result
        
    def get_current_cotizacion_data(self):
        """
        Obtiene los datos de la cotización actual para guardar como archivo
        Esta función debe ser implementada por la interfaz gráfica
        """
        # Esta es una implementación de ejemplo, debe ser sobrescrita
        return None
        
    def load_cotizacion_from_file(self, cotizacion_data):
        """
        Carga una cotización desde un archivo
        Esta función debe ser implementada por la interfaz gráfica
        """
        # Esta es una implementación de ejemplo, debe ser sobrescrita
        return None
        
    def save_cotizacion_to_file(self, cotizacion_data):
        """Guarda una cotización como archivo"""
        return self.file_manager.guardar_cotizacion(cotizacion_data)
        
    def load_cotizacion_from_file_path(self, filepath):
        """Carga una cotización desde un archivo"""
        return self.file_manager.cargar_cotizacion(filepath)
        
    def list_cotizacion_files(self):
        """Lista todos los archivos de cotización disponibles"""
        return self.file_manager.listar_cotizaciones()
