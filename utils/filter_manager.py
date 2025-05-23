from utils.database_manager import DatabaseManager
from typing import List, Dict, Optional

class FilterManager:
    """
    Clase para gestionar el filtrado y las relaciones entre actividades y productos.
    Proporciona métodos para búsqueda avanzada, filtrado contextual y sugerencias
    para facilitar la selección de actividades y productos.
    """
    
    def __init__(self, database_manager: DatabaseManager):
        """
        Inicializa el gestor de filtros con una conexión a la base de datos.
        
        Args:
            database_manager: Instancia de DatabaseManager para acceder a la base de datos
        """
        self.db_manager = database_manager
    
    def search_activities(self, search_text: str, categoria_id: Optional[int] = None) -> List[Dict]:
        """
        Busca actividades que coincidan con el texto de búsqueda y opcionalmente
        filtra por categoría.
        
        Args:
            search_text: Texto a buscar en las descripciones de actividades
            categoria_id: ID de categoría para filtrar (opcional)
            
        Returns:
            Lista de actividades que coinciden con los criterios de búsqueda
        """
        try:
            if not search_text and categoria_id is None:
                return self.db_manager.get_all_activities()
                
            if search_text and categoria_id is not None:
                # Buscar por texto y categoría
                activities = self.db_manager.search_activities_by_text(search_text)
                return [a for a in activities if a.get('categoria_id') == categoria_id]
            elif search_text:
                # Buscar solo por texto
                return self.db_manager.search_activities_by_text(search_text)
            else:
                # Buscar solo por categoría
                return self.db_manager.get_activities_by_category(categoria_id)
        except Exception as e:
            print(f"Error en search_activities: {str(e)}")
            return []
    
    def get_related_activities(self, activity_id: int) -> List[Dict]:
        """
        Obtiene las actividades relacionadas con una actividad específica.
        
        Args:
            activity_id: ID de la actividad principal
            
        Returns:
            Lista de actividades relacionadas
        """
        try:
            return self.db_manager.get_related_activities(activity_id)
        except Exception as e:
            print(f"Error en get_related_activities: {str(e)}")
            return []
    
    def suggest_activities(self, context: str) -> List[Dict]:
        """
        Sugiere actividades basadas en un contexto (por ejemplo, "pintura").
        
        Args:
            context: Contexto para sugerir actividades
            
        Returns:
            Lista de actividades sugeridas
        """
        try:
            # Buscar categoría que coincida con el contexto
            categories = self.db_manager.get_all_categories()
            matching_categories = [c for c in categories if context.lower() in c['nombre'].lower()]
            
            if matching_categories:
                # Si hay categorías que coinciden, devolver actividades de esas categorías
                result = []
                for category in matching_categories:
                    result.extend(self.db_manager.get_activities_by_category(category['id']))
                return result
            else:
                # Si no hay categorías que coincidan, buscar por texto
                return self.db_manager.search_activities_by_text(context)
        except Exception as e:
            print(f"Error en suggest_activities: {str(e)}")
            return []
    
    def search_products(self, search_text: str, categoria_id: Optional[int] = None) -> List[Dict]:
        """
        Busca productos que coincidan con el texto de búsqueda y opcionalmente
        filtra por categoría.
        
        Args:
            search_text: Texto a buscar en los nombres y descripciones de productos
            categoria_id: ID de categoría para filtrar (opcional)
            
        Returns:
            Lista de productos que coinciden con los criterios de búsqueda
        """
        try:
            if not search_text and categoria_id is None:
                return self.db_manager.get_all_products()
                
            if search_text and categoria_id is not None:
                # Buscar por texto y categoría
                products = self.db_manager.search_products_by_text(search_text)
                return [p for p in products if p.get('categoria_id') == categoria_id]
            elif search_text:
                # Buscar solo por texto
                return self.db_manager.search_products_by_text(search_text)
            else:
                # Filtrar productos por categoría
                products = self.db_manager.get_all_products()
                return [p for p in products if p.get('categoria_id') == categoria_id]
        except Exception as e:
            print(f"Error en search_products: {str(e)}")
            return []
    
    def get_products_for_activity(self, activity_id: int) -> List[Dict]:
        """
        Obtiene los productos asociados a una actividad específica.
        
        Args:
            activity_id: ID de la actividad
            
        Returns:
            Lista de productos asociados a la actividad
        """
        try:
            return self.db_manager.get_products_by_activity(activity_id)
        except Exception as e:
            print(f"Error en get_products_for_activity: {str(e)}")
            return []
    
    def advanced_search(self, search_text: str) -> Dict[str, List[Dict]]:
        """
        Realiza una búsqueda avanzada que devuelve tanto actividades como productos
        que coinciden con el texto de búsqueda.
        
        Args:
            search_text: Texto de búsqueda
            
        Returns:
            Diccionario con actividades y productos que coinciden
        """
        try:
            # Si no hay texto de búsqueda, devolver listas vacías
            if not search_text:
                return {'actividades': [], 'productos': []}
            
            # Buscar actividades y productos
            activities = self.db_manager.search_activities_by_text(search_text)
            products = self.db_manager.search_products_by_text(search_text)
            
            return {
                'actividades': activities,
                'productos': products
            }
        except Exception as e:
            print(f"Error en advanced_search: {str(e)}")
            return {'actividades': [], 'productos': []}
