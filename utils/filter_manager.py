from utils.database_manager import DatabaseManager
from typing import List, Dict, Optional

class FilterManager:
    def __init__(self, database_manager):
        self.database_manager = database_manager

    def search_activities(self, search_text, category_id=None):
        """Busca actividades por texto y categoría."""
        try:
            cursor = self.database_manager.connection.cursor()
            query = """
                SELECT
                    a.id, a.descripcion, a.unidad, a.valor_unitario, a.categoria_id, c.nombre as categoria_nombre
                FROM
                    actividades a
                LEFT JOIN
                    categorias c ON a.categoria_id = c.id
                WHERE
                    a.descripcion LIKE ?
            """
            params = [f'%{search_text}%']

            if category_id:
                query += " AND a.categoria_id = ?"
                params.append(category_id)

            cursor.execute(query, tuple(params))
            activities = []
            for row in cursor.fetchall():
                activities.append({
                    'id': row[0],
                    'descripcion': row[1],
                    'unidad': row[2],
                    'valor_unitario': row[3],
                    'categoria_id': row[4],
                    'categoria_nombre': row[5]
                })
            return activities
        except Exception as e:
            print(f"Error al buscar actividades: {str(e)}")
            return []