import sqlite3
from typing import List, Dict, Optional


class CapitulosManager:
    """
    Extensión para DatabaseManager que maneja la funcionalidad de capítulos de obra
    sin modificar la estructura original.
    """

    def __init__(self, database_manager):
        """
        Inicializa el gestor de capítulos con una referencia al DatabaseManager existente.
        """
        self.db_manager = database_manager
        self.conn = database_manager.conn
        self.cursor = database_manager.cursor

    def create_capitulos_table(self):
        """
        Crea la tabla de capítulos si no existe.
        """
        try:
            self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS capitulos (
                id INTEGER PRIMARY KEY,
                nombre TEXT NOT NULL,
                descripcion TEXT,
                orden INTEGER DEFAULT 0
            )
            ''')

            # Verificar si la columna capitulo_id ya existe en la tabla actividades
            self.cursor.execute("PRAGMA table_info(actividades)")
            columns = [column[1] for column in self.cursor.fetchall()]

            if 'capitulo_id' not in columns:
                # Añadir columna capitulo_id a la tabla actividades
                self.cursor.execute('''
                ALTER TABLE actividades ADD COLUMN capitulo_id INTEGER
                REFERENCES capitulos(id)
                ''')

            # Verificar si la columna orden ya existe en la tabla actividades
            if 'orden' not in columns:
                # Añadir columna orden a la tabla actividades
                self.cursor.execute('''
                ALTER TABLE actividades ADD COLUMN orden INTEGER DEFAULT 0
                ''')

            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al crear tabla de capítulos: {e}")
            return False

    def add_capitulo(self, capitulo_data: Dict) -> int:
        """
        Agrega un nuevo capítulo a la base de datos.
        """
        try:
            self.cursor.execute('''
            INSERT INTO capitulos (nombre, descripcion, orden)
            VALUES (?, ?, ?)
            ''', (capitulo_data['nombre'], capitulo_data.get('descripcion', ''),
                  capitulo_data.get('orden', 0)))
            self.conn.commit()
            return self.cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Error al agregar capítulo: {e}")
            return -1

    def update_capitulo(self, capitulo_id: int, capitulo_data: Dict) -> bool:
        """
        Actualiza un capítulo existente.
        """
        try:
            self.cursor.execute('''
            UPDATE capitulos
            SET nombre = ?, descripcion = ?, orden = ?
            WHERE id = ?
            ''', (capitulo_data['nombre'], capitulo_data.get('descripcion', ''),
                  capitulo_data.get('orden', 0), capitulo_id))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al actualizar capítulo: {e}")
            return False

    def delete_capitulo(self, capitulo_id: int) -> bool:
        """
        Elimina un capítulo y actualiza las actividades asociadas.
        """
        try:
            # Desasociar actividades del capítulo
            self.cursor.execute('''
            UPDATE actividades SET capitulo_id = NULL
            WHERE capitulo_id = ?
            ''', (capitulo_id,))

            # Eliminar el capítulo
            self.cursor.execute('DELETE FROM capitulos WHERE id = ?', (capitulo_id,))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al eliminar capítulo: {e}")
            return False

    def get_capitulo_by_id(self, capitulo_id: int) -> Optional[Dict]:
        """
        Obtiene un capítulo por su ID.
        """
        try:
            self.cursor.execute('SELECT * FROM capitulos WHERE id = ?', (capitulo_id,))
            row = self.cursor.fetchone()
            if row:
                columns = [column[0] for column in self.cursor.description]
                return dict(zip(columns, row))
            return None
        except sqlite3.Error as e:
            print(f"Error al obtener capítulo: {e}")
            return None

    def get_all_capitulos(self) -> List[Dict]:
        """
        Obtiene todos los capítulos ordenados por el campo orden.
        """
        try:
            self.cursor.execute('SELECT * FROM capitulos ORDER BY orden')
            columns = [column[0] for column in self.cursor.description]
            return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Error al obtener capítulos: {e}")
            return []

    def get_actividades_by_capitulo(self, capitulo_id: int) -> List[Dict]:
        """
        Obtiene todas las actividades asociadas a un capítulo.
        """
        try:
            self.cursor.execute('''
            SELECT a.*, c.nombre as categoria_nombre 
            FROM actividades a
            LEFT JOIN categorias c ON a.categoria_id = c.id
            WHERE a.capitulo_id = ?
            ORDER BY a.orden
            ''', (capitulo_id,))
            columns = [column[0] for column in self.cursor.description]
            return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Error al obtener actividades por capítulo: {e}")
            return []

    def assign_actividad_to_capitulo(self, actividad_id: int, capitulo_id: Optional[int]) -> bool:
        """
        Asigna una actividad a un capítulo.
        """
        try:
            # Obtener el orden máximo actual en el capítulo destino
            if capitulo_id is not None:
                self.cursor.execute('''
                SELECT MAX(orden) FROM actividades WHERE capitulo_id = ?
                ''', (capitulo_id,))
                max_orden = self.cursor.fetchone()[0]
                nuevo_orden = 1 if max_orden is None else max_orden + 1
            else:
                nuevo_orden = 0

            # Actualizar la actividad
            self.cursor.execute('''
            UPDATE actividades SET capitulo_id = ?, orden = ?
            WHERE id = ?
            ''', (capitulo_id, nuevo_orden, actividad_id))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al asignar actividad a capítulo: {e}")
            return False

    def update_actividad_orden(self, actividad_id: int, nuevo_orden: int) -> bool:
        """
        Actualiza el orden de una actividad dentro de su capítulo.
        """
        try:
            self.cursor.execute('''
            UPDATE actividades SET orden = ?
            WHERE id = ?
            ''', (nuevo_orden, actividad_id))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al actualizar orden de actividad: {e}")
            return False

    def intercambiar_orden_actividades(self, actividad_id1: int, actividad_id2: int) -> bool:
        """
        Intercambia el orden de dos actividades.
        """
        try:
            # Obtener orden actual de ambas actividades
            self.cursor.execute('SELECT orden, capitulo_id FROM actividades WHERE id = ?', (actividad_id1,))
            orden1, capitulo_id1 = self.cursor.fetchone()

            self.cursor.execute('SELECT orden, capitulo_id FROM actividades WHERE id = ?', (actividad_id2,))
            orden2, capitulo_id2 = self.cursor.fetchone()

            # Verificar que ambas actividades pertenezcan al mismo capítulo
            if capitulo_id1 != capitulo_id2:
                return False

            # Intercambiar órdenes
            self.cursor.execute('UPDATE actividades SET orden = ? WHERE id = ?', (orden2, actividad_id1))
            self.cursor.execute('UPDATE actividades SET orden = ? WHERE id = ?', (orden1, actividad_id2))

            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al intercambiar orden de actividades: {e}")
            return False

    def update_capitulo_orden(self, capitulo_id: int, nuevo_orden: int) -> bool:
        """
        Actualiza el orden de un capítulo.
        """
        try:
            self.cursor.execute('''
            UPDATE capitulos SET orden = ?
            WHERE id = ?
            ''', (nuevo_orden, capitulo_id))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al actualizar orden de capítulo: {e}")
            return False

    def intercambiar_orden_capitulos(self, capitulo_id1: int, capitulo_id2: int) -> bool:
        """
        Intercambia el orden de dos capítulos.
        """
        try:
            # Obtener orden actual de ambos capítulos
            self.cursor.execute('SELECT orden FROM capitulos WHERE id = ?', (capitulo_id1,))
            orden1 = self.cursor.fetchone()[0]

            self.cursor.execute('SELECT orden FROM capitulos WHERE id = ?', (capitulo_id2,))
            orden2 = self.cursor.fetchone()[0]

            # Intercambiar órdenes
            self.cursor.execute('UPDATE capitulos SET orden = ? WHERE id = ?', (orden2, capitulo_id1))
            self.cursor.execute('UPDATE capitulos SET orden = ? WHERE id = ?', (orden1, capitulo_id2))

            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al intercambiar orden de capítulos: {e}")
            return False
