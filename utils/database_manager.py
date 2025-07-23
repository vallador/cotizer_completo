# utils/database_manager.py
from typing import List, Dict

import sqlite3





class DatabaseManager:
    def __init__(self, db_path="data/cotizaciones.db"): # Ruta corregida para ser más robusta
        self.db_path = db_path
        self.connection = None
        self.connect()
        self.create_tables()

    def connect(self):
        """Establece la conexión a la base de datos."""
        try:
            # Asegurarse de que el directorio de datos exista
            import os
            os.makedirs(os.path.dirname(self.db_path), exist_ok=True)
            self.connection = sqlite3.connect(self.db_path)
            print("Conexión a la base de datos establecida.")
        except sqlite3.Error as e:
            print(f"Error al conectar a la base de datos: {e}")

    def close(self):
        """Cierra la conexión a la base de datos."""
        if self.connection:
            self.connection.close()
            print("Conexión a la base de datos cerrada.")

    def create_tables(self):
        """Crea las tablas necesarias si no existen."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS clientes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    tipo TEXT NOT NULL,
                    nombre TEXT NOT NULL,
                    direccion TEXT,
                    nit TEXT,
                    telefono TEXT,
                    email TEXT
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS categorias (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL UNIQUE
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS actividades (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    descripcion TEXT NOT NULL,
                    unidad TEXT NOT NULL,
                    valor_unitario REAL NOT NULL,
                    categoria_id INTEGER,
                    FOREIGN KEY (categoria_id) REFERENCES categorias(id)
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS capitulos (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nombre TEXT NOT NULL,
                    descripcion TEXT,
                    orden INTEGER DEFAULT 0
                )
            """)
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS aiu_values (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    administracion REAL,
                    imprevistos REAL,
                    utilidad REAL,
                    iva_sobre_utilidad REAL
                )
            """)
            self.connection.commit()
            print("Tablas verificadas/creadas correctamente.")

            # Insertar valores AIU por defecto si no existen
            cursor.execute("SELECT COUNT(*) FROM aiu_values")
            if cursor.fetchone()[0] == 0:
                cursor.execute(
                    "INSERT INTO aiu_values (administracion, imprevistos, utilidad, iva_sobre_utilidad) VALUES (?, ?, ?, ?)",
                    (10.0, 5.0, 5.0, 19.0))
                self.connection.commit()
                print("Valores AIU por defecto insertados.")

        except sqlite3.Error as e:
            print(f"Error al crear tablas: {e}")

    # Métodos para clientes
    def add_client(self, tipo, nombre, direccion, nit, telefono, email):
        """Agrega un nuevo cliente a la base de datos."""
        try:
            cursor = self.connection.cursor()
            cursor.execute(
                "INSERT INTO clientes (tipo, nombre, direccion, nit, telefono, email) VALUES (?, ?, ?, ?, ?, ?)",
                (tipo, nombre, direccion, nit, telefono, email))
            self.connection.commit()
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Error al agregar cliente: {e}")
            return None

    def get_all_clients(self):
        """Obtiene todos los clientes de la base de datos."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT id, tipo, nombre, direccion, nit, telefono, email FROM clientes")
            clients = []
            for row in cursor.fetchall():
                clients.append({
                    'id': row[0],
                    'tipo': row[1],
                    'nombre': row[2],
                    'direccion': row[3],
                    'nit': row[4],
                    'telefono': row[5],
                    'email': row[6]
                })
            return clients
        except sqlite3.Error as e:
            print(f"Error al obtener clientes: {e}")
            return []

    def get_client_by_id(self, client_id):
        """Obtiene un cliente por su ID."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT id, tipo, nombre, direccion, nit, telefono, email FROM clientes WHERE id = ?",
                           (client_id,))
            row = cursor.fetchone()
            if row:
                return {
                    'id': row[0],
                    'tipo': row[1],
                    'nombre': row[2],
                    'direccion': row[3],
                    'nit': row[4],
                    'telefono': row[5],
                    'email': row[6]
                }
            return None
        except sqlite3.Error as e:
            print(f"Error al obtener cliente por ID: {e}")
            return None

    def update_client(self, client_id, client_data):
        """Actualiza los datos de un cliente existente."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                UPDATE clientes
                SET tipo = ?, nombre = ?, direccion = ?, nit = ?, telefono = ?, email = ?
                WHERE id = ?
            """, (client_data['tipo'], client_data['nombre'], client_data['direccion'],
                  client_data['nit'], client_data['telefono'], client_data['email'], client_id))
            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al actualizar cliente: {e}")
            return False
    def get_related_activities(self, activity_id):
        """Obtiene las actividades relacionadas con una actividad por su ID."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                        SELECT
                            a.id, a.descripcion, a.unidad, a.valor_unitario, a.categoria_id, c.nombre as categoria_nombre
                        FROM
                            actividades a
                        LEFT JOIN
                            categorias c ON a.categoria_id = c.id
                        WHERE
                            a.id IN (
                                SELECT actividad_relacionada_id 
                                FROM actividad_relacionada 
                                WHERE actividad_principal_id = ?
                            )
                    """, (activity_id,))
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
            print(f"Error al obtener actividades relacionadas: {str(e)}")
            return []


    def delete_client(self, client_id):
        """Elimina un cliente de la base de datos."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM clientes WHERE id = ?", (client_id,))
            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al eliminar cliente: {e}")
            return False

    # Métodos para categorías
    def add_category(self, nombre):
        """Agrega una nueva categoría."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("INSERT INTO categorias (nombre) VALUES (?) ", (nombre,))
            self.connection.commit()
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Error al agregar categoría: {e}")
            return None

    def get_all_categories(self):
        """Obtiene todas las categorías."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT id, nombre FROM categorias")
            categories = []
            for row in cursor.fetchall():
                categories.append({
                    'id': row[0],
                    'nombre': row[1]
                })
            return categories
        except sqlite3.Error as e:
            print(f"Error al obtener categorías: {e}")
            return []

    def update_category(self, category_id, new_name):
        """Actualiza el nombre de una categoría."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("UPDATE categorias SET nombre = ? WHERE id = ?", (new_name, category_id))
            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al actualizar categoría: {e}")
            return False

    def delete_category(self, category_id):
        """Elimina una categoría."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM categorias WHERE id = ?", (category_id,))
            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al eliminar categoría: {e}")
            return False

    # Métodos para actividades
    def add_activity(self, descripcion, unidad, valor_unitario, categoria_id=None):
        """Agrega una nueva actividad."""
        try:
            cursor = self.connection.cursor()
            cursor.execute(
                "INSERT INTO actividades (descripcion, unidad, valor_unitario, categoria_id) VALUES (?, ?, ?, ?)",
                (descripcion, unidad, valor_unitario, categoria_id))
            self.connection.commit()
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Error al agregar actividad: {e}")
            return None

    def get_all_activities(self):
        """Obtiene todas las actividades con su categoría."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                SELECT
                    a.id, a.descripcion, a.unidad, a.valor_unitario, a.categoria_id, c.nombre as categoria_nombre
                FROM
                    actividades a
                LEFT JOIN
                    categorias c ON a.categoria_id = c.id
            """)
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
        except sqlite3.Error as e:
            print(f"Error al obtener actividades: {e}")
            return []

    def get_activity_by_id(self, activity_id):
        """Obtiene una actividad por su ID con su categoría."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                SELECT
                    a.id, a.descripcion, a.unidad, a.valor_unitario, a.categoria_id, c.nombre as categoria_nombre
                FROM
                    actividades a
                LEFT JOIN
                    categorias c ON a.categoria_id = c.id
                WHERE
                    a.id = ?
            """, (activity_id,))
            row = cursor.fetchone()
            if row:
                return {
                    'id': row[0],
                    'descripcion': row[1],
                    'unidad': row[2],
                    'valor_unitario': row[3],
                    'categoria_id': row[4],
                    'categoria_nombre': row[5]
                }
            return None
        except sqlite3.Error as e:
            print(f"Error al obtener actividad por ID: {e}")
            return None

    def update_activity(self, activity_id, descripcion, unidad, valor_unitario, categoria_id=None):
        """Actualiza una actividad existente."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                UPDATE actividades
                SET descripcion = ?, unidad = ?, valor_unitario = ?, categoria_id = ?
                WHERE id = ?
            """, (descripcion, unidad, valor_unitario, categoria_id, activity_id))
            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al actualizar actividad: {e}")
            return False

    def delete_activity(self, activity_id):
        """Elimina una actividad."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM actividades WHERE id = ?", (activity_id,))
            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al eliminar actividad: {e}")
            return False

    # Métodos para AIU
    def get_aiu_values(self):
        """Obtiene los valores de AIU."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT administracion, imprevistos, utilidad, iva_sobre_utilidad FROM aiu_values LIMIT 1")
            row = cursor.fetchone()
            if row:
                return {
                    'administracion': row[0],
                    'imprevistos': row[1],
                    'utilidad': row[2],
                    'iva_sobre_utilidad': row[3]
                }
            return None
        except sqlite3.Error as e:
            print(f"Error al obtener valores AIU: {e}")
            return None

    def update_aiu_values(self, administracion, imprevistos, utilidad, iva_sobre_utilidad):
        """Actualiza los valores de AIU."""
        try:
            cursor = self.connection.cursor()
            cursor.execute(
                "UPDATE aiu_values SET administracion = ?, imprevistos = ?, utilidad = ?, iva_sobre_utilidad = ? WHERE id = 1",
                (administracion, imprevistos, utilidad, iva_sobre_utilidad))
            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al actualizar valores AIU: {e}")
            return False

    # Métodos para capítulos (agregados recientemente)
    def get_all_chapters(self):
        """Obtiene todos los capítulos de la base de datos."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT id, nombre FROM capitulos ORDER BY orden ASC")
            chapters = []
            for row in cursor.fetchall():
                chapters.append({
                    'id': row[0],
                    'nombre': row[1]
                })
            return chapters
        except Exception as e:
            print(f"Error al obtener todos los capítulos: {str(e)}")
            return []

    def get_chapter_by_id(self, chapter_id):
        """Obtiene un capítulo por su ID."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT id, nombre, descripcion, orden FROM capitulos WHERE id = ?", (chapter_id,))
            row = cursor.fetchone()

            if row:
                return {
                    'id': row[0],
                    'nombre': row[1],
                    'descripcion': row[2],
                    'orden': row[3]
                }
            return None
        except Exception as e:
            print(f"Error al obtener capítulo por ID: {str(e)}")
            return None

    def add_chapter(self, nombre, descripcion):
        sql = "INSERT INTO capitulos (nombre, descripcion) VALUES (?, ?)"
        try:
            cursor = self.connection.cursor()
            cursor.execute(sql, (nombre, descripcion))
            self.connection.commit()
            last_id = cursor.lastrowid
            print(f"DEBUG (DB): Capítulo '{nombre}' agregado con ID: {last_id}")
            return last_id
        except Exception as e:
            print(f"ERROR (DB): No se pudo agregar el capítulo. Error: {e}")
            return None

    def update_chapter(self, chapter_id, nombre, descripcion):
        """Actualiza el nombre y la descripción de un capítulo existente."""
        sql = "UPDATE capitulos SET nombre = ?, descripcion = ? WHERE id = ?"
        try:
            cursor = self.connection.cursor()
            cursor.execute(sql, (nombre, descripcion, chapter_id))
            self.connection.commit()
            print(f"DEBUG (DB): Capítulo ID {chapter_id} actualizado.")
            return True
        except sqlite3.Error as e:
            print(f"ERROR (DB): No se pudo actualizar el capítulo {chapter_id}. Error: {e}")
            return False

    def delete_chapter(self, chapter_id):
        """Elimina un capítulo de la base de datos usando su ID."""
        sql = "DELETE FROM capitulos WHERE id = ?"
        try:
            cursor = self.connection.cursor()
            cursor.execute(sql, (chapter_id,))
            self.connection.commit()
            print(f"DEBUG (DB): Capítulo con ID {chapter_id} eliminado.")
            return True
        except Exception as e:
            print(f"ERROR (DB): No se pudo eliminar el capítulo {chapter_id}. Error: {e}")
            return False

    def add_product(self, product_data):
        """Agrega un nuevo producto a la base de datos."""
        try:
            cursor = self.connection.cursor()
            cursor.execute(
                "INSERT INTO productos (nombre, descripcion, unidad, precio_unitario, categoria_id) VALUES (?, ?, ?, ?, ?)",
                (product_data['nombre'], product_data['descripcion'], product_data['unidad'],
                 product_data['precio_unitario'], product_data.get('categoria_id'))
            )
            self.connection.commit()
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Error al agregar producto: {e}")
            return None

    def get_all_products(self):
        """Obtiene todos los productos de la base de datos."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT id, nombre, unidad, precio_unitario FROM productos")
            return [{'id': r[0], 'nombre': r[1], 'unidad': r[2], 'precio_unitario': r[3]} for r in cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Error al obtener productos: {e}")
            return []

    def get_product_by_id(self, product_id):
        """Obtiene un producto por su ID."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("SELECT * FROM productos WHERE id = ?", (product_id,))
            row = cursor.fetchone()
            # Asumiendo que las columnas son id, nombre, descripcion, unidad, precio_unitario, categoria_id
            if row:
                return {'id': row[0], 'nombre': row[1], 'descripcion': row[2], 'unidad': row[3],
                        'precio_unitario': row[4], 'categoria_id': row[5]}
            return None
        except sqlite3.Error as e:
            print(f"Error al obtener producto por ID: {e}")
            return None

    def update_product(self, product_id, product_data):
        """Actualiza un producto existente."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                UPDATE productos SET nombre = ?, descripcion = ?, unidad = ?, precio_unitario = ?, categoria_id = ?
                WHERE id = ?
            """, (product_data['nombre'], product_data['descripcion'], product_data['unidad'],
                  product_data['precio_unitario'], product_data.get('categoria_id'), product_id))
            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al actualizar producto: {e}")
            return False

    def delete_product(self, product_id):
        """Elimina un producto."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM productos WHERE id = ?", (product_id,))
            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al eliminar producto: {e}")
            return False

    # --- MÉTODOS PARA RELACIONES ---

    def add_activity_product(self, activity_id, product_id, quantity):
        """Crea una relación entre una actividad y un producto."""
        try:
            cursor = self.connection.cursor()
            cursor.execute(
                "INSERT INTO actividad_producto (actividad_id, producto_id, cantidad) VALUES (?, ?, ?)",
                (activity_id, product_id, quantity)
            )
            self.connection.commit()
            return cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Error al agregar relación actividad-producto: {e}")
            return None

    def delete_activity_product(self, relation_id):
        """Elimina una relación actividad-producto por su ID."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("DELETE FROM actividad_producto WHERE id = ?", (relation_id,))
            self.connection.commit()
            return True
        except sqlite3.Error as e:
            print(f"Error al eliminar relación actividad-producto: {e}")
            return False

    def get_products_by_activity(self, activity_id):
        """Obtiene los productos y sus cantidades para una actividad específica."""
        try:
            cursor = self.connection.cursor()
            cursor.execute("""
                SELECT p.id, p.nombre, p.unidad, p.precio_unitario, ap.cantidad, ap.id as relation_id
                FROM productos p
                JOIN actividad_producto ap ON p.id = ap.producto_id
                WHERE ap.actividad_id = ?
            """, (activity_id,))
            return [{'id': r[0], 'nombre': r[1], 'unidad': r[2], 'precio_unitario': r[3],
                     'cantidad': r[4], 'relation_id': r[5]} for r in cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"Error al obtener productos por actividad: {e}")
            return []