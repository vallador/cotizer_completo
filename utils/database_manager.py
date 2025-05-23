# utils/database_manager.py

import sqlite3
from typing import List, Dict

class DatabaseManager:
    def __init__(self, db_file: str):
        self.db_file = db_file
        self.conn = None
        self.cursor = None


    def connect(self):
        # Conectar a la base de datos
        self.conn = sqlite3.connect(self.db_file)
        # Crear el cursor para ejecutar comandos SQL
        self.cursor = self.conn.cursor()
        # Crear tablas si no existen
        self.create_tables()

    def disconnect(self):
        # Cerrar el cursor si existe
        if self.cursor:
            self.cursor.close()
        # Cerrar la conexión si existe
        if self.conn:
            self.conn.close()

    def create_tables(self):
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS clientes (
            id INTEGER PRIMARY KEY,
            nombre TEXT NOT NULL,
            tipo TEXT NOT NULL,
            direccion TEXT,
            nit TEXT,
            telefono TEXT,
            email TEXT
        )
        ''')
        
        # Crear tabla de categorías
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS categorias (
            id INTEGER PRIMARY KEY,
            nombre TEXT NOT NULL,
            descripcion TEXT
        )
        ''')
        
        # Crear tabla de actividades con referencia a categoría
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS actividades (
            id INTEGER PRIMARY KEY,
            descripcion TEXT NOT NULL,
            unidad TEXT NOT NULL,
            valor_unitario REAL NOT NULL,
            categoria_id INTEGER,
            FOREIGN KEY (categoria_id) REFERENCES categorias (id)
        )
        ''')
        
        # Crear tabla de productos
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS productos (
            id INTEGER PRIMARY KEY,
            nombre TEXT NOT NULL,
            descripcion TEXT,
            unidad TEXT NOT NULL,
            precio_unitario REAL NOT NULL,
            categoria_id INTEGER,
            FOREIGN KEY (categoria_id) REFERENCES categorias (id)
        )
        ''')
        
        # Crear tabla de relación entre actividades y productos
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS actividad_producto (
            id INTEGER PRIMARY KEY,
            actividad_id INTEGER NOT NULL,
            producto_id INTEGER NOT NULL,
            cantidad REAL NOT NULL,
            FOREIGN KEY (actividad_id) REFERENCES actividades (id),
            FOREIGN KEY (producto_id) REFERENCES productos (id)
        )
        ''')
        
        # Crear tabla de relaciones entre actividades
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS actividad_relacionada (
            id INTEGER PRIMARY KEY,
            actividad_principal_id INTEGER NOT NULL,
            actividad_relacionada_id INTEGER NOT NULL,
            FOREIGN KEY (actividad_principal_id) REFERENCES actividades (id),
            FOREIGN KEY (actividad_relacionada_id) REFERENCES actividades (id)
        )
        ''')
        
        # Crear tabla de cotizaciones
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS cotizaciones (
            id INTEGER PRIMARY KEY,
            numero TEXT NOT NULL,
            fecha TEXT NOT NULL,
            cliente_id INTEGER NOT NULL,
            subtotal REAL NOT NULL,
            iva REAL NOT NULL,
            total REAL NOT NULL,
            administracion REAL,
            imprevistos REAL,
            utilidad REAL,
            iva_utilidad REAL,
            FOREIGN KEY (cliente_id) REFERENCES clientes (id)
        )
        ''')
        
        # Crear tabla de detalles de cotización (actividades incluidas)
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS cotizacion_detalles (
            id INTEGER PRIMARY KEY,
            cotizacion_id INTEGER NOT NULL,
            actividad_id INTEGER NOT NULL,
            cantidad REAL NOT NULL,
            valor_unitario REAL NOT NULL,
            total REAL NOT NULL,
            FOREIGN KEY (cotizacion_id) REFERENCES cotizaciones (id),
            FOREIGN KEY (actividad_id) REFERENCES actividades (id)
        )
        ''')
        
        self.conn.commit()

    def add_client(self, client_data: Dict):
        self.cursor.execute('''
        INSERT INTO clientes (nombre, tipo, direccion, nit, telefono, email)
        VALUES (?, ?, ?, ?, ?, ?)
        ''', (client_data['nombre'], client_data['tipo'], client_data['direccion'],
              client_data.get('nit'), client_data.get('telefono'), client_data.get('email')))
        self.conn.commit()
        return self.cursor.lastrowid

    def get_all_clients(self) -> List[Dict]:
        self.cursor.execute('SELECT * FROM clientes')
        columns = [column[0] for column in self.cursor.description]
        return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
        
    def get_all_activities(self) -> List[Dict]:
        self.cursor.execute('''
        SELECT a.*, c.nombre as categoria_nombre 
        FROM actividades a
        LEFT JOIN categorias c ON a.categoria_id = c.id
        ''')
        columns = [column[0] for column in self.cursor.description]
        return [dict(zip(columns, row)) for row in self.cursor.fetchall()]

    def get_client_by_id(self, client_id: int) -> Dict:
        self.cursor.execute('SELECT * FROM clientes WHERE id = ?', (client_id,))
        columns = [column[0] for column in self.cursor.description]
        row = self.cursor.fetchone()
        return dict(zip(columns, row)) if row else None

    def update_client(self, client_id: int, client_data: Dict):
        self.cursor.execute('''
        UPDATE clientes
        SET nombre = ?, tipo = ?, direccion = ?, nit = ?, telefono = ?, email = ?
        WHERE id = ?
        ''', (client_data['nombre'], client_data['tipo'], client_data['direccion'],
              client_data.get('nit'), client_data.get('telefono'), client_data.get('email'),
              client_id))
        self.conn.commit()

    def delete_client(self, client_id: int):
        self.cursor.execute('DELETE FROM clientes WHERE id = ?', (client_id,))
        self.conn.commit()
        
    # Métodos para actividades
    def add_activity(self, activity_data: Dict):
        self.cursor.execute('''
        INSERT INTO actividades (descripcion, unidad, valor_unitario, categoria_id)
        VALUES (?, ?, ?, ?)
        ''', (activity_data['descripcion'], activity_data['unidad'], 
              activity_data['valor_unitario'], activity_data.get('categoria_id')))
        self.conn.commit()
        return self.cursor.lastrowid
        
    def update_activity(self, activity_id: int, activity_data: Dict):
        self.cursor.execute('''
        UPDATE actividades
        SET descripcion = ?, unidad = ?, valor_unitario = ?, categoria_id = ?
        WHERE id = ?
        ''', (activity_data['descripcion'], activity_data['unidad'], 
              activity_data['valor_unitario'], activity_data.get('categoria_id'), activity_id))
        self.conn.commit()
        
    def delete_activity(self, activity_id: int):
        # Primero eliminar relaciones
        self.cursor.execute('DELETE FROM actividad_producto WHERE actividad_id = ?', (activity_id,))
        self.cursor.execute('DELETE FROM actividad_relacionada WHERE actividad_principal_id = ? OR actividad_relacionada_id = ?', 
                           (activity_id, activity_id))
        # Luego eliminar la actividad
        self.cursor.execute('DELETE FROM actividades WHERE id = ?', (activity_id,))
        self.conn.commit()
        
    def get_activity_by_id(self, activity_id: int) -> Dict:
        self.cursor.execute('''
        SELECT a.*, c.nombre as categoria_nombre 
        FROM actividades a
        LEFT JOIN categorias c ON a.categoria_id = c.id
        WHERE a.id = ?
        ''', (activity_id,))
        columns = [column[0] for column in self.cursor.description]
        row = self.cursor.fetchone()
        return dict(zip(columns, row)) if row else None
    
    # Métodos para productos
    def add_product(self, product_data: Dict):
        self.cursor.execute('''
        INSERT INTO productos (nombre, descripcion, unidad, precio_unitario, categoria_id)
        VALUES (?, ?, ?, ?, ?)
        ''', (product_data['nombre'], product_data.get('descripcion'), 
              product_data['unidad'], product_data['precio_unitario'], 
              product_data.get('categoria_id')))
        self.conn.commit()
        return self.cursor.lastrowid
        
    def update_product(self, product_id: int, product_data: Dict):
        self.cursor.execute('''
        UPDATE productos
        SET nombre = ?, descripcion = ?, unidad = ?, precio_unitario = ?, categoria_id = ?
        WHERE id = ?
        ''', (product_data['nombre'], product_data.get('descripcion'), 
              product_data['unidad'], product_data['precio_unitario'], 
              product_data.get('categoria_id'), product_id))
        self.conn.commit()
        
    def delete_product(self, product_id: int):
        # Primero eliminar relaciones
        self.cursor.execute('DELETE FROM actividad_producto WHERE producto_id = ?', (product_id,))
        # Luego eliminar el producto
        self.cursor.execute('DELETE FROM productos WHERE id = ?', (product_id,))
        self.conn.commit()
        
    def get_product_by_id(self, product_id: int) -> Dict:
        self.cursor.execute('''
        SELECT p.*, c.nombre as categoria_nombre 
        FROM productos p
        LEFT JOIN categorias c ON p.categoria_id = c.id
        WHERE p.id = ?
        ''', (product_id,))
        columns = [column[0] for column in self.cursor.description]
        row = self.cursor.fetchone()
        return dict(zip(columns, row)) if row else None
        
    def get_all_products(self) -> List[Dict]:
        self.cursor.execute('''
        SELECT p.*, c.nombre as categoria_nombre 
        FROM productos p
        LEFT JOIN categorias c ON p.categoria_id = c.id
        ''')
        columns = [column[0] for column in self.cursor.description]
        return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
    
    # Métodos para categorías
    def add_category(self, category_data: Dict):
        self.cursor.execute('''
        INSERT INTO categorias (nombre, descripcion)
        VALUES (?, ?)
        ''', (category_data['nombre'], category_data.get('descripcion')))
        self.conn.commit()
        return self.cursor.lastrowid
        
    def update_category(self, category_id: int, category_data: Dict):
        self.cursor.execute('''
        UPDATE categorias
        SET nombre = ?, descripcion = ?
        WHERE id = ?
        ''', (category_data['nombre'], category_data.get('descripcion'), category_id))
        self.conn.commit()
        
    def delete_category(self, category_id: int):
        self.cursor.execute('DELETE FROM categorias WHERE id = ?', (category_id,))
        self.conn.commit()
        
    def get_category_by_id(self, category_id: int) -> Dict:
        self.cursor.execute('SELECT * FROM categorias WHERE id = ?', (category_id,))
        columns = [column[0] for column in self.cursor.description]
        row = self.cursor.fetchone()
        return dict(zip(columns, row)) if row else None
        
    def get_all_categories(self) -> List[Dict]:
        self.cursor.execute('SELECT * FROM categorias')
        columns = [column[0] for column in self.cursor.description]
        return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
    
    # Métodos para relaciones entre actividades y productos
    def add_activity_product(self, activity_id: int, product_id: int, cantidad: float):
        self.cursor.execute('''
        INSERT INTO actividad_producto (actividad_id, producto_id, cantidad)
        VALUES (?, ?, ?)
        ''', (activity_id, product_id, cantidad))
        self.conn.commit()
        return self.cursor.lastrowid
        
    def update_activity_product(self, relation_id: int, cantidad: float):
        self.cursor.execute('''
        UPDATE actividad_producto
        SET cantidad = ?
        WHERE id = ?
        ''', (cantidad, relation_id))
        self.conn.commit()
        
    def delete_activity_product(self, relation_id: int):
        self.cursor.execute('DELETE FROM actividad_producto WHERE id = ?', (relation_id,))
        self.conn.commit()
        
    def get_products_by_activity(self, activity_id: int) -> List[Dict]:
        self.cursor.execute('''
        SELECT p.*, ap.cantidad, ap.id as relation_id
        FROM productos p
        JOIN actividad_producto ap ON p.id = ap.producto_id
        WHERE ap.actividad_id = ?
        ''', (activity_id,))
        columns = [column[0] for column in self.cursor.description]
        return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
    
    # Métodos para relaciones entre actividades
    def add_related_activity(self, main_activity_id: int, related_activity_id: int):
        self.cursor.execute('''
        INSERT INTO actividad_relacionada (actividad_principal_id, actividad_relacionada_id)
        VALUES (?, ?)
        ''', (main_activity_id, related_activity_id))
        self.conn.commit()
        return self.cursor.lastrowid
        
    def delete_related_activity(self, relation_id: int):
        self.cursor.execute('DELETE FROM actividad_relacionada WHERE id = ?', (relation_id,))
        self.conn.commit()
        
    def get_related_activities(self, activity_id: int) -> List[Dict]:
        self.cursor.execute('''
        SELECT a.*, ar.id as relation_id
        FROM actividades a
        JOIN actividad_relacionada ar ON a.id = ar.actividad_relacionada_id
        WHERE ar.actividad_principal_id = ?
        ''', (activity_id,))
        columns = [column[0] for column in self.cursor.description]
        return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
    
    # Métodos para búsqueda y filtrado
    def search_activities_by_text(self, search_text: str) -> List[Dict]:
        search_pattern = f"%{search_text}%"
        self.cursor.execute('''
        SELECT a.*, c.nombre as categoria_nombre 
        FROM actividades a
        LEFT JOIN categorias c ON a.categoria_id = c.id
        WHERE a.descripcion LIKE ?
        ''', (search_pattern,))
        columns = [column[0] for column in self.cursor.description]
        return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
        
    def search_products_by_text(self, search_text: str) -> List[Dict]:
        search_pattern = f"%{search_text}%"
        self.cursor.execute('''
        SELECT p.*, c.nombre as categoria_nombre 
        FROM productos p
        LEFT JOIN categorias c ON p.categoria_id = c.id
        WHERE p.nombre LIKE ? OR p.descripcion LIKE ?
        ''', (search_pattern, search_pattern))
        columns = [column[0] for column in self.cursor.description]
        return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
        
    def get_activities_by_category(self, category_id: int) -> List[Dict]:
        self.cursor.execute('''
        SELECT a.*, c.nombre as categoria_nombre 
        FROM actividades a
        LEFT JOIN categorias c ON a.categoria_id = c.id
        WHERE a.categoria_id = ?
        ''', (category_id,))
        columns = [column[0] for column in self.cursor.description]
        return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
        
    # Métodos para cotizaciones
    def add_quotation(self, quotation_data: Dict, activities: List[Dict]):
        # Insertar la cotización
        self.cursor.execute('''
        INSERT INTO cotizaciones (numero, fecha, cliente_id, subtotal, iva, total, 
                                 administracion, imprevistos, utilidad, iva_utilidad)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (quotation_data['numero'], quotation_data['fecha'], quotation_data['cliente_id'],
              quotation_data['subtotal'], quotation_data['iva'], quotation_data['total'],
              quotation_data.get('administracion'), quotation_data.get('imprevistos'),
              quotation_data.get('utilidad'), quotation_data.get('iva_utilidad')))
        
        quotation_id = self.cursor.lastrowid
        
        # Insertar los detalles de la cotización (actividades)
        for activity in activities:
            self.cursor.execute('''
            INSERT INTO cotizacion_detalles (cotizacion_id, actividad_id, cantidad, valor_unitario, total)
            VALUES (?, ?, ?, ?, ?)
            ''', (quotation_id, activity['actividad_id'], activity['cantidad'],
                  activity['valor_unitario'], activity['total']))
        
        self.conn.commit()
        return quotation_id
        
    def get_all_quotations(self) -> List[Dict]:
        self.cursor.execute('''
        SELECT c.*, cl.nombre as cliente_nombre 
        FROM cotizaciones c
        JOIN clientes cl ON c.cliente_id = cl.id
        ''')
        columns = [column[0] for column in self.cursor.description]
        return [dict(zip(columns, row)) for row in self.cursor.fetchall()]
        
    def get_quotation_by_id(self, quotation_id: int) -> Dict:
        self.cursor.execute('''
        SELECT c.*, cl.nombre as cliente_nombre 
        FROM cotizaciones c
        JOIN clientes cl ON c.cliente_id = cl.id
        WHERE c.id = ?
        ''', (quotation_id,))
        columns = [column[0] for column in self.cursor.description]
        row = self.cursor.fetchone()
        quotation = dict(zip(columns, row)) if row else None
        
        if quotation:
            # Obtener los detalles de la cotización
            self.cursor.execute('''
            SELECT cd.*, a.descripcion, a.unidad
            FROM cotizacion_detalles cd
            JOIN actividades a ON cd.actividad_id = a.id
            WHERE cd.cotizacion_id = ?
            ''', (quotation_id,))
            columns = [column[0] for column in self.cursor.description]
            details = [dict(zip(columns, row)) for row in self.cursor.fetchall()]
            quotation['detalles'] = details
            
        return quotation
        
    def delete_quotation(self, quotation_id: int):
        # Eliminar primero los detalles
        self.cursor.execute('DELETE FROM cotizacion_detalles WHERE cotizacion_id = ?', (quotation_id,))
        # Luego eliminar la cotización
        self.cursor.execute('DELETE FROM cotizaciones WHERE id = ?', (quotation_id,))
        self.conn.commit()
