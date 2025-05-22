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
        
        # Crear tabla de actividades
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS actividades (
            id INTEGER PRIMARY KEY,
            descripcion TEXT NOT NULL,
            unidad TEXT NOT NULL,
            valor_unitario REAL NOT NULL
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
        self.cursor.execute('SELECT * FROM actividades')
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
        INSERT INTO actividades (descripcion, unidad, valor_unitario)
        VALUES (?, ?, ?)
        ''', (activity_data['descripcion'], activity_data['unidad'], activity_data['valor_unitario']))
        self.conn.commit()
        return self.cursor.lastrowid
        
    def update_activity(self, activity_id: int, activity_data: Dict):
        self.cursor.execute('''
        UPDATE actividades
        SET descripcion = ?, unidad = ?, valor_unitario = ?
        WHERE id = ?
        ''', (activity_data['descripcion'], activity_data['unidad'], 
              activity_data['valor_unitario'], activity_id))
        self.conn.commit()
        
    def delete_activity(self, activity_id: int):
        self.cursor.execute('DELETE FROM actividades WHERE id = ?', (activity_id,))
        self.conn.commit()
        
    def get_activity_by_id(self, activity_id: int) -> Dict:
        self.cursor.execute('SELECT * FROM actividades WHERE id = ?', (activity_id,))
        columns = [column[0] for column in self.cursor.description]
        row = self.cursor.fetchone()
        return dict(zip(columns, row)) if row else None
        
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
