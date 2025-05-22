import sqlite3
import os

def crear_tablas_iniciales():
    """Crea las tablas iniciales y agrega datos de ejemplo si no existen"""
    # Asegurarse de que el directorio data existe
    os.makedirs('data', exist_ok=True)
    
    # Conectar a la base de datos
    conn = sqlite3.connect('data/cotizaciones.db')
    cursor = conn.cursor()
    
    # Crear tablas
    cursor.execute('''
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
    
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS actividades (
        id INTEGER PRIMARY KEY,
        descripcion TEXT NOT NULL,
        unidad TEXT NOT NULL,
        valor_unitario REAL NOT NULL
    )
    ''')
    
    cursor.execute('''
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
    
    cursor.execute('''
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
    
    # Verificar si ya hay datos en las tablas
    cursor.execute("SELECT COUNT(*) FROM clientes")
    clientes_count = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(*) FROM actividades")
    actividades_count = cursor.fetchone()[0]
    
    # Agregar datos de ejemplo si no hay registros
    if clientes_count == 0:
        # Insertar clientes de ejemplo
        clientes = [
            ('Empresa ABC S.A.', 'jurídica', 'Calle Principal 123', '900123456-7', '3001234567', 'contacto@empresaabc.com'),
            ('Juan Pérez', 'natural', 'Avenida Central 45', None, '3109876543', 'juan.perez@email.com'),
            ('Constructora XYZ Ltda.', 'jurídica', 'Carrera 67 #89-12', '800987654-3', '3157894561', 'info@constructoraxyz.com')
        ]
        
        cursor.executemany('''
        INSERT INTO clientes (nombre, tipo, direccion, nit, telefono, email)
        VALUES (?, ?, ?, ?, ?, ?)
        ''', clientes)
    
    if actividades_count == 0:
        # Insertar actividades de ejemplo
        actividades = [
            ('Excavación manual', 'm³', 45000),
            ('Suministro e instalación de tubería PVC 4"', 'ml', 35000),
            ('Pintura vinilo sobre muros', 'm²', 12000),
            ('Instalación de piso cerámico', 'm²', 65000),
            ('Construcción de muro en ladrillo', 'm²', 85000),
            ('Instalación de ventana en aluminio', 'und', 250000),
            ('Instalación de puerta metálica', 'und', 350000),
            ('Suministro e instalación de punto eléctrico', 'und', 75000)
        ]
        
        cursor.executemany('''
        INSERT INTO actividades (descripcion, unidad, valor_unitario)
        VALUES (?, ?, ?)
        ''', actividades)
    
    # Guardar cambios y cerrar conexión
    conn.commit()
    conn.close()
    
    print("Base de datos inicializada correctamente.")

if __name__ == "__main__":
    crear_tablas_iniciales()
