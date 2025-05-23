import sqlite3
import os

def init_db():
    """Inicializa la base de datos con tablas y datos de ejemplo"""
    # Verificar si existe el directorio data
    if not os.path.exists('data'):
        os.makedirs('data')
    
    # Conectar a la base de datos (la crea si no existe)
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
    
    # Crear tabla de categorías
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS categorias (
        id INTEGER PRIMARY KEY,
        nombre TEXT NOT NULL,
        descripcion TEXT
    )
    ''')
    
    # Crear tabla de actividades con referencia a categoría
    cursor.execute('''
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
    cursor.execute('''
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
    cursor.execute('''
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
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS actividad_relacionada (
        id INTEGER PRIMARY KEY,
        actividad_principal_id INTEGER NOT NULL,
        actividad_relacionada_id INTEGER NOT NULL,
        FOREIGN KEY (actividad_principal_id) REFERENCES actividades (id),
        FOREIGN KEY (actividad_relacionada_id) REFERENCES actividades (id)
    )
    ''')
    
    # Crear tabla de cotizaciones
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
    
    # Crear tabla de detalles de cotización (actividades incluidas)
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
    
    # Insertar datos de ejemplo para categorías
    categorias = [
        ('Pintura', 'Trabajos de pintura y acabados'),
        ('Plomería', 'Instalaciones y reparaciones de plomería'),
        ('Electricidad', 'Instalaciones y reparaciones eléctricas'),
        ('Albañilería', 'Trabajos de construcción y reparación'),
        ('Carpintería', 'Trabajos en madera y muebles')
    ]
    
    for categoria in categorias:
        cursor.execute('INSERT OR IGNORE INTO categorias (nombre, descripcion) VALUES (?, ?)', categoria)
    
    # Obtener IDs de categorías
    cursor.execute('SELECT id, nombre FROM categorias')
    categorias_dict = {nombre: id for id, nombre in cursor.fetchall()}
    
    # Insertar datos de ejemplo para actividades
    actividades = [
        ('Pintura interior', 'm²', 15000, categorias_dict['Pintura']),
        ('Pintura exterior', 'm²', 18000, categorias_dict['Pintura']),
        ('Instalación de tubería PVC', 'm', 25000, categorias_dict['Plomería']),
        ('Instalación de sanitario', 'und', 120000, categorias_dict['Plomería']),
        ('Instalación de lavamanos', 'und', 90000, categorias_dict['Plomería']),
        ('Instalación de cableado eléctrico', 'm', 12000, categorias_dict['Electricidad']),
        ('Instalación de tomacorrientes', 'und', 35000, categorias_dict['Electricidad']),
        ('Instalación de interruptores', 'und', 30000, categorias_dict['Electricidad']),
        ('Construcción de muro en ladrillo', 'm²', 85000, categorias_dict['Albañilería']),
        ('Instalación de piso cerámico', 'm²', 65000, categorias_dict['Albañilería']),
        ('Fabricación de mueble de cocina', 'm', 350000, categorias_dict['Carpintería']),
        ('Instalación de puerta', 'und', 180000, categorias_dict['Carpintería'])
    ]
    
    for actividad in actividades:
        cursor.execute('INSERT OR IGNORE INTO actividades (descripcion, unidad, valor_unitario, categoria_id) VALUES (?, ?, ?, ?)', actividad)
    
    # Obtener IDs de actividades
    cursor.execute('SELECT id, descripcion FROM actividades')
    actividades_dict = {descripcion: id for id, descripcion in cursor.fetchall()}
    
    # Insertar datos de ejemplo para productos
    productos = [
        ('Pintura blanca', 'Pintura vinilo tipo 1', 'galón', 120000, categorias_dict['Pintura']),
        ('Pintura de colores', 'Pintura vinilo tipo 1', 'galón', 135000, categorias_dict['Pintura']),
        ('Rodillo', 'Rodillo para pintura', 'und', 15000, categorias_dict['Pintura']),
        ('Brocha', 'Brocha de 3 pulgadas', 'und', 8000, categorias_dict['Pintura']),
        ('Tubo PVC 1/2"', 'Tubo PVC para agua potable', 'm', 5000, categorias_dict['Plomería']),
        ('Codo PVC 1/2"', 'Codo PVC 90 grados', 'und', 1200, categorias_dict['Plomería']),
        ('Sanitario', 'Sanitario completo con tanque', 'und', 250000, categorias_dict['Plomería']),
        ('Lavamanos', 'Lavamanos de pedestal', 'und', 150000, categorias_dict['Plomería']),
        ('Cable #12', 'Cable eléctrico calibre 12', 'm', 3500, categorias_dict['Electricidad']),
        ('Tomacorriente', 'Tomacorriente doble', 'und', 12000, categorias_dict['Electricidad']),
        ('Interruptor', 'Interruptor sencillo', 'und', 10000, categorias_dict['Electricidad']),
        ('Ladrillo', 'Ladrillo farol', 'und', 1200, categorias_dict['Albañilería']),
        ('Cemento', 'Cemento gris', 'kg', 800, categorias_dict['Albañilería']),
        ('Arena', 'Arena de río', 'm³', 120000, categorias_dict['Albañilería']),
        ('Cerámica', 'Cerámica para piso', 'm²', 35000, categorias_dict['Albañilería']),
        ('Madera', 'Madera pino', 'm²', 45000, categorias_dict['Carpintería']),
        ('Tornillos', 'Tornillos para madera', 'und', 200, categorias_dict['Carpintería']),
        ('Bisagras', 'Bisagras para puerta', 'par', 8000, categorias_dict['Carpintería']),
        ('Puerta', 'Puerta de madera', 'und', 120000, categorias_dict['Carpintería'])
    ]
    
    for producto in productos:
        cursor.execute('INSERT OR IGNORE INTO productos (nombre, descripcion, unidad, precio_unitario, categoria_id) VALUES (?, ?, ?, ?, ?)', producto)
    
    # Obtener IDs de productos
    cursor.execute('SELECT id, nombre FROM productos')
    productos_dict = {nombre: id for id, nombre in cursor.fetchall()}
    
    # Insertar relaciones entre actividades y productos
    relaciones_actividad_producto = [
        (actividades_dict['Pintura interior'], productos_dict['Pintura blanca'], 0.1),
        (actividades_dict['Pintura interior'], productos_dict['Rodillo'], 0.02),
        (actividades_dict['Pintura interior'], productos_dict['Brocha'], 0.01),
        (actividades_dict['Pintura exterior'], productos_dict['Pintura de colores'], 0.12),
        (actividades_dict['Pintura exterior'], productos_dict['Rodillo'], 0.02),
        (actividades_dict['Pintura exterior'], productos_dict['Brocha'], 0.01),
        (actividades_dict['Instalación de tubería PVC'], productos_dict['Tubo PVC 1/2"'], 1),
        (actividades_dict['Instalación de tubería PVC'], productos_dict['Codo PVC 1/2"'], 0.2),
        (actividades_dict['Instalación de sanitario'], productos_dict['Sanitario'], 1),
        (actividades_dict['Instalación de lavamanos'], productos_dict['Lavamanos'], 1),
        (actividades_dict['Instalación de cableado eléctrico'], productos_dict['Cable #12'], 1),
        (actividades_dict['Instalación de tomacorrientes'], productos_dict['Tomacorriente'], 1),
        (actividades_dict['Instalación de interruptores'], productos_dict['Interruptor'], 1),
        (actividades_dict['Construcción de muro en ladrillo'], productos_dict['Ladrillo'], 50),
        (actividades_dict['Construcción de muro en ladrillo'], productos_dict['Cemento'], 10),
        (actividades_dict['Construcción de muro en ladrillo'], productos_dict['Arena'], 0.05),
        (actividades_dict['Instalación de piso cerámico'], productos_dict['Cerámica'], 1),
        (actividades_dict['Instalación de piso cerámico'], productos_dict['Cemento'], 5),
        (actividades_dict['Fabricación de mueble de cocina'], productos_dict['Madera'], 2),
        (actividades_dict['Fabricación de mueble de cocina'], productos_dict['Tornillos'], 50),
        (actividades_dict['Instalación de puerta'], productos_dict['Puerta'], 1),
        (actividades_dict['Instalación de puerta'], productos_dict['Bisagras'], 1.5)
    ]
    
    for relacion in relaciones_actividad_producto:
        cursor.execute('INSERT OR IGNORE INTO actividad_producto (actividad_id, producto_id, cantidad) VALUES (?, ?, ?)', relacion)
    
    # Insertar relaciones entre actividades
    relaciones_actividades = [
        (actividades_dict['Pintura interior'], actividades_dict['Pintura exterior']),
        (actividades_dict['Instalación de tubería PVC'], actividades_dict['Instalación de sanitario']),
        (actividades_dict['Instalación de tubería PVC'], actividades_dict['Instalación de lavamanos']),
        (actividades_dict['Instalación de cableado eléctrico'], actividades_dict['Instalación de tomacorrientes']),
        (actividades_dict['Instalación de cableado eléctrico'], actividades_dict['Instalación de interruptores']),
        (actividades_dict['Construcción de muro en ladrillo'], actividades_dict['Pintura interior']),
        (actividades_dict['Instalación de piso cerámico'], actividades_dict['Construcción de muro en ladrillo']),
        (actividades_dict['Fabricación de mueble de cocina'], actividades_dict['Instalación de puerta'])
    ]
    
    for relacion in relaciones_actividades:
        cursor.execute('INSERT OR IGNORE INTO actividad_relacionada (actividad_principal_id, actividad_relacionada_id) VALUES (?, ?)', relacion)
    
    # Insertar datos de ejemplo para clientes
    clientes = [
        ('Juan Pérez', 'natural', 'Calle 123 #45-67', None, '3101234567', 'juan@example.com'),
        ('María López', 'natural', 'Carrera 78 #90-12', None, '3109876543', 'maria@example.com'),
        ('Constructora XYZ', 'jurídica', 'Avenida Principal #123', '900123456-7', '6011234567', 'info@constructoraxyz.com'),
        ('Inmobiliaria ABC', 'jurídica', 'Calle Comercial #456', '800987654-3', '6019876543', 'contacto@inmobiliariaabc.com')
    ]
    
    for cliente in clientes:
        cursor.execute('INSERT OR IGNORE INTO clientes (nombre, tipo, direccion, nit, telefono, email) VALUES (?, ?, ?, ?, ?, ?)', cliente)
    
    # Guardar cambios y cerrar conexión
    conn.commit()
    conn.close()
    
    print("Base de datos inicializada correctamente con datos de ejemplo.")

if __name__ == "__main__":
    init_db()
