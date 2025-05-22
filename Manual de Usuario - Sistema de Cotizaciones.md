# Manual de Usuario - Sistema de Cotizaciones

## Introducción

El Sistema de Cotizaciones es una aplicación de escritorio diseñada para facilitar la creación, gestión y exportación de cotizaciones para proyectos de construcción y servicios. Permite manejar clientes, actividades, aplicar valores de AIU (Administración, Imprevistos y Utilidad) y generar documentos en formato Excel.

## Requisitos del Sistema

- Python 3.6 o superior
- PyQt5
- Openpyxl
- SQLite3

## Instalación

1. Descomprima el archivo `cotizer.zip` en la ubicación deseada
2. Abra una terminal o línea de comandos
3. Navegue hasta la carpeta donde descomprimió el archivo
4. Ejecute el siguiente comando para inicializar la base de datos (solo la primera vez):
   ```
   python init_db.py
   ```
5. Inicie la aplicación con:
   ```
   python main.py
   ```

## Funcionalidades Principales

### Gestión de Clientes

- **Agregar Cliente**: Complete los campos de nombre, tipo (natural o jurídica), dirección, NIT (para personas jurídicas), teléfono y email, luego haga clic en "Guardar Cliente".
- **Seleccionar Cliente**: Use el menú desplegable para seleccionar un cliente existente. Sus datos se cargarán automáticamente en los campos correspondientes.

### Gestión de Actividades

- **Agregar Actividad**: Seleccione una actividad del menú desplegable, especifique la cantidad y haga clic en "Agregar Actividad". La actividad se añadirá a la tabla con su valor unitario y total calculado.
- **Eliminar Actividad**: Haga clic en el botón "X" en la fila correspondiente para eliminar una actividad de la cotización.
- **Editar Cantidad/Precio**: Haga doble clic en los campos de cantidad o precio unitario para modificarlos. El total se recalculará automáticamente.

### Configuración de AIU

- **Ajustar Valores**: Modifique los valores de Administración, Imprevistos, Utilidad e IVA según sus necesidades. Estos valores se aplicarán automáticamente en los cálculos de la cotización.

### Generación de Cotizaciones

- **Guardar Cotización**: Haga clic en "Guardar Cotización" para almacenar la cotización en la base de datos. Se generará automáticamente un número único de cotización.
- **Generar Excel**: Haga clic en "Generar Excel" para crear un archivo Excel con el detalle de la cotización. El archivo se guardará en la carpeta del programa con un nombre que incluye la fecha y hora.

### Gestión de Datos

- **Acceder a Gestión de Datos**: Haga clic en "Gestionar Datos" para abrir una ventana que permite administrar clientes y actividades de forma más detallada.

## Flujo de Trabajo Recomendado

1. Seleccione un cliente existente o cree uno nuevo
2. Ajuste los valores de AIU según el proyecto
3. Agregue las actividades necesarias a la cotización
4. Verifique los totales calculados
5. Guarde la cotización en la base de datos
6. Genere el archivo Excel para enviar al cliente

## Solución de Problemas

- **Base de datos no encontrada**: Ejecute `python init_db.py` para crear la base de datos inicial.
- **Errores en cálculos**: Verifique que los valores ingresados sean numéricos y que los porcentajes de AIU estén correctamente configurados.
- **Problemas al generar Excel**: Asegúrese de tener permisos de escritura en la carpeta del programa.

## Contacto y Soporte

Para soporte técnico o consultas sobre el funcionamiento del sistema, contacte al desarrollador a través de los canales establecidos.

---

# Cambios y Mejoras Implementadas

## Mejoras en la Base de Datos
- Implementación completa de tablas para cotizaciones y detalles
- Creación de métodos para gestionar cotizaciones (guardar, consultar, eliminar)
- Inicialización automática de la base de datos con datos de ejemplo

## Mejoras en el Modelo de Cotización
- Integración completa con el gestor de AIU
- Cálculo correcto de totales según el tipo de cliente (natural o jurídica)
- Métodos para obtener desglose detallado de valores

## Mejoras en Controladores
- Eliminación de duplicidades de código
- Implementación completa de métodos para generar y guardar cotizaciones
- Integración con archivo de configuración para valores predeterminados

## Mejoras en la Interfaz de Usuario
- Validación de entradas para evitar errores
- Botones para eliminar actividades de la tabla
- Cálculo automático de totales con AIU según tipo de cliente
- Opción para guardar cotizaciones en la base de datos

## Mejoras en la Exportación
- Generación de archivos Excel con formato profesional
- Nombres de archivo con fecha y número de cotización
- Aplicación correcta de AIU según tipo de cliente

## Configuración
- Implementación del archivo config.json para almacenar valores predeterminados
- Carga y guardado automático de configuraciones
