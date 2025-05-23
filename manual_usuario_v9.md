# Manual de Usuario - Sistema de Cotizaciones v9.0

## Introducción

El Sistema de Cotizaciones es una aplicación diseñada para facilitar la creación, gestión y seguimiento de cotizaciones para servicios de construcción y remodelación. Esta versión 9.0 incluye importantes mejoras en la interfaz de usuario, correcciones de errores y nuevas funcionalidades.

## Requisitos del Sistema

- Sistema operativo: Windows, macOS o Linux
- Python 3.8 o superior
- Bibliotecas requeridas: PyQt5, openpyxl, python-docx

## Instalación

1. Descomprima el archivo ZIP en la ubicación deseada
2. Abra una terminal o línea de comandos
3. Navegue hasta la carpeta donde descomprimió el archivo
4. Ejecute el siguiente comando para inicializar la base de datos (solo la primera vez):
   ```
   python reset_db.py --auto
   ```
5. Para iniciar la aplicación, ejecute:
   ```
   python main.py
   ```

## Interfaz Principal

La interfaz principal del sistema se divide en tres secciones principales:

1. **Información del Cliente**: Permite seleccionar o crear clientes.
2. **Actividades**: Muestra las actividades incluidas en la cotización actual.
3. **Selección de Actividades**: Ubicada a la derecha, permite buscar, filtrar y agregar actividades a la cotización.

### Gestión de Clientes

- **Seleccionar Cliente**: Use el menú desplegable para elegir un cliente existente.
- **Crear Cliente**: Complete los campos del formulario y haga clic en "Guardar Cliente".
- **Gestionar Datos**: Haga clic en este botón para abrir la ventana de gestión de datos donde puede administrar clientes, actividades, productos y categorías.

### Gestión de Actividades

La tabla de actividades muestra todas las actividades incluidas en la cotización actual. Puede:

- **Agregar Actividades**: Seleccione una actividad del panel derecho y haga clic en "Agregar Actividad".
- **Agregar Actividades Relacionadas**: Seleccione una actividad relacionada y haga clic en "Agregar Relacionada".
- **Agregar Actividades Manualmente**: Complete los campos en la sección "Entrada Manual" y haga clic en "Agregar Actividad Manual".
- **Eliminar Actividades**: Haga clic en el botón "Eliminar" junto a la actividad que desea quitar.
- **Editar Actividades**: Haga doble clic en cualquier celda editable para modificar su valor.

### Panel de Selección de Actividades (Lado Derecho)

Este panel mejorado permite:

- **Buscar Actividades**: Ingrese texto en el campo de búsqueda para filtrar actividades.
- **Filtrar por Categoría**: Seleccione una categoría del menú desplegable.
- **Ver Descripción Completa**: La descripción completa de la actividad seleccionada se muestra en un área de texto dedicada.
- **Ver Actividades Relacionadas**: Las actividades relacionadas con la selección actual se muestran en un menú desplegable.

## Funcionalidades Principales

### Generar Excel

Para generar un archivo Excel con la cotización actual:

1. Asegúrese de haber agregado al menos una actividad a la cotización.
2. Haga clic en el botón "Generar Excel".
3. El sistema creará un archivo Excel con formato profesional que incluye:
   - Lista de actividades con cantidades y precios
   - Cálculos de AIU (Administración, Imprevistos, Utilidad)
   - Total de la cotización

### Generar Word

Para generar un documento Word personalizado:

1. Asegúrese de haber agregado al menos una actividad y seleccionado un cliente.
2. Haga clic en el botón "Generar Word".
3. Configure los parámetros del documento en el diálogo que aparece.
4. El sistema creará un documento Word con:
   - Información del cliente
   - Detalles del proyecto
   - Tabla de actividades y precios
   - Condiciones comerciales

### Guardar y Cargar Cotizaciones

El sistema permite guardar cotizaciones como archivos independientes:

- **Guardar como Archivo**: Haga clic en este botón para guardar la cotización actual como un archivo .cotiz.
- **Abrir Archivo**: Haga clic en este botón para cargar una cotización previamente guardada.

Esto facilita la edición posterior de cotizaciones sin afectar la base de datos central.

### Enviar por Correo

Para enviar una cotización por correo electrónico:

1. Complete la cotización con al menos una actividad.
2. Haga clic en "Enviar por Correo".
3. Configure los parámetros del correo en el diálogo que aparece.
4. El sistema enviará un correo con la cotización adjunta en formato Excel.

## Gestión de Datos

La ventana de Gestión de Datos permite administrar:

- **Clientes**: Agregar, editar o eliminar clientes.
- **Actividades**: Gestionar el catálogo de actividades disponibles.
- **Productos**: Administrar productos asociados a las actividades.
- **Categorías**: Organizar actividades y productos por categorías.
- **Relaciones**: Establecer relaciones entre actividades para sugerencias automáticas.

Para acceder a esta ventana, haga clic en el botón "Gestionar Datos" en la pantalla principal.

## Modo Oscuro

El sistema incluye un modo oscuro para reducir la fatiga visual:

- Haga clic en el botón "Modo Oscuro" para activarlo.
- Haga clic en "Modo Claro" para volver al tema predeterminado.

## Solución de Problemas

Si encuentra algún problema con la aplicación:

1. Asegúrese de tener instaladas todas las dependencias requeridas.
2. Verifique que la base de datos esté correctamente inicializada.
3. Si la aplicación no responde, ciérrela y vuelva a abrirla.
4. Para problemas persistentes, puede reiniciar la base de datos ejecutando `python reset_db.py --auto`.

## Novedades en la Versión 9.0

- **Interfaz mejorada**: Panel de selección de actividades reubicado al lado derecho para mejor usabilidad.
- **Visualización completa**: Ahora se muestra la descripción completa de las actividades en áreas de texto dedicadas.
- **Corrección de errores**: Solucionados problemas con el filtrado de actividades y la generación de Excel.
- **Modo oscuro mejorado**: Mejor contraste y visibilidad en todos los elementos de la interfaz.
- **Gestión de archivos**: Sistema mejorado para guardar y cargar cotizaciones como archivos independientes.

## Próximas Funcionalidades

En futuras versiones se planea implementar:

- Asistente de voz para control por comandos hablados
- Integración con servicios en la nube
- Aplicación móvil complementaria
- Análisis estadístico de cotizaciones
