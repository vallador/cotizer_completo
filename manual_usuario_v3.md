# Manual de Usuario - Sistema de Cotizaciones v3.0

## Índice
1. [Introducción](#introducción)
2. [Instalación y Configuración](#instalación-y-configuración)
3. [Interfaz Principal](#interfaz-principal)
4. [Gestión de Datos](#gestión-de-datos)
   - [Clientes](#clientes)
   - [Actividades](#actividades)
   - [Productos](#productos)
   - [Categorías](#categorías)
   - [Relaciones](#relaciones)
5. [Creación de Cotizaciones](#creación-de-cotizaciones)
6. [Sistema de Filtrado y Relaciones](#sistema-de-filtrado-y-relaciones)
7. [Exportación a Excel](#exportación-a-excel)
8. [Generación de Documentos Word](#generación-de-documentos-word)
9. [Envío por Correo Electrónico](#envío-por-correo-electrónico)
10. [Asistente de Voz (Próximamente)](#asistente-de-voz-próximamente)
11. [Solución de Problemas](#solución-de-problemas)

## Introducción

El Sistema de Cotizaciones es una aplicación diseñada para facilitar la creación, gestión y seguimiento de cotizaciones para empresas de construcción y servicios. Esta versión 3.0 incluye importantes mejoras como:

- Gestión completa de productos con categorías
- Sistema avanzado de filtrado y búsqueda
- Relaciones automáticas entre actividades relacionadas
- Generación de documentos Word personalizados
- Envío de cotizaciones por correo electrónico
- Base para integración con asistente de voz

## Instalación y Configuración

### Requisitos del Sistema
- Python 3.6 o superior
- Bibliotecas: PyQt5, openpyxl, python-docx, sqlite3, smtplib

### Pasos de Instalación
1. Descomprima el archivo ZIP en la ubicación deseada
2. Abra una terminal o línea de comandos
3. Navegue hasta la carpeta del programa
4. Ejecute el siguiente comando para inicializar la base de datos (solo la primera vez):
   ```
   python reset_db.py --auto
   ```
5. Inicie la aplicación con:
   ```
   python main.py
   ```

### Configuración
El archivo `config.json` contiene la configuración del sistema:
- Valores AIU (Administración, Imprevistos, Utilidad)
- Configuración de correo electrónico
- Otras preferencias del sistema

## Interfaz Principal

La interfaz principal del sistema está dividida en varias secciones:

1. **Información del Cliente**: Permite seleccionar o crear un cliente para la cotización.
2. **Actividades**: Muestra las actividades incluidas en la cotización actual.
3. **Búsqueda y Filtrado**: Permite buscar y filtrar actividades por texto o categoría.
4. **Resumen**: Muestra el resumen de costos de la cotización.
5. **Acciones**: Botones para guardar, exportar y enviar la cotización.

## Gestión de Datos

Para acceder a la gestión de datos, haga clic en el botón "Gestionar Datos" en la interfaz principal. Se abrirá una ventana con pestañas para gestionar diferentes tipos de datos.

### Clientes

En esta pestaña puede:
- Ver la lista de clientes existentes
- Agregar nuevos clientes
- Modificar clientes existentes
- Eliminar clientes

Para agregar un cliente, complete los campos requeridos y haga clic en "Agregar Cliente".

### Actividades

En esta pestaña puede:
- Ver la lista de actividades existentes
- Buscar actividades por texto o categoría
- Agregar nuevas actividades
- Modificar actividades existentes
- Eliminar actividades

Cada actividad tiene:
- Descripción
- Unidad de medida
- Valor unitario
- Categoría (opcional)

### Productos

Esta nueva funcionalidad permite gestionar los productos utilizados en las actividades:
- Ver la lista de productos existentes
- Buscar productos por texto o categoría
- Agregar nuevos productos
- Modificar productos existentes
- Eliminar productos

Cada producto tiene:
- Nombre
- Descripción
- Unidad de medida
- Precio unitario
- Categoría (opcional)

### Categorías

Las categorías permiten organizar tanto actividades como productos:
- Ver la lista de categorías existentes
- Agregar nuevas categorías
- Modificar categorías existentes
- Eliminar categorías

### Relaciones

Esta nueva pestaña permite establecer relaciones entre:
- Actividades relacionadas entre sí
- Actividades y los productos que utilizan

Para relacionar actividades:
1. Seleccione una actividad principal
2. Seleccione una actividad relacionada
3. Haga clic en "Agregar"

Para asociar productos a una actividad:
1. Seleccione una actividad
2. Seleccione un producto
3. Especifique la cantidad
4. Haga clic en "Agregar"

## Creación de Cotizaciones

Para crear una nueva cotización:

1. Seleccione un cliente de la lista desplegable o cree uno nuevo
2. Agregue actividades a la cotización:
   - Utilice el sistema de búsqueda y filtrado para encontrar actividades
   - Haga doble clic en una actividad para agregarla
   - Especifique la cantidad
3. Revise el resumen de costos
4. Guarde la cotización haciendo clic en "Guardar Cotización"

## Sistema de Filtrado y Relaciones

El nuevo sistema de filtrado y relaciones facilita la creación de cotizaciones:

### Filtrado de Actividades
- Busque actividades por texto en la descripción
- Filtre por categoría
- Combine ambos criterios para búsquedas más precisas

### Actividades Relacionadas
Al seleccionar una actividad, el sistema sugerirá automáticamente actividades relacionadas. Por ejemplo, si selecciona "Pintura interior", el sistema podría sugerir "Pintura exterior" como actividad relacionada.

### Productos Asociados
Al agregar una actividad, puede ver los productos asociados a esa actividad y sus cantidades recomendadas.

## Exportación a Excel

Para exportar una cotización a Excel:

1. Complete la cotización con cliente y actividades
2. Haga clic en "Generar Excel"
3. Seleccione la ubicación donde guardar el archivo
4. El sistema generará un archivo Excel con:
   - Información del cliente
   - Lista de actividades con cantidades y valores
   - Cálculos de AIU según el tipo de cliente
   - Totales y subtotales

## Generación de Documentos Word

Esta nueva funcionalidad permite generar documentos Word personalizados:

1. Complete la cotización con cliente y actividades
2. Haga clic en "Generar Word"
3. En el diálogo que aparece, configure:
   - Tipo de documento (largo para empresas, corto para empresas, corto para personas naturales)
   - Fecha de generación
   - Nombre del edificio o cliente
   - NIT (para cotizaciones jurídicas)
   - Dirección
   - Referencia del trabajo
   - Validez de la oferta
   - Personal de obra
   - Plazo de ejecución
   - Forma de pago
4. Haga clic en "Generar"
5. Seleccione la ubicación donde guardar el archivo

## Envío por Correo Electrónico

Para enviar una cotización por correo electrónico:

1. Complete la cotización y genere los archivos Excel y/o Word
2. Haga clic en "Enviar por Correo"
3. En el diálogo que aparece, configure:
   - Dirección de correo del destinatario
   - Asunto del correo
   - Mensaje
   - Archivos adjuntos (seleccione los archivos generados)
4. Haga clic en "Enviar"

## Asistente de Voz (Próximamente)

El sistema incluye la base para un asistente de voz que permitirá:
- Crear cotizaciones mediante comandos de voz
- Buscar actividades por voz
- Agregar actividades a la cotización
- Generar documentos y enviar correos

Esta funcionalidad estará disponible en próximas actualizaciones.

## Solución de Problemas

### Problemas con la Base de Datos
Si experimenta problemas con la base de datos, puede reiniciarla ejecutando:
```
python reset_db.py --auto
```
Esto creará una nueva base de datos con datos de ejemplo.

### Errores al Guardar Cotizaciones
Asegúrese de que:
- Ha seleccionado un cliente
- Ha agregado al menos una actividad
- Las cantidades son números válidos

### Problemas con la Gestión de Datos
Si la ventana de gestión de datos se cierra inesperadamente:
- Reinicie la aplicación
- Verifique que la base de datos no esté corrupta

### Soporte Técnico
Para obtener ayuda adicional, contacte al soporte técnico en:
soporte@cotizaciones.com
