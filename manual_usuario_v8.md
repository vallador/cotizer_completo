# Manual de Usuario - Sistema de Cotizaciones v8.0

## Introducción

El Sistema de Cotizaciones es una aplicación diseñada para facilitar la creación, gestión y seguimiento de cotizaciones para servicios de construcción y mantenimiento. Esta versión 8.0 incluye múltiples mejoras de estabilidad, nuevas funcionalidades y correcciones de errores.

## Requisitos del Sistema

- Sistema operativo: Windows, macOS o Linux
- Python 3.6 o superior
- Bibliotecas requeridas (incluidas en el archivo requirements.txt):
  - PyQt5
  - openpyxl
  - python-docx
  - sqlite3
  - smtplib

## Instalación

1. Descomprima el archivo ZIP en una ubicación de su preferencia
2. Abra una terminal o línea de comandos en la carpeta descomprimida
3. Ejecute el siguiente comando para inicializar la base de datos con datos de ejemplo:
   ```
   python reset_db.py --auto
   ```
4. Ejecute la aplicación con:
   ```
   python main.py
   ```

## Funcionalidades Principales

### Gestión de Clientes

- **Selección de cliente**: Utilice el desplegable para seleccionar un cliente existente
- **Creación de cliente**: Complete el formulario y haga clic en "Guardar Cliente"
- **Edición de cliente**: Seleccione un cliente, modifique sus datos y haga clic en "Guardar Cliente"
- **Gestión avanzada**: Haga clic en "Gestionar Datos" para acceder a funciones adicionales

### Gestión de Actividades

- **Agregar actividad manual**: Complete los campos de descripción, cantidad, unidad y valor unitario, luego haga clic en "Agregar Actividad"
- **Buscar actividades**: Utilice el campo de búsqueda y el filtro por categoría para encontrar actividades predefinidas
- **Agregar actividad predefinida**: Seleccione una actividad del desplegable y haga clic en "Agregar Actividad"
- **Agregar actividad relacionada**: Seleccione una actividad relacionada y haga clic en "Agregar Relacionada"
- **Eliminar actividad**: Haga clic en el botón "Eliminar" en la fila correspondiente

### Gestión de Cotizaciones

- **Guardar cotización**: Haga clic en "Guardar Cotización" para almacenar en la base de datos
- **Guardar como archivo**: Haga clic en "Guardar como Archivo" para crear un archivo .cotiz que puede editar posteriormente
- **Abrir archivo**: Haga clic en "Abrir Archivo" para cargar una cotización guardada como archivo

### Exportación y Comunicación

- **Generar Excel**: Crea una hoja de cálculo con el detalle de la cotización
- **Generar Word**: Crea un documento Word basado en plantillas (largo o corto para empresas, corto para personas naturales)
- **Enviar por correo**: Envía la cotización por correo electrónico con archivos adjuntos

### Personalización

- **Modo oscuro/claro**: Alterne entre temas visuales haciendo clic en el botón "Modo Oscuro/Claro"

## Flujo de Trabajo Recomendado

1. Seleccione o cree un cliente
2. Agregue actividades a la cotización (manualmente o desde actividades predefinidas)
3. Revise y ajuste cantidades y valores si es necesario
4. Guarde la cotización en la base de datos o como archivo separado
5. Genere documentos Excel y/o Word según necesite
6. Envíe la cotización por correo electrónico si lo desea

## Gestión de Datos

Para acceder a la gestión avanzada de datos, haga clic en el botón "Gestionar Datos". Desde allí podrá:

- Administrar clientes (crear, editar, eliminar)
- Administrar actividades (crear, editar, eliminar)
- Administrar productos (crear, editar, eliminar)
- Administrar categorías (crear, editar, eliminar)
- Establecer relaciones entre actividades y productos

## Gestión de Archivos de Cotización

La nueva funcionalidad de archivos de cotización permite:

- Guardar cotizaciones como archivos independientes (.cotiz)
- Abrir y modificar cotizaciones guardadas
- Compartir cotizaciones entre diferentes instalaciones del sistema

Para utilizar esta función:

1. Complete una cotización normalmente
2. Haga clic en "Guardar como Archivo"
3. Seleccione la ubicación donde desea guardar el archivo
4. Para abrir una cotización guardada, haga clic en "Abrir Archivo" y seleccione el archivo .cotiz

## Solución de Problemas

### La aplicación no inicia

- Verifique que Python esté correctamente instalado
- Asegúrese de haber instalado todas las dependencias
- Compruebe que la base de datos exista ejecutando `python reset_db.py --auto`

### No se muestran los clientes o actividades

- Verifique que la base de datos se haya inicializado correctamente
- Intente reiniciar la aplicación
- Si persiste el problema, ejecute `python reset_db.py --auto` para reinicializar la base de datos

### Error al guardar cotización

- Asegúrese de haber seleccionado un cliente
- Verifique que haya agregado al menos una actividad
- Compruebe que todos los campos obligatorios estén completos

## Contacto y Soporte

Para soporte técnico o consultas, contacte a:
- Email: soporte@cotizaciones.com
- Teléfono: +123 456 7890

---

© 2025 Sistema de Cotizaciones - Todos los derechos reservados
