# Manual de Usuario - Sistema de Cotizaciones (Versión 2.0)

## Introducción

El Sistema de Cotizaciones es una aplicación de escritorio diseñada para facilitar la creación, gestión y exportación de cotizaciones para proyectos de construcción y servicios. Esta versión 2.0 incluye importantes mejoras como la generación de documentos Word personalizados y el envío de cotizaciones por correo electrónico.

## Requisitos del Sistema

- Python 3.6 o superior
- PyQt5
- Openpyxl
- Python-docx
- Matplotlib
- Pillow
- Smtplib (incluido en Python estándar)
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

## Nuevas Funcionalidades

### 1. Generación de Documentos Word

Ahora puede generar documentos Word profesionales a partir de sus cotizaciones con las siguientes características:

- **Tres formatos diferentes**:
  - Formato largo para empresas (jurídicas)
  - Formato corto para empresas (jurídicas)
  - Formato corto para personas naturales

- **Personalización completa**:
  - Referencia del trabajo
  - Validez de la oferta
  - Personal de obra (cuadrillas y operarios)
  - Plazo de ejecución
  - Forma de pago (contraentrega o porcentajes)

- **Inclusión automática** de la tabla de cotización como imagen

### 2. Envío de Cotizaciones por Correo Electrónico

Envíe sus cotizaciones directamente desde la aplicación:

- Configuración de servidor SMTP personalizable
- Envío de múltiples archivos adjuntos (Excel y Word)
- Personalización de asunto y mensaje
- Guardado de configuración para futuros envíos

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
- **Generar Excel**: Haga clic en "Generar Excel" para crear un archivo Excel con el detalle de la cotización.
- **Generar Word**: Haga clic en "Generar Word" para crear un documento Word personalizado basado en la cotización.
- **Enviar por Correo**: Haga clic en "Enviar por Correo" para enviar la cotización por correo electrónico.

### Gestión de Datos

- **Acceder a Gestión de Datos**: Haga clic en "Gestionar Datos" para abrir una ventana que permite administrar clientes y actividades de forma más detallada.

## Guía Paso a Paso

### Crear y Guardar una Cotización

1. Seleccione un cliente existente o cree uno nuevo
2. Ajuste los valores de AIU según el proyecto
3. Agregue las actividades necesarias a la cotización
4. Verifique los totales calculados
5. Haga clic en "Guardar Cotización"
6. Elija si desea generar un archivo Excel

### Generar un Documento Word

1. Cree y guarde una cotización como se indicó anteriormente
2. Haga clic en "Generar Word"
3. Seleccione el formato deseado (largo o corto)
4. Complete la información adicional:
   - Referencia del trabajo
   - Validez de la oferta
   - Personal y plazos
   - Forma de pago
5. Haga clic en "Generar Documento"
6. Elija si desea enviar el documento por correo electrónico

### Enviar Cotización por Correo

1. Haga clic en "Enviar por Correo" (o elija esta opción después de generar un documento)
2. Si no hay archivos generados, se le preguntará si desea crear un Excel primero
3. Complete los campos de destinatario, asunto y mensaje
4. Haga clic en "Enviar"

### Configurar Servidor de Correo

1. En el diálogo de envío de correo, haga clic en "Configuración"
2. Complete los datos del servidor SMTP:
   - Servidor (ej. smtp.gmail.com)
   - Puerto (ej. 587)
   - Usuario y contraseña
   - Remitente predeterminado
3. Haga clic en "Probar Conexión" para verificar
4. Haga clic en "Guardar Configuración"

## Solución de Problemas

- **Error al gestionar datos**: Asegúrese de que la base de datos esté correctamente inicializada ejecutando `python init_db.py`.
- **Error al guardar cotización**: Verifique que haya seleccionado un cliente y agregado al menos una actividad.
- **Error al enviar correo**: Compruebe la configuración del servidor SMTP y asegúrese de que su proveedor de correo permita el acceso de aplicaciones menos seguras o utilice contraseñas de aplicación.
- **Error al generar Word**: Verifique que las bibliotecas python-docx, matplotlib y pillow estén instaladas correctamente.

## Contacto y Soporte

Para soporte técnico o consultas sobre el funcionamiento del sistema, contacte al desarrollador a través de los canales establecidos.
