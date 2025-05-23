# Diseño de Funcionalidad para Generación de Documentos Word

## Requisitos del Usuario

El usuario ha solicitado implementar la generación de documentos Word con las siguientes características:

1. **Tres formatos diferentes de plantillas**:
   - Formato largo para empresas (jurídicas)
   - Formato corto para empresas (jurídicas)
   - Formato corto para personas naturales

2. **Reemplazo dinámico de textos** con los siguientes campos:
   - Fecha de generación del documento
   - Nombre del edificio o cliente
   - NIT (solo para cotizaciones jurídicas)
   - Dirección (solo para cotizaciones jurídicas)
   - Referencia del trabajo a realizar
   - Validez de la oferta
   - Personal de obra (número de cuadrillas: uno, dos, tres)
   - Número de operarios
   - Plazo de ejecución (días hábiles)
   - Forma de pago:
     - Contraentrega, o
     - Tres porcentajes: % al inicio, % al avance, % al finalizar

3. **Inserción de tabla de Excel** como imagen en el documento Word

## Diseño de la Solución

### 1. Estructura de Plantillas

Se crearán tres plantillas Word base:
- `plantilla_juridica_larga.docx`
- `plantilla_juridica_corta.docx`
- `plantilla_natural_corta.docx`

Estas plantillas contendrán marcadores de posición para los campos a reemplazar, utilizando el formato `{{campo}}`.

### 2. Campos de Reemplazo

Los siguientes marcadores se utilizarán en las plantillas:

```
{{fecha}}                  - Fecha de generación (formato: DD de Mes de AAAA)
{{cliente}}                - Nombre del cliente o edificio
{{nit}}                    - NIT del cliente (solo jurídicas)
{{direccion}}              - Dirección del cliente (solo jurídicas)
{{referencia}}             - Referencia del trabajo
{{validez}}                - Validez de la oferta en días
{{cuadrillas}}             - Número de cuadrillas (uno, dos, tres)
{{operarios}}              - Número de operarios
{{plazo}}                  - Plazo de ejecución en días hábiles
{{forma_pago}}             - Descripción de la forma de pago
{{porcentaje_inicio}}      - Porcentaje al inicio (si aplica)
{{porcentaje_avance}}      - Porcentaje al avance (si aplica)
{{avance_requerido}}       - Porcentaje de avance requerido (si aplica)
{{porcentaje_final}}       - Porcentaje al finalizar (si aplica)
{{tabla_cotizacion}}       - Marcador para insertar la tabla de Excel
```

### 3. Flujo de Generación de Documentos

1. **Selección de Plantilla**:
   - Determinar el tipo de cliente (natural o jurídica)
   - Seleccionar formato (largo o corto) según preferencia del usuario
   - Cargar la plantilla correspondiente

2. **Recopilación de Datos**:
   - Obtener datos de la cotización (cliente, fecha, etc.)
   - Solicitar al usuario los datos adicionales específicos para el documento Word
   - Validar que todos los campos requeridos estén completos

3. **Generación de Tabla Excel**:
   - Generar o utilizar el archivo Excel de la cotización
   - Capturar la tabla como imagen

4. **Reemplazo de Textos**:
   - Reemplazar todos los marcadores en la plantilla con los datos correspondientes
   - Insertar la imagen de la tabla en la posición del marcador `{{tabla_cotizacion}}`

5. **Guardado del Documento**:
   - Guardar el documento con un nombre que incluya el número de cotización y fecha
   - Ofrecer opción para enviar por correo electrónico

### 4. Interfaz de Usuario

Se implementará un diálogo que permita al usuario:
- Seleccionar el tipo de formato (largo/corto)
- Ingresar los datos adicionales requeridos
- Previsualizar y confirmar antes de generar
- Enviar por correo electrónico después de generar

### 5. Componentes Técnicos

1. **WordController**:
   - Métodos para cargar plantillas
   - Métodos para reemplazar texto
   - Métodos para insertar imágenes
   - Métodos para guardar documentos

2. **ExcelController (Extensión)**:
   - Método para capturar tabla como imagen

3. **Diálogo de Configuración Word**:
   - Interfaz para ingresar datos adicionales
   - Validación de campos
   - Opciones de formato

### 6. Consideraciones Adicionales

- Las plantillas deben ser configurables y almacenadas en una carpeta específica
- Los documentos generados deben guardarse en una ubicación accesible
- Se debe implementar manejo de errores robusto
- La funcionalidad debe ser extensible para futuros formatos o campos
