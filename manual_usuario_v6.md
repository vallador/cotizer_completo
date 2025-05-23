# Manual de Usuario - Sistema de Cotizaciones v6.0

## Introducción

El Sistema de Cotizaciones es una aplicación completa para la gestión de cotizaciones de servicios y productos. Permite crear, editar, guardar y exportar cotizaciones, así como generar documentos Excel y Word con formatos profesionales.

Esta versión 6.0 incluye importantes mejoras y correcciones que optimizan la experiencia de usuario y añaden nuevas funcionalidades.

## Requisitos del Sistema

- Sistema operativo: Windows 10/11, macOS o Linux
- Python 3.8 o superior
- Bibliotecas requeridas (incluidas en el archivo requirements.txt)
- Espacio en disco: 50 MB mínimo

## Instalación

1. Descomprima el archivo ZIP en una ubicación de su preferencia
2. Abra una terminal o línea de comandos en la carpeta del programa
3. Ejecute el siguiente comando para inicializar la base de datos (solo la primera vez):
   ```
   python reset_db.py --auto
   ```
4. Inicie la aplicación con:
   ```
   python main.py
   ```

## Nuevas Funcionalidades

### 1. Gestión de Archivos de Cotización

Ahora puede guardar y abrir cotizaciones como archivos independientes, lo que permite:
- Editar cotizaciones existentes fácilmente
- Compartir cotizaciones entre diferentes instalaciones
- Mantener un respaldo de cotizaciones importantes
- Modificar cotizaciones sin afectar la base de datos central

Para utilizar esta función:
- **Guardar como Archivo**: Guarda la cotización actual como un archivo JSON
- **Abrir Archivo**: Abre el diálogo de gestión de archivos donde puede:
  - Ver todas las cotizaciones guardadas
  - Abrir una cotización para edición
  - Importar cotizaciones desde otras ubicaciones
  - Exportar cotizaciones a ubicaciones externas
  - Eliminar cotizaciones que ya no necesite

### 2. Modo Oscuro Mejorado

El modo oscuro ha sido completamente rediseñado para mejorar la visibilidad y reducir la fatiga visual:
- Mayor contraste en textos y botones
- Colores optimizados para todos los elementos de la interfaz
- Transición suave entre modos claro y oscuro
- Compatible con todas las ventanas y diálogos

Para activar el modo oscuro, simplemente haga clic en el botón "Modo Oscuro" en la parte inferior de la ventana principal.

### 3. Sistema Avanzado de Filtrado y Relaciones

El sistema de filtrado y relaciones entre actividades ha sido mejorado:
- Búsqueda por texto en descripciones
- Filtrado por categorías
- Sugerencias automáticas de actividades relacionadas
- Integración con productos asociados

## Correcciones Importantes

Esta versión corrige varios errores críticos:

1. **Error en Gestión de Datos**: Se ha corregido el problema que impedía abrir la ventana de gestión de datos.

2. **Error al Guardar Cotización**: Se ha solucionado el error "NOT NULL constraint failed" que ocurría al guardar cotizaciones o generar documentos Word.

3. **Sincronización de Cliente**: Se ha mejorado la sincronización entre la selección de cliente y los datos mostrados en la interfaz.

## Guía de Uso

### Creación de Cotizaciones

1. **Información del Cliente**:
   - Seleccione un cliente existente del menú desplegable, o
   - Complete los campos de información del cliente y guárdelo con el botón "Guardar Cliente"

2. **Agregar Actividades**:
   - Utilice los filtros de búsqueda y categoría para encontrar actividades
   - Seleccione una actividad del menú desplegable
   - Especifique la cantidad
   - Haga clic en "Agregar Actividad"
   - También puede agregar actividades relacionadas seleccionándolas del menú "Actividades Relacionadas"

3. **Editar Actividades**:
   - Modifique la cantidad o valor unitario directamente en la tabla
   - Elimine actividades con el botón "X" en cada fila

4. **Guardar y Exportar**:
   - **Guardar Cotización**: Guarda la cotización en la base de datos
   - **Guardar como Archivo**: Guarda la cotización como un archivo JSON independiente
   - **Generar Excel**: Crea una hoja de cálculo con formato profesional
   - **Generar Word**: Crea un documento Word basado en plantillas personalizables
   - **Enviar por Correo**: Envía la cotización por correo electrónico con archivos adjuntos

### Gestión de Datos

Acceda a la gestión de datos haciendo clic en el botón "Gestionar Datos". Desde allí podrá:

1. **Gestionar Clientes**:
   - Ver, agregar, editar y eliminar clientes

2. **Gestionar Actividades**:
   - Ver, agregar, editar y eliminar actividades
   - Asignar categorías a las actividades
   - Establecer relaciones entre actividades

3. **Gestionar Productos**:
   - Ver, agregar, editar y eliminar productos
   - Asignar productos a actividades

4. **Gestionar Categorías**:
   - Ver, agregar, editar y eliminar categorías para organizar actividades y productos

## Recomendación para Asistente de Voz

Se ha incluido un documento detallado con recomendaciones para la implementación del asistente de voz como una aplicación externa. Esta arquitectura ofrece:

- Mejor rendimiento al separar el procesamiento de voz del sistema principal
- Mayor flexibilidad para usar el asistente desde diferentes dispositivos
- Escalabilidad para mejorar el asistente sin modificar el código principal
- Mantenimiento simplificado al actualizar componentes por separado

Consulte el archivo "asistente_voz_recomendacion.md" para más detalles.

## Solución de Problemas

### La aplicación no inicia
- Verifique que Python 3.8 o superior esté instalado
- Asegúrese de haber instalado todas las dependencias
- Compruebe que la base de datos esté inicializada correctamente

### Error al guardar cotización
- Asegúrese de haber seleccionado un cliente
- Verifique que haya al menos una actividad en la tabla
- Compruebe que todos los valores numéricos sean válidos

### Problemas con el modo oscuro
- Intente alternar entre modo claro y oscuro varias veces
- Reinicie la aplicación si persisten problemas visuales

## Contacto y Soporte

Para soporte técnico o sugerencias, contacte a:
- Email: soporte@cotizaciones.com
- Teléfono: +57 123 456 7890

---

© 2025 Sistema de Cotizaciones - Todos los derechos reservados
