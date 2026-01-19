# 📊 QuotationMaster Pro: Sistema Inteligente de Gestión de Cotizaciones

[![Python Version](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![PyQt5](https://img.shields.io/badge/UI-PyQt5-green.svg)](https://www.riverbankcomputing.com/software/pyqt/)
[![Database](https://img.shields.io/badge/DB-SQLite-lightgrey.svg)](https://www.sqlite.org/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

**QuotationMaster Pro** es una solución empresarial de escritorio diseñada para automatizar el ciclo completo de creación, gestión y seguimiento de cotizaciones técnicas y comerciales. Aprovechando el poder de la automatización COM de Microsoft Office, transforma datos brutos en documentos PDF profesionales y listos para el cliente en segundos.

---

## 🔥 Funcionalidades Destacadas

### 🚀 Automatización Cross-Office
- **Generación One-Click**: Crea simultáneamente presupuestos en Excel y propuestas técnicas en Word perfectamente formateadas.
- **Inyección Dinámica**: Mapeo automático de datos desde la UI hacia marcadores (bookmarks) de Word y celdas de Excel.
- **Conversión Inteligente**: Exportación nativa a PDF manteniendo la fidelidad del diseño original.

### 📊 Dashboard de Negocio y Control
- **Visibilidad 360°**: Seguimiento en tiempo real de cotizaciones pendientes, ganadas, perdidas y canceladas.
- **Analytics Integrado**: Cálculo automático de tasa de conversión y volúmenes de cotización mensuales.
- **Modo Prueba**: Separa tus borradores y pruebas de las estadísticas reales de negocio con un solo toggle.

### 🛡️ Gestión Robusta de Proyectos
- **Snapshots de Recuperación**: Guarda el estado completo de cada cotización (actividades, AIU, pólizas) permitiendo su carga o duplicación posterior sin pérdida de datos.
- **Fusión de PDF (Smart Merge)**: Une automáticamente la propuesta, el presupuesto y anexos externos (certificaciones, pólizas) en un único archivo consolidado.
- **Gestión de Clientes**: Base de datos centralizada de clientes y proyectos para una precarga rápida.

---

## 🛠️ Stack Tecnológico

Entorno desarrollado con precisión y eficiencia:

- **Interfaz**: PyQt5 (Modern UI, responsive, modo oscuro disponible).
- **Core**: Python 3.10 con arquitectura orientada a objetos (OOP).
- **Automatización**: `win32com.client` para orquestación de Microsoft Office.
- **Documental**: `python-docx`, `openpyxl`, `PyPDF2`.
- **Persistencia**: SQLite3 con serialización JSON para snapshots complejos.

---

## 📂 Arquitectura del Proyecto

```text
├── main.py                 # Punto de entrada de la aplicación
├── controllers/            # Lógica de negocio y orquestación
├── views/                  # UI (Main Window, Dashboard, Dialogs)
├── utils/                  # Motores de DB, PDF Merger y Office Automation
├── data/                   # Base de datos local
└── plantillas/             # Plantillas base de Word y Excel (Master Templates)
```

---

## 🚀 Instalación y Uso

### Requisitos Previos
*   Windows OS (Requerido para automatización de Office).
*   Microsoft Office (Word y Excel) instalado.
*   Python 3.10+.

### Configuración
1. Clonar el repositorio:
   ```bash
   git clone https://github.com/tu-usuario/quotation-master-pro.git
   cd quotation-master-pro
   ```
2. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   ```
3. Ejecutar la aplicación:
   ```bash
   python main.py
   ```

---

## 🧩 Flujo de Trabajo Técnico

1.  **Carga**: Selecciona un cliente o crea uno nuevo.
2.  **Configuración**: Agrega actividades y capítulos. Ajusta el AIU (Administración, Imprevistos, Utilidad).
3.  **Personalización**: Define pólizas, personal y validez en el diálogo de Word.
4.  **Ejecución**: El motor COM abre Office transparentemente, procesa las plantillas y genera el PDF final unido.
5.  **Seguimiento**: La cotización se registra automáticamente en el Dashboard para su futura gestión comercial.

---

## 🤝 Contribuciones

¿Quieres mejorar QuotationMaster Pro? Las contribuciones son bienvenidas:

1. Realiza un **Fork** del proyecto.
2. Crea una **Rama** para tu funcionalidad (`git checkout -b feature/NuevaMejora`).
3. Haz un **Commit** de tus cambios (`git commit -m 'Add: Nueva mejora visual'`).
4. Haz un **Push** a la rama (`git push origin feature/NuevaMejora`).
5. Abre un **Pull Request**.

---

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo `LICENSE` para más detalles.

---

<p align="center">
  Hecho con ❤️ para optimizar la productividad empresarial.
</p>
