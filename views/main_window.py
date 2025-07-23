from PyQt5.QtWidgets import QMainWindow, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QLineEdit, QComboBox, QTableWidget, \
    QTableWidgetItem, QMessageBox, QWidget, QDoubleSpinBox, QHeaderView, QTextEdit, QFileDialog, QSplitter,\
    QScrollArea, QGridLayout, QApplication, QVBoxLayout, QHBoxLayout, QAbstractItemView, QGroupBox,QFormLayout
from PyQt5.QtCore import Qt, pyqtSlot, QMimeData, QByteArray
from PyQt5.QtGui import QPalette, QColor, QDrag, QFont
import os
from datetime import datetime
from controllers.word_controller import WordController
from views.data_management_window import DataManagementWindow
from views.email_dialog import SendEmailDialog
from views.word_dialog import WordConfigDialog
from views.cotizacion_file_dialog import CotizacionFileDialog

class EditableTableWidgetItem(QTableWidgetItem):
    """Un QTableWidgetItem que se asegura de que los valores numéricos se traten como texto."""
    def __init__(self, value, editable=True):
        super().__init__(str(value))
        if not editable:
            self.setFlags(self.flags() & ~Qt.ItemIsEditable)


class DraggableTableWidget(QTableWidget):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        # Es importante que el modo de drop sea InternalMove para que Qt gestione
        # el movimiento de la fila automáticamente si no lo haces tú manualmente.
        # Sin embargo, como lo haces manualmente, esto es más una indicación.
        self.setDragDropMode(self.DragDropMode.InternalMove)
        self.setSelectionBehavior(self.SelectionBehavior.SelectRows)
        self.setDropIndicatorShown(True)
        self.drag_row_data = None
        self.drag_row_index = -1  # Inicializar el índice

    def startDrag(self, supportedActions):
        """Inicia el arrastre de una fila."""
        try:
            self.drag_row_index = self.currentRow()
            if self.drag_row_index < 0:
                return

            # No permitir arrastrar filas de capítulo
            first_item = self.item(self.drag_row_index, 0)
            if first_item and first_item.data(Qt.UserRole) and first_item.data(Qt.UserRole).get('type') == 'chapter':
                return

            self.drag_row_data = []
            # Guardar los datos de la fila que se está arrastrando
            for col in range(self.columnCount()):  # Itera sobre todas las columnas
                item = self.item(self.drag_row_index, col)
                # Guardamos el texto y cualquier otro dato relevante (como UserRole)
                if item:
                    self.drag_row_data.append({'text': item.text(), 'data': item.data(Qt.UserRole)})
                else:
                    self.drag_row_data.append(None)

            drag = QDrag(self)
            # Definimos un MIME type personalizado para asegurarnos de que solo
            # nuestro widget acepte el drop.
            mimeData = QMimeData()
            # Es necesario poner algún dato, aunque sea vacío.
            mimeData.setData('application/x-delfos-table-row', QByteArray())
            drag.setMimeData(mimeData)

            # Inicia la operación de arrastre y espera a que termine.
            # El cursor cambiará según lo que retornen dragEnterEvent/dragMoveEvent.
            drag.exec(supportedActions)

        except Exception as e:
            print(f"[startDrag] ERROR: {e}")

    def dragEnterEvent(self, event):
        """
        Este evento se dispara cuando el cursor entra en el widget mientras se arrastra.
        Aquí debemos indicar si el widget puede aceptar el drop.
        """
        # Comprobamos si los datos arrastrados tienen nuestro formato personalizado.
        if event.mimeData().hasFormat('application/x-delfos-table-row'):
            event.accept()  # Aceptamos el evento, el cursor cambiará a una flecha.
        else:
            event.ignore()  # Lo rechazamos, el cursor mostrará el signo de prohibido.

    def dragMoveEvent(self, event):
        """
        Este evento se dispara continuamente mientras el cursor se mueve sobre el widget.
        También debe aceptar el evento para que el drop sea posible.
        """
        if event.mimeData().hasFormat('application/x-delfos-table-row'):
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        """Maneja el evento de soltar una fila."""
        if not event.mimeData().hasFormat('application/x-delfos-table-row'):
            event.ignore()
            return

        drop_row = self.indexAt(event.pos()).row()

        # Si se suelta fuera de cualquier fila, se añade al final.
        if drop_row == -1:
            drop_row = self.rowCount()

        # Si se suelta en la misma fila de origen, no hacemos nada.
        if drop_row == self.drag_row_index:
            return

        # Lógica para manejar el drop sobre un capítulo (impedirlo)
        drop_item = self.item(drop_row, 0)
        if drop_item and drop_item.data(Qt.UserRole) and drop_item.data(Qt.UserRole).get('type') == 'chapter':
            # Opcional: podrías buscar la siguiente fila válida o simplemente ignorar el drop
            event.ignore()
            return

        # Guardamos el índice de la fila de origen y la eliminamos.
        # Es importante hacerlo antes de calcular el nuevo índice de inserción.
        origin_row = self.drag_row_index
        self.removeRow(origin_row)

        # Si la fila se mueve hacia abajo, el índice de destino disminuye en uno
        # porque acabamos de eliminar una fila de antes.
        if origin_row < drop_row:
            drop_row -= 1

        # Insertamos una nueva fila en la posición de destino
        self.insertRow(drop_row)
        for col, item_data in enumerate(self.drag_row_data):
            if item_data:
                new_item = QTableWidgetItem(item_data['text'])
                new_item.setData(Qt.UserRole, item_data['data'])  # Restauramos los datos de rol
                self.setItem(drop_row, col, new_item)

        # Seleccionar la fila que acabamos de mover
        self.selectRow(drop_row)

        # Llamadas a métodos del padre para actualizar la UI
        if hasattr(self.parent(), "reconnect_delete_button"):
            self.parent().reconnect_delete_button(drop_row)
        if hasattr(self.parent(), "update_totals"):
            self.parent().update_totals()

        event.accept()

    def insert_chapter_header(self, chapter_id, chapter_name):
        """Método público para insertar una fila de encabezado de capítulo en la tabla."""
        row = self.rowCount()  # Siempre insertar al final
        self.insertRow(row)
        header_item = QTableWidgetItem(chapter_name.upper())
        header_item.setTextAlignment(Qt.AlignCenter)
        header_item.setData(Qt.UserRole, {'type': 'chapter', 'id': chapter_id, 'name': chapter_name})
        font = QFont()
        font.setBold(True)
        header_item.setFont(font)
        header_item.setBackground(QColor("#008080"))
        header_item.setForeground(QColor("white"))
        self.setSpan(row, 0, 1, self.columnCount())
        self.setItem(row, 0, header_item)
        # Hacemos que la fila de capítulo no se pueda seleccionar, editar ni arrastrar.
        header_item.setFlags(header_item.flags() & ~Qt.ItemIsSelectable & ~Qt.ItemIsEditable & ~Qt.ItemIsDragEnabled)
class MainWindow(QMainWindow):
    def __init__(self, cotizacion_controller, excel_controller):
        super().__init__()

        ### -----------------------------------------------------------------------------------------
        ### 1. INICIALIZACIÓN DE CONTROLADORES Y CONFIGURACIÓN DE LA VENTANA
        ### -----------------------------------------------------------------------------------------
        self.cotizacion_controller = cotizacion_controller
        self.excel_controller = excel_controller
        self.aiu_manager = self.cotizacion_controller.aiu_manager

        self.setWindowTitle("Sistema de Cotizaciones")
        self.setGeometry(100, 100, 1300, 850)

        ### -----------------------------------------------------------------------------------------
        ### 2. CREACIÓN DE WIDGETS Y LAYOUTS PRINCIPALES
        ### Se crea un widget central y un layout vertical que contendrá todo.
        ### -----------------------------------------------------------------------------------------
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        ### -----------------------------------------------------------------------------------------
        ### 3. SECCIÓN SUPERIOR: INFORMACIÓN DEL CLIENTE
        ### -----------------------------------------------------------------------------------------
        client_group = QGroupBox("Información del Cliente")
        client_main_layout = QVBoxLayout(client_group)

        client_selection_layout = QHBoxLayout()
        client_selection_layout.addWidget(QLabel("Cliente Existente:"))
        self.client_combo = QComboBox();
        self.client_combo.setMinimumWidth(300)
        self.client_combo.currentIndexChanged.connect(self.actualizar_datos_cliente)
        client_selection_layout.addWidget(self.client_combo)
        manage_data_btn = QPushButton("Gestionar Datos");
        manage_data_btn.clicked.connect(self.open_data_management)
        client_selection_layout.addWidget(manage_data_btn)
        client_selection_layout.addStretch()
        client_main_layout.addLayout(client_selection_layout)

        client_form_layout = QGridLayout()
        self.tipo_combo = QComboBox();
        self.tipo_combo.addItems(["Natural", "Jurídica"])
        self.nit_input = QLineEdit()
        self.nombre_input = QLineEdit()
        self.telefono_input = QLineEdit()
        self.direccion_input = QLineEdit()
        self.email_input = QLineEdit()
        client_form_layout.addWidget(QLabel("Tipo:"), 0, 0);
        client_form_layout.addWidget(self.tipo_combo, 0, 1)
        client_form_layout.addWidget(QLabel("NIT/CC:"), 0, 2);
        client_form_layout.addWidget(self.nit_input, 0, 3)
        client_form_layout.addWidget(QLabel("Nombre:"), 1, 0);
        client_form_layout.addWidget(self.nombre_input, 1, 1)
        client_form_layout.addWidget(QLabel("Teléfono:"), 1, 2);
        client_form_layout.addWidget(self.telefono_input, 1, 3)
        client_form_layout.addWidget(QLabel("Dirección:"), 2, 0);
        client_form_layout.addWidget(self.direccion_input, 2, 1)
        client_form_layout.addWidget(QLabel("Email:"), 2, 2);
        client_form_layout.addWidget(self.email_input, 2, 3)
        client_form_layout.setColumnStretch(4, 1)
        client_main_layout.addLayout(client_form_layout)

        save_client_btn = QPushButton("Guardar Cliente");
        save_client_btn.clicked.connect(self.guardar_cliente)
        client_main_layout.addWidget(save_client_btn, 0, Qt.AlignLeft)

        main_layout.addWidget(client_group)

        ### -----------------------------------------------------------------------------------------
        ### 4. SECCIÓN CENTRAL: DIVISOR (SPLITTER)
        ### Divide la ventana en un panel izquierdo (tabla) y un panel derecho (herramientas).
        ### -----------------------------------------------------------------------------------------
        main_splitter = QSplitter(Qt.Horizontal)

        # --- 4.1 PANEL IZQUIERDO: TABLA DE COTIZACIÓN Y TOTALES ---
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.addWidget(QLabel("Detalle de la Cotización", styleSheet="font-size: 16px; font-weight: bold;"))

        self.activities_table = DraggableTableWidget(0, 6)
        self.activities_table.setHorizontalHeaderLabels(
            ["Descripción", "Cantidad", "Unidad", "Valor Unitario", "Total", "Acción"])
        self.activities_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.activities_table.itemChanged.connect(self.on_item_changed)
        left_layout.addWidget(self.activities_table)

        totals_layout = QHBoxLayout()
        totals_layout.addStretch()
        totals_layout.addWidget(QLabel("Subtotal:"));
        self.subtotal_label = QLabel("0.00")
        totals_layout.addWidget(self.subtotal_label)
        totals_layout.addWidget(QLabel("IVA:"));
        self.iva_label = QLabel("0.00")
        totals_layout.addWidget(self.iva_label)
        totals_layout.addWidget(QLabel("Total:"));
        self.total_label = QLabel("0.00")
        totals_layout.addWidget(self.total_label)
        left_layout.addLayout(totals_layout)
        main_splitter.addWidget(left_panel)

        # --- 4.2 PANEL DERECHO: HERRAMIENTAS DE SELECCIÓN ---
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.addWidget(QLabel("Herramientas", styleSheet="font-size: 16px; font-weight: bold;"))

        # Grupo para Capítulos
        chapter_group = QGroupBox("Gestión de Capítulos")
        chapter_layout = QHBoxLayout(chapter_group)
        self.chapter_selection_combo = QComboBox()
        add_chapter_header_btn = QPushButton("Insertar Encabezado");
        add_chapter_header_btn.clicked.connect(self.insert_chapter_header)
        chapter_layout.addWidget(QLabel("Capítulo:"));
        chapter_layout.addWidget(self.chapter_selection_combo, 1);
        chapter_layout.addWidget(add_chapter_header_btn)
        right_layout.addWidget(chapter_group)

        # Grupo para Actividades Predefinidas
        activity_group = QGroupBox("Selección de Actividades")
        activity_layout = QVBoxLayout(activity_group)
        self.search_input = QLineEdit();
        self.search_input.setPlaceholderText("Buscar actividades...")
        self.search_input.textChanged.connect(self.filter_activities)
        self.category_combo = QComboBox();
        self.category_combo.currentIndexChanged.connect(self.filter_activities)
        self.activity_combo = QComboBox();
        self.activity_combo.currentIndexChanged.connect(self.update_related_activities)
        self.activity_description = QTextEdit();
        self.activity_description.setReadOnly(True);
        self.activity_description.setMaximumHeight(60)
        self.related_activities_combo = QComboBox()
        self.pred_quantity_spinbox = QDoubleSpinBox();
        self.pred_quantity_spinbox.setRange(0.01, 9999.99);
        self.pred_quantity_spinbox.setValue(1.0)
        add_activity_btn = QPushButton("Agregar Actividad");
        add_activity_btn.clicked.connect(self.add_selected_activity)
        add_related_btn = QPushButton("Agregar Relacionada");
        add_related_btn.clicked.connect(self.add_related_activity)

        activity_layout.addWidget(self.search_input)
        activity_layout.addWidget(self.category_combo)
        activity_layout.addWidget(QLabel("Actividad Predefinida:"));
        activity_layout.addWidget(self.activity_combo)
        activity_layout.addWidget(self.activity_description)
        activity_layout.addWidget(QLabel("Actividades Relacionadas:"));
        activity_layout.addWidget(self.related_activities_combo)

        pred_qty_layout = QHBoxLayout();
        pred_qty_layout.addWidget(QLabel("Cantidad:"));
        pred_qty_layout.addWidget(self.pred_quantity_spinbox)
        activity_layout.addLayout(pred_qty_layout)
        add_btns_layout = QHBoxLayout();
        add_btns_layout.addWidget(add_activity_btn);
        add_btns_layout.addWidget(add_related_btn)
        activity_layout.addLayout(add_btns_layout)
        right_layout.addWidget(activity_group)

        # Grupo para Entrada Manual
        manual_group = QGroupBox("Entrada Manual")
        manual_layout = QFormLayout(manual_group)
        self.description_input = QLineEdit()
        self.quantity_spinbox = QDoubleSpinBox();
        self.quantity_spinbox.setRange(0.01, 9999.99);
        self.quantity_spinbox.setValue(1.0)
        self.unit_input = QLineEdit()
        self.value_spinbox = QDoubleSpinBox();
        self.value_spinbox.setRange(0.01, 9999999.99);
        self.value_spinbox.setValue(1000.0)
        add_manual_btn = QPushButton("Agregar Actividad Manual");
        add_manual_btn.clicked.connect(self.add_manual_activity)
        manual_layout.addRow("Descripción:", self.description_input)
        manual_layout.addRow("Cantidad:", self.quantity_spinbox)
        manual_layout.addRow("Unidad:", self.unit_input)
        manual_layout.addRow("Valor Unitario:", self.value_spinbox)
        manual_layout.addRow(add_manual_btn)
        right_layout.addWidget(manual_group)

        right_layout.addStretch()
        main_splitter.addWidget(right_panel)

        main_splitter.setSizes([800, 500])
        main_layout.addWidget(main_splitter)

        ### -----------------------------------------------------------------------------------------
        ### 5. SECCIÓN INFERIOR: BOTONES DE ACCIÓN
        ### -----------------------------------------------------------------------------------------
        actions_layout = QHBoxLayout()
        save_file_btn = QPushButton("Guardar como Archivo");
        save_file_btn.clicked.connect(self.save_as_file)
        open_file_btn = QPushButton("Abrir Archivo");
        open_file_btn.clicked.connect(self.open_file_dialog)
        generate_excel_btn = QPushButton("Generar Excel");
        generate_excel_btn.clicked.connect(self.generate_excel)
        generate_word_btn = QPushButton("Generar Word");
        generate_word_btn.clicked.connect(self.generate_word)
        send_email_btn = QPushButton("Enviar por Correo");
        send_email_btn.clicked.connect(self.send_email)
        clear_btn = QPushButton("Limpiar");
        clear_btn.clicked.connect(self.clear_form)
        self.theme_btn = QPushButton("Modo Oscuro");
        self.theme_btn.clicked.connect(self.toggle_theme)

        actions_layout.addWidget(save_file_btn);
        actions_layout.addWidget(open_file_btn)
        actions_layout.addWidget(generate_excel_btn);
        actions_layout.addWidget(generate_word_btn)
        actions_layout.addWidget(send_email_btn);
        actions_layout.addWidget(clear_btn)
        actions_layout.addStretch();
        actions_layout.addWidget(self.theme_btn)
        main_layout.addLayout(actions_layout)

        ### -----------------------------------------------------------------------------------------
        ### 6. CARGA DE DATOS INICIALES
        ### -----------------------------------------------------------------------------------------
        self.dark_mode = False
        self.load_initial_data()

    ### =============================================================================================
    ### MÉTODOS DE LÓGICA DE LA INTERFAZ
    ### El resto de tus métodos no necesitan cambios.
    ### =============================================================================================

    def load_initial_data(self):
        """Carga los datos iniciales en la interfaz"""
        try:
            self.refresh_client_combo()
            self.refresh_category_combo()
            self.load_chapters_into_selection_combo()
            self.refresh_activity_combo()
        except Exception as e:
            import traceback
            traceback.print_exc()
            QMessageBox.critical(self, "Error Fatal al Cargar Datos",
                                 f"No se pudieron cargar los datos iniciales de la aplicación.\n\n"
                                 f"Error: {e}\n\n"
                                 "Por favor, revise la consola para más detalles.")

    def load_chapters_into_selection_combo(self):
        """Carga los capítulos en el ComboBox de selección."""
        try:
            self.chapter_selection_combo.clear()
            self.chapter_selection_combo.addItem("--- Seleccione un Capítulo ---", None)
            chapters = self.cotizacion_controller.get_all_chapters()
            for chapter in chapters:
                self.chapter_selection_combo.addItem(chapter['nombre'], chapter['id'])
        except Exception as e:
            print(f"Error al cargar capítulos en el combo de selección: {e}")

    def insert_chapter_header(self):
        """Obtiene los datos del capítulo y le ordena a la tabla que inserte el encabezado."""
        chapter_id = self.chapter_selection_combo.currentData()
        chapter_name = self.chapter_selection_combo.currentText()
        if not chapter_id:
            QMessageBox.warning(self, "Sin Selección", "Por favor, seleccione un capítulo para insertar.")
            return
        self.activities_table.insert_chapter_header(chapter_id, chapter_name)

    def add_activity_to_table(self, descripcion, cantidad, unidad, valor_unitario):
        """Agrega una fila de actividad normal a la tabla."""
        row = self.activities_table.currentRow()
        if row == -1:
            row = self.activities_table.rowCount()
        else:
            row += 1

        self.activities_table.insertRow(row)
        desc_item = EditableTableWidgetItem(descripcion)
        desc_item.setData(Qt.UserRole, {'type': 'activity'})
        self.activities_table.setItem(row, 0, desc_item)
        self.activities_table.setItem(row, 1, EditableTableWidgetItem(cantidad))
        self.activities_table.setItem(row, 2, EditableTableWidgetItem(unidad))
        self.activities_table.setItem(row, 3, EditableTableWidgetItem(valor_unitario))
        total = float(cantidad) * float(valor_unitario)
        self.activities_table.setItem(row, 4, EditableTableWidgetItem(f"{total:.2f}", editable=False))
        self.reconnect_delete_button(row)
        self.update_totals()

    def reconnect_delete_button(self, row):
        """Crea y conecta el botón de eliminar para una fila específica."""
        delete_btn = QPushButton("Eliminar")
        delete_btn.clicked.connect(lambda _, r=row: self.delete_activity(r))
        self.activities_table.setCellWidget(row, 5, delete_btn)

    def delete_activity(self, row):
        """Elimina una fila de la tabla y re-conecta los botones."""
        self.activities_table.removeRow(row)
        for r in range(self.activities_table.rowCount()):
            first_item = self.activities_table.item(r, 0)
            if first_item and first_item.data(Qt.UserRole) and first_item.data(Qt.UserRole).get('type') == 'activity':
                self.reconnect_delete_button(r)
        self.update_totals()

    def on_item_changed(self, item):
        """Actualiza el total de una fila cuando se edita la cantidad o el valor."""
        row = item.row()
        first_item = self.activities_table.item(row, 0)
        if not first_item or not first_item.data(Qt.UserRole) or first_item.data(Qt.UserRole)['type'] != 'activity':
            return
        if item.column() in [1, 3]:
            try:
                cantidad = float(self.activities_table.item(row, 1).text())
                valor_unitario = float(self.activities_table.item(row, 3).text())
                total = cantidad * valor_unitario
                self.activities_table.itemChanged.disconnect(self.on_item_changed)
                self.activities_table.setItem(row, 4, EditableTableWidgetItem(f"{total:.2f}", editable=False))
                self.activities_table.itemChanged.connect(self.on_item_changed)
                self.update_totals()
            except (ValueError, AttributeError):
                pass

    def update_totals(self):
        """Actualiza los totales, ignorando las filas de capítulo."""
        subtotal = 0
        for row in range(self.activities_table.rowCount()):
            first_item = self.activities_table.item(row, 0)
            if first_item and first_item.data(Qt.UserRole) and first_item.data(Qt.UserRole).get('type') == 'activity':
                try:
                    total_item = self.activities_table.item(row, 4)
                    if total_item and total_item.text():
                        subtotal += float(total_item.text())
                except (ValueError, AttributeError):
                    pass

        aiu_values = self.aiu_manager.get_aiu_values()
        if not aiu_values: aiu_values = {'utilidad': 0, 'iva_sobre_utilidad': 19.0}

        iva = 0
        if self.tipo_combo.currentText().lower() == "jurídica":
            utilidad = subtotal * (aiu_values.get('utilidad', 0) / 100)
            iva = utilidad * (aiu_values.get('iva_sobre_utilidad', 0) / 100)
        else:
            iva = subtotal * (aiu_values.get('iva_sobre_utilidad', 19.0) / 100)

        total = subtotal + iva
        self.subtotal_label.setText(f"${subtotal:,.2f}")
        self.iva_label.setText(f"${iva:,.2f}")
        self.total_label.setText(f"${total:,.2f}")

    def generate_excel(self, show_message=True):
        """Prepara los datos y llama al controlador de Excel."""
        structured_items = []
        for row in range(self.activities_table.rowCount()):
            first_item = self.activities_table.item(row, 0)
            if not first_item or not first_item.data(Qt.UserRole):
                continue

            metadata = first_item.data(Qt.UserRole)
            if metadata['type'] == 'chapter':
                structured_items.append({'type': 'chapter', 'name': first_item.text()})
            elif metadata['type'] == 'activity':
                structured_items.append({
                    'type': 'activity',
                    'descripcion': first_item.text(),
                    'cantidad': float(self.activities_table.item(row, 1).text()),
                    'unidad': self.activities_table.item(row, 2).text(),
                    'valor_unitario': float(self.activities_table.item(row, 3).text()),
                })

        if not any(item['type'] == 'activity' for item in structured_items):
            QMessageBox.warning(self, "Vacío", "No hay actividades en la cotización para generar el Excel.")
            return

        # Obtener valores AIU
        aiu_values = self.aiu_manager.get_aiu_values()

        # Crear controlador de Excel
        try:
            excel_path = self.excel_controller.generate_excel(
                items=structured_items,
                activities=structured_items,  # <-- Tal vez querías pasar structured_items aquí
                tipo_persona=self.tipo_combo.currentText().lower(),
                administracion=aiu_values['administracion'],
                imprevistos=aiu_values['imprevistos'],
                utilidad=aiu_values['utilidad'],
                iva_utilidad=aiu_values['iva_sobre_utilidad']
            )
            if show_message:
                QMessageBox.information(self, "Éxito", f"Archivo Excel generado en:\n{excel_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo generar el Excel: {e}")
            return

        return excel_path

    def generate_word(self):
        """Genera un documento Word con la cotización actual"""
        try:
            # Verificar que haya actividades
            if self.activities_table.rowCount() == 0:
                QMessageBox.warning(self, "Error", "Debe agregar al menos una actividad a la cotización.")
                return

            # Verificar que haya un cliente seleccionado
            if self.client_combo.currentIndex() == -1:
                QMessageBox.warning(self, "Error", "Debe seleccionar un cliente.")
                return

            # Generar Excel primero (necesario para la imagen)
            excel_path = self.generate_excel(show_message=False)
            if not excel_path:
                return

            # Obtener datos del cliente
            client_id = self.client_combo.currentData()
            client = self.cotizacion_controller.database_manager.get_client_by_id(client_id)

            # Abrir diálogo para configurar documento Word
            dialog = WordConfigDialog(self, client['tipo'])
            if dialog.exec_():
                config = dialog.get_config()

                # Crear controlador de Word
                word_controller = WordController()

                # Generar Word
                word_path = word_controller.generate_word(
                    excel_path=excel_path,
                    client_data=client,
                    config=config
                )

                QMessageBox.information(self, "Éxito", f"Documento Word generado correctamente en:\n{word_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al generar el Word: {str(e)}")

    def clear_form(self):
        """Limpia el formulario"""
        try:
            # Limpiar tabla de actividades
            self.activities_table.setRowCount(0)

            # Limpiar totales
            self.subtotal_label.setText("0.00")
            self.iva_label.setText("0.00")
            self.total_label.setText("0.00")

            # Limpiar campos de búsqueda
            self.search_input.clear()
            self.category_combo.setCurrentIndex(0)

            # Limpiar campos de actividad manual
            self.description_input.clear()
            self.quantity_spinbox.setValue(1.0)
            self.unit_input.clear()
            self.value_spinbox.setValue(1000.0)

            # Actualizar combos
            self.refresh_activity_combo()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al limpiar formulario: {str(e)}")

    def save_as_file(self):
        """Guarda la cotización actual como un archivo"""
        try:
            # Verificar que haya actividades
            if self.activities_table.rowCount() == 0:
                QMessageBox.warning(self, "Error", "Debe agregar al menos una actividad a la cotización.")
                return

            # Verificar que haya un cliente seleccionado
            if self.client_combo.currentIndex() == -1:
                QMessageBox.warning(self, "Error", "Debe seleccionar un cliente.")
                return

            # Obtener datos del cliente
            client_id = self.client_combo.currentData()
            client_data = self.cotizacion_controller.database_manager.get_client_by_id(client_id)

            # Obtener actividades de la tabla
            activities = []
            for row in range(self.activities_table.rowCount()):
                activity = {
                    'descripcion': self.activities_table.item(row, 0).text(),
                    'cantidad': float(self.activities_table.item(row, 1).text()),
                    'unidad': self.activities_table.item(row, 2).text(),
                    'valor_unitario': float(self.activities_table.item(row, 3).text()),
                    'total': float(self.activities_table.item(row, 4).text()),
                    'id': row + 1  # Asignar un ID único basado en el índice
                }
                activities.append(activity)

            # Obtener valores AIU
            aiu_values = self.aiu_manager.get_aiu_values()

            # Calcular totales
            subtotal = sum(activity['total'] for activity in activities)
            administracion = subtotal * (aiu_values['administracion'] / 100)
            imprevistos = subtotal * (aiu_values['imprevistos'] / 100)
            utilidad = subtotal * (aiu_values['utilidad'] / 100)
            iva_utilidad = utilidad * (aiu_values['iva_sobre_utilidad'] / 100)
            total = subtotal + administracion + imprevistos + utilidad + iva_utilidad

            # Crear datos de la cotización
            cotizacion_data = {
                'fecha': datetime.now().strftime('%Y-%m-%d'),
                'cliente': client_data,
                'actividades': activities,
                'subtotal': subtotal,
                'administracion': administracion,
                'imprevistos': imprevistos,
                'utilidad': utilidad,
                'iva_utilidad': iva_utilidad,
                'total': total,
                'aiu_values': aiu_values
            }

            # Abrir diálogo para guardar archivo
            file_path, _ = QFileDialog.getSaveFileName(self, "Guardar Cotización", "", "Archivos de Cotización (*.cotiz);;Todos los archivos (*)")

            if file_path:
                # Asegurar que tenga la extensión correcta
                if not file_path.endswith('.cotiz'):
                    file_path += '.cotiz'

                # Guardar archivo
                self.cotizacion_controller.file_manager.guardar_cotizacion(cotizacion_data, file_path)
                QMessageBox.information(self, "Éxito", f"Cotización guardada correctamente en:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al guardar archivo: {str(e)}")

    def open_file_dialog(self):
        """Abre el diálogo para seleccionar un archivo de cotización"""
        try:
            dialog = CotizacionFileDialog(self.cotizacion_controller, self)
            if dialog.exec_():
                cotizacion_data = dialog.get_selected_cotizacion()
                if cotizacion_data:
                    self.load_cotizacion_from_file(cotizacion_data)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al abrir archivo: {str(e)}")

    def load_cotizacion_from_file(self, cotizacion_data):
        """Carga una cotización desde un archivo"""
        try:
            # Limpiar formulario
            self.clear_form()

            # Cargar datos del cliente
            client_data = cotizacion_data['cliente']

            # Buscar cliente en la base de datos o crearlo si no existe
            clients = self.cotizacion_controller.get_all_clients()
            client_exists = False
            client_id = None

            for client in clients:
                if client['nombre'] == client_data['nombre'] and client['tipo'] == client_data['tipo']:
                    client_exists = True
                    client_id = client['id']
                    break

            if not client_exists:
                # Crear nuevo cliente
                client_id = self.cotizacion_controller.add_client(client_data)
                self.refresh_client_combo()

            # Seleccionar cliente en el combo
            index = self.client_combo.findData(client_id)
            if index >= 0:
                self.client_combo.setCurrentIndex(index)

            # Cargar actividades
            for activity in cotizacion_data['actividades']:
                row = self.activities_table.rowCount()
                self.activities_table.insertRow(row)

                # Descripción
                self.activities_table.setItem(row, 0, EditableTableWidgetItem(activity['descripcion']))

                # Cantidad
                self.activities_table.setItem(row, 1, EditableTableWidgetItem(activity['cantidad']))

                # Unidad
                self.activities_table.setItem(row, 2, EditableTableWidgetItem(activity['unidad']))

                # Valor unitario
                self.activities_table.setItem(row, 3, EditableTableWidgetItem(activity['valor_unitario']))

                # Total
                total = activity['cantidad'] * activity['valor_unitario']
                self.activities_table.setItem(row, 4, EditableTableWidgetItem(total, False))

                # Botón eliminar
                delete_btn = QPushButton("Eliminar")
                delete_btn.clicked.connect(lambda _, r=row: self.delete_activity(r))
                self.activities_table.setCellWidget(row, 5, delete_btn)

            # Actualizar totales
            self.update_totals()

            QMessageBox.information(self, "Éxito", "Cotización cargada correctamente.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al cargar cotización: {str(e)}")

    def send_email(self):
        """Abre el diálogo para enviar cotización por correo"""
        # Verificar si hay archivos para enviar
        attachments = []

        # Si hay un Excel generado, añadirlo
        if hasattr(self, 'last_excel_path') and self.last_excel_path and os.path.exists(self.last_excel_path):
            attachments.append(self.last_excel_path)

        # Si hay un Word generado, añadirlo
        if hasattr(self, 'last_word_path') and self.last_word_path and os.path.exists(self.last_word_path):
            attachments.append(self.last_word_path)

        # Si no hay archivos, preguntar si desea generar Excel
        if not attachments:
            reply = QMessageBox.question(self, "Sin Archivos",
                                         "No hay archivos para enviar. ¿Desea generar un Excel primero?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)
            if reply == QMessageBox.Yes:
                excel_path = self.generate_excel(show_message=False)
                if excel_path:
                    attachments.append(excel_path)
                else:
                    return
            else:
                return

        # Abrir diálogo de envío de correo
        email_dialog = SendEmailDialog(attachments, self)
        email_dialog.exec_()


    def open_data_management(self):
        """Abre la ventana de gestión de datos"""
        try:
            self.data_management_window = DataManagementWindow(self.cotizacion_controller, self)
            self.data_management_window.closed.connect(self.on_data_management_closed)
            self.data_management_window.show()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al abrir la ventana de gestión de datos: {str(e)}")

    def actualizar_datos_cliente(self):
        """
        Se activa cuando se selecciona un cliente del ComboBox.
        Obtiene los datos del cliente y rellena el formulario.
        """
        try:
            # Obtener el ID del cliente seleccionado, que guardamos con setData
            client_id = self.client_combo.currentData()

            # Si no hay un cliente seleccionado (ej. "Seleccione un cliente"), no hacer nada
            if not client_id:
                # Opcional: Limpiar el formulario si se deselecciona
                self.nombre_input.clear()
                self.direccion_input.clear()
                self.nit_input.clear()
                self.telefono_input.clear()
                self.email_input.clear()
                return

            # Usar el controlador para obtener los datos del cliente desde la base de datos
            client = self.cotizacion_controller.database_manager.get_client_by_id(client_id)

            if client:
                # Rellenar los campos del formulario con los datos obtenidos
                self.tipo_combo.setCurrentText(client.get('tipo', 'Natural').capitalize())
                self.nombre_input.setText(client.get('nombre', ''))
                self.direccion_input.setText(client.get('direccion', ''))
                self.nit_input.setText(client.get('nit', ''))
                self.telefono_input.setText(client.get('telefono', ''))
                self.email_input.setText(client.get('email', ''))
        except Exception as e:
            print(f"Error al actualizar datos del cliente: {e}")
            import traceback
            traceback.print_exc()

        # Pega este método dentro de la clase MainWindow en views/main_window.py

    def guardar_cliente(self):
        """
        Recolecta los datos del formulario del cliente y los guarda en la base de datos.
        Actualiza un cliente existente si hay uno seleccionado, o crea uno nuevo si no.
        """
        try:
            # Recolectar los datos de los campos de la interfaz
            client_data = {
                'tipo': self.tipo_combo.currentText().lower(),
                'nombre': self.nombre_input.text(),
                'direccion': self.direccion_input.text(),
                'nit': self.nit_input.text(),
                'telefono': self.telefono_input.text(),
                'email': self.email_input.text()
            }

            # Validar que el nombre no esté vacío
            if not client_data['nombre']:
                QMessageBox.warning(self, "Dato Requerido", "El nombre del cliente es obligatorio.")
                return

            # Determinar si estamos actualizando o creando
            # Si hay un cliente seleccionado en el combo, usamos su ID para actualizar
            client_id = self.client_combo.currentData()

            if client_id:
                # Actualizar el cliente existente
                self.cotizacion_controller.database_manager.update_client(client_id, client_data)
                QMessageBox.information(self, "Éxito", "Cliente actualizado correctamente.")
            else:
                # Agregar un nuevo cliente
                # El método add_client del controlador devuelve el ID del nuevo cliente
                client_id = self.cotizacion_controller.add_client(client_data)
                QMessageBox.information(self, "Éxito", "Cliente agregado correctamente.")

            # Refrescar la lista de clientes para que se muestren los cambios
            self.refresh_client_combo()

            # Opcional: Seleccionar automáticamente el cliente que acabamos de guardar
            if client_id:
                index = self.client_combo.findData(client_id)
                if index >= 0:
                    self.client_combo.setCurrentIndex(index)

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error al guardar el cliente: {e}")
            import traceback
            traceback.print_exc()

        # Pega este método dentro de la clase MainWindow en views/main_window.py

    def filter_activities(self):
            """
            Filtra la lista de actividades predefinidas (en el QComboBox)
            basándose en el texto de búsqueda y la categoría seleccionada.
            """
            try:
                # 1. Obtener los valores de los filtros de la interfaz
                search_text = self.search_input.text()
                category_id = self.category_combo.currentData()

                # 2. Usar el controlador para obtener la lista de actividades filtrada
                activities = self.cotizacion_controller.search_activities(search_text, category_id)

                # 3. Actualizar el ComboBox de actividades con los resultados

                # Guardar la selección actual para intentar restaurarla después
                current_activity_id = self.activity_combo.currentData()

                # Desconectar temporalmente la señal para evitar llamadas en cascada
                self.activity_combo.currentIndexChanged.disconnect(self.update_related_activities)

                self.activity_combo.clear()
                for activity in activities:
                    self.activity_combo.addItem(activity['descripcion'], activity['id'])

                # Intentar restaurar la selección anterior
                if current_activity_id:
                    index = self.activity_combo.findData(current_activity_id)
                    if index >= 0:
                        self.activity_combo.setCurrentIndex(index)

                # Volver a conectar la señal
                self.activity_combo.currentIndexChanged.connect(self.update_related_activities)

                # Forzar la actualización de la descripción y las actividades relacionadas
                # para la nueva lista filtrada.
                self.update_related_activities()

            except Exception as e:
                print(f"Error al filtrar actividades: {e}")
                import traceback
                traceback.print_exc()
            # Pega este método dentro de la clase MainWindow en views/main_window.py

    def update_related_activities(self):
        """
        Se activa cuando se selecciona una actividad principal.
        Actualiza la descripción de la actividad y la lista de actividades relacionadas.
        """
        try:
            # --- 1. Actualizar la descripción de la actividad principal ---
            activity_id = self.activity_combo.currentData()
            if activity_id:
                activity = self.cotizacion_controller.database_manager.get_activity_by_id(activity_id)
                if activity:
                    desc_text = (f"Descripción: {activity.get('descripcion', '')}\n"
                                 f"Unidad: {activity.get('unidad', '')}\n"
                                 f"Valor: ${activity.get('valor_unitario', 0):,.2f}")
                    self.activity_description.setText(desc_text)
                else:
                    self.activity_description.clear()
            else:
                self.activity_description.clear()

            # --- 2. Actualizar el ComboBox de actividades relacionadas ---
            self.related_activities_combo.clear()

            if activity_id:
                related_activities = self.cotizacion_controller.get_related_activities(activity_id)
                for rel_activity in related_activities:
                    self.related_activities_combo.addItem(rel_activity['descripcion'], rel_activity['id'])

        except Exception as e:
            print(f"Error al actualizar actividades relacionadas: {e}")
            import traceback
            traceback.print_exc()

        # Pega este método dentro de la clase MainWindow en views/main_window.py

    def add_selected_activity(self):
            """
            Toma la actividad predefinida seleccionada en el ComboBox y la
            añade a la tabla de la cotización.
            """
            try:
                # 1. Obtener el ID de la actividad seleccionada
                activity_id = self.activity_combo.currentData()

                if not activity_id:
                    QMessageBox.warning(self, "Sin Selección", "Por favor, seleccione una actividad de la lista.")
                    return

                # 2. Obtener los datos completos de la actividad desde la base de datos
                activity = self.cotizacion_controller.database_manager.get_activity_by_id(activity_id)

                if not activity:
                    QMessageBox.critical(self, "Error",
                                         "No se pudieron encontrar los datos de la actividad seleccionada.")
                    return

                # 3. Obtener la cantidad del SpinBox
                cantidad = self.pred_quantity_spinbox.value()

                # 4. Llamar al método que realmente añade la fila a la tabla
                self.add_activity_to_table(
                    descripcion=activity['descripcion'],
                    cantidad=cantidad,
                    unidad=activity['unidad'],
                    valor_unitario=activity['valor_unitario']
                )
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Ocurrió un error al agregar la actividad: {e}")
                import traceback
                traceback.print_exc()

    def add_related_activity(self):
        """
        Toma la actividad seleccionada en el ComboBox de actividades RELACIONADAS
        y la añade a la tabla de la cotización.
        """
        try:
            # 1. Obtener el ID de la actividad del ComboBox de relacionadas
            activity_id = self.related_activities_combo.currentData()

            if not activity_id:
                QMessageBox.warning(self, "Sin Selección",
                                    "Por favor, seleccione una actividad relacionada de la lista.")
                return

            # 2. Obtener los datos completos de la actividad desde la base de datos
            activity = self.cotizacion_controller.database_manager.get_activity_by_id(activity_id)

            if not activity:
                QMessageBox.critical(self, "Error", "No se pudieron encontrar los datos de la actividad relacionada.")
                return

            # 3. Obtener la cantidad del SpinBox (usa el mismo que la actividad principal)
            cantidad = self.pred_quantity_spinbox.value()

            # 4. Llamar al método que realmente añade la fila a la tabla
            self.add_activity_to_table(
                descripcion=activity['descripcion'],
                cantidad=cantidad,
                unidad=activity['unidad'],
                valor_unitario=activity['valor_unitario']
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un error al agregar la actividad relacionada: {e}")
            import traceback
            traceback.print_exc()

        # Pega este método dentro de la clase MainWindow en views/main_window.py

    def add_manual_activity(self):
            """
            Recolecta los datos de los campos de entrada manual y los
            añade como una nueva actividad en la tabla de la cotización.
            """
            try:
                # 1. Obtener los datos de los campos del formulario manual
                descripcion = self.description_input.text()
                cantidad = self.quantity_spinbox.value()
                unidad = self.unit_input.text()
                valor_unitario = self.value_spinbox.value()

                # 2. Validar que los campos obligatorios no estén vacíos
                if not descripcion:
                    QMessageBox.warning(self, "Dato Requerido", "La descripción de la actividad es obligatoria.")
                    return
                if not unidad:
                    QMessageBox.warning(self, "Dato Requerido", "La unidad de la actividad es obligatoria.")
                    return

                # 3. Llamar al método que realmente añade la fila a la tabla
                self.add_activity_to_table(
                    descripcion=descripcion,
                    cantidad=cantidad,
                    unidad=unidad,
                    valor_unitario=valor_unitario
                )

                # 4. Opcional: Limpiar los campos después de agregar la actividad
                self.description_input.clear()
                self.quantity_spinbox.setValue(1.0)
                self.unit_input.clear()
                self.value_spinbox.setValue(1000.0)

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Ocurrió un error al agregar la actividad manual: {e}")
                import traceback
                traceback.print_exc()

    def toggle_theme(self):
        """
        Alterna entre el tema claro y el oscuro de la aplicación.
        """
        self.dark_mode = not self.dark_mode  # Invierte el estado actual
        if self.dark_mode:
            self.theme_btn.setText("Modo Claro")
            self.apply_dark_theme()
        else:
            self.theme_btn.setText("Modo Oscuro")
            self.apply_light_theme()

    def apply_dark_theme(self):
        """Aplica una paleta de colores oscuros a la aplicación."""
        dark_palette = QPalette()
        dark_palette.setColor(QPalette.Window, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.WindowText, Qt.white)
        dark_palette.setColor(QPalette.Base, QColor(25, 25, 25))
        dark_palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ToolTipBase, Qt.white)
        dark_palette.setColor(QPalette.ToolTipText, Qt.white)
        dark_palette.setColor(QPalette.Text, Qt.white)
        dark_palette.setColor(QPalette.Button, QColor(53, 53, 53))
        dark_palette.setColor(QPalette.ButtonText, Qt.white)
        dark_palette.setColor(QPalette.BrightText, Qt.red)
        dark_palette.setColor(QPalette.Link, QColor(42, 130, 218))
        dark_palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
        dark_palette.setColor(QPalette.HighlightedText, Qt.black)

        QApplication.instance().setPalette(dark_palette)

    def apply_light_theme(self):
        """Restaura la paleta de colores por defecto de la aplicación."""
        # Simplemente creamos una nueva paleta vacía para que Qt use la del sistema
        QApplication.instance().setPalette(QPalette())

    def refresh_client_combo(self):
        """
        Limpia y vuelve a llenar el ComboBox de clientes con los datos
        actualizados de la base de datos.
        """
        try:
            # Guardar el ID del cliente que está seleccionado actualmente
            current_client_id = self.client_combo.currentData()

            # Desconectar la señal temporalmente para evitar que se dispare
            # el método 'actualizar_datos_cliente' mientras llenamos la lista.
            self.client_combo.currentIndexChanged.disconnect(self.actualizar_datos_cliente)

            self.client_combo.clear()
            self.client_combo.addItem("--- Seleccione un Cliente ---", None)

            # Obtener todos los clientes desde el controlador
            clients = self.cotizacion_controller.get_all_clients()
            for client in clients:
                self.client_combo.addItem(client['nombre'], client['id'])

            # Intentar restaurar la selección anterior
            if current_client_id:
                index = self.client_combo.findData(current_client_id)
                if index >= 0:
                    self.client_combo.setCurrentIndex(index)

            # Volver a conectar la señal una vez que la lista está llena
            self.client_combo.currentIndexChanged.connect(self.actualizar_datos_cliente)

        except Exception as e:
            print(f"Error al refrescar el combo de clientes: {e}")
            import traceback
            traceback.print_exc()

    def refresh_category_combo(self):
        """
        Limpia y vuelve a llenar el ComboBox de categorías con los datos
        actualizados de la base de datos.
        """
        try:
            # Guardar la selección actual para intentar restaurarla
            current_category_id = self.category_combo.currentData()

            # Desconectar la señal para evitar llamadas innecesarias
            self.category_combo.currentIndexChanged.disconnect(self.filter_activities)

            self.category_combo.clear()
            self.category_combo.addItem("Todas las Categorías", None)

            # Obtener todas las categorías desde el controlador
            categories = self.cotizacion_controller.get_all_categories()
            for category in categories:
                self.category_combo.addItem(category['nombre'], category['id'])

            # Intentar restaurar la selección
            if current_category_id:
                index = self.category_combo.findData(current_category_id)
                if index >= 0:
                    self.category_combo.setCurrentIndex(index)

            # Volver a conectar la señal
            self.category_combo.currentIndexChanged.connect(self.filter_activities)

        except Exception as e:
            print(f"Error al refrescar el combo de categorías: {e}")
            import traceback
            traceback.print_exc()

    def refresh_activity_combo(self):
        """
        Limpia y vuelve a llenar el ComboBox de actividades con los datos
        actualizados de la base de datos.
        """
        try:
            # Guardar la selección actual para intentar restaurarla
            current_activity_id = self.activity_combo.currentData()

            # Desconectar la señal para evitar llamadas en cascada
            self.activity_combo.currentIndexChanged.disconnect(self.update_related_activities)

            self.activity_combo.clear()
            self.activity_combo.addItem("--- Seleccione una Actividad ---", None)

            # Obtener todas las actividades desde el controlador
            activities = self.cotizacion_controller.get_all_activities()
            for activity in activities:
                self.activity_combo.addItem(activity['descripcion'], activity['id'])

            # Intentar restaurar la selección
            if current_activity_id:
                index = self.activity_combo.findData(current_activity_id)
                if index >= 0:
                    self.activity_combo.setCurrentIndex(index)

            # Volver a conectar la señal
            self.activity_combo.currentIndexChanged.connect(self.update_related_activities)

            # Forzar una actualización inicial de la descripción y relacionadas
            self.update_related_activities()

        except Exception as e:
            print(f"Error al refrescar el combo de actividades: {e}")
            import traceback
            traceback.print_exc()

    @pyqtSlot()
    def on_data_management_closed(self):
        """Se ejecuta cuando se cierra la ventana de gestión de datos"""
        # Actualizar combos
        self.refresh_client_combo()
        self.refresh_category_combo()
        self.refresh_activity_combo()