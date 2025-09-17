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
from views.word_dialog import ImprovedWordConfigDialog
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


        self.cotizacion_controller = cotizacion_controller
        self.excel_controller = excel_controller
        self.aiu_manager = self.cotizacion_controller.aiu_manager

        self.setWindowTitle("Sistema de Cotizaciones")
        self.setGeometry(100, 100, 1300, 850)


        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

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
        self.tipo_combo.addItems(["Natural", "Juridica"])
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

        # Priorizar datos de cotización importada si están disponibles
        if hasattr(self, 'selected_cotizacion') and self.selected_cotizacion:
            print("Usando datos de cotización importada")  # Debug
            cotizacion = self.selected_cotizacion

            # Usar table_rows si existe, sino usar actividades (compatibilidad)
            if 'table_rows' in cotizacion:
                for row_data in cotizacion['table_rows']:
                    if row_data['type'] == 'chapter_header':
                        structured_items.append({
                            'type': 'chapter',
                            'name': row_data['descripcion']
                        })
                    elif row_data['type'] == 'activity':
                        structured_items.append({
                            'type': 'activity',
                            'descripcion': row_data['descripcion'],
                            'cantidad': float(row_data.get('cantidad', 0)),
                            'unidad': row_data.get('unidad', ''),
                            'valor_unitario': float(row_data.get('valor_unitario', 0)),
                        })
            elif 'actividades' in cotizacion:
                # Formato anterior por compatibilidad
                for activity in cotizacion['actividades']:
                    structured_items.append({
                        'type': 'activity',
                        'descripcion': activity.get('descripcion', ''),
                        'cantidad': float(activity.get('cantidad', 0)),
                        'unidad': activity.get('unidad', ''),
                        'valor_unitario': float(activity.get('valor_unitario', 0)),
                    })
        else:
            # Usar datos de la tabla de la interfaz (comportamiento original)
            print("Usando datos de la tabla de interfaz")  # Debug
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

        print(f"Total de items estructurados: {len(structured_items)}")  # Debug
        print(f"Actividades encontradas: {sum(1 for item in structured_items if item['type'] == 'activity')}")  # Debug

        if not any(item['type'] == 'activity' for item in structured_items):
            QMessageBox.warning(self, "Vacío", "No hay actividades en la cotización para generar el Excel.")
            return

        # Obtener valores AIU
        aiu_values = self.aiu_manager.get_aiu_values()

        # Crear controlador de Excel
        try:
            excel_path = self.excel_controller.generate_excel(
                items=structured_items,
                activities=structured_items,
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
            print(f"Error detallado: {e}")  # Debug
            import traceback
            traceback.print_exc()  # Debug
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

            # Obtener datos del cliente directamente de los campos de la interfaz
            client_type = self.tipo_combo.currentText().lower()
            client_data = {
                'tipo': client_type,
                'nombre': self.nombre_input.text(),
                'nit': self.nit_input.text(),
                'direccion': self.direccion_input.text(),
                'telefono': self.telefono_input.text(),
                'email': self.email_input.text()
            }

            # Crear datos precargados para el diálogo
            precarga_data = {
                'referencia': f"Cotización para {client_data['nombre']}",
                'titulo': f"COTIZACIÓN DE SERVICIOS PARA {client_data['nombre'].upper()}",
                'lugar': client_data['direccion'],
                'concepto': "Servicios especializados según especificaciones técnicas."
            }

            # Abrir diálogo para configurar documento Word con datos precargados
            dialog = ImprovedWordConfigDialog(
                self,
                client_type=client_type,
                client_data=client_data,
                precarga_data=precarga_data
            )

            if dialog.exec_():
                config = dialog.get_config()

                # Crear controlador de Word
                word_controller = WordController(self.cotizacion_controller)

                # Generar Word según el tipo de cliente
                if config["client_type"] == "natural":
                    word_path = word_controller.generate_natural_cotizacion(config, excel_path)
                elif config["client_type"] == "juridica":
                    word_path = word_controller.generate_juridica_cotizacion(config, excel_path)
                else:
                    QMessageBox.warning(self, "Error", "Tipo de cliente no reconocido en la configuración del diálogo.")
                    return

                QMessageBox.information(self, "Éxito", f"Documento Word generado correctamente en:\n{word_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al generar el Word: {str(e)}")

    def load_imported_cotizacion_to_ui(self, cotizacion_data):
        """
        Carga los datos de una cotización importada en la interfaz de usuario

        Args:
            cotizacion_data (dict): Datos de la cotización importada
        """
        try:
            # Limpiar tabla actual
            self.activities_table.setRowCount(0)

            # Cargar información del cliente si existe
            if 'cliente' in cotizacion_data and cotizacion_data['cliente']:
                cliente = cotizacion_data['cliente']
                # Aquí puedes actualizar los campos del cliente en tu interfaz
                # Por ejemplo, si tienes campos como self.client_name, self.client_nit, etc.

            # Cargar actividades en la tabla
            if 'table_rows' in cotizacion_data:
                for row_data in cotizacion_data['table_rows']:
                    self.add_imported_row_to_table(row_data)
            elif 'actividades' in cotizacion_data:
                # Formato anterior por compatibilidad
                for activity in cotizacion_data['actividades']:
                    self.add_imported_activity_to_table(activity)

            # Actualizar totales y otros cálculos
            self.calculate_totals()

            print("Datos cargados en la interfaz correctamente")

        except Exception as e:
            print(f"Error cargando datos en la interfaz: {e}")
            import traceback
            traceback.print_exc()

    def add_imported_row_to_table(self, row_data):
        """
        Añade una fila importada a la tabla de actividades

        Args:
            row_data (dict): Datos de la fila (chapter_header o activity)
        """
        row_position = self.activities_table.rowCount()
        self.activities_table.insertRow(row_position)

        if row_data['type'] == 'chapter_header':
            # Crear ítem de capítulo
            item = QTableWidgetItem(row_data['descripcion'])
            item.setData(Qt.UserRole, {'type': 'chapter'})
            item.setBackground(QColor(200, 200, 200))  # Color de fondo para capítulos
            self.activities_table.setItem(row_position, 0, item)

            # Llenar celdas vacías para el capítulo
            for col in range(1, self.activities_table.columnCount()):
                empty_item = QTableWidgetItem("")
                empty_item.setBackground(QColor(200, 200, 200))
                self.activities_table.setItem(row_position, col, empty_item)

        elif row_data['type'] == 'activity':
            # Crear ítem de actividad
            desc_item = QTableWidgetItem(row_data['descripcion'])
            desc_item.setData(Qt.UserRole, {'type': 'activity'})
            self.activities_table.setItem(row_position, 0, desc_item)

            # Cantidad
            self.activities_table.setItem(row_position, 1,
                                          QTableWidgetItem(str(row_data.get('cantidad', 0))))

            # Unidad
            self.activities_table.setItem(row_position, 2,
                                          QTableWidgetItem(row_data.get('unidad', '')))

            # Valor unitario
            self.activities_table.setItem(row_position, 3,
                                          QTableWidgetItem(str(row_data.get('valor_unitario', 0))))

            # Total (calculado)
            cantidad = float(row_data.get('cantidad', 0))
            valor_unitario = float(row_data.get('valor_unitario', 0))
            total = cantidad * valor_unitario
            self.activities_table.setItem(row_position, 4,
                                          QTableWidgetItem(str(total)))

    def add_imported_activity_to_table(self, activity_data):
        """
        Añade una actividad importada (formato anterior) a la tabla

        Args:
            activity_data (dict): Datos de la actividad
        """
        row_position = self.activities_table.rowCount()
        self.activities_table.insertRow(row_position)

        # Descripción
        desc_item = QTableWidgetItem(activity_data.get('descripcion', ''))
        desc_item.setData(Qt.UserRole, {'type': 'activity'})
        self.activities_table.setItem(row_position, 0, desc_item)

        # Cantidad
        self.activities_table.setItem(row_position, 1,
                                      QTableWidgetItem(str(activity_data.get('cantidad', 0))))

        # Unidad
        self.activities_table.setItem(row_position, 2,
                                      QTableWidgetItem(activity_data.get('unidad', '')))

        # Valor unitario
        self.activities_table.setItem(row_position, 3,
                                      QTableWidgetItem(str(activity_data.get('valor_unitario', 0))))

        # Total
        cantidad = float(activity_data.get('cantidad', 0))
        valor_unitario = float(activity_data.get('valor_unitario', 0))
        total = cantidad * valor_unitario
        self.activities_table.setItem(row_position, 4,
                                      QTableWidgetItem(str(total)))
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
        """Guarda la cotización actual como un archivo incluyendo encabezados de capítulo"""
        try:
            # ✅ VALIDACIÓN 1: Verificar que haya filas en la tabla
            if self.activities_table.rowCount() == 0:
                QMessageBox.warning(
                    self,
                    "Sin Actividades",
                    "Debe agregar al menos una actividad a la cotización antes de guardar.\n\n"
                    "Para agregar actividades:\n"
                    "• Use el botón 'Agregar Actividad'\n"
                    "• O inserte encabezados de capítulo si es necesario"
                )
                return

            # ✅ VALIDACIÓN 2: Verificar que haya un cliente seleccionado
            if self.client_combo.currentIndex() == -1 or self.client_combo.currentData() is None:
                reply = QMessageBox.warning(
                    self,
                    "Cliente Requerido",
                    "Debe seleccionar un cliente antes de guardar la cotización.\n\n"
                    "¿Desea seleccionar un cliente ahora?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )

                if reply == QMessageBox.Yes:
                    # Enfocar el combo de clientes para que el usuario seleccione
                    self.client_combo.setFocus()
                    self.client_combo.showPopup()
                return

            # ✅ VALIDACIÓN 3: Verificar que el cliente seleccionado sea válido
            client_id = self.client_combo.currentData()
            if client_id is None:
                QMessageBox.critical(
                    self,
                    "Error de Cliente",
                    "El cliente seleccionado no es válido.\n\n"
                    "Por favor:\n"
                    "• Seleccione un cliente diferente de la lista\n"
                    "• O actualice la lista de clientes"
                )
                return

            # ✅ VALIDACIÓN 4: Verificar que se pueden obtener los datos del cliente
            try:
                client_data = self.cotizacion_controller.database_manager.get_client_by_id(client_id)
                if not client_data:
                    QMessageBox.critical(
                        self,
                        "Cliente No Encontrado",
                        f"No se encontraron los datos del cliente seleccionado (ID: {client_id}).\n\n"
                        "Esto puede suceder si:\n"
                        "• El cliente fue eliminado de la base de datos\n"
                        "• Hay un problema con la conexión a la base de datos\n\n"
                        "Por favor seleccione un cliente diferente."
                    )
                    return
            except Exception as e:
                QMessageBox.critical(
                    self,
                    "Error de Base de Datos",
                    f"Error al obtener los datos del cliente:\n{str(e)}\n\n"
                    "Verifique la conexión a la base de datos e intente nuevamente."
                )
                return

            # ✅ VALIDACIÓN 5: Verificar que haya actividades reales (no solo encabezados)
            activities = []
            activity_id = 1

            for row in range(self.activities_table.rowCount()):
                # Obtener los elementos de la fila
                desc_item = self.activities_table.item(row, 0)
                cant_item = self.activities_table.item(row, 1)
                unid_item = self.activities_table.item(row, 2)
                valor_item = self.activities_table.item(row, 3)
                total_item = self.activities_table.item(row, 4)

                # Verificar si es un encabezado de capítulo
                is_chapter_header = False

                # Método 1: Verificar si tiene el atributo personalizado
                if desc_item and hasattr(desc_item, 'is_chapter_header'):
                    is_chapter_header = desc_item.is_chapter_header

                # Método 2: Verificar por patrón de datos si no tiene el atributo
                if not is_chapter_header:
                    try:
                        cant_text = cant_item.text().strip() if cant_item else ""
                        valor_text = valor_item.text().strip() if valor_item else ""
                        total_text = total_item.text().strip() if total_item else ""

                        if not cant_text and not valor_text and not total_text:
                            is_chapter_header = True
                        else:
                            cantidad = float(cant_text) if cant_text else 0.0
                            valor_unitario = float(valor_text) if valor_text else 0.0
                            total = float(total_text) if total_text else 0.0

                            if cantidad == 0 and valor_unitario == 0 and total == 0:
                                descripcion = desc_item.text().strip() if desc_item else ""
                                if descripcion and (
                                        "CAPÍTULO" in descripcion.upper() or "CAPITULO" in descripcion.upper()):
                                    is_chapter_header = True

                    except (ValueError, AttributeError):
                        is_chapter_header = True

                # Si no es encabezado, es una actividad
                if not is_chapter_header:
                    try:
                        descripcion = desc_item.text() if desc_item else ''
                        cantidad = float(cant_item.text()) if cant_item and cant_item.text().strip() else 0.0
                        unidad = unid_item.text() if unid_item else ''
                        valor_unitario = float(valor_item.text()) if valor_item and valor_item.text().strip() else 0.0
                        total = float(total_item.text()) if total_item and total_item.text().strip() else 0.0

                        activity = {
                            'descripcion': descripcion,
                            'cantidad': cantidad,
                            'unidad': unidad,
                            'valor_unitario': valor_unitario,
                            'total': total,
                            'id': activity_id
                        }
                        activities.append(activity)
                        activity_id += 1

                    except (ValueError, AttributeError) as e:
                        print(f"Error procesando fila {row}: {e}")
                        continue

            # ✅ VALIDACIÓN 6: Verificar que haya al menos una actividad válida
            if not activities:
                QMessageBox.warning(
                    self,
                    "Sin Actividades Válidas",
                    "No se encontraron actividades válidas para guardar.\n\n"
                    "La cotización solo contiene encabezados de capítulo.\n"
                    "Agregue al menos una actividad con:\n"
                    "• Descripción\n"
                    "• Cantidad mayor a 0\n"
                    "• Valor unitario mayor a 0"
                )
                return

            # Si llegamos aquí, todas las validaciones pasaron
            # Continuar con el guardado usando la lógica existente mejorada

            # Obtener todas las filas de la tabla (actividades y encabezados)
            table_rows = []
            for row in range(self.activities_table.rowCount()):
                desc_item = self.activities_table.item(row, 0)
                cant_item = self.activities_table.item(row, 1)
                unid_item = self.activities_table.item(row, 2)
                valor_item = self.activities_table.item(row, 3)
                total_item = self.activities_table.item(row, 4)

                # Determinar tipo
                is_chapter_header = False
                if desc_item and hasattr(desc_item, 'is_chapter_header'):
                    is_chapter_header = desc_item.is_chapter_header
                else:
                    # Lógica de detección (ya implementada arriba)
                    try:
                        cant_text = cant_item.text().strip() if cant_item else ""
                        valor_text = valor_item.text().strip() if valor_item else ""
                        total_text = total_item.text().strip() if total_item else ""

                        if not cant_text and not valor_text and not total_text:
                            is_chapter_header = True
                        elif all(float(text) == 0.0 for text in [cant_text, valor_text, total_text] if text):
                            descripcion = desc_item.text().strip() if desc_item else ""
                            if "CAPÍTULO" in descripcion.upper() or "CAPITULO" in descripcion.upper():
                                is_chapter_header = True
                    except:
                        is_chapter_header = True

                if is_chapter_header:
                    chapter_row = {
                        'type': 'chapter_header',
                        'descripcion': desc_item.text() if desc_item else '',
                        'chapter_id': getattr(desc_item, 'chapter_id', None) if desc_item else None,
                        'row_index': row
                    }
                    table_rows.append(chapter_row)
                else:
                    # Buscar la actividad correspondiente
                    for activity in activities:
                        if activity.get('row_index') == row:
                            activity['type'] = 'activity'
                            table_rows.append(activity)
                            break
                    else:
                        # Si no se encontró, crear una nueva
                        try:
                            activity = {
                                'type': 'activity',
                                'descripcion': desc_item.text() if desc_item else '',
                                'cantidad': float(cant_item.text()) if cant_item and cant_item.text().strip() else 0.0,
                                'unidad': unid_item.text() if unid_item else '',
                                'valor_unitario': float(
                                    valor_item.text()) if valor_item and valor_item.text().strip() else 0.0,
                                'total': float(total_item.text()) if total_item and total_item.text().strip() else 0.0,
                                'id': len(activities) + 1,
                                'row_index': row
                            }
                            table_rows.append(activity)
                        except:
                            continue

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
                'table_rows': table_rows,
                'actividades': activities,
                'subtotal': subtotal,
                'administracion': administracion,
                'imprevistos': imprevistos,
                'utilidad': utilidad,
                'iva_utilidad': iva_utilidad,
                'total': total,
                'aiu_values': aiu_values,
                'version': '2.0'
            }

            # Abrir diálogo para guardar archivo
            file_path, _ = QFileDialog.getSaveFileName(
                self,
                "Guardar Cotización",
                f"cotizacion_{client_data.get('nombre', 'cliente').replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}",
                "Archivos de Cotización (*.cotiz);;Todos los archivos (*)"
            )

            if file_path:
                # Asegurar que tenga la extensión correcta
                if not file_path.endswith('.cotiz'):
                    file_path += '.cotiz'

                # Guardar archivo
                self.cotizacion_controller.file_manager.guardar_cotizacion(cotizacion_data, file_path)

                QMessageBox.information(
                    self,
                    "Guardado Exitoso",
                    f"Cotización guardada correctamente.\n\n"
                    f"Archivo: {os.path.basename(file_path)}\n"
                    f"Cliente: {client_data.get('nombre', 'N/A')}\n"
                    f"Actividades: {len(activities)}\n"
                    f"Total: ${total:,.2f}"
                )

        except Exception as e:
            QMessageBox.critical(
                self,
                "Error al Guardar",
                f"Ocurrió un error inesperado al guardar el archivo:\n\n{str(e)}\n\n"
                "Verifique que:\n"
                "• Tenga permisos de escritura en la ubicación seleccionada\n"
                "• No esté intentando sobrescribir un archivo en uso\n"
                "• Todos los datos de la cotización sean válidos"
            )
            print(f"Error detallado: {e}")
            import traceback
            traceback.print_exc()

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
        """
        Carga los datos de una cotización desde archivo en la interfaz

        Args:
            cotizacion_data (dict): Datos completos de la cotización
        """
        try:
            print("Iniciando carga de cotización desde archivo...")

            # Guardar referencia a los datos importados
            self.selected_cotizacion = cotizacion_data

            # 1. Cargar información del cliente
            if 'cliente' in cotizacion_data and cotizacion_data['cliente']:
                cliente = cotizacion_data['cliente']
                print(f"Cargando cliente: {cliente.get('nombre', 'Sin nombre')}")

                # Actualizar campos del cliente en la interfaz si existen
                if hasattr(self, 'client_name_field'):
                    self.client_name_field.setText(cliente.get('nombre', ''))
                if hasattr(self, 'client_nit_field'):
                    self.client_nit_field.setText(cliente.get('nit', ''))
                # Agregar más campos según tu interfaz

            # 2. Limpiar tabla actual
            self.activities_table.setRowCount(0)
            print("Tabla limpiada")

            # 3. Cargar actividades en la tabla
            activities_loaded = 0

            if 'table_rows' in cotizacion_data and cotizacion_data['table_rows']:
                print(f"Cargando {len(cotizacion_data['table_rows'])} filas de table_rows")
                for row_data in cotizacion_data['table_rows']:
                    self.add_imported_row_to_table(row_data)
                    if row_data.get('type') == 'activity':
                        activities_loaded += 1

            elif 'actividades' in cotizacion_data and cotizacion_data['actividades']:
                print(f"Cargando {len(cotizacion_data['actividades'])} actividades")
                for activity in cotizacion_data['actividades']:
                    self.add_imported_activity_to_table(activity)
                    activities_loaded += 1

            print(f"Actividades cargadas en tabla: {activities_loaded}")

            # 4. Cargar valores AIU si existen
            if 'aiu_values' in cotizacion_data and cotizacion_data['aiu_values']:
                try:
                    # Usar método alternativo si set_aiu_values no existe
                    if hasattr(self.aiu_manager, 'set_aiu_values'):
                        self.aiu_manager.set_aiu_values(cotizacion_data['aiu_values'])
                    elif hasattr(self.aiu_manager, 'update_from_imported_data'):
                        self.aiu_manager.update_from_imported_data(cotizacion_data['aiu_values'])
                    else:
                        # Método manual para actualizar AIU
                        aiu_data = cotizacion_data['aiu_values']
                        print(f"Actualizando AIU manualmente: {aiu_data}")

                        # Aquí puedes actualizar los campos AIU directamente si conoces sus nombres
                        # Ejemplo (ajusta según los nombres reales de tus campos):
                        # if hasattr(self, 'admin_spinbox'):
                        #     self.admin_spinbox.setValue(float(aiu_data.get('administracion', 0)))

                except Exception as aiu_error:
                    print(f"Error cargando valores AIU: {aiu_error}")
                    # Continuar sin AIU si falla

            # 5. Actualizar totales
            if hasattr(self, 'calculate_totals'):
                self.calculate_totals()

            # 6. Actualizar otros campos si existen
            if 'fecha' in cotizacion_data and hasattr(self, 'date_field'):
                self.date_field.setText(cotizacion_data['fecha'])

            print("Cotización cargada exitosamente en la interfaz")
            return True

        except Exception as e:
            print(f"Error cargando cotización en interfaz: {e}")
            import traceback
            traceback.print_exc()
            return False

    def add_imported_row_to_table(self, row_data):
        """
        Añade una fila importada a la tabla de actividades
        """
        try:
            row_position = self.activities_table.rowCount()
            self.activities_table.insertRow(row_position)

            if row_data['type'] == 'chapter_header':
                # Crear ítem de capítulo
                item = QTableWidgetItem(row_data['descripcion'])
                item.setData(Qt.UserRole, {'type': 'chapter'})
                item.setBackground(QColor(200, 200, 200))
                self.activities_table.setItem(row_position, 0, item)

                # Llenar celdas vacías para el capítulo
                for col in range(1, self.activities_table.columnCount()):
                    empty_item = QTableWidgetItem("")
                    empty_item.setBackground(QColor(200, 200, 200))
                    self.activities_table.setItem(row_position, col, empty_item)

            elif row_data['type'] == 'activity':
                # Descripción
                desc_item = QTableWidgetItem(row_data['descripcion'])
                desc_item.setData(Qt.UserRole, {'type': 'activity'})
                self.activities_table.setItem(row_position, 0, desc_item)

                # Cantidad
                self.activities_table.setItem(row_position, 1,
                                              QTableWidgetItem(str(row_data.get('cantidad', 0))))

                # Unidad
                self.activities_table.setItem(row_position, 2,
                                              QTableWidgetItem(row_data.get('unidad', '')))

                # Valor unitario
                self.activities_table.setItem(row_position, 3,
                                              QTableWidgetItem(str(row_data.get('valor_unitario', 0))))

                # Total
                total = row_data.get('total', 0)
                if not total:  # Calcular si no existe
                    cantidad = float(row_data.get('cantidad', 0))
                    valor_unitario = float(row_data.get('valor_unitario', 0))
                    total = cantidad * valor_unitario

                self.activities_table.setItem(row_position, 4, QTableWidgetItem(str(total)))

        except Exception as e:
            print(f"Error añadiendo fila a tabla: {e}")

    def add_imported_activity_to_table(self, activity_data):
        """
        Añade una actividad importada (formato anterior) a la tabla
        """
        try:
            row_position = self.activities_table.rowCount()
            self.activities_table.insertRow(row_position)

            # Descripción
            desc_item = QTableWidgetItem(activity_data.get('descripcion', ''))
            desc_item.setData(Qt.UserRole, {'type': 'activity'})
            self.activities_table.setItem(row_position, 0, desc_item)

            # Cantidad
            self.activities_table.setItem(row_position, 1,
                                          QTableWidgetItem(str(activity_data.get('cantidad', 0))))

            # Unidad
            self.activities_table.setItem(row_position, 2,
                                          QTableWidgetItem(activity_data.get('unidad', '')))

            # Valor unitario
            self.activities_table.setItem(row_position, 3,
                                          QTableWidgetItem(str(activity_data.get('valor_unitario', 0))))

            # Total
            cantidad = float(activity_data.get('cantidad', 0))
            valor_unitario = float(activity_data.get('valor_unitario', 0))
            total = cantidad * valor_unitario
            self.activities_table.setItem(row_position, 4, QTableWidgetItem(str(total)))

        except Exception as e:
            print(f"Error añadiendo actividad a tabla: {e}")

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

    def _load_table_rows_with_headers(self, table_rows):
        """Carga las filas de la tabla incluyendo encabezados de capítulo (formato v2.0)"""
        for row_data in table_rows:
            if row_data['type'] == 'chapter_header':
                # Es un encabezado de capítulo
                self._insert_chapter_header_from_file(row_data)
            elif row_data['type'] == 'activity':
                # Es una actividad normal
                self._insert_activity_from_file(row_data)

    def _load_activities_only(self, activities):
        """Carga solo actividades (formato v1.0 - compatibilidad hacia atrás)"""
        for activity in activities:
            self._insert_activity_from_file(activity)

    def _insert_chapter_header_from_file(self, chapter_data):
        """Inserta un encabezado de capítulo desde el archivo"""
        row = self.activities_table.rowCount()
        self.activities_table.insertRow(row)

        # Crear item para el encabezado del capítulo
        chapter_name = chapter_data.get('descripcion', '')
        chapter_item = EditableTableWidgetItem(chapter_name, False)  # No editable

        # Aplicar estilo visual para distinguirlo
        from PyQt5.QtCore import Qt
        from PyQt5.QtGui import QColor, QFont

        # Fondo diferente para el encabezado
        chapter_item.setBackground(QColor(230, 230, 250))  # Color lavanda claro

        # Fuente en negrita
        font = QFont()
        font.setBold(True)
        chapter_item.setFont(font)

        # Centrar el texto
        chapter_item.setTextAlignment(Qt.AlignCenter)

        # Guardar información del capítulo como atributo personalizado
        chapter_id = chapter_data.get('chapter_id')
        if chapter_id:
            chapter_item.chapter_id = chapter_id
            chapter_item.is_chapter_header = True

        # Colocar en la primera columna
        self.activities_table.setItem(row, 0, chapter_item)

        # Para las demás columnas, crear items vacíos con el mismo estilo
        for col in range(1, 5):  # Columnas 1-4 (cantidad, unidad, valor, total)
            empty_item = EditableTableWidgetItem("", False)
            empty_item.setBackground(QColor(230, 230, 250))
            empty_item.setTextAlignment(Qt.AlignCenter)
            self.activities_table.setItem(row, col, empty_item)

        # No agregar botón eliminar en la columna 5 para encabezados
        # O agregar un botón especial si lo deseas

    def _insert_activity_from_file(self, activity):
        """Inserta una actividad desde el archivo"""
        row = self.activities_table.rowCount()
        self.activities_table.insertRow(row)

        # Descripción
        self.activities_table.setItem(row, 0, EditableTableWidgetItem(str(activity.get('descripcion', ''))))

        # Cantidad
        cantidad = activity.get('cantidad', 0)
        self.activities_table.setItem(row, 1, EditableTableWidgetItem(str(cantidad)))

        # Unidad
        self.activities_table.setItem(row, 2, EditableTableWidgetItem(str(activity.get('unidad', ''))))

        # Valor unitario
        valor_unitario = activity.get('valor_unitario', 0)
        self.activities_table.setItem(row, 3, EditableTableWidgetItem(str(valor_unitario)))

        # Total (calculado o desde archivo)
        if 'total' in activity:
            total = activity['total']
        else:
            total = float(cantidad) * float(valor_unitario)

        self.activities_table.setItem(row, 4, EditableTableWidgetItem(str(total), False))

        # Botón eliminar
        delete_btn = QPushButton("Eliminar")
        delete_btn.clicked.connect(lambda _, r=row: self.delete_activity(r))
        self.activities_table.setCellWidget(row, 5, delete_btn)


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