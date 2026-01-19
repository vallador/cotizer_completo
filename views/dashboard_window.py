from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
                             QPushButton, QLabel, QComboBox, QCheckBox, QDateEdit, QLineEdit,
                             QGroupBox, QHeaderView, QWidget, QMessageBox, QMenu, QAction, QTextEdit,
                             QGridLayout)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QColor, QFont, QIcon
from datetime import datetime
import os


class DashboardWindow(QDialog):
    """Dashboard window for viewing and managing quotations"""
    
    def __init__(self, database_manager, parent=None):
        super().__init__(parent)
        self.db = database_manager
        self.parent_window = parent
        self.current_quotations = []
        
        self.setWindowTitle("📊 Dashboard de Cotizaciones")
        self.setGeometry(100, 100, 1200, 700)
        self.setModal(False)
        
        self.setup_ui()
        self.load_quotations()
    
    def setup_ui(self):
        """Setup the UI components"""
        main_layout = QHBoxLayout()
        
        # ===== LEFT PANEL: FILTERS =====
        filter_panel = self.create_filter_panel()
        main_layout.addWidget(filter_panel, 0)
        
        # ===== RIGHT PANEL: TABLE AND STATS =====
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        
        # Statistics panel
        stats_panel = self.create_stats_panel()
        right_layout.addWidget(stats_panel)
        
        # Toolbar
        toolbar = self.create_toolbar()
        right_layout.addLayout(toolbar)
        
        # Table
        self.quotations_table = self.create_quotations_table()
        right_layout.addWidget(self.quotations_table)
        
        # Action buttons
        actions_layout = self.create_action_buttons()
        right_layout.addLayout(actions_layout)
        
        main_layout.addWidget(right_panel, 1)
        
        self.setLayout(main_layout)
    
    def create_filter_panel(self):
        """Creates the left filter panel"""
        panel = QGroupBox("Filtros")
        panel.setMaximumWidth(250)
        layout = QVBoxLayout()
        
        # Include tests checkbox
        layout.addWidget(QLabel("Tipo:"))
        self.filter_pruebas_check = QCheckBox("Incluir Pruebas")
        self.filter_pruebas_check.setChecked(False)
        self.filter_pruebas_check.stateChanged.connect(self.apply_filters)
        layout.addWidget(self.filter_pruebas_check)
        
        self.filter_solo_reales_check = QCheckBox("Solo Reales")
        self.filter_solo_reales_check.setChecked(True)
        self.filter_solo_reales_check.stateChanged.connect(self.apply_filters)
        layout.addWidget(self.filter_solo_reales_check)
        
        layout.addWidget(QLabel(""))  # Spacer
        
        # State filters
        layout.addWidget(QLabel("Estado:"))
        self.filter_pendiente_check = QCheckBox("⏱ Pendiente")
        self.filter_pendiente_check.setChecked(True)
        self.filter_pendiente_check.stateChanged.connect(self.apply_filters)
        layout.addWidget(self.filter_pendiente_check)
        
        self.filter_ganada_check = QCheckBox("✓ Ganada")
        self.filter_ganada_check.setChecked(True)
        self.filter_ganada_check.stateChanged.connect(self.apply_filters)
        layout.addWidget(self.filter_ganada_check)
        
        self.filter_perdida_check = QCheckBox("✗ Perdida")
        self.filter_perdida_check.setChecked(True)
        self.filter_perdida_check.stateChanged.connect(self.apply_filters)
        layout.addWidget(self.filter_perdida_check)
        
        self.filter_cancelada_check = QCheckBox("🚫 Cancelada")
        self.filter_cancelada_check.setChecked(False)
        self.filter_cancelada_check.stateChanged.connect(self.apply_filters)
        layout.addWidget(self.filter_cancelada_check)
        
        layout.addWidget(QLabel(""))  # Spacer
        
        # Date range
        layout.addWidget(QLabel("Rango de Fechas:"))
        self.filter_fecha_inicio = QDateEdit()
        self.filter_fecha_inicio.setCalendarPopup(True)
        self.filter_fecha_inicio.setDate(QDate.currentDate().addMonths(-1))
        self.filter_fecha_inicio.dateChanged.connect(self.apply_filters)
        layout.addWidget(self.filter_fecha_inicio)
        
        self.filter_fecha_fin = QDateEdit()
        self.filter_fecha_fin.setCalendarPopup(True)
        self.filter_fecha_fin.setDate(QDate.currentDate())
        self.filter_fecha_fin.dateChanged.connect(self.apply_filters)
        layout.addWidget(self.filter_fecha_fin)
        
        layout.addWidget(QLabel(""))  # Spacer
        
        # Client filter
        layout.addWidget(QLabel("Cliente:"))
        self.filter_cliente_combo = QComboBox()
        self.filter_cliente_combo.addItem("Todos", None)
        self.load_clients_to_filter()
        self.filter_cliente_combo.currentIndexChanged.connect(self.apply_filters)
        layout.addWidget(self.filter_cliente_combo)
        
        layout.addWidget(QLabel(""))  # Spacer
        
        # Clear filters button
        clear_btn = QPushButton("Limpiar Filtros")
        clear_btn.clicked.connect(self.clear_filters)
        layout.addWidget(clear_btn)
        
        layout.addStretch()
        
        panel.setLayout(layout)
        return panel
    
    def create_stats_panel(self):
        """Creates the statistics panel"""
        panel = QGroupBox("Resumen del Período Seleccionado")
        layout = QGridLayout()
        
        # Labels for statistics
        self.stats_labels = {
            'pendiente_count': QLabel("0"),
            'pendiente_total': QLabel("$0"),
            'ganada_count': QLabel("0"),
            'ganada_total': QLabel("$0"),
            'perdida_count': QLabel("0"),
            'perdida_total': QLabel("$0"),
            'tasa_conversion': QLabel("0%")
        }
        
        # Setup grid
        layout.addWidget(QLabel("<b>Estado</b>"), 0, 0)
        layout.addWidget(QLabel("<b>Cantidad</b>"), 0, 1)
        layout.addWidget(QLabel("<b>Monto Total</b>"), 0, 2)
        
        layout.addWidget(QLabel("⏱ Pendientes:"), 1, 0)
        layout.addWidget(self.stats_labels['pendiente_count'], 1, 1)
        layout.addWidget(self.stats_labels['pendiente_total'], 1, 2)
        
        layout.addWidget(QLabel("✓ Ganadas:"), 2, 0)
        layout.addWidget(self.stats_labels['ganada_count'], 2, 1)
        layout.addWidget(self.stats_labels['ganada_total'], 2, 2)
        
        layout.addWidget(QLabel("✗ Perdidas:"), 3, 0)
        layout.addWidget(self.stats_labels['perdida_count'], 3, 1)
        layout.addWidget(self.stats_labels['perdida_total'], 3, 2)
        
        layout.addWidget(QLabel("<b>Tasa de Conversión:</b>"), 4, 0)
        layout.addWidget(self.stats_labels['tasa_conversion'], 4, 1, 1, 2)
        
        # Style the labels
        for label in self.stats_labels.values():
            font = label.font()
            font.setPointSize(10)
            label.setFont(font)
        
        panel.setLayout(layout)
        return panel
    
    def create_toolbar(self):
        """Creates the toolbar with quick actions"""
        layout = QHBoxLayout()
        
        refresh_btn = QPushButton("🔄 Actualizar")
        refresh_btn.clicked.connect(self.load_quotations)
        layout.addWidget(refresh_btn)
        
        new_btn = QPushButton("➕ Nueva Cotización")
        new_btn.clicked.connect(self.new_quotation)
        layout.addWidget(new_btn)
        
        # Search box
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Buscar proyecto...")
        self.search_input.textChanged.connect(self.apply_filters)
        layout.addWidget(self.search_input)
        
        layout.addStretch()
        
        return layout
    
    def create_quotations_table(self):
        """Creates the main quotations table"""
        table = QTableWidget(0, 8)
        table.setHorizontalHeaderLabels([
            "ID", "Cliente", "Proyecto", "Fecha", "Monto", "Estado", "Tipo", "Ver"
        ])
        
        # Set column widths
        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)  # ID
        header.setSectionResizeMode(1, QHeaderView.Stretch)  # Cliente
        header.setSectionResizeMode(2, QHeaderView.Stretch)  # Proyecto
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)  # Fecha
        header.setSectionResizeMode(4, QHeaderView.ResizeToContents)  # Monto
        header.setSectionResizeMode(5, QHeaderView.ResizeToContents)  # Estado
        header.setSectionResizeMode(6, QHeaderView.ResizeToContents)  # Tipo
        header.setSectionResizeMode(7, QHeaderView.ResizeToContents)  # Ver
        
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.setContextMenuPolicy(Qt.CustomContextMenu)
        table.customContextMenuRequested.connect(self.show_context_menu)
        table.doubleClicked.connect(self.open_quotation)
        
        return table
    
    def create_action_buttons(self):
        """Creates bottom action buttons"""
        layout = QHBoxLayout()
        
        layout.addStretch()
        
        close_btn = QPushButton("Cerrar")
        close_btn.clicked.connect(self.close)
        layout.addWidget(close_btn)
        
        return layout
    
    def load_clients_to_filter(self):
        """Loads clients into the filter combo"""
        try:
            clients = self.db.get_all_clients()
            for client in clients:
                self.filter_cliente_combo.addItem(client['nombre'], client['id'])
        except Exception as e:
            print(f"Error loading clients: {e}")
    
    def load_quotations(self):
        """Loads quotations from database with current filters"""
        try:
            filters = self.get_current_filters()
            include_test = self.filter_pruebas_check.isChecked()
            
            self.current_quotations = self.db.get_all_quotations(
                include_test=include_test,
                filters=filters
            )
            
            self.populate_table()
            self.update_statistics()
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al cargar cotizaciones: {e}")
            print(f"Error loading quotations: {e}")
    
    def get_current_filters(self):
        """Gets current filter values as a dictionary"""
        filters = {}
        
        # Estado filter
        estados = []
        if self.filter_pendiente_check.isChecked():
            estados.append('pendiente')
        if self.filter_ganada_check.isChecked():
            estados.append('ganada')
        if self.filter_perdida_check.isChecked():
            estados.append('perdida')
        if self.filter_cancelada_check.isChecked():
            estados.append('cancelada')
        
        if estados:
            filters['estados'] = estados  # Will need to adjust query for multiple states
        
        # Date range
        filters['fecha_inicio'] = self.filter_fecha_inicio.date().toString("yyyy-MM-dd")
        filters['fecha_fin'] = self.filter_fecha_fin.date().toString("yyyy-MM-dd")
        
        # Client
        cliente_id = self.filter_cliente_combo.currentData()
        if cliente_id:
            filters['cliente_id'] = cliente_id
        
        return filters
    
    def populate_table(self):
        """Populates the table with quotations"""
        self.quotations_table.setRowCount(0)
        
        # Filter by search text
        search_text = self.search_input.text().lower()
        
        for quotation in self.current_quotations:
            # Apply search filter
            if search_text and search_text not in quotation['nombre_proyecto'].lower():
                continue
            
            # Check estado filter
            estado_checks = {
                'pendiente': self.filter_pendiente_check,
                'ganada': self.filter_ganada_check,
                'perdida': self.filter_perdida_check,
                'cancelada': self.filter_cancelada_check
            }
            
            if quotation['estado'] in estado_checks:
                if not estado_checks[quotation['estado']].isChecked():
                    continue
            
            row = self.quotations_table.rowCount()
            self.quotations_table.insertRow(row)
            
            # ID
            id_item = QTableWidgetItem(f"COT-{quotation['id']:03d}")
            id_item.setData(Qt.UserRole, quotation)
            self.quotations_table.setItem(row, 0, id_item)
            
            # Cliente
            self.quotations_table.setItem(row, 1, QTableWidgetItem(quotation['cliente_nombre'] or "N/A"))
            
            # Proyecto
            proyecto_item = QTableWidgetItem(quotation['nombre_proyecto'])
            if quotation['es_prueba']:
                proyecto_item.setBackground(QColor("#E3F2FD"))  # Light blue for tests
            self.quotations_table.setItem(row, 2, proyecto_item)
            
            # Fecha
            try:
                fecha = datetime.fromisoformat(quotation['fecha_creacion']).strftime("%d/%m/%Y")
            except:
                fecha = quotation['fecha_creacion']
            self.quotations_table.setItem(row, 3, QTableWidgetItem(fecha))
            
            # Monto
            monto_item = QTableWidgetItem(f"${quotation['monto_total']:,.0f}")
            monto_item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
            self.quotations_table.setItem(row, 4, monto_item)
            
            # Estado
            estado_item = self.create_estado_item(quotation['estado'])
            self.quotations_table.setItem(row, 5, estado_item)
            
            # Tipo
            tipo_text = "🧪 Prueba" if quotation['es_prueba'] else "📄 Real"
            self.quotations_table.setItem(row, 6, QTableWidgetItem(tipo_text))
            
            # Ver button (would be better as QPushButton in cell widget)
            ver_item = QTableWidgetItem("👁")
            ver_item.setTextAlignment(Qt.AlignCenter)
            self.quotations_table.setItem(row, 7, ver_item)
    
    def create_estado_item(self, estado):
        """Creates a colored estado item"""
        icons_and_colors = {
            'pendiente': ("⏱ Pendiente", "#FFF9C4"),  # Yellow
            'ganada': ("✓ Ganada", "#C8E6C9"),  # Green
            'perdida': ("✗ Perdida", "#FFCDD2"),  # Red
            'cancelada': ("🚫 Cancelada", "#E0E0E0")  # Gray
        }
        
        text, color = icons_and_colors.get(estado, (estado, "#FFFFFF"))
        item = QTableWidgetItem(text)
        item.setBackground(QColor(color))
        item.setTextAlignment(Qt.AlignCenter)
        
        font = item.font()
        font.setBold(True)
        item.setFont(font)
        
        return item
    
    def update_statistics(self):
        """Updates the statistics panel"""
        try:
            # Get date range from filters
            fecha_inicio = self.filter_fecha_inicio.date().toPyDate()
            fecha_fin = self.filter_fecha_fin.date().toPyDate()
            
            # Calculate stats from loaded quotations
            stats = {
                'pendiente': {'count': 0, 'total': 0},
                'ganada': {'count': 0, 'total': 0},
                'perdida': {'count': 0, 'total': 0}
            }
            
            for quotation in self.current_quotations:
                estado = quotation['estado']
                if estado in stats:
                    stats[estado]['count'] += 1
                    stats[estado]['total'] += quotation['monto_total']
            
            # Update labels
            self.stats_labels['pendiente_count'].setText(str(stats['pendiente']['count']))
            self.stats_labels['pendiente_total'].setText(f"${stats['pendiente']['total']:,.0f}")
            
            self.stats_labels['ganada_count'].setText(str(stats['ganada']['count']))
            self.stats_labels['ganada_total'].setText(f"${stats['ganada']['total']:,.0f}")
            
            self.stats_labels['perdida_count'].setText(str(stats['perdida']['count']))
            self.stats_labels['perdida_total'].setText(f"${stats['perdida']['total']:,.0f}")
            
            # Conversion rate
            total_count = stats['ganada']['count'] + stats['perdida']['count']
            if total_count > 0:
                tasa = (stats['ganada']['count'] / total_count) * 100
                self.stats_labels['tasa_conversion'].setText(f"{tasa:.1f}%")
            else:
                self.stats_labels['tasa_conversion'].setText("N/A")
            
        except Exception as e:
            print(f"Error updating statistics: {e}")
    
    def apply_filters(self):
        """Reloads quotations with current filters"""
        self.load_quotations()
    
    def clear_filters(self):
        """Clears all filters to defaults"""
        self.filter_pruebas_check.setChecked(False)
        self.filter_solo_reales_check.setChecked(True)
        self.filter_pendiente_check.setChecked(True)
        self.filter_ganada_check.setChecked(True)
        self.filter_perdida_check.setChecked(True)
        self.filter_cancelada_check.setChecked(False)
        self.filter_fecha_inicio.setDate(QDate.currentDate().addMonths(-1))
        self.filter_fecha_fin.setDate(QDate.currentDate())
        self.filter_cliente_combo.setCurrentIndex(0)
        self.search_input.clear()
    
    def show_context_menu(self, position):
        """Shows context menu for table row"""
        row = self.quotations_table.rowAt(position.y())
        if row < 0:
            return
        
        item = self.quotations_table.item(row, 0)
        quotation = item.data(Qt.UserRole)
        
        menu = QMenu()
        
        open_action = QAction("📂 Abrir Cotización", self)
        open_action.triggered.connect(lambda: self.open_quotation_by_id(quotation['id']))
        menu.addAction(open_action)
        
        duplicate_action = QAction("📋 Duplicar", self)
        duplicate_action.triggered.connect(lambda: self.duplicate_quotation(quotation['id']))
        menu.addAction(duplicate_action)
        
        menu.addSeparator()
        
        # State changes
        if quotation['estado'] != 'ganada':
            ganada_action = QAction("✓ Marcar como Ganada", self)
            ganada_action.triggered.connect(lambda: self.change_state(quotation['id'], 'ganada'))
            menu.addAction(ganada_action)
        
        if quotation['estado'] != 'perdida':
            perdida_action = QAction("✗ Marcar como Perdida", self)
            perdida_action.triggered.connect(lambda: self.change_state(quotation['id'], 'perdida'))
            menu.addAction(perdida_action)
        
        if quotation['estado'] != 'pendiente':
            pendiente_action = QAction("⏱ Marcar como Pendiente", self)
            pendiente_action.triggered.connect(lambda: self.change_state(quotation['id'], 'pendiente'))
            menu.addAction(pendiente_action)
        
        menu.addSeparator()
        
        # View files
        if quotation['ruta_pdf'] and os.path.exists(quotation['ruta_pdf']):
            pdf_action = QAction("📄 Ver PDF", self)
            pdf_action.triggered.connect(lambda: os.startfile(quotation['ruta_pdf']))
            menu.addAction(pdf_action)
        
        if quotation['ruta_excel'] and os.path.exists(quotation['ruta_excel']):
            excel_action = QAction("📊 Ver Excel", self)
            excel_action.triggered.connect(lambda: os.startfile(quotation['ruta_excel']))
            menu.addAction(excel_action)
        
        menu.addSeparator()
        
        delete_action = QAction("🗑 Eliminar", self)
        delete_action.triggered.connect(lambda: self.delete_quotation(quotation['id']))
        menu.addAction(delete_action)
        
        menu.exec_(self.quotations_table.mapToGlobal(position))
    
    def open_quotation(self):
        """Opens selected quotation for editing"""
        row = self.quotations_table.currentRow()
        if row >= 0:
            item = self.quotations_table.item(row, 0)
            quotation = item.data(Qt.UserRole)
            self.open_quotation_by_id(quotation['id'])
    
    def open_quotation_by_id(self, quotation_id):
        """Loads a quotation into the main window"""
        if self.parent_window:
            try:
                # Get snapshot
                snapshot = self.db.get_latest_snapshot(quotation_id)
                if snapshot:
                    self.parent_window.load_quotation_from_snapshot(snapshot)
                    self.close()
                else:
                    QMessageBox.warning(self, "Sin Snapshot", 
                                      "No se encontró un snapshot para esta cotización.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al cargar cotización: {e}")
                print(f"Error opening quotation: {e}")
    
    def duplicate_quotation(self, quotation_id):
        """Duplicates a quotation"""
        try:
            snapshot = self.db.get_latest_snapshot(quotation_id)
            if snapshot and self.parent_window:
                self.parent_window.load_quotation_from_snapshot(snapshot, as_new=True)
                self.close()
                QMessageBox.information(self, "Éxito", "Cotización duplicada. Modifique y genere nuevamente.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al duplicar: {e}")
    
    def change_state(self, quotation_id, new_state):
        """Changes the state of a quotation"""
        try:
            self.db.update_quotation_state(quotation_id, new_state)
            self.load_quotations()  # Reload
            QMessageBox.information(self, "Éxito", f"Estado actualizado a: {new_state}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al cambiar estado: {e}")
    
    def delete_quotation(self, quotation_id):
        """Deletes a quotation after confirmation"""
        reply = QMessageBox.question(
            self, "Confirmar Eliminación",
            "¿Está seguro de eliminar esta cotización?\nEsta acción no se puede deshacer.",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            try:
                self.db.delete_quotation(quotation_id)
                self.load_quotations()
                QMessageBox.information(self, "Éxito", "Cotización eliminada.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al eliminar: {e}")
    
    def new_quotation(self):
        """Closes dashboard and prepares for new quotation"""
        if self.parent_window:
            self.parent_window.clear_form()
        self.close()
