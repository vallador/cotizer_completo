from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFormLayout, 
                             QMessageBox, QGroupBox, QScrollArea, QWidget, QCheckBox)
from PyQt5.QtGui import QFont
from utils.email_manager import EmailManager
import os


class EmailConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Configuración de Correo Electrónico")
        self.setMinimumWidth(500)

        # Crear gestor de correo
        self.email_manager = EmailManager()

        # Crear layout principal
        main_layout = QVBoxLayout(self)

        # Grupo de configuración SMTP
        smtp_group = QGroupBox("Configuración del Servidor SMTP")
        smtp_layout = QFormLayout()

        self.server_input = QLineEdit(self.email_manager.config['smtp_server'])
        self.port_input = QLineEdit(str(self.email_manager.config['smtp_port']))
        self.user_input = QLineEdit(self.email_manager.config['smtp_user'])
        self.password_input = QLineEdit(self.email_manager.config['smtp_password'])
        self.password_input.setEchoMode(QLineEdit.Password)

        smtp_layout.addRow("Servidor SMTP:", self.server_input)
        smtp_layout.addRow("Puerto:", self.port_input)
        smtp_layout.addRow("Usuario:", self.user_input)
        smtp_layout.addRow("Contraseña:", self.password_input)

        smtp_group.setLayout(smtp_layout)
        main_layout.addWidget(smtp_group)

        # Grupo de configuración de correo predeterminado
        email_group = QGroupBox("Configuración de Correo Predeterminado")
        email_layout = QFormLayout()

        self.sender_input = QLineEdit(self.email_manager.config['default_sender'])
        self.recipients_input = QLineEdit(', '.join(self.email_manager.config['default_recipients']))
        self.subject_input = QLineEdit(self.email_manager.config['default_subject'])

        email_layout.addRow("Remitente:", self.sender_input)
        email_layout.addRow("Destinatarios (separados por coma):", self.recipients_input)
        email_layout.addRow("Asunto predeterminado:", self.subject_input)

        email_group.setLayout(email_layout)
        main_layout.addWidget(email_group)

        # Botones de acción
        buttons_layout = QHBoxLayout()

        self.test_btn = QPushButton("Probar Conexión")
        self.test_btn.clicked.connect(self.test_connection)

        self.save_btn = QPushButton("Guardar Configuración")
        self.save_btn.clicked.connect(self.save_config)

        self.cancel_btn = QPushButton("Cancelar")
        self.cancel_btn.clicked.connect(self.reject)

        buttons_layout.addWidget(self.test_btn)
        buttons_layout.addWidget(self.save_btn)
        buttons_layout.addWidget(self.cancel_btn)

        main_layout.addLayout(buttons_layout)

    def test_connection(self):
        # Actualizar configuración con valores actuales
        self.update_config_from_inputs()

        # Probar conexión
        if self.email_manager.test_connection():
            QMessageBox.information(self, "Éxito", "Conexión exitosa al servidor SMTP.")
        else:
            QMessageBox.critical(self, "Error", "No se pudo conectar al servidor SMTP. Verifique la configuración.")

    def save_config(self):
        # Actualizar configuración con valores actuales
        self.update_config_from_inputs()

        # Guardar configuración
        if self.email_manager.save_config():
            QMessageBox.information(self, "Éxito", "Configuración guardada correctamente.")
            self.accept()
        else:
            QMessageBox.critical(self, "Error", "No se pudo guardar la configuración.")

    def update_config_from_inputs(self):
        # Actualizar configuración con valores de los campos
        self.email_manager.update_config(
            smtp_server=self.server_input.text(),
            smtp_port=int(self.port_input.text()),
            smtp_user=self.user_input.text(),
            smtp_password=self.password_input.text(),
            default_sender=self.sender_input.text(),
            default_recipients=[r.strip() for r in self.recipients_input.text().split(',') if r.strip()],
            default_subject=self.subject_input.text()
        )


class SendEmailDialog(QDialog):
    def __init__(self, attachments=None, parent=None, client_email=""):
        super().__init__(parent)
        self.setWindowTitle("Enviar Cotización por Correo")
        self.setMinimumWidth(600)
        self.resize(600, 500)

        # Crear gestor de correo
        self.email_manager = EmailManager()
        self.attachments = attachments or []
        self.checkboxes = []

        # Crear layout principal
        main_layout = QVBoxLayout(self)

        # --- SECCIÓN 1: Formulario de Correo ---
        form_group = QGroupBox("Detalles del Correo")
        form_layout = QFormLayout()

        # Usar correo del cliente si está disponible, si no, el default
        destinatarios = client_email if client_email else ', '.join(self.email_manager.config['default_recipients'])
        self.to_input = QLineEdit(destinatarios)
        self.subject_input = QLineEdit(self.email_manager.config['default_subject'])
        self.message_input = QLineEdit(self.email_manager.config['default_message'])

        form_layout.addRow("Para:", self.to_input)
        form_layout.addRow("Asunto:", self.subject_input)
        form_layout.addRow("Mensaje:", self.message_input)

        form_group.setLayout(form_layout)
        main_layout.addWidget(form_group)

        # --- SECCIÓN 2: Selección de Adjuntos ---
        attachments_group = QGroupBox("Archivos Adjuntos")
        attachments_layout = QVBoxLayout()
        
        # Área de scroll para adjuntos
        scroll = QScrollArea()
        scroll_widget = QWidget()
        scroll_inner = QVBoxLayout()
        scroll_inner.setContentsMargins(0, 0, 0, 0)

        if not self.attachments:
            scroll_inner.addWidget(QLabel("No hay archivos adjuntos disponibles."))
        else:
            title_lbl = QLabel("Seleccione los archivos a enviar:")
            title_lbl.setStyleSheet("font-weight: bold; color: #555;")
            scroll_inner.addWidget(title_lbl)
            
            for path in self.attachments:
                if os.path.exists(path):
                    filename = os.path.basename(path)
                    cb = QCheckBox(filename)
                    cb.setChecked(True) # Por defecto todos marcados
                    cb.setProperty("file_path", path)
                    cb.setToolTip(path)
                    self.checkboxes.append(cb)
                    scroll_inner.addWidget(cb)
        
        scroll_inner.addStretch()
        scroll_widget.setLayout(scroll_inner)
        scroll.setWidget(scroll_widget)
        scroll.setWidgetResizable(True)
        scroll.setMaximumHeight(150) # Limitar altura
        
        attachments_layout.addWidget(scroll)
        attachments_group.setLayout(attachments_layout)
        main_layout.addWidget(attachments_group)

        # --- SECCIÓN 3: Botones ---
        buttons_layout = QHBoxLayout()

        self.config_btn = QPushButton("Configuración SMTP")
        self.config_btn.clicked.connect(self.open_config)

        self.send_btn = QPushButton("Enviar Correo")
        self.send_btn.clicked.connect(self.send_email)
        self.send_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")

        self.cancel_btn = QPushButton("Cancelar")
        self.cancel_btn.clicked.connect(self.reject)

        buttons_layout.addWidget(self.config_btn)
        buttons_layout.addStretch()
        buttons_layout.addWidget(self.cancel_btn)
        buttons_layout.addWidget(self.send_btn)

        main_layout.addLayout(buttons_layout)

    def open_config(self):
        config_dialog = EmailConfigDialog(self)
        if config_dialog.exec_() == QDialog.Accepted:
            # Recargar configuración
            self.email_manager = EmailManager()
            # No sobreescribimos el destinatario si ya había uno puesto manualmente
            if not self.to_input.text():
                 self.to_input.setText(', '.join(self.email_manager.config['default_recipients']))
            self.subject_input.setText(self.email_manager.config['default_subject'])
            self.message_input.setText(self.email_manager.config['default_message'])

    def get_selected_attachments(self):
        """Devuelve la lista de rutas de archivo seleccionadas."""
        selected_paths = []
        for cb in self.checkboxes:
            if cb.isChecked():
                selected_paths.append(cb.property("file_path"))
        return selected_paths

    def send_email(self):
        # Obtener destinatarios
        recipients = [r.strip() for r in self.to_input.text().split(',') if r.strip()]
        if not recipients:
            QMessageBox.warning(self, "Error", "Debe especificar al menos un destinatario.")
            return

        # Obtener adjuntos seleccionados
        selected_attachments = self.get_selected_attachments()

        # Enviar correo
        if self.email_manager.send_email(
                recipients=recipients,
                subject=self.subject_input.text(),
                message=self.message_input.text(),
                attachments=selected_attachments
        ):
            QMessageBox.information(self, "Éxito", "Correo enviado correctamente.")
            self.accept()
        else:
            QMessageBox.critical(self, "Error", "No se pudo enviar el correo. Verifique la configuración (SMTP).")
