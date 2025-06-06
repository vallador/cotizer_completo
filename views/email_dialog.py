from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QFormLayout, QCheckBox, \
    QMessageBox, QFileDialog, QGroupBox, QComboBox
from PyQt5.QtCore import Qt
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
    def __init__(self, attachments=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Enviar Cotización por Correo")
        self.setMinimumWidth(500)

        # Crear gestor de correo
        self.email_manager = EmailManager()
        self.attachments = attachments or []

        # Crear layout principal
        main_layout = QVBoxLayout(self)

        # Formulario de correo
        form_layout = QFormLayout()

        self.to_input = QLineEdit(', '.join(self.email_manager.config['default_recipients']))
        self.subject_input = QLineEdit(self.email_manager.config['default_subject'])
        self.message_input = QLineEdit(self.email_manager.config['default_message'])

        form_layout.addRow("Para:", self.to_input)
        form_layout.addRow("Asunto:", self.subject_input)
        form_layout.addRow("Mensaje:", self.message_input)

        main_layout.addLayout(form_layout)

        # Sección de archivos adjuntos
        attachments_group = QGroupBox("Archivos Adjuntos")
        attachments_layout = QVBoxLayout()

        # Mostrar archivos adjuntos actuales
        for attachment in self.attachments:
            attachments_layout.addWidget(QLabel(os.path.basename(attachment)))

        # Botón para agregar más archivos
        add_attachment_btn = QPushButton("Agregar Archivo Adjunto")
        add_attachment_btn.clicked.connect(self.add_attachment)
        attachments_layout.addWidget(add_attachment_btn)

        attachments_group.setLayout(attachments_layout)
        main_layout.addWidget(attachments_group)

        # Botones de acción
        buttons_layout = QHBoxLayout()

        self.config_btn = QPushButton("Configuración")
        self.config_btn.clicked.connect(self.open_config)

        self.send_btn = QPushButton("Enviar")
        self.send_btn.clicked.connect(self.send_email)

        self.cancel_btn = QPushButton("Cancelar")
        self.cancel_btn.clicked.connect(self.reject)

        buttons_layout.addWidget(self.config_btn)
        buttons_layout.addWidget(self.send_btn)
        buttons_layout.addWidget(self.cancel_btn)

        main_layout.addLayout(buttons_layout)

    def add_attachment(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Seleccionar Archivo", "", "Todos los Archivos (*)")
        if file_path:
            self.attachments.append(file_path)
            # Actualizar la lista de archivos adjuntos
            layout = self.findChild(QGroupBox, "").layout()
            layout.insertWidget(layout.count() - 1, QLabel(os.path.basename(file_path)))

    def open_config(self):
        config_dialog = EmailConfigDialog(self)
        if config_dialog.exec_() == QDialog.Accepted:
            # Recargar configuración
            self.email_manager = EmailManager()
            self.to_input.setText(', '.join(self.email_manager.config['default_recipients']))
            self.subject_input.setText(self.email_manager.config['default_subject'])
            self.message_input.setText(self.email_manager.config['default_message'])

    def send_email(self):
        # Obtener destinatarios
        recipients = [r.strip() for r in self.to_input.text().split(',') if r.strip()]
        if not recipients:
            QMessageBox.warning(self, "Error", "Debe especificar al menos un destinatario.")
            return

        # Enviar correo
        if self.email_manager.send_email(
                recipients=recipients,
                subject=self.subject_input.text(),
                message=self.message_input.text(),
                attachments=self.attachments
        ):
            QMessageBox.information(self, "Éxito", "Correo enviado correctamente.")
            self.accept()
        else:
            QMessageBox.critical(self, "Error", "No se pudo enviar el correo. Verifique la configuración.")
