import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
import json
from typing import List, Optional

class EmailManager:
    def __init__(self, config_file=None):
        """
        Inicializa el gestor de correos electrónicos.
        
        Args:
            config_file: Ruta al archivo de configuración JSON (opcional)
        """
        self.config = {
            'smtp_server': 'smtp.gmail.com',
            'smtp_port': 587,
            'smtp_user': '',
            'smtp_password': '',
            'default_sender': '',
            'default_recipients': [],
            'default_subject': 'Cotización',
            'default_message': 'Adjunto encontrará la cotización solicitada.'
        }
        
        # Cargar configuración desde archivo si existe
        if config_file and os.path.exists(config_file):
            self.load_config(config_file)
        else:
            # Intentar cargar desde config.json en el directorio actual
            default_config = os.path.join(os.getcwd(), 'config.json')
            if os.path.exists(default_config):
                self.load_config(default_config)
    
    def load_config(self, config_file):
        """
        Carga la configuración desde un archivo JSON.
        
        Args:
            config_file: Ruta al archivo de configuración
        """
        try:
            with open(config_file, 'r') as f:
                config_data = json.load(f)
                
            # Actualizar configuración si existe la sección 'email'
            if 'email' in config_data:
                self.config.update(config_data['email'])
        except Exception as e:
            print(f"Error al cargar la configuración de email: {e}")
    
    def save_config(self, config_file=None):
        """
        Guarda la configuración actual en un archivo JSON.
        
        Args:
            config_file: Ruta al archivo de configuración (opcional)
        """
        if not config_file:
            config_file = os.path.join(os.getcwd(), 'config.json')
            
        try:
            # Leer configuración existente si el archivo existe
            if os.path.exists(config_file):
                with open(config_file, 'r') as f:
                    config_data = json.load(f)
            else:
                config_data = {}
            
            # Actualizar sección 'email'
            config_data['email'] = self.config
            
            # Guardar configuración
            with open(config_file, 'w') as f:
                json.dump(config_data, f, indent=4)
                
            return True
        except Exception as e:
            print(f"Error al guardar la configuración de email: {e}")
            return False
    
    def send_email(self, 
                  recipients: List[str] = None, 
                  subject: str = None, 
                  message: str = None, 
                  attachments: List[str] = None,
                  sender: str = None) -> bool:
        """
        Envía un correo electrónico con archivos adjuntos.
        
        Args:
            recipients: Lista de destinatarios
            subject: Asunto del correo
            message: Cuerpo del mensaje
            attachments: Lista de rutas a archivos adjuntos
            sender: Dirección de correo del remitente
            
        Returns:
            bool: True si el correo se envió correctamente, False en caso contrario
        """
        # Usar valores predeterminados si no se proporcionan
        recipients = recipients or self.config['default_recipients']
        subject = subject or self.config['default_subject']
        message = message or self.config['default_message']
        sender = sender or self.config['default_sender'] or self.config['smtp_user']
        
        # Verificar que haya destinatarios
        if not recipients:
            print("Error: No se especificaron destinatarios")
            return False
        
        # Verificar que haya credenciales SMTP
        if not self.config['smtp_user'] or not self.config['smtp_password']:
            print("Error: Faltan credenciales SMTP")
            return False
        
        try:
            # Crear mensaje
            msg = MIMEMultipart()
            msg['From'] = sender
            msg['To'] = ', '.join(recipients)
            msg['Subject'] = subject
            
            # Agregar cuerpo del mensaje
            msg.attach(MIMEText(message, 'plain'))
            
            # Agregar archivos adjuntos
            if attachments:
                for file_path in attachments:
                    if os.path.exists(file_path):
                        with open(file_path, 'rb') as file:
                            part = MIMEApplication(file.read(), Name=os.path.basename(file_path))
                            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
                            msg.attach(part)
                    else:
                        print(f"Advertencia: No se encontró el archivo {file_path}")
            
            # Conectar al servidor SMTP
            server = smtplib.SMTP(self.config['smtp_server'], self.config['smtp_port'])
            server.starttls()  # Habilitar conexión segura
            server.login(self.config['smtp_user'], self.config['smtp_password'])
            
            # Enviar correo
            server.send_message(msg)
            server.quit()
            
            print(f"Correo enviado correctamente a {', '.join(recipients)}")
            return True
            
        except Exception as e:
            print(f"Error al enviar correo: {e}")
            return False
    
    def test_connection(self) -> bool:
        """
        Prueba la conexión al servidor SMTP.
        
        Returns:
            bool: True si la conexión es exitosa, False en caso contrario
        """
        try:
            server = smtplib.SMTP(self.config['smtp_server'], self.config['smtp_port'])
            server.starttls()
            server.login(self.config['smtp_user'], self.config['smtp_password'])
            server.quit()
            return True
        except Exception as e:
            print(f"Error al conectar con el servidor SMTP: {e}")
            return False
    
    def update_config(self, **kwargs):
        """
        Actualiza la configuración con los valores proporcionados.
        
        Args:
            **kwargs: Pares clave-valor para actualizar la configuración
        """
        for key, value in kwargs.items():
            if key in self.config:
                self.config[key] = value
