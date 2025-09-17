import speech_recognition as sr
from PyQt5.QtCore import QThread, pyqtSignal
import time
import re

class VoiceRecognition(QThread):
    """
    Clase para manejar el reconocimiento de voz en un hilo separado.
    Permite controlar la aplicación mediante comandos de voz.
    """
    
    # Señales para comunicarse con la interfaz gráfica
    command_detected = pyqtSignal(str)
    activity_suggestion = pyqtSignal(list)
    status_update = pyqtSignal(str)
    
    def __init__(self, main_window, controller):
        """
        Inicializa el reconocimiento de voz.
        
        Args:
            main_window: Referencia a la ventana principal
            controller: Controlador de cotizaciones
        """
        super().__init__()
        self.main_window = main_window
        self.controller = controller
        self.recognizer = sr.Recognizer()
        self.is_listening = False
        self.commands = {
            'nueva cotización': self._nueva_cotizacion,
            'agregar actividad': self._agregar_actividad,
            'buscar actividad': self._buscar_actividad,
            'seleccionar cliente': self._seleccionar_cliente,
            'guardar cotización': self._guardar_cotizacion,
            'generar excel': self._generar_excel,
            'generar word': self._generar_word,
            'enviar por correo': self._enviar_correo,
            'ayuda': self._mostrar_ayuda
        }
    
    def run(self):
        """Método principal que se ejecuta en el hilo"""
        self.is_listening = True
        self.status_update.emit("Asistente de voz activado")
        
        while self.is_listening:
            try:
                with sr.Microphone() as source:
                    self.status_update.emit("Escuchando...")
                    self.recognizer.adjust_for_ambient_noise(source, duration=0.5)
                    audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=5)
                
                self.status_update.emit("Procesando comando...")
                text = self.recognizer.recognize_google(audio, language="es-ES")
                self.status_update.emit(f"Comando detectado: {text}")
                
                # Procesar el comando
                self._process_command(text.lower())
                
            except sr.WaitTimeoutError:
                # Timeout, seguir escuchando
                pass
            except sr.UnknownValueError:
                # No se pudo entender el audio
                self.status_update.emit("No se pudo entender el comando")
            except sr.RequestError as e:
                # Error en la solicitud a Google
                self.status_update.emit(f"Error en el servicio de reconocimiento: {e}")
            except Exception as e:
                # Otro error
                self.status_update.emit(f"Error: {e}")
            
            # Pequeña pausa para no saturar la CPU
            time.sleep(0.1)
    
    def stop(self):
        """Detiene el reconocimiento de voz"""
        self.is_listening = False
        self.status_update.emit("Asistente de voz desactivado")
    
    def _process_command(self, text):
        """
        Procesa el comando de voz detectado.
        
        Args:
            text: Texto del comando detectado
        """
        # Verificar si el texto coincide con algún comando conocido
        for command, function in self.commands.items():
            if command in text:
                function(text)
                return
        
        # Si no coincide con ningún comando conocido, buscar actividades
        if "actividad" in text or "actividades" in text:
            self._buscar_actividad(text)
        else:
            self.status_update.emit("Comando no reconocido")
    
    def _nueva_cotizacion(self, text):
        """Inicia una nueva cotización"""
        self.command_detected.emit("nueva_cotizacion")
        self.status_update.emit("Iniciando nueva cotización")
    
    def _agregar_actividad(self, text):
        """Agrega una actividad a la cotización actual"""
        # Extraer detalles de la actividad del texto
        description_match = re.search(r'actividad\s+(.+?)(?:\s+cantidad|\s+por|\s+con|\s+$)', text)
        quantity_match = re.search(r'cantidad\s+(\d+(?:\.\d+)?)', text)
        
        if description_match:
            description = description_match.group(1).strip()
            # Buscar actividades que coincidan con la descripción
            activities = self.controller.get_activity_suggestions_for_voice(description)
            
            if activities:
                # Enviar sugerencias a la interfaz
                self.activity_suggestion.emit(activities)
                
                # Si se especificó una cantidad, establecerla
                if quantity_match:
                    quantity = float(quantity_match.group(1))
                    self.command_detected.emit(f"establecer_cantidad:{quantity}")
            else:
                self.status_update.emit(f"No se encontraron actividades para: {description}")
        else:
            self.status_update.emit("Por favor, especifique la actividad")
    
    def _buscar_actividad(self, text):
        """Busca actividades según el texto"""
        # Extraer términos de búsqueda
        search_terms = text.replace("buscar", "").replace("actividad", "").replace("actividades", "").strip()
        
        if search_terms:
            # Buscar actividades
            activities = self.controller.get_activity_suggestions_for_voice(search_terms)
            
            if activities:
                # Enviar sugerencias a la interfaz
                self.activity_suggestion.emit(activities)
                self.status_update.emit(f"Se encontraron {len(activities)} actividades")
            else:
                self.status_update.emit(f"No se encontraron actividades para: {search_terms}")
        else:
            self.status_update.emit("Por favor, especifique términos de búsqueda")
    
    def _seleccionar_cliente(self, text):
        """Selecciona un cliente para la cotización"""
        # Extraer nombre del cliente
        client_name = text.replace("seleccionar", "").replace("cliente", "").strip()
        
        if client_name:
            self.command_detected.emit(f"seleccionar_cliente:{client_name}")
        else:
            self.status_update.emit("Por favor, especifique el nombre del cliente")
    
    def _guardar_cotizacion(self, text):
        """Guarda la cotización actual"""
        self.command_detected.emit("guardar_cotizacion")
        self.status_update.emit("Guardando cotización")
    
    def _generar_excel(self, text):
        """Genera el archivo Excel de la cotización"""
        self.command_detected.emit("generar_excel")
        self.status_update.emit("Generando Excel")
    
    def _generar_word(self, text):
        """Genera el documento Word de la cotización"""
        # Extraer tipo de documento
        if "largo" in text or "completo" in text:
            doc_type = "largo"
        elif "corto" in text:
            doc_type = "corto"
        else:
            doc_type = "default"
        
        self.command_detected.emit(f"generar_word:{doc_type}")
        self.status_update.emit(f"Generando documento Word {doc_type}")
    
    def _enviar_correo(self, text):
        """Envía la cotización por correo electrónico"""
        # Extraer dirección de correo si se menciona
        email_match = re.search(r'correo\s+(?:a|para)?\s+(.+@.+\..+)', text)
        
        if email_match:
            email = email_match.group(1)
            self.command_detected.emit(f"enviar_correo:{email}")
        else:
            self.command_detected.emit("enviar_correo")
        
        self.status_update.emit("Preparando envío de correo")
    
    def _mostrar_ayuda(self, text):
        """Muestra la ayuda del asistente de voz"""
        help_text = """
        Comandos disponibles:
        - Nueva cotización
        - Agregar actividad [descripción] cantidad [número]
        - Buscar actividad [términos]
        - Seleccionar cliente [nombre]
        - Guardar cotización
        - Generar Excel
        - Generar Word [largo/corto]
        - Enviar por correo [a correo@ejemplo.com]
        - Ayuda
        """
        self.status_update.emit(help_text)
