import json
import os
from datetime import datetime

class CotizacionFileManager:
    """
    Clase para gestionar el guardado y carga de cotizaciones como archivos separados
    que pueden ser editados fácilmente.
    """
    
    def __init__(self, base_dir="cotizaciones"):
        """
        Inicializa el gestor de archivos de cotización.
        
        Args:
            base_dir (str): Directorio base donde se guardarán las cotizaciones
        """
        self.base_dir = base_dir
        
        # Crear el directorio si no existe
        if not os.path.exists(base_dir):
            os.makedirs(base_dir)
    
    def guardar_cotizacion(self, cotizacion_data, filepath=None):
        """
        Guarda una cotización como un archivo JSON.
        
        Args:
            cotizacion_data (dict): Datos de la cotización a guardar
            filepath (str, optional): Ruta personalizada donde guardar el archivo.
                                     Si no se proporciona, se genera automáticamente.
            
        Returns:
            str: Ruta del archivo guardado
        """
        if filepath is None:
            # Generar nombre de archivo basado en cliente y fecha
            cliente_nombre = cotizacion_data.get('cliente', {}).get('nombre', 'cliente')
            fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
            numero = cotizacion_data.get('numero', '000')
            
            # Sanitizar nombre de cliente para el archivo
            cliente_nombre = ''.join(c if c.isalnum() or c in [' ', '_'] else '_' for c in cliente_nombre)
            cliente_nombre = cliente_nombre.replace(' ', '_')
            
            filename = f"cotizacion_{numero}_{cliente_nombre}_{fecha}.json"
            filepath = os.path.join(self.base_dir, filename)
        
        # Asegurar que el directorio exista
        os.makedirs(os.path.dirname(os.path.abspath(filepath)), exist_ok=True)
        
        # Guardar como JSON
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(cotizacion_data, f, ensure_ascii=False, indent=4)
        
        return filepath
    
    def cargar_cotizacion(self, filepath):
        """
        Carga una cotización desde un archivo JSON.
        
        Args:
            filepath (str): Ruta del archivo a cargar
            
        Returns:
            dict: Datos de la cotización cargada
        """
        with open(filepath, 'r', encoding='utf-8') as f:
            cotizacion_data = json.load(f)
        
        return cotizacion_data
    
    def listar_cotizaciones(self):
        """
        Lista todas las cotizaciones disponibles.
        
        Returns:
            list: Lista de diccionarios con información de las cotizaciones
        """
        cotizaciones = []
        
        for filename in os.listdir(self.base_dir):
            if filename.endswith('.json'):
                filepath = os.path.join(self.base_dir, filename)
                try:
                    # Intentar cargar los datos básicos
                    with open(filepath, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    
                    # Extraer información relevante
                    info = {
                        'filename': filename,
                        'filepath': filepath,
                        'cliente': data.get('cliente', {}).get('nombre', 'Desconocido'),
                        'fecha': data.get('fecha', 'Desconocida'),
                        'numero': data.get('numero', '000'),
                        'total': data.get('total', 0)
                    }
                    cotizaciones.append(info)
                except Exception as e:
                    # Si hay error al cargar, incluir información básica
                    cotizaciones.append({
                        'filename': filename,
                        'filepath': filepath,
                        'error': str(e)
                    })
        
        # Ordenar por fecha de modificación (más reciente primero)
        cotizaciones.sort(key=lambda x: os.path.getmtime(x['filepath']), reverse=True)
        
        return cotizaciones
