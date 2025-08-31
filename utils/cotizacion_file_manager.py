import os
import json
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
        try:
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

        except Exception as e:
            print(f"Error guardando cotización: {e}")
            raise e

    def cargar_cotizacion(self, filepath):
        """
        Carga una cotización desde un archivo JSON con manejo robusto de errores.

        Args:
            filepath (str): Ruta del archivo a cargar

        Returns:
            dict: Datos de la cotización cargada o None si hay error
        """
        try:
            print(f"Intentando cargar: {filepath}")  # Debug temporal

            # Verificar que el archivo existe
            if not os.path.exists(filepath):
                print(f"Error: El archivo no existe: {filepath}")
                raise FileNotFoundError(f"El archivo no existe: {filepath}")

            # Verificar que no esté vacío
            file_size = os.path.getsize(filepath)
            if file_size == 0:
                print(f"Error: El archivo está vacío: {filepath}")
                raise ValueError(f"El archivo está vacío: {filepath}")

            print(f"Tamaño del archivo: {file_size} bytes")  # Debug temporal

            # Intentar leer con diferentes codificaciones
            encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252']

            for encoding in encodings:
                try:
                    print(f"Intentando con codificación: {encoding}")  # Debug temporal

                    with open(filepath, 'r', encoding=encoding) as f:
                        content = f.read().strip()

                        if not content:
                            raise ValueError("El archivo está vacío después de leer")

                        print(f"Primeros 100 caracteres: {repr(content[:100])}")  # Debug temporal

                        # Parsear JSON
                        cotizacion_data = json.loads(content)

                        # Validar que sea un diccionario
                        if not isinstance(cotizacion_data, dict):
                            raise ValueError(f"El contenido no es un objeto JSON válido, es: {type(cotizacion_data)}")

                        # Validaciones básicas de estructura
                        if 'cliente' not in cotizacion_data:
                            print("Advertencia: El archivo no contiene información de cliente")

                        print(f"Cotización cargada exitosamente con {encoding}. Claves: {list(cotizacion_data.keys())}")
                        return cotizacion_data

                except UnicodeDecodeError:
                    print(f"Error de codificación con {encoding}, intentando siguiente...")
                    continue
                except json.JSONDecodeError as e:
                    print(f"Error JSON con {encoding}: {e}")
                    if encoding == encodings[-1]:  # Si es la última codificación, relanzar el error
                        raise e
                    continue

            # Si llegamos aquí, ninguna codificación funcionó
            raise ValueError("No se pudo leer el archivo con ninguna codificación soportada")

        except json.JSONDecodeError as e:
            print(f"Error de formato JSON en {filepath}: {e}")
            raise ValueError(f"El archivo no tiene formato JSON válido: {str(e)}")

        except FileNotFoundError as e:
            print(f"Archivo no encontrado: {e}")
            raise e

        except PermissionError as e:
            print(f"Sin permisos para leer: {e}")
            raise e

        except Exception as e:
            print(f"Error inesperado cargando {filepath}: {e}")
            raise e

    def listar_cotizaciones(self):
        """
        Lista todas las cotizaciones disponibles.

        Returns:
            list: Lista de diccionarios con información de las cotizaciones
        """
        cotizaciones = []

        # Verificar que el directorio existe
        if not os.path.exists(self.base_dir):
            print(f"Directorio base no existe: {self.base_dir}")
            return cotizaciones

        # Buscar archivos .json y .cotiz
        valid_extensions = ['.json', '.cotiz']

        for filename in os.listdir(self.base_dir):
            if any(filename.endswith(ext) for ext in valid_extensions):
                filepath = os.path.join(self.base_dir, filename)
                try:
                    # Intentar cargar los datos básicos usando nuestro método mejorado
                    data = self.cargar_cotizacion(filepath)

                    if data:  # Solo si se cargó correctamente
                        # Extraer información relevante
                        info = {
                            'filename': filename,
                            'filepath': filepath,
                            'cliente': data.get('cliente', {}).get('nombre', 'Desconocido') if isinstance(
                                data.get('cliente'), dict) else 'Cliente inválido',
                            'fecha': data.get('fecha', 'Desconocida'),
                            'numero': data.get('numero', '000'),
                            'total': data.get('total', 0)
                        }
                        cotizaciones.append(info)
                    else:
                        # Si cargar_cotizacion retorna None
                        cotizaciones.append({
                            'filename': filename,
                            'filepath': filepath,
                            'error': 'No se pudo cargar el archivo'
                        })

                except Exception as e:
                    # Si hay error al cargar, incluir información básica
                    print(f"Error listando {filename}: {e}")
                    cotizaciones.append({
                        'filename': filename,
                        'filepath': filepath,
                        'error': str(e)
                    })

        # Ordenar por fecha de modificación (más reciente primero)
        try:
            cotizaciones.sort(key=lambda x: os.path.getmtime(x['filepath']), reverse=True)
        except Exception as e:
            print(f"Error ordenando cotizaciones: {e}")

        return cotizaciones