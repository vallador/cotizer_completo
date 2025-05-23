# controllers/cotizacion_controller.py

from models.cliente import Cliente
from models.cotizacion import Cotizacion
from models.actividad import Actividad
from models.producto import Producto
from utils.database_manager import DatabaseManager
from utils.aiu_manager import AIUManager
from utils.filter_manager import FilterManager
from typing import List, Dict
import datetime
import os
import json


class CotizacionController:
    def __init__(self):
        self.database_manager = DatabaseManager('data/cotizaciones.db')
        self.database_manager.connect()
        self.aiu_manager = AIUManager()
        self.filter_manager = FilterManager(self.database_manager)
        self.cargar_configuracion()

    def cargar_configuracion(self):
        """Carga la configuración desde el archivo config.json"""
        config_path = os.path.join(os.getcwd(), 'config.json')
        if os.path.exists(config_path) and os.path.getsize(config_path) > 0:
            try:
                with open(config_path, 'r') as f:
                    config = json.load(f)
                    if 'aiu' in config:
                        self.aiu_manager.modificar_aiu(
                            config['aiu'].get('administracion', 7),
                            config['aiu'].get('imprevistos', 3),
                            config['aiu'].get('utilidad', 5),
                            config['aiu'].get('iva_sobre_utilidad', 19)
                        )
                    if 'email' in config:
                        self.email_config = config['email']
            except Exception as e:
                print(f"Error al cargar la configuración: {e}")
                # Usar valores predeterminados si hay error

    def guardar_configuracion(self):
        """Guarda la configuración actual en el archivo config.json"""
        config_path = os.path.join(os.getcwd(), 'config.json')
        config = {
            'aiu': self.aiu_manager.obtener_aiu()
        }
        try:
            with open(config_path, 'w') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            print(f"Error al guardar la configuración: {e}")

    def crear_cliente(self, datos_cliente: Dict) -> Cliente:
        """Crea un objeto Cliente a partir de un diccionario de datos"""
        return Cliente(**datos_cliente)

    def crear_actividad(self, datos_actividad: Dict) -> Actividad:
        """Crea un objeto Actividad a partir de un diccionario de datos"""
        productos = [Producto(**p) for p in datos_actividad.get('productos', [])]
        return Actividad(**{**datos_actividad, 'productos': productos})

    def crear_cotizacion(self, datos_cotizacion: Dict) -> Cotizacion:
        """Crea un objeto Cotización a partir de un diccionario de datos"""
        # Obtener el cliente
        if 'cliente_id' in datos_cotizacion:
            cliente_data = self.database_manager.get_client_by_id(datos_cotizacion['cliente_id'])
            cliente = self.crear_cliente(cliente_data)
        else:
            cliente = self.crear_cliente(datos_cotizacion['cliente'])
        
        # Crear actividades
        actividades = []
        for act_data in datos_cotizacion.get('actividades', []):
            if 'actividad_id' in act_data:
                # Si se proporciona un ID de actividad, obtener los datos de la BD
                actividad_base = self.database_manager.get_activity_by_id(act_data['actividad_id'])
                # Combinar con los datos proporcionados (cantidad, etc.)
                actividad_data = {**actividad_base, **act_data}
                actividad = self.crear_actividad(actividad_data)
                actividad.cantidad = act_data.get('cantidad', 1)
                actividad.calcular_total()
            else:
                # Si se proporcionan todos los datos de la actividad
                actividad = self.crear_actividad(act_data)
            
            actividades.append(actividad)
        
        # Crear la cotización
        cotizacion = Cotizacion(
            numero=datos_cotizacion.get('numero', self.generar_numero_cotizacion()),
            fecha=datos_cotizacion.get('fecha', datetime.date.today()),
            cliente=cliente,
            actividades=actividades,
            administracion=datos_cotizacion.get('administracion', self.aiu_manager.administracion),
            imprevistos=datos_cotizacion.get('imprevistos', self.aiu_manager.imprevistos),
            utilidad=datos_cotizacion.get('utilidad', self.aiu_manager.utilidad),
            iva_utilidad=datos_cotizacion.get('iva_utilidad', self.aiu_manager.iva_sobre_utilidad)
        )
        
        # Calcular el total
        cotizacion.calcular_total()
        
        return cotizacion

    def generar_numero_cotizacion(self) -> str:
        """Genera un número único para la cotización"""
        # Formato: COT-YYYYMMDD-XXX donde XXX es un número secuencial
        fecha = datetime.date.today().strftime("%Y%m%d")
        
        # Obtener todas las cotizaciones para contar
        cotizaciones = self.database_manager.get_all_quotations()
        secuencial = len(cotizaciones) + 1
        
        return f"COT-{fecha}-{secuencial:03d}"

    def guardar_cotizacion(self, cotizacion: Cotizacion) -> int:
        """Guarda una cotización en la base de datos"""
        # Convertir la cotización a diccionario
        cotizacion_dict = cotizacion.to_dict()
        
        # Preparar las actividades para guardar
        actividades = []
        for actividad in cotizacion.actividades:
            actividades.append({
                'actividad_id': actividad.id if hasattr(actividad, 'id') else None,
                'cantidad': actividad.cantidad,
                'valor_unitario': actividad.valor_unitario,
                'total': actividad.calcular_total()  # Usar el método en lugar de la propiedad
            })
        
        # Guardar en la base de datos
        return self.database_manager.add_quotation(cotizacion_dict, actividades)

    def obtener_cotizacion(self, cotizacion_id: int) -> Cotizacion:
        """Obtiene una cotización de la base de datos por su ID"""
        cotizacion_data = self.database_manager.get_quotation_by_id(cotizacion_id)
        if not cotizacion_data:
            return None
        
        # Obtener el cliente
        cliente_data = self.database_manager.get_client_by_id(cotizacion_data['cliente_id'])
        cliente = self.crear_cliente(cliente_data)
        
        # Crear actividades a partir de los detalles
        actividades = []
        for detalle in cotizacion_data.get('detalles', []):
            actividad_data = {
                'id': detalle['actividad_id'],
                'descripcion': detalle['descripcion'],
                'unidad': detalle['unidad'],
                'valor_unitario': detalle['valor_unitario'],
                'cantidad': detalle['cantidad']
            }
            actividad = self.crear_actividad(actividad_data)
            actividades.append(actividad)
        
        # Crear la cotización
        cotizacion = Cotizacion(
            numero=cotizacion_data['numero'],
            fecha=cotizacion_data['fecha'],
            cliente=cliente,
            actividades=actividades,
            subtotal=cotizacion_data['subtotal'],
            iva=cotizacion_data['iva'],
            total=cotizacion_data['total'],
            administracion=cotizacion_data.get('administracion', 7),
            imprevistos=cotizacion_data.get('imprevistos', 3),
            utilidad=cotizacion_data.get('utilidad', 5),
            iva_utilidad=cotizacion_data.get('iva_utilidad', 19)
        )
        
        return cotizacion

    def obtener_todas_cotizaciones(self) -> List[Dict]:
        """Obtiene todas las cotizaciones de la base de datos"""
        return self.database_manager.get_all_quotations()

    def eliminar_cotizacion(self, cotizacion_id: int) -> bool:
        """Elimina una cotización de la base de datos"""
        try:
            self.database_manager.delete_quotation(cotizacion_id)
            return True
        except Exception as e:
            print(f"Error al eliminar cotización: {e}")
            return False

    def add_client(self, client_data):
        """Añade un cliente a la base de datos"""
        return self.database_manager.add_client(client_data)

    def get_all_clients(self):
        """Obtiene todos los clientes de la base de datos"""
        return self.database_manager.get_all_clients()

    def get_all_activities(self):
        """Obtiene todas las actividades de la base de datos"""
        return self.database_manager.get_all_activities()
    
    def get_all_products(self):
        """Obtiene todos los productos de la base de datos"""
        return self.database_manager.get_all_products()
    
    def get_all_categories(self):
        """Obtiene todas las categorías de la base de datos"""
        return self.database_manager.get_all_categories()

    def add_activity(self, activity_data):
        """Añade una actividad a la base de datos"""
        return self.database_manager.add_activity(activity_data)

    def update_activity(self, activity_id, activity_data):
        """Actualiza una actividad en la base de datos"""
        return self.database_manager.update_activity(activity_id, activity_data)

    def delete_activity(self, activity_id):
        """Elimina una actividad de la base de datos"""
        return self.database_manager.delete_activity(activity_id)

    def generate_quotation(self, client_id, activities, aiu_values=None):
        """Genera una cotización a partir de un cliente y actividades"""
        # Obtener el cliente
        cliente_data = self.database_manager.get_client_by_id(client_id)
        if not cliente_data:
            raise ValueError("Cliente no encontrado")
        
        cliente = self.crear_cliente(cliente_data)
        
        # Crear actividades
        actividades_obj = []
        for act in activities:
            actividad = Actividad(
                descripcion=act['descripcion'],
                unidad=act['unidad'],
                valor_unitario=float(act['valor_unitario']),
                cantidad=float(act['cantidad']),
                id=act.get('id')
            )
            actividades_obj.append(actividad)
        
        # Valores AIU
        aiu = aiu_values or self.aiu_manager.obtener_aiu()
        
        # Crear la cotización
        cotizacion = Cotizacion(
            numero=self.generar_numero_cotizacion(),
            fecha=datetime.date.today(),
            cliente=cliente,
            actividades=actividades_obj,
            administracion=aiu['administracion'],
            imprevistos=aiu['imprevistos'],
            utilidad=aiu['utilidad'],
            iva_utilidad=aiu['iva_sobre_utilidad']
        )
        
        # Calcular el total
        cotizacion.calcular_total()
        
        # Guardar en la base de datos
        cotizacion_id = self.guardar_cotizacion(cotizacion)
        
        return {
            'id': cotizacion_id,
            'numero': cotizacion.numero,
            'subtotal': cotizacion.subtotal,
            'iva': cotizacion.iva,
            'total': cotizacion.total,
            'desglose': cotizacion.obtener_desglose()
        }
    
    # Métodos para filtrado y relaciones
    def search_activities(self, search_text, category_id=None):
        """
        Busca actividades que coincidan con el texto de búsqueda y opcionalmente
        filtra por categoría.
        
        Args:
            search_text: Texto a buscar en las descripciones de actividades
            category_id: ID de categoría para filtrar (opcional)
            
        Returns:
            Lista de actividades que coinciden con los criterios de búsqueda
        """
        return self.filter_manager.search_activities(search_text, category_id)
    
    def search_products(self, search_text, category_id=None):
        """
        Busca productos que coincidan con el texto de búsqueda y opcionalmente
        filtra por categoría.
        
        Args:
            search_text: Texto a buscar en los nombres y descripciones de productos
            category_id: ID de categoría para filtrar (opcional)
            
        Returns:
            Lista de productos que coinciden con los criterios de búsqueda
        """
        return self.filter_manager.search_products(search_text, category_id)
    
    def get_related_activities(self, activity_id):
        """
        Obtiene las actividades relacionadas con una actividad específica.
        
        Args:
            activity_id: ID de la actividad principal
            
        Returns:
            Lista de actividades relacionadas
        """
        return self.filter_manager.get_related_activities(activity_id)
    
    def suggest_activities(self, context):
        """
        Sugiere actividades basadas en un contexto (por ejemplo, "pintura").
        
        Args:
            context: Contexto para sugerir actividades
            
        Returns:
            Lista de actividades sugeridas
        """
        return self.filter_manager.suggest_activities(context)
    
    def get_products_for_activity(self, activity_id):
        """
        Obtiene los productos asociados a una actividad específica.
        
        Args:
            activity_id: ID de la actividad
            
        Returns:
            Lista de productos asociados a la actividad
        """
        return self.filter_manager.get_products_for_activity(activity_id)
    
    def get_activity_suggestions_for_voice(self, context):
        """
        Obtiene sugerencias de actividades optimizadas para reconocimiento de voz.
        
        Args:
            context: Contexto o frase reconocida por voz
            
        Returns:
            Lista de actividades sugeridas
        """
        return self.filter_manager.get_activity_suggestions_for_voice(context)

    def disconnect(self):
        """Desconecta de la base de datos"""
        self.database_manager.disconnect()
