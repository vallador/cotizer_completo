from dataclasses import dataclass, field
from datetime import date
from typing import List
from .cliente import Cliente
from .actividad import Actividad
from utils.aiu_manager import AIUManager


@dataclass
class Cotizacion:
    numero: str
    fecha: date
    cliente: Cliente
    actividades: List[Actividad] = field(default_factory=list)
    subtotal: float = 0
    iva: float = 0
    total: float = 0
    administracion: float = 7
    imprevistos: float = 3
    utilidad: float = 5
    iva_utilidad: float = 19
    
    def __post_init__(self):
        # Inicializar el gestor AIU con los valores proporcionados
        self.aiu_manager = AIUManager(
            administracion=self.administracion,
            imprevistos=self.imprevistos,
            utilidad=self.utilidad,
            iva_sobre_utilidad=self.iva_utilidad
        )

    def calcular_total(self):
        # Calcular el subtotal sumando el valor total de cada actividad
        self.subtotal = sum(actividad.valor_total for actividad in self.actividades)
        
        # Aplicar AIU si el cliente es jurídico
        if self.cliente.tipo == 'jurídica':
            admin_value, imprevistos_value, utilidad_value, iva_value = self.aiu_manager.calcular_aiu(self.subtotal)
            self.total = self.subtotal + admin_value + imprevistos_value + utilidad_value + iva_value
        else:
            # Para clientes naturales, solo aplicar IVA general
            self.iva = self.subtotal * (self.iva_utilidad / 100)
            self.total = self.subtotal + self.iva
        
        return self.total
    
    def actualizar_valores_aiu(self, administracion, imprevistos, utilidad, iva_utilidad):
        """Actualiza los valores de AIU y recalcula los totales"""
        self.administracion = administracion
        self.imprevistos = imprevistos
        self.utilidad = utilidad
        self.iva_utilidad = iva_utilidad
        
        # Actualizar el gestor AIU
        self.aiu_manager.modificar_aiu(administracion, imprevistos, utilidad, iva_utilidad)
        
        # Recalcular totales
        return self.calcular_total()
    
    def obtener_desglose(self):
        """Retorna un desglose detallado de los valores de la cotización"""
        desglose = {
            'subtotal': self.subtotal,
            'iva': self.iva,
            'total': self.total
        }
        
        # Si es cliente jurídico, incluir desglose de AIU
        if self.cliente.tipo == 'jurídica':
            admin_value, imprevistos_value, utilidad_value, iva_value = self.aiu_manager.calcular_aiu(self.subtotal)
            desglose.update({
                'administracion_porcentaje': self.administracion,
                'administracion_valor': admin_value,
                'imprevistos_porcentaje': self.imprevistos,
                'imprevistos_valor': imprevistos_value,
                'utilidad_porcentaje': self.utilidad,
                'utilidad_valor': utilidad_value,
                'iva_utilidad_porcentaje': self.iva_utilidad,
                'iva_utilidad_valor': iva_value
            })
        
        return desglose
    
    def to_dict(self):
        """Convierte la cotización a un diccionario para almacenamiento en BD"""
        return {
            'numero': self.numero,
            'fecha': self.fecha.isoformat() if isinstance(self.fecha, date) else self.fecha,
            'cliente_id': self.cliente.id if hasattr(self.cliente, 'id') else None,
            'subtotal': self.subtotal,
            'iva': self.iva,
            'total': self.total,
            'administracion': self.administracion,
            'imprevistos': self.imprevistos,
            'utilidad': self.utilidad,
            'iva_utilidad': self.iva_utilidad
        }
