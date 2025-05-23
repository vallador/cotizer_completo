from dataclasses import dataclass, field
from typing import List
from .producto import Producto

@dataclass
class Actividad:
    descripcion: str
    unidad: str
    cantidad: float
    valor_unitario: float
    productos: List[Producto] = field(default_factory=list)
    id: int = None
    categoria: str = None

    @property
    def valor_total(self):
        return self.cantidad * self.valor_unitario
        
    def calcular_total(self):
        """Calcula el valor total de la actividad (cantidad * valor unitario)"""
        return self.cantidad * self.valor_unitario
