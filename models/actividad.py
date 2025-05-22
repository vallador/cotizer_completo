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

    @property
    def valor_total(self):
        return self.cantidad * self.valor_unitario
