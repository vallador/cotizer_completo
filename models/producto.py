from dataclasses import dataclass
from typing import Optional

@dataclass
class Producto:
    nombre: str
    unidad: str
    precio_unitario: float
    descripcion: Optional[str] = None
    marca: Optional[str] = None
    categoria: Optional[str] = None
    id: Optional[int] = None
