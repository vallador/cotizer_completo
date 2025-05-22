from dataclasses import dataclass

@dataclass
class Producto:
    nombre: str
    marca: str = None
    descripcion: str = None