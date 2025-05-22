
from dataclasses import dataclass

@dataclass
class Cliente:
    nombre: str
    tipo: str
    direccion: str
    nit: int
    telefono: str
    email: str