from dataclasses import dataclass
from typing import Optional

@dataclass
class Cliente:
    nombre: str
    tipo: str
    direccion: str
    nit: Optional[str] = None
    telefono: Optional[str] = None
    email: Optional[str] = None
    id: Optional[int] = None
