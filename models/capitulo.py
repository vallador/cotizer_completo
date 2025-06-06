from dataclasses import dataclass, field
from typing import List, Optional

@dataclass
class Capitulo:
    """
    Clase que representa un capítulo de obra, que agrupa actividades relacionadas.
    """
    nombre: str
    descripcion: str = ""
    id: Optional[int] = None
    orden: int = 0
    
    def __str__(self):
        return self.nombre
