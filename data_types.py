from enum import Enum
from typing import List

from dataclasses import dataclass, field

dielectrics = ['np0', 'x5r', 'x7r']
units_cap = ['u', 'n', 'pf']


class ComponentType(Enum):
    RESISTOR = 0
    CAPACITOR = 1
    CRYSTAL = 3
    INDUCTOR = 3
    DIOD = 4
    CHIP = 5
    MODULE = 6
    EMI_FILTER = 7
    CONNECTOR = 8
    BUTTON = 9
    SWITCH = 10
    TRANSISTOR = 11


@dataclass
class SimpleComponent:
    component_type: ComponentType
    pn: str = ""
    footprint: str = ""
    manufacturer: str = ""
    pn_alt: List[str] = field(default_factory=list)
    designator: List[str] = field(default_factory=list)
    description: str = ""


@dataclass
class Capacitor:
    value: float
    unit: str
    dielectric: str = ""
    voltage: float = 6.3
    tolerance: int = 100


@dataclass
class Resistor:
    value: float
    tolerance: 1


@dataclass
class Inductor:
    value: str
