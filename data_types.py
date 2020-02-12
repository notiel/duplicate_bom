from enum import Enum
from typing import List, Union, Optional

from dataclasses import dataclass, field

dielectrics = ['np0', 'x5r', 'x7r']
units_cap = ['u', 'n', 'pf']
types = ['resistor', 'capacitor', 'crystal', 'inductor', 'diode', 'chip', 'module', 'emifil', 'connector', 'button',
         'switch', 'transistor']


class ComponentType(Enum):
    RESISTOR = 0
    CAPACITOR = 1
    CRYSTAL = 2
    INDUCTOR = 3
    DIODE = 4
    CHIP = 5
    MODULE = 6
    EMI_FILTER = 7
    CONNECTOR = 8
    BUTTON = 9
    SWITCH = 10
    TRANSISTOR = 11
    OTHER = 12


@dataclass
class SimpleComponent:
    component_type: ComponentType
    pn: str = ""
    footprint: str = ""
    manufacturer: str = ""
    pn_alt: List[str] = field(default_factory=list)
    designator: List[str] = field(default_factory=list)
    description: str = ""
    details: Optional[Union['Capacitor', 'Inductor', 'Resistor']] = None


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
