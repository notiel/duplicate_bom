# module with data tyoes for comparing boms
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


# list of component types that could be compared by their parameters, not pn only
parametrized = [ComponentType.RESISTOR, ComponentType.INDUCTOR, ComponentType.CAPACITOR]


class Dielectric(Enum):
    NP0 = 0
    X5R = 1
    X7R = 2


class CapUnits:
    U = 0
    N = 1
    PF = 2


@dataclass
class Component:
    row: int
    component_type: ComponentType
    pn: str = ""
    footprint: str = ""
    manufacturer: str = ""
    filename = ""
    pn_alt: List[str] = field(default_factory=list)
    designator: List[str] = field(default_factory=list)
    description: str = ""
    details: Optional[Union['Capacitor', 'Inductor', 'Resistor']] = None


@dataclass
class Capacitor:
    value: float
    unit: CapUnits
    absolute_pf_value: Optional[Union[int, float]]  # value in pfs
    dielectric: List[Dielectric]
    voltage: Optional[float] = 6.3
    tolerance: Optional[float] = 1.0


@dataclass
class Resistor:
    value: float
    tolerance: 1


@dataclass
class Inductor:
    value: str
