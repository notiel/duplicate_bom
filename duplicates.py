# srevice module for finding similar components for find_duplicates script

import data_types
from typing import List, Tuple

# number of symbols that must be equal in pns
root = 8
tail = 4


def compare_pns(components: List[data_types.Component], root_len: int = 8, precise: bool = False) \
        -> Tuple[List[Tuple[str, int, int, str, int, int]], List[Tuple[str, int, int, str, int, int]],
                 List[Tuple[str, int, int, str, int, int]], str]:
    """
    compares all rows and gets similar pns
    :param precise: use precise comparizon of pns only
    :param root_len: len of common components part
    :param components:
    :return: 
    """
    equal: List[Tuple[str, int, int, str, int, int]] = list()
    similar: List[Tuple[str, int, int, str, int, int]] = list()
    alternative: List[Tuple[str, int, int, str, int, int]] = list()
    warning: str = ""
    sorted_components = sorted(components, key=lambda x: x.pn)
    # remove components without pns, we can not compare them by pn
    sorted_components = [component for component in sorted_components if component.pn]
    for (index, component) in enumerate(sorted_components[:-1]):
        next_comp = sorted_components[index + 1]
        if component.pn.lower() == next_comp.pn.lower():
            # similar.append((component.filename, component.row, index, next_comp.filename, next_comp.row, index+1))
            if component.footprint != next_comp.footprint:
                warning += "Components in file %s row %i and  file %s row %i have same partnumbers but different " \
                           "footprint\n" % (component.filename, component.row, next_comp.filename, next_comp.row)
            if component.component_type != next_comp.component_type:
                if component.component_type.name.lower() not in data_types.the_same.keys() \
                        or data_types.the_same[component.component_type.name.lower()] \
                        != next_comp.component_type.name.lower():
                    warning += "Components in file %s row %i and  file %s row %i have same partnumbers but different " \
                               "type\n" % (component.filename, component.row, next_comp.filename, next_comp.row)
                equal.append((component.filename, component.row, index, next_comp.filename, next_comp.row, index+1))
        if not precise:
            if component.component_type not in data_types.parametrized and len(component.footprint) >= 4:
                next_index = index + 1
                while next_index < len(sorted_components) and \
                        sorted_components[next_index].pn.startswith(component.pn[:root_len + 1]):
                    if component.footprint == sorted_components[next_index].footprint:
                        if component.component_type == sorted_components[next_index].component_type:
                            similar.append((component.filename, component.row, index,
                                            sorted_components[next_index].filename, sorted_components[next_index].row,
                                            next_index))
                    next_index += 1
            for alternative_comp in components:
                crosses = [alt for alt in alternative_comp.pn_alt if component.pn in alt]
                if crosses and component.row != alternative_comp.row:
                    if component.component_type == alternative_comp.component_type:
                        if (component.row, next_comp.row) not in equal:
                            alternative.append((component.filename, component.row, index,
                                                alternative_comp.filename, alternative_comp.row,
                                                components.index(alternative_comp)))
    return equal, similar, alternative, warning


def compare_capacitors(components: List[data_types.Component]) -> List[Tuple[str, str, int, str, str, int]]:
    """
    compare capacitors by value and footprint
    :param components:
    :return:
    """
    similar_caps: List[Tuple[str, str, int, str, str, int]] = list()
    cap_sorted: List[data_types.Component] = sorted([component for component in components
                                                     if component.component_type == data_types.ComponentType.CAPACITOR
                                                     and component.details.absolute_pf_value],
                                                    key=lambda x: x.details.absolute_pf_value)
    if len(cap_sorted) < 2:
        return list()
    for (index, cap) in enumerate(cap_sorted):
        next_index = index + 1
        while next_index < len(cap_sorted) and cap_sorted[next_index].details.absolute_pf_value \
                == cap.details.absolute_pf_value and cap_sorted[next_index].footprint == cap.footprint:
            similar_caps.append((cap.filename, str(cap.details.value)+data_types.units_cap[cap.details.unit], cap.row,
                                 cap_sorted[next_index].filename,
                                 str(cap_sorted[next_index].details.value) +
                                 data_types.units_cap[cap_sorted[next_index].details.unit],
                                 cap_sorted[next_index].row))
            next_index += 1
    return similar_caps


def compare_resistors(components: List[data_types.Component]) -> List[Tuple[str, str, int, str, str, int]]:
    """
    compare capacitors by value and footprint
    :param components: list of components
    :return:
    """
    similar_resistors: List[Tuple[str, str, int, str, str, int]] = list()
    res_sorted: List[data_types.Component] = sorted([component for component in components
                                                     if component.component_type == data_types.ComponentType.RESISTOR
                                                     and component.details.value != -1],
                                                    key=lambda x: x.details.value)
    if len(res_sorted) < 2:
        return list()
    for (index, res) in enumerate(res_sorted):
        next_index = index + 1
        while next_index < len(res_sorted) and res_sorted[next_index].details.value == res.details.value \
                and res_sorted[next_index].footprint == res.footprint:
            similar_resistors.append((res.filename, str(res.details.value)+'R', res.row,
                                      res_sorted[next_index].filename, str(res_sorted[next_index].details.value)+'R',
                                      res_sorted[next_index].row))
            next_index += 1
    return similar_resistors
