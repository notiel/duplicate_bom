import xlsx_parce
import data_types
from typing import List, Tuple

# number of symbols that must be equal in pns
root = 8
tail = 4


def compare_pns(components: List[data_types.Component]):
    """
    compares all rows and gets similar pns
    :param components: 
    :return: 
    """
    equal: List[Tuple[int, int]] = list()
    similar: List[Tuple[int, int]] = list()
    alternative: List[Tuple[int, int]] = list()
    warning: str = ""
    sorted_components = sorted(components, key=lambda x: x.pn)
    # remove components without pns, we can not compare them by pn
    sorted_components = [component for component in sorted_components if component.pn]
    for (index, component) in enumerate(sorted_components[:-1]):
        next_comp = sorted_components[index+1]
        if component.pn.lower() == next_comp.pn.lower():
            similar.append((component.row, next_comp.row))
            if component.component_type != next_comp.component_type or component.footprint != next_comp.footprint:
                warning += "Components in rows %i and %i have same partnumbers but different type or footprint" % \
                           (component.row, next_comp.row)
            equal.append((component.row, next_comp.row))
        if component.component_type not in data_types.parametrized and len(component.footprint) >= 4:
            next_index = index + 1
            while next_index < len(sorted_components) and \
                    sorted_components[next_index].pn.startswith(component.pn[:root+1]):
                if component.footprint == sorted_components[next_index].footprint:
                    if component.component_type == sorted_components[next_index].component_type:
                        similar.append((component.row, sorted_components[next_index].row))
                next_index += 1
        for alternative_comp in components:
            crosses = [alt for alt in alternative_comp.pn_alt if component.pn in alt]
            if crosses and component.row != alternative_comp.row:
                if component.component_type == alternative_comp.component_type:
                    if (component.row, next_comp.row) not in equal:
                        alternative.append((component.row, alternative_comp.row))
    print("These rows have equal pns")
    print(equal)
    print("These rows have similar pns")
    print(similar)
    print('These rows have similar alternative pn')
    print(alternative)
    print(warning)


def compare_capacitors(components: List[data_types.Component]):
    """
    compare capacitors by value and footprint
    :param components:
    :return:
    """
    similar_caps: List[Tuple[int, int]] = list()
    cap_sorted: List[data_types.Component] = sorted([component for component in components
                                                     if component.component_type == data_types.ComponentType.CAPACITOR
                                                     and component.details.absolute_pf_value],
                                                    key=lambda x: x.details.absolute_pf_value)
    if len(cap_sorted) < 2:
        return
    for (index, cap) in enumerate(cap_sorted):
        next_index = index + 1
        while next_index < len(cap_sorted) and cap_sorted[next_index].details.absolute_pf_value \
                == cap.details.absolute_pf_value and cap_sorted[next_index].footprint == cap.footprint:
            similar_caps.append((cap.row, cap_sorted[next_index].row))
            next_index += 1
    print("Similar capacitor rows: ")
    print(similar_caps)


components_list: List[data_types.Component] = xlsx_parce.get_components_from_xlxs('BOM_Test.xlsx')
compare_pns(components_list)
compare_capacitors(components_list)
