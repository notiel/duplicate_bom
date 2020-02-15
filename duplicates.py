import xlsx_parce
import data_types
from typing import List, Tuple


def compare_pns(components: List[data_types.Component]):
    """
    
    :param components: 
    :return: 
    """
    equal: List[Tuple[int, int]] = list()
    similar: List[Tuple[int, int]] = list()
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
        if component.component_type not in data_types.parametrized and len(component.footprint) >= 4:
            next_index = index + 1
            while next_index < len(sorted_components) and sorted_components[next_index].pn.startswith(component.pn[:7]):
                if component.footprint == sorted_components[next_index].footprint:
                    if component.component_type == sorted_components[next_index].component_type:
                        similar.append((component.row, sorted_components[next_index].row))
                next_index += 1
    print(similar)
    print(equal)


components_list: List[data_types.Component] = xlsx_parce.get_components_from_xlxs('BOM_Test.xlsx')
compare_pns(components_list)
