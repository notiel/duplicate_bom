import xlsx_parce
import data_types
from typing import List, Tuple
import sys
import os

# number of symbols that must be equal in pns
root = 8
tail = 4


def compare_pns(components: List[data_types.Component]):
    """
    compares all rows and gets similar pns
    :param components: 
    :return: 
    """
    equal: List[Tuple[str, int, str, int]] = list()
    similar: List[Tuple[str, int, str, int]] = list()
    alternative: List[Tuple[str, int, str, int]] = list()
    warning: str = ""
    sorted_components = sorted(components, key=lambda x: x.pn)
    # remove components without pns, we can not compare them by pn
    sorted_components = [component for component in sorted_components if component.pn]
    for (index, component) in enumerate(sorted_components[:-1]):
        next_comp = sorted_components[index + 1]
        if component.pn.lower() == next_comp.pn.lower():
            similar.append((component.filename, component.row, next_comp.filename, next_comp.row))
            if component.component_type != next_comp.component_type or component.footprint != next_comp.footprint:
                warning += "Components in file %s row %i and  file %s row %i have same partnumbers but different " \
                           "type or footprint" % (component.filename, component.row, next_comp.filename, next_comp.row)
            equal.append((component.filename, component.row, next_comp.filename, next_comp.row))
        if component.component_type not in data_types.parametrized and len(component.footprint) >= 4:
            next_index = index + 1
            while next_index < len(sorted_components) and \
                    sorted_components[next_index].pn.startswith(component.pn[:root + 1]):
                if component.footprint == sorted_components[next_index].footprint:
                    if component.component_type == sorted_components[next_index].component_type:
                        similar.append((component.filename, component.row, sorted_components[next_index].filename,
                                        sorted_components[next_index].row))
                next_index += 1
        for alternative_comp in components:
            crosses = [alt for alt in alternative_comp.pn_alt if component.pn in alt]
            if crosses and component.row != alternative_comp.row:
                if component.component_type == alternative_comp.component_type:
                    if (component.row, next_comp.row) not in equal:
                        alternative.append((component.filename, component.row,
                                            alternative_comp.filename, alternative_comp.row))
    if equal:
        print("These rows have equal pns")
        print(equal)
    if similar:
        print("These rows have similar pns")
        print(similar)
    if alternative:
        print('These rows have similar alternative pn')
        print(alternative)
    print(warning)


def compare_capacitors(components: List[data_types.Component]):
    """
    compare capacitors by value and footprint
    :param components:
    :return:
    """
    similar_caps: List[Tuple[str, int, str, int]] = list()
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
            similar_caps.append((cap.filename, cap.row,
                                 cap_sorted[next_index].filename, cap_sorted[next_index].row))
            next_index += 1
    if similar_caps:
        print("Similar capacitor rows: ")
        print(similar_caps)


def compare_resistors(components: List[data_types.Component]):
    """
    compare capacitors by value and footprint
    :param components: list of components
    :return:
    """
    similar_resistors: List[Tuple[str, int, str, int]] = list()
    res_sorted: List[data_types.Component] = sorted([component for component in components
                                                     if component.component_type == data_types.ComponentType.RESISTOR
                                                     and component.details.value],
                                                    key=lambda x: x.details.value)
    if len(res_sorted) < 2:
        return
    for (index, res) in enumerate(res_sorted):
        next_index = index + 1
        while next_index < len(res_sorted) and res_sorted[next_index].details.value == res.details.value \
                and res_sorted[next_index].footprint == res.footprint:
            similar_resistors.append((res.filename, res.row,
                                      res_sorted[next_index].filename, res_sorted[next_index].row))
            next_index += 1
    if similar_resistors:
        print("Similar resistor rows: ")
        print(similar_resistors)


if __name__ == '__main__':
    if len(sys.argv) >= 2:
        if os.path.exists(sys.argv[1]):
            if os.path.isdir(sys.argv[1]):
                folder = sys.argv[1]
                components_list: List[data_types.Component] = list()
                for filename in os.listdir(folder):
                    if os.path.splitext(filename)[1].lower() == '.xlsx' \
                            and os.access(os.path.join(folder, filename), os.R_OK):
                        try:
                            components_list.extend(xlsx_parce.get_components_from_xlxs(os.path.join(folder, filename)))
                        except PermissionError:
                            pass
            else:
                components_list: List[data_types.Component] = xlsx_parce.get_components_from_xlxs(sys.argv[1])
            compare_pns(components_list)
            compare_capacitors(components_list)
            compare_resistors(components_list)
            print("Search complited")
        else:
            print("Incorrect filename")
    else:
        print("No file")
