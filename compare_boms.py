import data_types
from typing import List, Tuple, Union, Optional
import duplicates

ParamData = Tuple[Union[float, str], str]


def get_comp_list_precise(components: List[data_types.Component]) -> Tuple[List[ParamData], List[ParamData],
                                                                           List[ParamData], List[ParamData]]:

    """
    gets lists of partnumbers, capacitors, inductors, resistors with their footprints
    :param components: list of components
    :return: list of pn+fp, cap+fp, res+fp, ind+fp
    """
    pns = [(component.pn, component.footprint) for component in components
           if component.component_type not in data_types.parametrized]
    capacitors = [(component.details.value, component.footprint) for component in components
                  if component.component_type == data_types.ComponentType.CAPACITOR]
    resistors = [(component.details.value, component.footprint) for component in components
                 if component.component_type == data_types.ComponentType.RESISTOR]
    inductors = [(component.details.value, component.footprint) for component in components
                 if component.component_type == data_types.ComponentType.INDUCTOR]
    return pns, capacitors, resistors, inductors


def get_diff(first: List[ParamData], second: List[ParamData]):
    """
    gets difference between two sets
    :param first: first set
    :param second: second set
    :return: two sets, one first without secons and then vice versa
    """
    f_set = set(first)
    s_set = set(second)
    return f_set.difference(s_set), s_set.difference(f_set)


def print_diff_data(first: List[ParamData], second: List[ParamData], type_str: str):
    """
    prints added and deleted elements of two lists
    :param first: list with first data
    :param second: list with second data
    :param type_str: description of data compared
    :return:
    """
    plus, minus = get_diff(second, first)
    print("Added %s:" % type_str)
    print(plus)
    print("Delete %s:" % type_str)
    print(minus)


def find_new_pns(old: List[data_types.Component], new: List[data_types.Component]):
    """
    compare two lists to find new positions
    :param old: list of old components
    :param new: list of new components
    :return:
    """
    old_pn, old_cap, old_res, old_ind = get_comp_list_precise(old)
    new_pn, new_cap, new_res, new_ind = get_comp_list_precise(new)
    print_diff_data(old_pn, new_pn, "partnumbers")
    print_diff_data(old_cap, new_cap, "capacitors")
    print_diff_data(old_res, new_res, "resistors")
    print_diff_data(old_ind, new_ind, "inductors")


def join_the_same(components: List[data_types.Component]):
    """
    joins the same components and changes list at the place
    :param components: 
    :return: 
    """
    similar = duplicates.compare_pns(components, precise=True)
    similar_sorted = sorted(similar, key=lambda x: x[3], reverse=True)
    for pair in similar_sorted:
        first_index = pair[2]
        second_index = pair[5]
        components[first_index].designator.extend(components[second_index].designator)
        components[first_index].description += components[second_index].description
    for pair in similar_sorted:
        components.pop(pair[5])


def find_components_in_list(key_component: data_types.Component, components: List[data_types.Component]) \
        -> Optional[data_types.Component]:
    """
    finds component in components and returns its index
    :param key_component: component to find
    :param components: list of components
    :return: component or None
    """
    if key_component.component_type not in data_types.parametrized:
        for component in components:
            if component.pn.lower() == key_component.pn.lower():
                if component.footprint.lower() == key_component.footprint.lower():
                    return component
        return None
    if key_component.component_type == data_types.ComponentType.CAPACITOR:
        for component in components:
            if component.details.absolute_pf_value == key_component.details.absolute_pf_value:
                if component.footprint == key_component.footprint:
                    return component
        return None
    if key_component.component_type in [data_types.ComponentType.RESISTOR, data_types.ComponentType.INDUCTOR]:
        for component in components:
            if component.details.value == key_component.details.value:
                if component.footprint == key_component.footprint:
                    if component.component_type == key_component.component_type:
                        return component
        return None


def detail_compare(old: List[data_types.Component], new: List[data_types.Component]):
    """
    compare two list
    :param old:
    :param new:
    :return:
    """
    join_the_same(old)
    join_the_same(new)
    old_sorted = sorted(old, key=lambda x: x.row)
    for component in old_sorted:
        paired = find_components_in_list(component, new)
        if paired:
            warning = ""
            if component.component_type != paired.component_type:
                warning += "Type: was %s and now %s\n" % \
                           (data_types.types[component.component_type], data_types.types[paired.component_type])
            if component.manufacturer != paired.manufacturer:
                warning += "Manufacturer: was %s and now %s\n" % (component.manufacturer, paired.manufacturer)

            if component.description.lower() != paired.description.lower():
                warning += "Description: was %s and now %s\n" % (component.description, paired.description)
            if warning:
                print("For component with pn %s, old row %i, new row %i changes are the following: \n" %
                      (component.pn, component.row, paired.row) + warning)