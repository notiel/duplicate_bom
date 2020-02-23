import data_types
from typing import List, Tuple, Union

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
