import duplicates
import compare_boms
import xlsx_parce
import data_types
import os
import sys
from typing import List


def find_similar():
    path = sys.argv[1] if len(sys.argv) >= 2 else '.'
    if os.path.exists(path):
        if os.path.isdir(path):
            components_list: List[data_types.Component] = list()
            for filename in os.listdir(path):
                if os.path.splitext(filename)[1].lower() == '.xlsx' \
                        and os.access(os.path.join(path, filename), os.R_OK):
                    try:
                        components_list.extend(xlsx_parce.get_components_from_xlxs(os.path.join(path, filename)))
                    except PermissionError:
                        pass
        else:
            components_list: List[data_types.Component] = xlsx_parce.get_components_from_xlxs(path)
        duplicates.compare_pns(components_list, root_len=8, precise=False)
        duplicates.compare_capacitors(components_list)
        duplicates.compare_resistors(components_list)
        print("Search in %s complited" % path)
    else:
        print("Incorrect filename")


def compare_boms_new_pns(detailed=False):
    """

    :return:
    """
    if len(sys.argv) != 3:
        print("Wrong params: two bom files to compare expected")
        return
    if not os.path.exists(sys.argv[1]) or not os.path.exists(sys.argv[2]) or os.path.isdir(sys.argv[1]) \
            or os.path.isdir(sys.argv[2]):
        print("Both files should be existing files (not folders)")
        return
    old = xlsx_parce.get_components_from_xlxs(sys.argv[1])
    new = xlsx_parce.get_components_from_xlxs(sys.argv[2])
    if detailed:
        compare_boms.detail_compare(old, new)
    else:
        compare_boms.find_new_pns(old, new)


if __name__ == '__main__':
    # compare_boms_new_pns(True)
    # find_similar()
    compare_boms_new_pns()

