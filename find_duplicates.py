# main script for comparing BOM's and finding duplicates
# run with parameters «--compare filename1 filename2» to compare two BOMs by pns
#                     «--compare filename1 filename2» to compare two BOMs by pns and quantity
#                     «--detailed filename1 filename2 to get detailed compare
#                     «--duplicate path» to find duplicates in path

import duplicates
import compare_boms
import xlsx_parce
import data_types
import os
import sys
from typing import List


def find_similar():
    """
    main function for finding similar positions
    :return:
    """
    path = sys.argv[2] if len(sys.argv) >= 3 else '.'
    if os.path.exists(path):
        if os.path.isdir(path):
            components_list: List[data_types.Component] = list()
            for filename in os.listdir(path):
                if os.path.splitext(filename)[1].lower() == '.xlsx' \
                        and os.access(os.path.join(path, filename), os.R_OK):
                    try:
                        data, _ = xlsx_parce.get_components_from_xlxs(os.path.join(path, filename))
                        components_list.extend(data)
                    except PermissionError:
                        pass
        else:
            components_list: List[data_types.Component] = xlsx_parce.get_components_from_xlxs(path)[0]
        equal, similar, alternative, error = duplicates.compare_pns(components_list, root_len=8, precise=False)
        if error:
            print(error)
        if equal:
            print("These rows have equal pns:\n")
            print([(pn1, row1, pn2, row2) for (pn1, row1, in1, pn2, row2, ind2) in equal])
        if similar:
            print("These rows have similar pns:\n")
            print([(pn1, row1, pn2, row2) for (pn1, row1, in1, pn2, row2, ind2) in similar])
        if alternative:
            print('These rows have similar alternative pn:\n')
            print([(pn1, row1, pn2, row2) for (pn1, row1, in1, pn2, row2, ind2) in alternative])
        similar_caps = duplicates.compare_capacitors(components_list)
        if similar_caps:
            print("Those capacitors are similar:")
            print(similar_caps)
        similar_resistors = duplicates.compare_resistors(components_list)
        if similar_resistors:
            print("Similar resistor rows: ")
            print(similar_resistors)

        print("Search in %s complited" % path)
    else:
        print("Incorrect filename")


def compare_boms_new_pns(quantity=False, detailed=False):
    """
    main function for bom comparing. Use detailed=True to get detailed comparing and False to compare PNs only
    :return:
    """
    if len(sys.argv) != 4:
        print("Wrong params: two bom files to compare expected")
        return
    if not os.path.exists(sys.argv[2]) or not os.path.exists(sys.argv[3]) or os.path.isdir(sys.argv[2]) \
            or os.path.isdir(sys.argv[3]):
        print("Both files should be existing files (not folders)")
        return
    old, warning = xlsx_parce.get_components_from_xlxs(sys.argv[2])
    print(warning)
    new, warning = xlsx_parce.get_components_from_xlxs(sys.argv[3])
    print(warning)
    compare_boms.find_new_pns(old, new, not quantity)
    if detailed:
        compare_boms.detail_compare(old, new, False)
    if quantity:
        compare_boms.detail_compare(old, new, True)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Parameters are missing, need type parameters and 1 or 2 filenames")
    if sys.argv[1].lower() == '--compare':
        compare_boms_new_pns()
    elif sys.argv[1].lower() == '--detailed':
        compare_boms_new_pns(detailed=True)
    elif sys.argv[1].lower() == '--quantity':
        compare_boms_new_pns(quantity=True)
    elif sys.argv[1].lower() == '--duplicates':
        find_similar()
    else:
        print("Wrong parameter")

