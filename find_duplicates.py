import duplicates
import xlsx_parce
import data_types
import os
import sys
from typing import List

if __name__ == '__main__':
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
        duplicates.compare_pns(components_list)
        duplicates.compare_capacitors(components_list)
        duplicates.compare_resistors(components_list)
        print("Search in %s complited" % path)
    else:
        print("Incorrect filename")
