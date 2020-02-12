from openpyxl import Workbook
from openpyxl import load_workbook
from typing import List
from string import ascii_uppercase

import data_types

headers = ['type', 'pn', 'manufacturer', 'pn alternative 1', 'pn alternative 2', 'designator', 'footprint',
           'dielectric', 'value', 'voltage', 'tolerance', 'description']
recommended = ['type', 'pn', 'designator', 'footprint', 'value']


def get_value(key: str, row: int, sheet, header_index):
    """

    :param key:
    :param sheet:
    :param header_index:
    :return:
    """
    if header_index[key] and sheet.cell(row=row, column=header_index[key]):
        return sheet.cell(row=row, column=header_index[key]).value
    return ""


def get_components_from_xlxs(filename):
    wb = load_workbook(filename=filename)
    sheet = wb.active
    header_index = dict.fromkeys(headers, None)
    for col in ascii_uppercase:
        col_header = sheet['%s1' % col].value
        if col_header and col_header.lower() in headers:
            header_index[col_header.lower()] = ascii_uppercase.index(col) + 1
    warning = 'The following columns are recommended but absent: ' + \
              ', '.join([header for header in recommended if not header_index[header]])
    print(warning)
    print(header_index)
    print(sheet.max_row)
    for row in range(1, sheet.max_row):
        component = data_types.SimpleComponent(component_type=get_value('type', row, sheet, header_index),
                                               footprint=get_value('footprint', row, sheet, header_index),
                                               pn=get_value('pn', row, sheet, header_index),
                                               manufacturer=get_value('manufacturer', row, sheet, header_index),
                                               designator=get_value('designator', row, sheet, header_index),
                                               description=get_value('description', row, sheet, header_index))
        print(component)




get_components_from_xlxs('BOM_StPlus_SOM_V1.5.xlsx')
