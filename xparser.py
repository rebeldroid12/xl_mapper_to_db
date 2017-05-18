import openpyxl
from pprint import pprint
import string


def column_num_to_str(num):
    """
    Converts number (1) to the column string ('A') as it would appear in an excel doc
    :param num:
    :return:
    """
    div = num
    result = ""
    while div > 0:
        module = (div - 1) % 26
        result = chr(65 + module) + result
        div = int((div - module)/26)
    return result


def column_str_to_num(col):
    """
    Converts column string ('A') to the column number (1)
    :param col: column
    :return: column number
    """
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def get_workbook(xlsx):
    """
    Grabs excel workbook object
    :param xlsx: excel file
    :return: workbook object
    """

    wb = openpyxl.load_workbook(xlsx)

    return wb


def get_value(workbook, sheet, cell):
    """
    Return value of cell based on workbook, sheet and cell
    :param workbook: workbook
    :param sheet: workbook sheet
    :param cell: cell
    :return: value
    """
    return str(workbook.get_sheet_by_name(sheet)[cell].value)


def build_out_map_dict(map_workbook, field_col, field_rows):
    """
    From map workbook, column of data fields and rows data fields are in - return the map dict to be used
    :param map_workbook: map workbook
    :param field_col: column all data fields are in
    :param field_rows: list of row nums or range data fields are found at
    :return: map dict
    """

    # data fields
    data_fields = {}

    for num in field_rows:
        data_fields[num] = get_value(map_workbook, 'Map', '{}{}'.format(field_col, num))      # key (num) = value (data field)

    # map
    map_dict = {}

    # get cell coordinates per version - key to dict
    for row in map_workbook['Map'].iter_rows(max_row=1, min_col=2):  # get all row 1, skip column 1

        for cell in row:
            if cell.value:
                map_dict[str(cell.value)] = {
                    'Location': cell.coordinate[0][0]  # get letter
                }
        # ends with -- version : { 'Location': 'A' }

    # loop through and add data field info - version : { 'Location' : A, 'Data Field' : [Tab, Cell] }
    for num in field_rows:
        for version in map_dict.keys():
            current_col = column_str_to_num(map_dict[version]['Location'])        # current column holds the tabs
            next_col = column_num_to_str(current_col + 1)         # next column holds the cell coordinate info

            # if multiple cells...
            if ',' in get_value(map_workbook, 'Map', '{}{}'.format(next_col, num)):
                cell = get_value(map_workbook, 'Map', '{}{}'.format(next_col, num)).split(',')

            else:
                cell = [get_value(map_workbook, 'Map', '{}{}'.format(next_col, num))]

            map_dict[version][data_fields[num]] = [get_value(map_workbook, 'Map', '{}{}'.format(map_dict[version]['Location'], num)), cell]

    return map_dict


def build_out_result_dict(report_wb, map_dict):
    """
    Based on workbook and map dictionary, build dictionary of results (to be pushed to db later)
    :param report_wb: report workbook
    :param map_dict: mapping dictionary
    :return: result dictionary
    """

    result = {}

    # all versions are in 1 place: tab Intro & cell B1
    version = get_value(report_wb, 'Intro', 'B1')

    # grab all the elements (data fields) based on version
    for field in map_dict[version].keys():
        if 'Location' not in field:

            # if value in multiple cells, and must be concat'd
            if len(map_dict[version][field][1]) > 1:

                concat_str = ''

                for cell in map_dict[version][field][1]:
                    concat_str += ' {}'.format(get_value(report_wb, map_dict[version][field][0], cell))

                result[field] = concat_str

            else:
                result[field] = get_value(report_wb, map_dict[version][field][0], map_dict[version][field][1][0])

    return result


###################################

if __name__ == '__main__':
    # report file
    report_wb = get_workbook('A_report.xlsx')

    report_wb2 = get_workbook('B_report.xlsx')

    # map file
    map_wb = get_workbook('DataMap.xlsx')

    # map dict
    mapper = build_out_map_dict(map_wb, 'A', range(3, 5))

    pprint(build_out_result_dict(report_wb, mapper))

    pprint(build_out_result_dict(report_wb2, mapper))
