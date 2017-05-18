import openpyxl
from pprint import pprint
import string


# given number, calculate column in excel
def colnum_string(n):
    div=n
    str=""
    while div>0:
        module=(div-1)%26
        str=chr(65+module)+str
        div=int((div-module)/26)
    return str


# converts from col to num
def col2num(col):
    num = 0
    for c in col:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


def get_workbook(xlsx):
    """
    Grabs excel workbook object
    :param xlsx:
    :return: workbook object
    """

    wb = openpyxl.load_workbook(xlsx)

    return wb


def get_value(wb, sheet, cell):
    """
    Return value of cell based on workbook, sheet and cell
    :param wb:
    :param sheet:
    :param cell:
    :return: value
    """
    return str(wb.get_sheet_by_name(sheet)[cell].value)


def build_out_map_dict(map_wb, field_col, field_rows):
    """
    From map workbook and row of fields return the map dict to be used
    :param map_wb:
    :param field_col: column all data fields are in
    :param field_rows: list of row nums or range
    :return: map dict
    """

    # data fields
    data_fields = dict()

    for num in field_rows:
        data_fields[num] = get_value(map_wb, 'Map', '{}{}'.format(field_col, num))      # key (num) = value (data field)

    # map
    map_dict = dict()

    # get cell coordinates per version - key to dict
    for row in map_wb['Map'].iter_rows(max_row=1, min_col=2):  # get all row 1, skip column 1

        for cell in row:
            if cell.value:
                map_dict[str(cell.value)] = {
                    'Location': cell.coordinate[0][0]  # get letter
                }
        # ends with -- version : { 'Location': 'A' }

    # loop through and add data field info - version : { 'Location' : A, 'Data Field' : [Tab, Cell] }
    for num in field_rows:
        for version in map_dict.keys():
            current_col = col2num(map_dict[version]['Location'])        # current column holds the tabs
            next_col = colnum_string(current_col+1)         # next column holds the cell coordinate info

            # if multiple cells...
            if ',' in get_value(map_wb, 'Map', '{}{}'.format(next_col, num)):
                cell = get_value(map_wb, 'Map', '{}{}'.format(next_col, num)).split(',')

            else:
                cell = [get_value(map_wb, 'Map', '{}{}'.format(next_col, num))]

            map_dict[version][data_fields[num]] = [get_value(map_wb, 'Map', '{}{}'.format(map_dict[version]['Location'], num)), cell]

    return map_dict


# results
def build_out_result_dict(report_wb, map_dict):
    """
    Based on wb and map, build dictionary of results to be pushed to db
    :param report_wb:
    :param map_dict:
    :return:
    """

    result = dict()

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

# report file
report_wb = get_workbook('A_report.xlsx')


report_wb2 = get_workbook('B_report.xlsx')

# map file
map_wb = get_workbook('DataMap.xlsx')

# map dict
map = build_out_map_dict(map_wb, 'A', range(3, 5))

pprint(build_out_result_dict(report_wb, map))

pprint(build_out_result_dict(report_wb2, map))