import openpyxl
from pprint import pprint

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
    return wb.get_sheet_by_name(sheet)[cell].value


def build_out_result_dict(wb, map):

    # all versions are in 1 place: tab Intro & cell B1
    version = get_value(wb, 'Intro', 'B1')



# report file
report_wb = get_workbook('A_report.xlsx')


# map file
map_wb = get_workbook('DataMap.xlsx')

#print(map_wb.get_sheet_by_name('Map').columns)

map_cols = dict()

# cell coordinates per versions
for row in map_wb['Map'].iter_rows('A1:H1'):
    print(row[1].coordinate + 1)
    print(row[2])
    for cell in row:
        if cell.value:
            map_cols[str(cell.value)] = {'location': {
                                                cell.coordinate
                                                      }
                                         }

pprint(map_cols)

#build dict with version: [cell]

# grab results - manual

# results = dict()
#
# results['version'] = str(get_value(report_wb, 'Intro', 'B1'))
# results['name'] = '{} {}'.format(get_value(report_wb, 'Personal', 'B1'), get_value(report_wb, 'Personal', 'B2'))
# results['job'] = str(get_value(report_wb, 'Work', 'B2'))
#
# pprint(results)
