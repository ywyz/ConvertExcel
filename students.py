from ProcessForm import deleteRows, newSheet, divideForm, deleteNullLines
from openpyxl.utils import column_index_from_string


def Process(wb, ws, teacher=False):
    deleteRows(ws, column_index_from_string('B'), ord('F') - ord('B') + 1)
    print("删除2-5列成功")
    newSheet(wb, teacher)
    divideForm(wb, teacher)
    deleteNullLines(wb)
    for sheet in wb.get_sheet_names():
        if sheet == '教职工' or sheet == 'Sheet1':
            continue
        deleteRows(wb[sheet], column_index_from_string('J'), ord('M') - ord('J') + 1)
        deleteRows(wb[sheet], column_index_from_string('C'), ord('H') - ord('C') + 1)
