import openpyxl
from ProcessForm import deleteRows


def Process(ws):
    deleteRows(ws, 2, 5)
    print("删除2-5列成功")
    return

