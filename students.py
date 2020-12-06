from ProcessForm import deleteRows, newSheet, divideForm


def Process(wb, ws, teacher=False):
    deleteRows(ws, 2, 5)
    print("删除2-5列成功")
    newSheet(wb, teacher)
    divideForm(wb, teacher)

