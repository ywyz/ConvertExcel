from ProcessForm import deleteRows, newSheet


def Process(wb, ws, teacher=False):
    deleteRows(ws, 2, 5)
    newSheet(wb, teacher)
    print("删除2-5列成功")
