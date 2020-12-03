import openpyxl


def CompareTime(today,ws):
    """比较时间"""
    row = 1;
    for cell in ws['B']:
        if cell.value == "提交答卷时间":
            row += 1
            continue
        elif today in cell.value:
            break
        else:
            row += 1

    print("删除", row - 1, "行")
    ws.delete_rows(2, row - 2)


def deleteRows(ws, column_a, column_b):
    """删除指定列"""
    ws.delete_cols(column_a, column_b)