def CompareTime(today, ws):
    """比较时间"""
    row = 1
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


def newSheet(wb1, teacher=False):
    """创建新表"""
    wb1.create_sheet("小一班")
    wb1.create_sheet("小二班")
    wb1.create_sheet("中一班")
    wb1.create_sheet("中二班")
    wb1.create_sheet("大一班")
    wb1.create_sheet("大二班")
    if teacher:
        wb1.create_sheet("教职工")
