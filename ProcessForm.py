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
    wb1.create_sheet("错误")
    if teacher:
        wb1.create_sheet("教职工")


def copyCell(ws1, ws2, rows):
    for n in range(1, ws1.max_column + 1):
        cell1 = ws1.cell(row=rows, column=n).value
        ws2.cell(row=rows, column=n).value = cell1


def divideForm(wb, teacher=False):
    """分到相应单元格"""
    ws1 = wb["Sheet1"]
    rows = 1
    for cell in ws1['C']:
        if cell.value == "学生":
            if ws1.cell(row=rows, column=17).value == '小班' and ws1.cell(row=rows, column=18).value == '一班':
                copyCell(ws1, wb["小一班"], rows)
            elif ws1.cell(row=rows, column=17).value == '小班' and ws1.cell(row=rows, column=18).value == '二班':
                copyCell(ws1, wb["小二班"], rows)
            elif ws1.cell(row=rows, column=17).value == '中班' and ws1.cell(row=rows, column=18).value == '一班':
                copyCell(ws1, wb["中一班"], rows)
            elif ws1.cell(row=rows, column=17).value == '中班' and ws1.cell(row=rows, column=18).value == '二班':
                copyCell(ws1, wb["中二班"], rows)
            elif ws1.cell(row=rows, column=17).value == '大班' and ws1.cell(row=rows, column=18).value == '一班':
                copyCell(ws1, wb["大一班"], rows)
            elif ws1.cell(row=rows, column=17).value == '大班' and ws1.cell(row=rows, column=18).value == '二班':
                copyCell(ws1, wb["大二班"], rows)
            rows += 1
            continue
        elif cell.value == "教职员工":
            if teacher:
                copyCell(ws1, wb["教职工"], rows)
                rows += 1
                continue
            else:
                copyCell(ws1, wb["错误"], rows)
                rows += 1
                continue
        else:
            copyCell(ws1, wb["小一班"], rows)
            copyCell(ws1, wb["小二班"], rows)
            copyCell(ws1, wb["错误"], rows)
            copyCell(ws1, wb["中一班"], rows)
            copyCell(ws1, wb["中二班"], rows)
            copyCell(ws1, wb["大一班"], rows)
            copyCell(ws1, wb["大二班"], rows)
            if teacher:
                copyCell(ws1, wb["教职工"], rows)
            rows += 1
            continue

    print("划分成功")
