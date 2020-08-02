def DeleteNoneLines(sheet):
    empty = []
    rows = 1
    print(sheet.title)
    for cell in sheet['B']:
        if cell.value is not None:
            rows += 1
            continue
        else:
            empty.append(rows)
            rows += 1
            continue

    for x in reversed(empty):
        sheet.delete_rows(x)

    print("删除空行成功")

