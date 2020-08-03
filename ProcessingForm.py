"""
寻找表格最大行、列
"""


class ProcessingForm:
    def __init__(self, sheet):
        self.sheet = sheet

    def GetMaxROW(self):
        """获取今日表格最大行"""
        return self.sheet.max_row

    def GetMaxColumn(self):
        """获取今日表格最大列"""
        return self.sheet.max_column

    def DeleteColumn(self, start, num):
        self.sheet.delete_cols(start, num)
        print("删除列成功")

    def DeleteNoneLines(self):
        """删除空行"""
        empty = []
        rows = 1
        print(self.sheet.title)
        for cell in self.sheet['B']:
            if cell.value is not None:
                rows += 1
                continue
            else:
                empty.append(rows)
                rows += 1
                continue

        for x in reversed(empty):
            self.sheet.delete_rows(x)

        print("删除空行成功")

    def DeleteRedundantData(self):
        """删除多次填写的数据"""

        names = []
        rows = 1
        delete_rows = []
        num = 0
        for cell in self.sheet['D']:
            if cell.value in names:
                delete_rows.append(rows)
                rows += 1

            else:
                names.append(cell.value)
                rows += 1

        for x in reversed(delete_rows):
            self.sheet.delete_rows(x, 1)
            num += 1

        print("删除", num, "行数据")
