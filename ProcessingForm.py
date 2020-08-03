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

    def CompareStudentNumber(self):
        """
        比较各班学生数量
        :return:
        """
        small_one = 0
        small_two = 0
        middle_one = 0
        middle_two = 0
        big_one = 0
        big_two = 0
        rows = 1
        for cell in self.sheet['H']:
            if cell.value == '大班':
                if self.sheet.cell(rows, 9).value == '一班':
                    big_one += 1
                    rows += 1
                    continue
                elif self.sheet.cell(rows, 9).value == '二班':
                    big_two += 1
                    rows += 1
                    continue
            elif cell.value == '中班':
                if self.sheet.cell(rows, 9).value == '一班':
                    middle_one += 1
                    rows += 1
                    continue
                elif self.sheet.cell(rows, 9).value == '二班':
                    middle_two += 1
                    rows += 1
                    continue
            elif cell.value == '小班':
                if self.sheet.cell(rows, 9).value == '一班':
                    small_one += 1
                    rows += 1
                    continue
                elif self.sheet.cell(rows, 9).value == '二班':
                    small_two += 1
                    rows += 1
                    continue

            else:
                rows += 1
                continue

        print("小一班：", small_one)
        print("小二班：", small_two)
        print("中一班：", middle_one)
        print("中二班：", middle_two)
        print("大一班：", big_one)
        print("大二班：", big_two)
