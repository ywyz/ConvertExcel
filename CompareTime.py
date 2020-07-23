"""
比较时间并删除时间早的一行
"""


class CompareTime:
    def __init__(self, sheet, strings, max_row):
        self.sheet = sheet  # 工作簿对象
        self.strings = strings  # 时间字符串
        self.max_row = max_row  # 最大行数

    def CompareStrings(self):
        """比较填写时间中是否有今天，没有今天就删除该行"""
        row = 1
        for cell in self.sheet['B']:
            if self.strings == "提交答卷时间":
                continue
            elif self.strings in cell.value:
                break

            else:
                row += 1
                continue

        self.sheet.delete_rows(2, row -2)
