"""
寻找表格最大行、列
"""


class ProcessingForm:
    def __init__(self, sheet):
        self.sheet = sheet

    def GetMaxROW(self):
        return self.sheet.max_row

    def GetMaxColumn(self):
        return self.sheet.max_column
