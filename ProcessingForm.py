"""
寻找表格最大行、列
"""


class ProcessingForm:
    def __init__(self, sheet, teacher, student):
        self.sheet = sheet
        self.teacher = teacher
        self.student = student

    def GetMaxROW(self):
        """获取今日表格最大行"""
        return self.sheet.max_row

    def GetMaxColumn(self):
        """获取今日表格最大列"""
        return self.sheet.max_column

