class CopyCell:
    def __init__(self, sheet, teacher_sheet, student_sheet, max_row, max_column):
        self.sheet = sheet
        self.teacher_sheet = teacher_sheet
        self.student_sheet = student_sheet
        self.max_row = max_row
        self.max_column = max_column

    def CopyValue(self):
        """
        当表格值为教职员工，或学生时，遍历从A到最后列，将单元格值复制入新工作簿
        """
        for m in range(1, self.max_row + 1):  # 遍历行
            for n in range(1, self.max_column):
                if self.sheet.cell(row=m, column=8).value == '教职员工':
                    cell1 = self.sheet.cell(row=m, column=n).value
                    self.teacher_sheet.cell(row=m, column=n).value = cell1

                elif self.sheet.cell(row=m, column=8).value == '学生':
                    cell1 = self.sheet.cell(row=m, column=n).value
                    self.student_sheet.cell(row=m, column=n).value = cell1

                else:

                    cell1 = self.sheet.cell(row=m, column=n).value
                    self.teacher_sheet.cell(row=m, column=n).value = cell1
                    self.student_sheet.cell(row=m, column=n).value = cell1

        print("复制成功")
