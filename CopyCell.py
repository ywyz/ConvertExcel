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
        rows = 1

        for cell in self.sheet['H']:
            if cell.value == "教职员工":
                for num in range(1, self.max_column):
                    self.teacher_sheet.cell(row=rows, column=num).value = list(self.sheet.rows)[rows][num].value
                print("复制成功", rows)
                rows += 1

            elif cell.value == "学生":
                for num in range(1, self.max_column):
                    self.student_sheet.cell(row=rows, column=num).value = list(self.sheet.rows)[rows][num].value
                print("复制成功", rows)
                rows += 1

            else:
                for num in range(1, self.max_column):
                    self.teacher_sheet.cell(row=rows, column=num).value = list(self.sheet.rows)[rows][num].value
                    self.student_sheet.cell(row=rows, column=num).value = list(self.sheet.rows)[rows][num].value
                print("复制成功", rows)
                rows += 1

            if rows == self.max_row:
                break
