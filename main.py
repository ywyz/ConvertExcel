'''
@Date: 2020-07-22 20:27:33
@Author: ywyz
@LastModifiedBy: ywyz
@Github: https://github.com/ywyz
@LastEditors: ywyz
@LastEditTime: 2020-07-22 22:02:53
'''

import openpyxl
import GetTime
import datetime
import ProcessingForm
import CompareTime
import CopyCell

current_time = datetime.datetime.now()
# 获取当前时间
time = GetTime.GetTime(current_time)
time_strings = time.ReturnTimeStrings()
# 获取当前时间字符串
print(time_strings)

# names = input("输入文件名字：")
wb = openpyxl.load_workbook('example.xlsx')  # 打开文件
wb_sheet = wb["Sheet1"]  # 打开工作簿1
# 建立教职工和学生工作簿
teachers = wb.create_sheet("教职工")
students = wb.create_sheet("学生")

# 获取表格最多行数
processForm = ProcessingForm.ProcessingForm(wb_sheet)
max_row = processForm.GetMaxROW()


# 比较时间，删除过时的行
time = CompareTime.CompareTime(wb_sheet, time_strings, max_row)
time.CompareStrings()

# 删除通用列
processForm.DeleteColumn(2, 5)

# 获取删除后的行数和列数
max_row = processForm.GetMaxROW()
max_column = processForm.GetMaxColumn()

CopyCell = CopyCell.CopyCell(wb_sheet, teachers, students, max_row, max_column)
CopyCell.CopyValue()

# 获取学生老师最多行数
t_process = ProcessingForm.ProcessingForm(teachers)
s_process = ProcessingForm.ProcessingForm(students)
t_max_row = t_process.GetMaxROW()
s_max_row = s_process.GetMaxROW()
print(s_max_row)
print(t_max_row)

# 删除空行
t_process.DeleteNoneLines()
s_process.DeleteNoneLines()

# 删除教师表中幼儿数据列和幼儿表中教师数据列
t_process.DeleteColumn(9, 1)
t_process.DeleteColumn(13, 20)
s_process.DeleteColumn(4, 5)
s_process.DeleteColumn(5, 4)
wb.save("修改1.xlsx")
t_process.DeleteRedundantData()
s_process.DeleteRedundantData()

wb.save("已修改.xlsx")
