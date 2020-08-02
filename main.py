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
from DeleteNoneLines import DeleteNoneLines

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
processform = ProcessingForm.ProcessingForm(wb_sheet)
max_row = processform.GetMaxROW()
max_column = processform.GetMaxColumn()

# 比较时间，删除不在时间的单元格
time = CompareTime.CompareTime(wb_sheet, time_strings, max_row)
time.CompareStrings()
max_row = processform.GetMaxROW()

CopyCell = CopyCell.CopyCell(wb_sheet, teachers, students, max_row, max_column)
CopyCell.CopyValue()

# 获取学生老师最多行数
t_process = ProcessingForm.ProcessingForm(teachers)
s_process = ProcessingForm.ProcessingForm(students)
t_max_row = t_process.GetMaxROW()
s_max_row = s_process.GetMaxROW()
print(s_max_row)
print(t_max_row)

# DeleteNoneLines(teachers)
DeleteNoneLines(students)
DeleteNoneLines(teachers)

wb.save("已修改.xlsx")
