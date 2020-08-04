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
print("今日时间：", time_strings)

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
print("删除第二列到第五列成功")

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


# 删除空行
t_process.DeleteNoneLines()
print("删除教师表空行成功")
s_process.DeleteNoneLines()
print("删除学生表空行成功")

# 删除教师表中幼儿数据列和幼儿表中教师数据列
t_process.DeleteColumn(9, 1)
t_process.DeleteColumn(13, 20)
print("删除教师表第9列，13-33列成功")
s_process.DeleteColumn(4, 5)
s_process.DeleteColumn(5, 4)
print("删除教师表第4-8，10-13列成功")
t_process.DeleteRedundantData()
s_process.DeleteRedundantData()
print("删除冗余数据成功")
# 输出各班已填表人数
s_process.CompareStudentNumber()
wb.save("没填人数1.xlsx")
# 查找各班没填人数和错误人数
s_process.FindStudentFilledError()
wb.save("没填人数2.xlsx")
# 查找教职工没填人数
t_process.FindTeacherFilledError()
wb.save("没填人数final.xlsx")
small_one = wb.create_sheet("小一")
small_two = wb.create_sheet("小二")
middle_one = wb.create_sheet("中一")
middle_two = wb.create_sheet("中二")
big_one = wb.create_sheet("大一")
big_two = wb.create_sheet("大二")
s_process.CopyClass(small_one, small_two, middle_one, middle_two, big_one, big_two)
wb.save("复制分班.xlsx")
# 删除分班后空行
small_one = ProcessingForm.ProcessingForm(small_one)
small_two = ProcessingForm.ProcessingForm(small_two)
middle_one = ProcessingForm.ProcessingForm(middle_one)
middle_two = ProcessingForm.ProcessingForm(middle_two)
big_one = ProcessingForm.ProcessingForm(big_one)
big_two = ProcessingForm.ProcessingForm(big_two)
small_one.DeleteNoneLines()
small_two.DeleteNoneLines()
middle_one.DeleteNoneLines()
middle_two.DeleteNoneLines()
big_one.DeleteNoneLines()
big_two.DeleteNoneLines()
print("分班成功")
wb.save("已修改.xlsx")
print("保存成功")
