'''
@Date: 2020-07-22 20:27:33
@Author: ywyz
@LastModifiedBy: ywyz
@Github: https://github.com/ywyz
@LastEditors: ywyz
@LastEditTime: 2020-07-22 22:02:53
'''

import datetime

import openpyxl

import CompareTime
import GetTime
import ProcessingForm

current_time = datetime.datetime.now()
# 获取当前时间
time = GetTime.GetTime(current_time)
time_strings = time.ReturnTimeStrings()
# 获取当前时间字符串
print("今日时间：", time_strings)
namestrings = time.TimeStrings()

names = input("输入文件名或拖动文件到此处：")
wb = openpyxl.load_workbook(names)  # 打开文件
wb_sheet = wb["Sheet1"]  # 打开工作簿1

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

# 删除教师表中幼儿数据列和幼儿表中教师数据列
processForm.DeleteRedundantData()
print("删除冗余数据成功")

# 查找各班没填人数和错误人数
processForm.FindStudentFilledError()

# 更改序号
processForm.ChangeNum()
processForm.DeleteNoneLines()

# 删除sheet1
names = namestrings + "新生防疫信息表.xlsx"
wb.save(names)
print("保存成功")
input("按任意键退出")
