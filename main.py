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
current_time = datetime.datetime.now()
# 获取当前时间
time = GetTime.GetTime(current_time)
time_strings = time.ReturnTimeStrings()
# 获取当前时间字符串
print(time_strings)

names = input("输入文件名字：")
wb = openpyxl.load_workbook(names)  # 打开文件
wb_sheet = wb["Sheet1"] # 打开工作簿1

process = ProcessingForm.ProcessingForm(wb_sheet)
max_row = process.GetMaxROW()
print(max_row)
