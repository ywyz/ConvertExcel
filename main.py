'''
@Date: 2020-07-22 20:27:33
@Author: ywyz
@LastModifiedBy: ywyz
@Github: https://github.com/ywyz
@LastEditors: ywyz
@LastEditTime: 2020-07-22 22:02:53
'''

import openpyxl

names = input("输入文件名字：")
wb = openpyxl.load_workbook(names)  # 打开文件
wb_sheet = wb.get_sheet_by_name('sheet1')  #打开工作簿1
