'''
@Date: 2020-07-22 20:27:33
@Author: ywyz
@LastModifiedBy: ywyz
@Github: https://github.com/ywyz
@LastEditors: ywyz
@LastEditTime: 2020-07-22 21:45:09
'''

import openpyxl

names = input("输入文件名字：")
wb = openpyxl.load_workbook(names)
