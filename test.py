'''
@Date: 2020-07-22 20:55:55
@Author: ywyz
@LastModifiedBy: ywyz
@Github: https://github.com/ywyz
@LastEditors: ywyz
@LastEditTime: 2020-07-22 21:25:50
'''
import times
import datetime

current_time = datetime.datetime.now()
time = times.GetTime(current_time)

day = time.day
month = time.month
year = time.year
strings = time.ReturnTimeStrings()
print(strings)