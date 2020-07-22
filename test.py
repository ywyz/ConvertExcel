'''
@Date: 2020-07-22 20:55:55
@Author: ywyz
@LastModifiedBy: ywyz
@Github: https://github.com/ywyz
@LastEditors: ywyz
@LastEditTime: 2020-07-22 21:42:50
'''
import GetTime
import datetime
current_time = datetime.datetime.now()

time = GetTime.GetTime(current_time)
time_strings = time.ReturnTimeStrings()
