'''
@Date: 2020-07-22 21:31:02
@Author: ywyz
@LastModifiedBy: ywyz
@Github: https://github.com/ywyz
@LastEditors: ywyz
@LastEditTime: 2020-07-23 11:07:00
获取时间字符串
'''
import GetTime
import datetime
current_time = datetime.datetime.now()

time = GetTime.GetTime(current_time)
time_strings = time.ReturnTimeStrings()


def time_string():
    return time_strings
