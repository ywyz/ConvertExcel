"""
获取时间并删除旧的时间
"""


def GetTime(current_time):
    year = current_time.year
    month = current_time.month
    day = current_time.day
    strings = (str(year) + "/" + str(month) + "/" + str(day))
    return strings


def GetNameTime(current_time):
    year = current_time.year
    month = current_time.month
    day = current_time.day
    strings = (str(year) + "." + str(month) + "." + str(day))
    return strings
