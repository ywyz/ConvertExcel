import datetime

today = datetime.date.today()
yesterday = datetime.date.today() + datetime.timedelta(days=-1)


def GetToday():
    TodayStrings = (str(today.year) + "/" + str(today.month) + "/" + str(today.day))
    return TodayStrings


def GetName():
    nameStrings = (str(today.year) + "." + str(today.month) + "." + str(today.day) + "防疫信息表.xlsx")
    return nameStrings


def GetYesterday():
    return yesterday
