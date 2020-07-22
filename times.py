'''
@Date: 2020-07-22 20:41:27
@Author: ywyz
@LastModifiedBy: ywyz
@Github: https://github.com/ywyz
@LastEditors: ywyz
@LastEditTime: 2020-07-22 20:55:41
'''


class GetTime():
    """获取时间类"""
    def __init__(self, current_time):
        self.year = current_time.year
        self.month = current_time.month
        self.day = current_time.day

    def GetYear(self):
        return self.year

    def GetMonth(self):
        return self.month

    def GetDay(self):
        return self.day
