'''
@Date: 2020-07-22 20:41:27
@Author: ywyz
@LastModifiedBy: ywyz
@Github: https://github.com/ywyz
@LastEditors: ywyz
@LastEditTime: 2020-07-22 21:26:57
'''


class GetTime():
    """时间类"""
    def __init__(self, current_time):
        """"获取当天时间"""
        self.year = current_time.year
        self.month = current_time.month
        self.day = current_time.day

    def ReturnTimeStrings(self):
        """返回时间字符串"""
        strings = (str(self.year) + "/" + str(self.month) + "/" +
                   str(self.day))
        return strings
