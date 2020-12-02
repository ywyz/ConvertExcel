import openpyxl
from ReturnTime import GetToday, GetYesterday, GetName
from ProcessForm import CompareTime

# fileName = input("将文件拖动到这儿：")
fileName = 'test.xlsx'
newWb = openpyxl.load_workbook(filename=fileName)
print("打开成功！")
ws1 = newWb["Sheet1"]

# 删除非今日时间行数
CompareTime(GetToday(), ws1)

newName = GetName()

newWb.save(newName)
print("保存成功")





