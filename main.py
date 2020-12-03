import openpyxl
from ReturnTime import GetToday, GetYesterday, GetName
from ProcessForm import CompareTime
import students

# fileName = input("将文件拖动到这儿：")
fileName = 'test.xlsx'
newWb = openpyxl.load_workbook(filename=fileName)
print("打开成功！")
ws1 = newWb["Sheet1"]

# 删除非今日时间行数
CompareTime(GetToday(), ws1)

# 判断今日是否统计教师
while True:
    choose = input("今日是否统计教师(y/n):")
    if choose == 'y' or choose == 'Y':

        break
    elif choose == 'n' or choose == 'N':
        students.Process(ws1)
        break

    else:
        print("输入有误，请重新输入。")
newName = GetName()

newWb.save(newName)
print("保存成功")





