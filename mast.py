"""

小学一二年级加减法基础题
ver 1.0 
作者：xy

"""
import random
import re
import xlwings as xw  # 导入excel模块

app = xw.App(visible=True, add_book=False)
#app.display_alerts = False
#app.screen_updating = False
filepath = r'./小学数学计算题.xlsx'
wb = app.books.open(filepath)



list_ynj = []  # 一年级加减法
list_enj = []  # 二年级加减法

"""
# 一年级加法
"""
for i in range(1, 100):
    for j in range(1, 10):           # 10以内加数
        if i + j <= 100:
            list_ynj.append('%d + %d =' % (i, j))
    for j in range(10, 110, 10):      # 整十加数
        if i + j <= 100:
            list_ynj.append('%d + %d =' % (i, j))

"""
#一年级减法
"""
for i in range(100, 0, -1):
    for j in range(1, 10):  # 10以内被减数
        if i > j:
            list_ynj.append('%d - %d =' % (i, j))
    for j in range(10, 110, 10):  # 整十被减数
        if i > j:
            list_ynj.append('%d - %d =' % (i, j))

"""
#二年级加法
"""
for i in range(1, 100):
    for j in range(1, 100):
        if i + j <= 100:
            list_enj.append('%d + %d =' % (i, j))

"""
#二年级减法
"""
for i in range(100, 0, -1):
    for j in range(1, 100):
        if i > j:
            list_enj.append('%d - %d =' % (i, j))

"""
#打印列表
# print(list_ynj)
for x in list_ynj:
    print(x,end="\n")
"""

# print(random.sample(list_ynj,200))
"""
输出几年级加减题
"""


def dytimu(nj=2, sm=200):  # 打印题目，nj年级默认2，sm数目默认200
    #n = 0
    if sm<5:        #如果题目小于5就等于5
        sm=5
    if nj == 1:
        tm = random.sample(list_ynj, sm)
    else:
        tm = random.sample(list_enj, sm)
    
    """写入单元格"""
    #写入列
    #wb.sheets['sheet1'].range('A1').options(transpose=True).value=tm
    
    #写入行
    x1=0        #切片坐标1
    x2=5        #切片坐标2
    xh=sm//5    #循环次数
    for x in range(xh):
        wb.sheets['sheet1'].range('B'+str(x+1)).value=tm[x1:x2]
        x1+=5
        x2+=5


    """
    for x in tm:
        print(x)
        n += 1
        if n == 20:
            print("\n")
            n = 0
    """


nj1 = input('请输入要打印几年级的题目：')
while not re.findall('^[0-9]+$', nj1):
    nj1 = input('年级只能输入数字，请重新输入：')

"""
sm1 = input('请输入要打印多少道题: ')
while not re.findall('^[0-9]+$', sm1):
    sm1 = input('题目数量只能输入数字，请重新输入：')
"""

dytimu(int(nj1), 280)  #打印题目

wb.save()
# wb.close()
# app.quit()

