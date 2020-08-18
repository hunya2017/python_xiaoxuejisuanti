"""
小学数学计算题生成excel
ver 1.01
作者：xy
"""

import fnmatch
import os
import random
import re
import xlwings as xw

""" 
打开excel并返回打开文件
"""
def open_excel(flile_path):
    app = xw.App(visible=True, add_book=False)
    try:
        wb = app.books.open(flile_path)  # 尝试打开文件，没有就新建
    except FileNotFoundError:
        wb = app.books.add()  # 新建
        wb.save(flile_path)  # 保存
    return wb


"""
设置表格格式
"""
def sz_excel(_wb_, _sht_, _range_=None):

    if _range_ != None:         # 设置单元格格式
        _range_.color = 255,200,255
        _range_.api.Font.Size = 16

    if _sht_ != None:
        _sht_.range('b1:f56').column_width = 14     # 设置单元格宽度
        _sht_.range('b1:f56').row_height = 25       # 设置单元格高度
        _sht_.range('b1:f56').api.Font.Size = 16    #设置单元格字体大小
        _sht_.range()



"""
添加sheet
"""
def add_sht(_wb_):      #wb为打开的表格
    sht_s = _wb_.sheets.count  # 获取sheet数量
    _wb_.sheets.add(after=_wb_.sheets[sht_s - 1])  # 在最后一个sheet后添加一个sheet
    return  _wb_.sheets.count - 1       #返回最后一个sheet的下标

    """"
    # print(sht_s)  # 打印sheet数量
    try:
        _wb_.sheets[sht_s].range('A1').value = 234  # 如果有sht_s,在最后一个sheet的a1写入数据
    except IndexError:
        # _wb_.sheets[0].active()
        # _wb_.sheets.add(after=_wb_.sheets.active)
        _wb_.sheets.add(after=_wb_.sheets[sht_s - 1])  # 在最后一个sheet后添加一个sheet
        _wb_.sheets[sht_s].range('A1').value = 234
    """


"""
获取当前最后一个sheet
"""
def sht_end(_wb_):
    sht_s = _wb_.sheets.count       # 获取sheet数量
    sht = _wb_.sheets[sht_s-1]
    return sht



"""
生成题库
"""
def tiku(nj=1):
    if nj==1:       # 生成一年级题库
        list_1 = []  # 一年级题库
        for i in range(1, 100):         # 生成一年级加法
            for j in range(1, 10):      # 10以内加数
                if i + j <= 100:
                    list_1.append('%d + %d = %d' % (i, j, i + j))
            for j in range(10, 110, 10):  # 整十加数
                if i + j <= 100:
                    list_1.append('%d + %d = %d' % (i, j, i + j))
        for i in range(1, 100, -1):     # 生成一年级减法
            for j in range(1, 10):      # 10以内被减数
                if i > j:
                    list_1.append('%d - %d = %d' % (i, j,i - j ))
            for j in range(10, 110, 10):  # 整十被减数
                if i > j:
                    list_1.append('%d - %d = %d' % (i, j, i - j))
        return list_1
    elif nj==2:     # 生成二年级题库
        list_2 = []  # 二年级题库
        for i in range(1, 100):     # 二年级加法
            for j in range(1, 100):
                if i + j <= 100:
                    list_2.append('%d + %d = %d' % (i, j, i + j))
        for i in range(100, 0, -1):     # 二年级减法
            for j in range(1, 100):
                if i > j:
                    list_2.append('%d - %d = %d' % (i, j, i - j))
        return list_2

"""
生成随机的题目
"""
def tm_suiji(nj=2, sm=280):      # 打印题目,nj年级默认2,sm数目默认280
    if sm < 5:      # 如果题目少于5 就等于5 方便输出
        sm = 5
    if nj == 1:     # 生成一年级随机题目
        list_tm = tiku(1)
        tm = random.sample(list_tm, sm)
        return tm
    else:       # 如果输入年级非1,就生成2年级随机题目
        list_tm = tiku(2)
        tm = random.sample(list_tm, sm)
        return tm


"""
写入表格
"""
def xr_excel(nj, sm=280, daan=False):     #nj=年级 sm=打印题的数目 daan=是否打印答案
    wb = open_excel('计算题.xlsx')     # 打开或生成表格
    sht = add_sht(wb)     #在最后一个sheet后添加一个新的sheet
    # 设置格式
    sz_excel(wb, wb.sheets[sht])

    x1 = 0     #切片坐标1
    x2 = 5     #切片坐标2
    xh = sm // 5
    tm = tm_suiji(nj,sm)    #生成题目
    for x in range(xh):
        suanshi = tm[x1:x2]  #选择5道题
        suan_n = []     #不带答案算式
        suan_d = []     #答案
        for suan in suanshi:
            suan_n.append(suan[: suan.rfind('=')+1])    #分解出算式
            suan_d.append(suan[suan.rfind('=')+1:])     #分解出答案

        wb.sheets[sht].range('B'+str(x+1)).value = suan_n   # 写入算式
        if daan==True:
            wb.sheets[sht].range('H'+str(x+1)).value = suan_d   # 写入答案
        x1 += 5
        x2 += 5
    wb.save()   #保存


def main():
    # wb = open_excel('测试.xlsx')
    # add_sht(wb)
    # sht = sht_end(wb)
    # sz_excel(wb, sht)


    nj = input('请输入年级:')
    while not re.findall('^[0-9]+$', nj):
        nj = input('只能输入数字,请重新输入:')
    xr_excel(nj)

if __name__ == '__main__':
    main()




# wb1.sheets[0].range('A1').value = 123

# wb1.sheets.add()

# wb.save()
