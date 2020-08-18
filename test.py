"""测试xlwinges模块

import xlwings as xw 

app = xw.App(visible=True,add_book=False)
wb = app.books.add()
wb.save(r'./test.xlsx')
wb.close()
app.quit()

"""
import re

str1="1+2=123456"
print(str1[0: str1.rfind("=")+1])
print(str1[str1.rfind('=')+1: ])
