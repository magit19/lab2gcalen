# coding: utf-8
import xlrd
book = xlrd.open_workbook('imi2019.xls')
sh = book.sheets()[9]
n = 0
for r in range(3, 39):
    para = sh.cell(r, 23).value
    room = sh.cell(r, 25).value
    if para != "":
        para = para.replace("\n", " ")
        print(para, room)
        n = n + 1
print(n) 

     
        
        
    
    
