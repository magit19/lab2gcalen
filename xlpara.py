
import xlrd
book = xlrd.open_workbook('imi2019.xls')
sh = book.sheets()[9]
for r in range(3,39):
    time = sh.cell(r, 1).value
    para = sh.cell(r, 23).value
    room = sh.cell(r, 25).value
    if para != "" :
        para = para.replace("""
"""," ")
        print(time,para,room)

