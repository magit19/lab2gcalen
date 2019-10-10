import xlrd                                                   
                                                              
start = ['08:00', '09:50', '11:40', '14:00', '15:50', '17:40',]
end = ['09:35', '11:25', '13:15', '15:35', '17:25', '19:15',]  
                                                              
book = xlrd.open_workbook('imi2019.xls')                      
mag = book.sheet_by_index(9)                                  
                                                              
for i in range(3,39):                                         
    if mag.cell(i, 20).value != "":                            
        continue                                              
    para = mag.cell(i, 20).value                               
    l_pr = mag.cell(i, 21).value                               
    room = mag.cell(i, 22).value    
    day = 2 + (i-3)//6                          
    print(day, "сен", start[(i-3)%6], '-', end[(i-3)%6],    para, l_pr, room)  