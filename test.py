# -*- coding: utf-8 -*-
import re
import xlrd
from xlwt import *
 
regex = re.compile("0?\d{11}")
   
w = Workbook()
ws = w.add_sheet(u'目标')
data = xlrd.open_workbook(u'C:\\Users\\Administrator\\Desktop\\XXXXX.xlsx')
table = data.sheets()[0]
ws.write(0, 0, table.row_values(0)[0])
ws.write(0, 1, table.row_values(0)[1])
ws.write(0, 2, table.row_values(0)[2])
ws.write(0, 3, table.row_values(0)[3])
ws.write(0, 4, table.row_values(0)[4])
ws.write(0, 5, table.row_values(0)[5])
ws.write(0, 6, u'目标号码')
ws.write(0, 7, u'属性')


j = 1
for i in range(1, table.nrows):
    index =  table.row_values(i)[0]
    usernum = table.row_values(i)[1]
    address = table.row_values(i)[2]
    hometel = table.row_values(i)[3]
    mobiletel = table.row_values(i)[4]
    aliastel = table.row_values(i)[5]

    for x in regex.findall(hometel):
        ws.write(j, 0, index)
        ws.write(j, 1, usernum)
        ws.write(j, 2, address)
        ws.write(j, 3, hometel)
        ws.write(j, 4, mobiletel)
        ws.write(j, 5, aliastel)
        ws.write(j, 6, x)
        ws.write(j, 7, u'户主电话')
        j += 1

        
    for x in regex.findall(mobiletel):
        ws.write(j, 0, index)
        ws.write(j, 1, usernum)
        ws.write(j, 2, address)
        ws.write(j, 3, hometel)
        ws.write(j, 4, mobiletel)
        ws.write(j, 5, aliastel)
        ws.write(j, 6, x)
        ws.write(j, 7, u'联系电话')
        j += 1
        
    for x in regex.findall(aliastel):
        ws.write(j, 0, index)
        ws.write(j, 1, usernum)
        ws.write(j, 2, address)
        ws.write(j, 3, hometel)
        ws.write(j, 4, mobiletel)
        ws.write(j, 5, aliastel)
        ws.write(j, 6, x)
        ws.write(j, 7, u'别名电话')
        j += 1
    #j+=1

w.save('num_formats.xls') 
