import os
import os.path
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.pyplot import MultipleLocator
import math

import xlwt


def set_stlye(name,height,bold=False):
    #初始化样式
    style = xlwt.XFStyle()
    #创建字体
    font = xlwt.Font()
    font.bold = bold
    font.colour_index = 0
    font.height = height
    font.name = name
    style.font = font
    return style


#写入数据
def write_excel(ename, name, IP, Time):

    f = xlwt.Workbook()

    #创建sheet1
    sheet1 = f.add_sheet(u'sheet2',cell_overwrite_ok=True)
    row0 = [u'cases',u'IP',u'Time']


    for t in range(0,1):

        column0 = name
        column1 = IP
        column2 = Time

        lie = t*4


        #生成第一行
        for i in range(0,len(row0)):
            sheet1.write(0,lie+i,row0[i],set_stlye('Times New Roman',220,True))

        #生成第一列
        for i in range(0,len(column0)):
            sheet1.write(i+1,lie+0,column0[i],set_stlye('Times New Roman',220,True))

        # 生成第一列
        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie+1, column1[i], set_stlye('Times New Roman', 220, True))

        # 生成第一列
        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie+2, column2[i], set_stlye('Times New Roman', 220, True))

    f.save(ename+'.xls')


filepath = "C:\\Study\\TestCase\\2021_MEC\\result\\outPPOSsyn160-20210923.txt"

file = open(filepath)

s = file.readline()

name = []
ip = []
time = []

while s is not None and not s.__eq__(""):
    # print(s.split("."))
    if s.split(".").__len__()>1 and s.split(".")[1] == "xml\n":
        # print(s.split(".")[0])
        name.append(s.split(".")[0])
        # print(name)
        ss = file.readline()
        time.append(ss.split(";")[2])
        ip.append(ss.split(";")[1].split(",")[1].split(")")[0])
    s = file.readline()

file.close()

print(name)
print(ip)
print(time)

write_excel("PPOS",name,ip,time)