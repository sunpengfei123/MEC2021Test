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
def write_excel(ename, name, IP, Time, IP2, Time2, ipmin, ipmax):

    f = xlwt.Workbook()

    #创建sheet1
    sheet1 = f.add_sheet(u'sheet2',cell_overwrite_ok=True)
    row0 = [u'cases',u'PPOS_IP',u'PPOS_Time',u'MT_IP',u'MT_Time',u'MT_IP_min',u'MT_IP_max']


    for t in range(0,name.__len__()):

        column0 = name.__getitem__(t)
        column1 = IP.__getitem__(t)
        column2 = Time.__getitem__(t)
        column3 = IP2.__getitem__(t)
        column4 = Time2.__getitem__(t)
        column5 = ipmin.__getitem__(t)
        column6 = ipmax.__getitem__(t)
        print(column5.__len__())
        print(column3.__len__())

        lie = t*8

        acc = 0
        accmin = 0
        accmax = 0
        #生成第一行
        for i in range(0,len(row0)):
            sheet1.write(0,lie+i,row0[i],set_stlye('Times New Roman',220,True))

        #生成第一列

        for i in range(0,len(column0)):
            t = (column3[i] - column1[i])/column1[i]
            tmin = (column5[i] - column1[i])/column1[i]
            tmax = (column6[i] - column1[i])/column1[i]
            acc = (acc * i + t) / (i + 1)
            accmin = (accmin * i + tmin) / (i + 1)
            accmax = (accmax * i + tmax) / (i + 1)
            sheet1.write(i+1,lie+0,column0[i],set_stlye('Times New Roman',220,True))


        # 生成第一列
        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie+1, column1[i], set_stlye('Times New Roman', 220, True))


        average1 = 0
        # 生成第一列
        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie+2, column2[i], set_stlye('Times New Roman', 220, True))
            average1 = (average1 * i + column2[i]) / (i + 1)

        # 生成第一列
        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie + 3, column3[i], set_stlye('Times New Roman', 220, True))

        average2 = 0;
        # 生成第一列
        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie + 4, column4[i]/1000, set_stlye('Times New Roman', 220, True))
            average2 = (average2 * i + column4[i]/1000) / (i + 1)

        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie + 5, column5[i], set_stlye('Times New Roman', 220, True))

        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie + 6, column6[i], set_stlye('Times New Roman', 220, True))

        sheet1.write(len(column0) + 3, lie, 'average', set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0) + 3, lie + 2, average1, set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0) + 3, lie + 4, average2, set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0) + 4, lie, 'acc', set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0) + 4, lie + 1, acc, set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0) + 4, lie+2, 'acc_min', set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0) + 4, lie + 3, accmin, set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0) + 4, lie+4, 'acc_max', set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0) + 4, lie + 5, accmax, set_stlye('Times New Roman', 220, True))

    sheet1.write(37, 0, "此表格数据是PPOS算法与多线程加速的启发式算法时间以及准确率对比结果汇总，包括两个算法每个测试用例对应的Latency以及耗时，准确率比较取平均值。", set_stlye('Times New Roman', 220, True))
    sheet1.write(38, 0, "对每一个测试用例，启发式算法计算50次，取50次的Latency以及耗时平均值进行比较。",
                 set_stlye('Times New Roman', 220, True))

    f.save(ename+'.xls')


filepath = "C:\\Study\\TestCase\\2021_MEC\\result\\outPPOSsyn160-20210923.txt"

file = open(filepath)

s = file.readline()

name = []
ip = []
time = []
mt_ip = []
mt_ip_min = []
mt_ip_max = []
mt_time = []

last = ""

n = []
i = []
t = []
mt_i = []
mt_t = []
mt_i_min = []
mt_i_max = []

while s is not None and not s.__eq__(""):
    # print(s.split("."))
    if s.split(".").__len__()>1 and s.split(".")[1] == "xml\n":
        # print(s.split(".")[0])
        ff = open(os.path.join("C:\\Study\\TestCase\\2021_MEC\\result\\SDFG_Search_MT_CL_test\\synSC\\STEvsHCS\\"+s.split(".")[0].split("-")[0]+"\\"+s.split(".")[0]+".txt"))

        if last == "" or s.split(".")[0].split("-")[0] == last:
            n.append(s.split(".")[0])
            ss = file.readline()
            t.append((float)(ss.split(";")[2])/1000.0)
            i.append((int)(ss.split(";")[1].split(",")[1].split(")")[0]))
            mt_i.append((int)(ff.readline()))
            mt_t.append((float)(ff.readline()))
            mt_i_min.append((int)(ff.readline()))
            mt_i_max.append((int)(ff.readline()))
        else:
            name.append(n)
            ip.append(i)
            time.append(t)
            mt_ip.append(mt_i)
            mt_time.append(mt_t)
            mt_ip_min.append(mt_i_min)
            mt_ip_max.append(mt_i_max)
            n = []
            i = []
            t = []
            mt_i = []
            mt_t = []
            mt_i_min = []
            mt_i_max = []
            n.append(s.split(".")[0])
            ss = file.readline()
            t.append((float)(ss.split(";")[2])/1000.0)
            i.append((int)(ss.split(";")[1].split(",")[1].split(")")[0]))
            mt_i.append((int)(ff.readline()))
            mt_t.append((float)(ff.readline()))
            mt_i_min.append((int)(ff.readline()))
            mt_i_max.append((int)(ff.readline()))
        last = s.split(".")[0].split("-")[0]
    s = file.readline()

file.close()

name.append(n)
ip.append(i)
time.append(t)
mt_ip.append(mt_i)
mt_time.append(mt_t)
mt_ip_min.append(mt_i_min)
mt_ip_max.append(mt_i_max)

print(name)
print(ip)
print(time)

write_excel("PPOS_STE",name,ip,time,mt_ip,mt_time,mt_ip_min,mt_ip_max)