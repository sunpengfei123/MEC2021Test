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
def write_excel(ename, name, IP, Time,average):

    f = xlwt.Workbook()

    #创建sheet1
    sheet1 = f.add_sheet(u'sheet2',cell_overwrite_ok=True)
    row0 = [u'cases',u'IP',u'Time']


    for t in range(0,len(name)):

        column0 = name.__getitem__(t)
        column1 = IP.__getitem__(t)
        column2 = Time.__getitem__(t)

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

        sheet1.write(len(column0) + 1, lie , 'average', set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0)+1, lie + 2, average[t], set_stlye('Times New Roman', 220, True))

    sheet1.write(34, 0, "此表格数据是启发式算法测试结果汇总表，对每一个测试用例，启发式算法计算50次，取50次的Latency以及耗时平均值。",
                 set_stlye('Times New Roman', 220, True))

    f.save(ename+'.xls')



path = "C:\\Study\\TestCase\\2021_MEC\\result"

# xmlSSet = {"Test1_Muti_Pro","Test2_Muti_Pro","acyclic_degree4_Muti_Pro","acyclic_degree6_Muti_Pro","cyclic_degree2_Muti_Pro","cyclic_degree4_Muti_Pro","cyclic_degree6_Muti_Pro","sc_degree2_Muti_Pro","sc_degree4_Muti_Pro","sc_degree6_Muti_Pro"}

xmlSSet = {"synSC_time"}

# algotithms = ["MEC_IP_un_Test","MEC_IP_GA_Test",]

algotithms = ["SDFG_Search_MT_CL_test"]

Test_SSet = ["a5q20","a5q50","a10q50","a10q100","a20q100","a10q1000"]#

lie = 0

for algotithm in algotithms:

    for xmlSet in xmlSSet:
        average = []
        print(xmlSet)
        Test_SSet = ["a5q20","a5q50","a10q50","a10q100","a20q100","a10q1000"]#

        name = []
        GA_IP = []
        GA_IP_T = []

        for i in range(Test_SSet.__len__()):



            a = 0
            Test_Set = Test_SSet.pop()

            FA_path = path+"\\"+algotithm+"\\"+xmlSet+"\\MT"

            nname = []
            nGA_IP = []
            nGA_IP_T = []

            j = 1
            for filename in os.listdir(FA_path+"\\"+Test_Set):
                if filename.split(".")[1] == "png":
                    break
                # print(os.path.join(FA_path+"\\"+Test_Set, filename))
                file = open(os.path.join(FA_path+"\\"+Test_Set, filename))
                nname.append(filename)
                ip = (int)(file.readline())
                t = (float)(file.readline())
                # if t < 500:
                nGA_IP.append(ip)
                nGA_IP_T.append(t)
                a = ((a*j+t)/(j+1))
                j+=1

                file.close()
            average.append(a)

            name.append(nname)
            GA_IP_T.append(nGA_IP_T)
            GA_IP.append(nGA_IP)

        print(name)

        write_excel(algotithm+'_'+xmlSet+"_MT", name,GA_IP,GA_IP_T,average)

        lie = lie + 3

