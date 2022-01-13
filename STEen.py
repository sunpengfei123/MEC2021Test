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
def write_excel(ename, name, IP, EN, Time,TT):

    f = xlwt.Workbook()

    #创建sheet1
    sheet1 = f.add_sheet(u'sheet2',cell_overwrite_ok=True)
    row0 = [u'cases']
    row0.append(u'Latency')
    row0.append(u'Energy')
    # row0.append(u'Time')
    for i in range(IP.__getitem__(0)[0].__len__()-1):
        row0.append(u'Latency-rate')
        row0.append(u'Energy-rate')



    for t in range(0,len(name)):

        column0 = name.__getitem__(t)
        column1 = IP.__getitem__(t)
        column2 = EN.__getitem__(t)
        column3 = Time.__getitem__(t)
        column4 = TT.__getitem__(t)

        lie = t*25

        default = set_stlye('Times New Roman',220,True)

        #生成第一行
        for i in range(0,len(row0)):
            sheet1.write(0,lie+i,row0[i],default)

        #生成第一列
        for i in range(0,len(column0)):
            sheet1.write(i+1,lie+0,column0[i],default)

        # 生成第一列
        for i in range(0, len(column0)):
            # sheet1.write(i + 1, lie + 1, column1[i][0], default)
            for j in range(column1[i].__len__()):
                sheet1.write(i + 1, lie+1+j*5, column1[i][j], default)

        for i in range(0, len(column0)):
            for j in range(column1[i].__len__()):
                sheet1.write(i + 1, lie+3+j*5, (column1[i][j] - column1[i][0]+0.0)/column1[i][0], default)

                # 生成第一列
        for i in range(0, len(column0)):
            # sheet1.write(i + 1, lie + 2, column2[i][0], default)
            for j in range(column2[i].__len__()):
                sheet1.write(i + 1, lie + 2 + j * 5, column2[i][j], default)

        for i in range(0, len(column0)):
            for j in range(column2[i].__len__()):
                sheet1.write(i + 1, lie + 4 + j * 5, (column2[i][0] - column2[i][j]+0.0)/column2[i][0], default)

        # for i in range(0, len(column0)):
        #     for j in range(column3[i].__len__()):
        #         sheet1.write(i + 1, lie + 5 + j * 5, column2[i][j], default)

        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie + 5 + 26, column4[i], default)


    sheet1.write(34, 0, "此表格数据是启发式算法测试结果汇总表，对每一个测试用例，启发式算法计算50次，取50次的Latency以及耗时平均值。",
                 default)

    f.save(ename+'.xls')



path = "C:\\Study\\TestCase\\2021_MEC\\result"

# xmlSSet = {"Test1_Muti_Pro","Test2_Muti_Pro","acyclic_degree4_Muti_Pro","acyclic_degree6_Muti_Pro","cyclic_degree2_Muti_Pro","cyclic_degree4_Muti_Pro","cyclic_degree6_Muti_Pro","sc_degree2_Muti_Pro","sc_degree4_Muti_Pro","sc_degree6_Muti_Pro"}

xmlSSet = {"proc2","proc4","proc8","proc16"}

# algotithms = ["MEC_IP_un_Test","MEC_IP_GA_Test",]

algotithms = ["benchmarks"]

Test_SSet = ["HPSSen"]#

lie = 0

for algotithm in algotithms:

    for xmlSet in xmlSSet:
        average = []
        print(xmlSet)
        Test_SSet = ["HPSStran"]##"a5q50","a10q50","a5q20","a10q100","a20q100","a10q1000"

        name = []
        GA_IP = []
        GA_en = []
        GA_IP_T = []
        tT=[]

        for i in range(Test_SSet.__len__()):



            a = 0
            Test_Set = Test_SSet.pop()

            FA_path = path+"\\"+algotithm+"\\"+xmlSet

            nname = []
            nGA_IP = []
            nGA_en = []
            nGA_IP_T = []
            TT=[]

            j = 1
            for filename in os.listdir(FA_path+"\\"+Test_Set):
                if filename.split(".")[1] == "png":
                    break
                print(os.path.join(FA_path+"\\"+Test_Set, filename))
                file = open(os.path.join(FA_path+"\\"+Test_Set, filename))
                nname.append(filename)
                ipl = []
                enl = []
                tl = []
                for i in range(6):
                    ip = (int)(file.readline())
                    en = (int)(file.readline())/1000
                    t = (float)(file.readline())
                    # print(en)
                    ipl.append(ip)
                    enl.append(en)
                    tl.append(t)
                # if t < 500:
                tt = (float)(file.readline())
                nGA_IP.append(ipl)
                nGA_en.append(enl)
                nGA_IP_T.append(tl)
                TT.append(tt)
                # a = ((a*j+t)/(j+1))
                # j+=1

                file.close()
            # average.append(a)

            name.append(nname)
            GA_IP_T.append(nGA_IP_T)
            GA_en.append(nGA_en)
            GA_IP.append(nGA_IP)
            tT.append(TT)

        print(name)

        write_excel(xmlSet+"_HPSStran", name,GA_IP,GA_en,GA_IP_T,tT)

        lie = lie + 3

