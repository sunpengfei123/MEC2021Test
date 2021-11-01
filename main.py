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
def write_excel(ename, name, IP, Time1,Time2):

    f = xlwt.Workbook()

    #创建sheet1
    sheet1 = f.add_sheet(u'sheet2',cell_overwrite_ok=True)
    row0 = [u'cases',u're',u'CP_Time',u'Ant_Time']


    for t in range(0,len(name)):

        column0 = name.__getitem__(t)
        column1 = IP.__getitem__(t)
        column2 = Time1.__getitem__(t)
        column3 = Time2.__getitem__(t)

        lie = t*5


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
        ave1 = 0
        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie+2, column2[i], set_stlye('Times New Roman', 220, True))
            ave1 = (ave1*i+column2[i])/(i+1)

        ave2 = 0
        for i in range(0, len(column0)):
            sheet1.write(i + 1, lie+3, column3[i], set_stlye('Times New Roman', 220, True))
            ave2 = (ave2 * i + column3[i]) / (i + 1)

        sheet1.write(len(column0) + 2, lie , 'average', set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0)+2, lie + 2, ave1, set_stlye('Times New Roman', 220, True))
        sheet1.write(len(column0) + 2, lie + 3, ave2, set_stlye('Times New Roman', 220, True))

    f.save(ename+'.xls')



path = "C:\\Study\\TestCase\\2021_MEC\\result"

# xmlSSet = {"Test1_Muti_Pro","Test2_Muti_Pro","acyclic_degree4_Muti_Pro","acyclic_degree6_Muti_Pro","cyclic_degree2_Muti_Pro","cyclic_degree4_Muti_Pro","cyclic_degree6_Muti_Pro","sc_degree2_Muti_Pro","sc_degree4_Muti_Pro","sc_degree6_Muti_Pro"}

xmlSSet = {"ac_SDFG_mutipro"}

algotithms = ["SDFG_IP_CL_Test",]

Test_SSet = ["a10","a5"]#

lie = 0

for algotithm in algotithms:

    for xmlSet in xmlSSet:
        average = []
        print(xmlSet)
        Test_SSet = ["a10","a5"]#

        name = []
        GA_IP = []
        GA_IP_T = []
        GA_IP_T2 = []

        for i in range(Test_SSet.__len__()):

            a = 0
            Test_Set = Test_SSet.pop()

            FA_path = path+"\\"+algotithm+"\\"+xmlSet

            nname = []
            all_re = []
            all_t = []
            all_t2 = []

            for filename in os.listdir(FA_path+"\\"+Test_Set):
                if filename.split(".")[1] == "png":
                    break
                # print(os.path.join(FA_path+"\\"+Test_Set, filename))
                file = open(os.path.join(FA_path+"\\"+Test_Set, filename))
                nname.append(filename)
                re = (int)(file.readline().split(":")[1])
                ip1 = (int)(file.readline().split(":")[1])
                ip2 = (int)(file.readline().split(":")[1])
                t1 = (float)(file.readline().split(":")[1])
                t2 = (float)(file.readline().split(":")[1])
                # t = (float)(file.readline())
                # if t < 500:
                all_re.append(re)
                all_t.append(t1)
                all_t2.append(t2)
                file.close()

            name.append(nname)
            GA_IP_T.append(all_t)
            GA_IP.append(all_re)
            GA_IP_T2.append(all_t2)

        print(name)

        write_excel(algotithm+'_'+xmlSet, name,GA_IP,GA_IP_T,GA_IP_T2)

        lie = lie + 3
