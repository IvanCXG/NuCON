import xlrd
import datetime
#import time
from time import *
#from page_calc import p_calc
import win32api
data = xlrd.open_workbook('./datasrc.xlsx')
sheet_data = data.sheet_by_index(0) #读取sheet1中的全部内容
total = sheet_data.nrows  # sheet1中的总行数
ana_name = []   #存储模拟量点名
ana_desc = []   #存储模拟量描述
ana_unit = []   #存储模拟量单位
ana_uplim = []  #存储模拟量上限值
ana_lolim = []  #存储模拟量下限值
ana_alarm = []  #存储模拟量报警优先级
ana_alarm_v = []    #存储模拟量报警整定值
dig_name = []   #存储开关量点名
dig_desc = []   #存储开关量描述
dig_unit = []   #存储开关量单位
dig_alarm = []  #存开关量报警优先级
dig_alarm_v = []    #存储开关量报警整定值（0或1）
cell_value3 = []  # 单元格的值
ana_cst = 0  # 模拟量计数器
dig_cst = 0  # 开关量计数器

# 将sheet表中的数据按照每一行进行读取，然后将读取的每一行数据存储到列表row_data中
for i in range(1, total):
    cell_value3 = sheet_data.cell_value(i, 3)  # 读取第四列的数据类型，用于判断是模拟量还是开关量

    # 以下：判断模拟量和开关量点，并分别存储在ana_name列表和dig_name列表中
    if cell_value3 == 'Analog':
        ana_name.append(sheet_data.cell_value(i, 0))
        ana_desc.append(sheet_data.cell_value(i, 1))
        ana_unit.append(sheet_data.cell_value(i, 2))
        ana_uplim.append(sheet_data.cell_value(i, 4))
        ana_lolim.append(sheet_data.cell_value(i, 5))
        ana_alarm.append(sheet_data.cell_value(i, 6))
        ana_alarm_v.append(sheet_data.cell_value(i, 7))
        ana_cst += 1
    elif cell_value3 == 'Digital':
        dig_name.append(sheet_data.cell_value(i, 0))
        dig_desc.append(sheet_data.cell_value(i, 1))
        dig_unit.append(sheet_data.cell_value(i, 2))
        dig_alarm.append(sheet_data.cell_value(i, 6))
        dig_alarm_v.append(sheet_data.cell_value(i, 7))
        dig_cst += 1
#print(ana_uplim)
print(ana_unit)

#计算是否是整数页
def p_calc(num):
    if num % 100 != 0:
        page_cst = int(num / 100) + 1
        return (page_cst)
    elif num % 100 == 0:
        page_cst = int(num / 100)
        #print(page_cst)
        return (page_cst)

ana_page = p_calc(ana_cst)
dig_page = p_calc(dig_cst)
print('模拟量页数 =',ana_page)
print('开关量页数 =',dig_page)
def CU_Generate():
    # 写CU文件头
    now1 = datetime.datetime.now()
    RevTime1 = str(now1.strftime('%Y-%m-%d %H:%M:%S'))
    file = open('./CU09.txt', 'w',encoding='UTF-8')
    file.write('NuCon Cu File\n\n' 
                'FileHead \n'
                'Version=3.0.0.0\n'
                'Drop=9 \n'
                'Description=\n'
                'Project=\n'
                'Profile=10A319\n'
                'Temperature=80\n'
                'CpuLoad=80\n'
                'MemLoad=80\n'
                'MaxAxId=' + str(ana_cst) + '\n'
                'MaxDxId=' + str(dig_cst) + '\n'
                'MaxExchangeId=60\n'
                'NetworkRedundancy=2\n'
                'FileLastUpdate=' + str(RevTime1) + '\n'
               )
    #time.sleep(1)
    t = str(int(time()))
    now2 = datetime.datetime.now()
    RevTime2 = str(now2.strftime('%Y-%m-%d %H:%M:%S'))
    file.write('PointDirLastUpdate=' + str(RevTime2) + ' V1\n'
                'FileHeadEnd\n\n'
                'Class1OutputTimestamp='+ t + '\n'
                'Class1OutputExchange,1,1,1,'+ t +',0,40000000' + '\n\n\n'
                )

    # 模拟量标签点
    if ana_cst % 100 != 0:  #非整数页
        for i in range(0, ana_page-1):
            k = 165
            l = 68
            c = 0
            file.write('Page, ' + str(i + 1) + ':' + str(10 * (i + 1)) + ', 20 x10ms 3 0 0\n')
            file.write('	Description=' + str(i + 1) + '\n')
            file.write('	RevTime=' + RevTime2 + '\n')
            file.write('	Sub=\n')
            for j in range(0, 100): #
                val1 = int(j / 18)
                val3 = j % 18
                file.write('	Func, NetAO, ' + str(j + 2) + ':' + str(10 * (j + 2)) + ', (' + str(k + 100 * val1) \
                           + ',' + str(l + 30 * val3) + '), ' + str(c) + ', ' + str(c) + ' \n')
                file.write('		In= ,B102-0,' + '\n')
                file.write('		Para= ' + str(100 * i + j) + ',' + str(ana_name[100 * i + j]) + \
                           ',0,0,1,0,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,0,0,0,0,0,0,\n' )
                file.write('		Out= ,\n')
                file.write('Page=' + str(i + 1) + ',' + str(i + 1) + '\n')
                file.write('	FuncEnd\n')
            file.write('	Func, Signal, 102:1020, (75,218), 0, 0\n'
                           '		In= ,0d, 0d,\n'
                           '		Para= 3,50,100,50,\n'
                           '		Out= ,0, 0,\n'
                           )
            file.write('Page=' + str(i + 1) + ',' + str(i + 1) + '\n')
            file.write('	FuncEnd\n')
            file.write('	EndDesc=\n'
                           '	SubProfile=AFC1C033\n'
                           'PageEnd\n\n')

        for i in range(ana_page-1,ana_page):
            k = 165
            l = 68
            c = 0
            file.write('Page, ' + str(i + 1) + ':' + str(10 * (i + 1)) + ', 20 x10ms 3 0 0\n')
            file.write('	Description=' + str(i + 1) + '\n')
            file.write('	RevTime=' + RevTime2 + '\n')
            file.write('	Sub=' + '\n')
            for j in range(0, ana_cst % 100):
                val1 = int(j / 18)
                val3 = j % 18
                file.write('	Func, NetAO, ' + str(j + 2) + ':' + str(10 * (j + 2)) + ', (' + str(
                    k + 100 * val1) + ',' + str(l + 30 * val3) + '), ' + str(c) + ', ' + str(c) + ' \n')
                file.write('		In= ,B102-0,\n')
                file.write('		Para= ' + str(100 * i + j) + ',' + str(ana_name[100 * i + j]) + \
                           ',0,0,1,0,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,0,0,0,0,0,0,\n')
                file.write('		Out= ,\n')
                file.write('Page=' + str(i + 1) + ',' + str(i + 1) + '\n')
                file.write('	FuncEnd\n')
            file.write('	Func, Signal, 102:1020, (75,218), 0, 0\n'
                       '		In= ,0d, 0d,\n'
                       '		Para= 3,50,100,50,\n'
                       '		Out= ,0, 0,\n')
            file.write('Page=' + str(i + 1) + ',' + str(i + 1) + '\n')
            file.write('	FuncEnd\n')
            file.write('	EndDesc=\n'
                       '	SubProfile=AFC1C033\n'
                       'PageEnd\n\n'
                       )

    elif ana_cst % 100 == 0:    #整数页
        for i in range(0, ana_page):
            k = 165
            l = 68
            c = 0
            file.write('Page, ' + str(i + 1) + ':' + str(10 * (i + 1)) + ', 20 x10ms 3 0 0' + '\n')
            file.write('	Description=' + str(i + 1) + '\n')
            file.write('	RevTime=' + RevTime2 + '\n')
            file.write('	Sub=' + '\n')
            for j in range(0, 100):
                val1 = int(j / 18)
                val3 = j % 18
                file.write('	Func, NetAO, ' + str(j + 2) + ':' + str(10 * (j + 2)) + ', (' \
                           + str(k + 100 * val1) + ',' + str(l + 30 * val3) + '), ' + str(c) + ', ' + str(c) + ' \n')
                file.write('		In= ,B102-0,' + '\n')
                file.write('		Para= ' + str(100 * i + j) + ',' + str(ana_name[100 * i + j]) + \
                           ',0,0,1,0,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,0,0,0,0,0,0,' + '\n')
                file.write('		Out= ,\n')
                file.write('Page=' + str(i + 1) + ',' + str(i + 1) + '\n')
                file.write('	FuncEnd' + '\n')
            file.write('	Func, Signal, 102:1020, (75,218), 0, 0\n'
                           '		In= ,0d, 0d,\n'
                           '		Para= 3,50,100,50,\n'
                           '		Out= ,0, 0,\n'
                           )
            file.write('Page=' + str(i + 1) + ',' + str(i + 1) + '\n')
            file.write('	FuncEnd\n')
            file.write('	EndDesc=\n'
                           '	SubProfile=AFC1C033\n'
                           'PageEnd\n\n'
                       )

        # 开关量标签点
    if dig_cst % 100 != 0:  #点数不是页数的整数倍
        for i in range(0, dig_page-1):
            d = 105
            t = 68
            c1 = 0
            file.write('Page, ' + str(i + ana_page + 1) + ':' + str(10 * (i + 1)) + ', 20 x10ms 3 0 0' + '\n')
            file.write('	Description=DI' + str(i + 1) + '\n')
            file.write('	RevTime=' + RevTime2 + '\n')
            file.write('	Sub=' + '\n')
            for j in range(0, 100):
                val2 = int(j / 18)
                val4 = j % 18
                file.write('	Func, NetDO, ' + str(j + 2) + ':' + str(10 * (j + 2)) + ', (' + str(d + 100 * val2) \
                           + ',' + str(t + 30 * val4) + '), ' + str(c1) + ', ' + str(c1) + ' \n')
                file.write('		In= ,B102-0,' + '\n')
                file.write('		Para= ' + str(100 * i + j) + ',' + str(dig_name[100 * i + j]) +\
                           ',256,0,1,0,0,0,0,0,0,0,\n')
                file.write('		Out= ,' + '\n')
                file.write('Page=' + str(i + ana_page + 1) + ',DI' + str(i + 1) + '\n')
                file.write('	FuncEnd' + '\n')
            file.write('	Func, Not, 102:1020, (15,53), 0, 0\n'
                           '		In= ,B102-0, \n'
                           '		Para= \n'
                           '		Out= ,0,\n'
                       )
            file.write('Page=' + str(i + ana_page + 1) + ',DI' + str(i + 1) + '\n')
            file.write('	FuncEnd\n')
            file.write('	EndDesc= \n'
                           '	SubProfile=B5DC2017\n'
                           'PageEnd\n\n'
                       )

        for i in range(dig_page-1,dig_page):
            d = 105
            t = 68
            c1 = 0
            file.write('Page, ' + str(i + ana_page + 1) + ':' + str(10 * (i + 1)) + ', 20 x10ms 3 0 0\n')
            file.write('	Description=DI' + str(i + 1) + '\n')
            file.write('	RevTime=' + RevTime2 + '\n')
            file.write('	Sub=\n')
            for j in range(0, dig_cst % 100):
                val2 = int(j / 18)
                val4 = j % 18
                file.write('	Func, NetDO, ' + str(j + 2) + ':' + str(10 * (j + 2)) + ', (' + str(
                    d + 100 * val2) + ',' + str(t + 30 * val4) + '), ' + str(c1) + ', ' + str(c1) + ' \n')
                file.write('		In= ,B102-0,\n')
                file.write('		Para= ' + str(100 * i + j) + ',' + str(
                    dig_name[100 * i + j]) + ',256,0,1,0,0,0,0,0,0,0,\n')
                file.write('		Out= ,\n')
                file.write('Page=' + str(i + ana_page + 1) + ',DI' + str(i + 1) + '\n')
                file.write('	FuncEnd\n')
            file.write('	Func, Not, 102:1020, (15,53), 0, 0\n'
                        '		In= ,B102-0, \n'
                        '		Para= \n'
                        '		Out= ,0,\n'
                       )
            file.write('Page=' + str(i + ana_page + 1) + ',DI' + str(i + 1) + '\n')
            file.write('	FuncEnd\n')
            file.write('	EndDesc= \n'
                       '	SubProfile=B5DC2017\n'
                       'PageEnd\n\n')

    elif dig_cst % 100 == 0:    #点数是页数的整数倍
        for i in range(0, dig_page):
            d = 105
            t = 68
            c1 = 0
            file.write('Page, ' + str(i + ana_page + 1) + ':' + str(10 * (i + 1)) + ', 20 x10ms 3 0 0\n')
            file.write('	Description=DI' + str(i + 1) + '\n')
            file.write('	RevTime=' + RevTime2 + '\n')
            file.write('	Sub=\n')
            for j in range(0, 100):
                val2 = int(j / 18)
                val4 = j % 18
                file.write('	Func, NetDO, ' + str(j + 2) + ':' + str(10 * (j + 2)) + \
                           ', (' + str(d + 100 * val2) + ',' + str(t + 30 * val4) + '), ' + \
                           str(c1) + ', ' + str(c1) + ' \n')
                file.write('		In= ,B102-0,' + '\n')
                file.write('		Para= ' + str(100 * i + j) + ',' + str(dig_name[100 * i + j]) +\
                           ',256,0,1,0,0,0,0,0,0,0,\n')
                file.write('		Out= ,\n')
                file.write('Page=' + str(i + ana_page + 1) + ',DI' + str(i + 1) + '\n')
                file.write('	FuncEnd\n')
            file.write('	Func, Not, 102:1020, (15,53), 0, 0\n'
                       '		In= ,B102-0, \n'
                       '		Para= \n'
                       '		Out= ,0,\n'
                       )
            file.write('Page=' + str(i + ana_page + 1) + ',DI' + str(i + 1) + '\n')
            file.write('	FuncEnd\n')
            file.write('	EndDesc= \n'
                       '	SubProfile=B5DC2017\n'
                       'PageEnd\n\n'
                       )
        # 标签点参数

        # 模拟量数据源
    if ana_cst % 100 != 0:
        file.write('[POINT_DIR_INFO]\n')
        file.write('BEGIN_AX \n')
        for i in range(0, ana_page-1):
            for j in range(0, 100):
                a = 0
                b = 32
                c = 0
                d = 1
                e = 2
                file.write(str(ana_name[100 * i + j]) + '=20,' + str(ana_desc[100 * i + j]) +\
                           ',--------,--------,'+str(ana_unit[100*i+j])+',7.2,100,0,' + str(i + 1) + ',' + str(a + 80 * j) + \
                           ',' + str(b + 80 * j) + ',' + str(c + j) + ',' + str(i + 1) + ',' + str(e + j) + '\n')

        for i in range(ana_page-1,ana_page):
            for j in range(0, ana_cst % 100):
                a = 0
                b = 32
                c = 0
                d = 1
                e = 2
                file.write(str(ana_name[100 * i + j]) + '=20,' + str(ana_desc[100 * i + j]) +\
                           ',--------,--------,'+str(ana_unit[100*i+j])+',7.2,100,0,' + str(i + 1) + ',' + str(a + 80 * j) +\
                           ',' + str(b + 80 * j) + ',' + str(c + j) + ',' + str(i + 1) + ',' + str(e + j) + '\n')
        file.write('END_AX\n\n')

    elif ana_cst % 100 == 0:
        file.write('[POINT_DIR_INFO]\n')
        file.write('BEGIN_AX \n')
        for i in range(0, ana_page - 1):
            for j in range(0, 100):
                a = 0
                b = 32
                c = 0
                d = 1
                e = 2
                file.write(str(ana_name[100 * i + j]) + '=20,' + str(ana_desc[100 * i + j]) +\
                           ',--------,--------,'+str(ana_unit[100*i+j])+',7.2,100,0,' + str(i + 1) + ',' + str(a + 80 * j) + \
                           ',' + str(b + 80 * j) + ',' + str(c + j) + ',' + str(i + 1) + ',' + str(e + j) + '\n')
        file.write('END_AX\n\n')


    # 开关量数据源
    if dig_cst % 100 != 0:  #非整数页
        file.write('BEGIN_DX\n')
        for i in range(0, dig_page-1):
            for j in range(0, 100):
                a1 = 8000
                b1 = 8001
                c1 = 0
                d1 = 51
                e1 = 2
                file.write(str(dig_name[100 * i + j]) + '=20,' + str(dig_desc[100 * i + j]) + \
                       ',--------,--------,false,true,' + str(36 + j) + ',' + str(a1 + 32 * j) + \
                       ',' + str(b1 + 32 * j) + ',' + str(c1 + j) + ',' + str(d1 + i) + ',' + str(e1 + j) + '\n')
        #file.write('END_DX' + '\n')
        for i in range(dig_page-1, dig_page):
            for j in range(0, dig_cst % 100):
                a1 = 8000
                b1 = 8001
                c1 = 0
                d1 = 51
                e1 = 2
                file.write(str(dig_name[100 * i + j]) + '=20,' + str(dig_desc[100 * i + j]) +\
                           ',--------,--------,false,true,' + str(36 + j) + ',' + str(a1 + 32 * j) + \
                           ',' + str(b1 + 32 * j) + ',' + str(c1 + j) + ',' + str(d1 + i) + ',' + str(e1 + j) + '\n')
        file.write('END_DX\n')
    elif dig_cst % 100 == 0:    #整数页
        file.write('BEGIN_DX\n')
        for i in range(0, dig_page):
            for j in range(0, 100):
                a1 = 8000
                b1 = 8001
                c1 = 0
                d1 = 51
                e1 = 2
                file.write(str(dig_name[100 * i + j]) + '=20,' + str(dig_desc[100 * i + j]) + \
                           ',--------,--------,false,true,' + str(36 + j) + ',' + str(1 + 32 * j) + \
                           ',' + str(b1 + 32 * j) + ',' + str(c1 + j) + ',' + str(d1 + i) + ',' + str(e1 + j) + '\n')
        file.write('END_DX\n')

if __name__ == '__main__':
    CU_Generate()
    win32api.MessageBox(0, '组态生成完成！', '提示信息')
