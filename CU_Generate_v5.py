import xlrd
import datetime
#import time
import win32api
from time import *
analogPage=50
digitalPage=50
analogPageNum=100
digitalPageNum=100

def CU_Generate():    
    #写CU文件头
    now1=datetime.datetime.now()
    RevTime1=str(now1.strftime('%Y-%m-%d %H:%M:%S'))
    file = open('./CU06.txt', 'w',encoding='UTF-8')
    file.write('NuCon Cu File \n\n'
               'FileHead \n'
               'Version=2.0.0.0\n'
               'Drop=6 \n'
               'Description=\n'
               'Project=\n'
               'Profile=10A319\n'
               'Temperature=80\n'
               'CpuLoad=80\n'
               'MemLoad=80\n'
               'MaxAxId='+str(analogPage*analogPageNum-1)+'\n'
               'MaxDxId='+str(digitalPage*digitalPageNum-1)+'\n'
               'MaxExchangeId=60\n'
               'NetworkRedundancy=2\n'
               'FileLastUpdate='+str(RevTime1)+'\n')
    #time.sleep(1)
    t = str(int(time()))
    print(t)
    now2=datetime.datetime.now()
    RevTime2=str(now2.strftime('%Y-%m-%d %H:%M:%S'))
    file.write('PointDirLastUpdate='+str(RevTime2)+' V1\n'
               'FileHeadEnd\n\n'
               'Class1OutputTimestamp='+ t +'\n'
               'Class1OutputExchange,1,1,1,'+ t +',0,40000000\n\n\n')

    #读取Excel中的点名，并存储到list中
    data1 = xlrd.open_workbook('./test1.xlsx')
    table1 = data1.sheets()[0]
    table2 = data1.sheets()[1]
    table3 = data1.sheets()[2]
    table4 = data1.sheets()[3]
    ana_list = []
    dig_list = []
    ana_desc = []
    dig_desc = []
    for i in range(0,analogPage):#读取模拟量点名
        for j in range(0,analogPageNum):
            _list1=table1.row_values(100 * i + j)
            for item1 in _list1:
                ana_list.append(item1)
    for i in range(0,digitalPage):#读取开关量点名
        for j in range(0,digitalPageNum):
            _list2=table2.row_values(100 * i + j)
            for item2 in _list2:
                dig_list.append(item2)
    for i in range(0,analogPage):#读取模拟量点描述
        for j in range(0,analogPageNum):
            _list3=table3.row_values(100 * i + j)
            for item3 in _list3:
                ana_desc.append(item3)
    for i in range(0,digitalPage):#读取开关量点描述
        for j in range(0,digitalPageNum):
            _list4=table4.row_values(100 * i + j)
            for item4 in _list4:
                dig_desc.append(item4)

    #模拟量标签点
    for i in range(0,analogPage):
        k=165
        l=68
        c=0
        file.write('Page, '+str(i+1)+':'+str(10*(i+1))+', 4 x10ms 3 0 0\n')
        file.write('	Description='+str(i+1)+'\n')
        file.write('	RevTime='+RevTime2+'\n')
        file.write('	Sub=\n')
        for j in range(0,analogPageNum):      
            val1=int(j/18)
            val3=j%18
            file.write('	Func, NetAO, '+str(j+1)+':'+str(10*(j+1))+', ('+str(k+100*val1)+','+str(l+30*val3)+'), '+str(c)+', '+str(c)+' \n')
            file.write('		In= ,B102-0,\n')
            file.write('		Para= '+str(100*i+j)+','+str(ana_list[100*i+j])+',0,0,1,0,0,-99999.9,0,-99999.9,0,50,3,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,-99999.9,0,0,0,0,0,0,0,\n')
            file.write('		Out= ,\n')
            file.write('Page='+str(i+1)+','+str(i+1)+'\n')
            file.write('	FuncEnd\n')
        file.write('	Func, Signal, 101:1020, (75,218), 0, 0\n'
                   '		In= ,0d, 0d,\n'
                   '		Para= 3,50,2,50,\n'
                   '		Out= ,0, 0,\n'
                   )
        file.write('Page='+str(i+1)+','+str(i+1)+'\n')
        file.write('	FuncEnd\n')
        file.write('	EndDesc=\n'
                   '	SubProfile=AFC1C033\n'
                   'PageEnd\n\n')

    #开关量标签点
    for i in range(0,digitalPage):
        d=105
        t=68
        c1=0
        file.write('Page, '+str(i+51)+':'+str(10*(i+1))+', 100 x10ms 3 0 0'+'\n')
        file.write('	Description=DI'+str(i+1)+'\n')
        file.write('	RevTime='+RevTime2+'\n')
        file.write('	Sub=\n')
        for j in range(0,digitalPageNum):
            val2=int(j/18)
            val4=j%18
            file.write('	Func, NetDO, '+str(j+2)+':'+str(10*(j+2))+', ('+str(d+100*val2)+','+str(t+30*val4)+'), '+str(c1)+', '+str(c1)+' \n')
            file.write('		In= ,B102-0,'+'\n')
            file.write('		Para= '+str(100*i+j)+','+str(dig_list[100*i+j])+',256,0,1,0,0,2,0,0,0,0,'+'\n')
            file.write('		Out= ,'+'\n')
            file.write('Page='+str(i+51)+',DI'+str(i+1)+'\n')
            file.write('	FuncEnd'+'\n')
        file.write('	Func, Not, 102:1020, (15,53), 0, 0\n'
                   '		In= ,B102-0, \n'
                   '		Para= \n'
                   '		Out= ,0,\n'
                   )
        file.write('Page='+str(i+51)+',DI'+str(i+51)+'\n')
        file.write('	FuncEnd'+'\n')
        file.write('	EndDesc= \n'
                '	SubProfile=B5DC2017\n'
               'PageEnd\n\n')

    #标签点参数

    #模拟量数据源
    file.write('[POINT_DIR_INFO]\n')
    file.write('BEGIN_AX \n')
    for i in range(0,analogPage):
        for j in range(0,analogPageNum):
            a=0
            b=32
            c=0
            d=1
            e=2
            file.write(str(ana_list[100*i+j])+'=20,'+str(ana_desc[100*i+j])+',--------,--------,,7.2,100,0,'+str(i+1)+','+str(a+80*j)+','+str(b+80*j)+','+str(c+j)+','+str(i+1)+','+str(e+j)+'\n')
    file.write('END_AX\n')
    file.write('\n')

    #开关量数据源
    file.write('BEGIN_DX\n')
    for i in range(0,digitalPage):
        for j in range(0,digitalPageNum):
            a1=8000
            b1=8001
            c1=0
            d1=51
            e1=2
            file.write(str(dig_list[100*i+j])+'=20,'+str(dig_desc[100*i+j])+',--------,--------,false,true,'+str(36+j)+','+str(a1+32*j)+','+str(b1+32*j)+','+str(c1+j)+','+str(d1+i)+','+str(e1+j)+'\n')
    file.write('END_DX\n')
    
if __name__=='__main__':
    CU_Generate()
    win32api.MessageBox(0,'组态生成完成！','提示信息')
    
