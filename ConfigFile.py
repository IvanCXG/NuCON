import xlrd
# import tkinter
# from tkinter import *
# from time import ctime
# import datetime
import win32api,win32con
#读取excel表中的sheet，和每个sheet中的行数
data1 = xlrd.open_workbook('./test2.xlsx')
table1 = data1.sheets()[0]
table2 = data1.sheets()[1]
table3 = data1.sheets()[2]
table4 = data1.sheets()[3]
r_anacnt=table1.nrows
r_digcnt=table2.nrows
w_anacnt=table3.nrows
w_digcnt=table4.nrows
print(table1)
#空列表
rana_list = []
rdig_list = []
wana_list = []
wdig_list = []


def ConfigFile():

    #将每个sheet中的每一行的值存放到对应的列表中
    for i in range(0,r_anacnt):#读功能的模拟量点名
        _list1=table1.row_values(i)
        for item in _list1:
            rana_list.append(item)
    for i in range(0,r_digcnt):#读功能的开关量点名
        _list2=table2.row_values(i)
        for item in _list2:
            rdig_list.append(item)
    for i in range(0,w_anacnt):#回写功能的模拟量点名
        _list3=table3.row_values(i)
        for item in _list3:
            wana_list.append(item)
    for i in range(0,w_digcnt):#回写功能的开关量点名
        _list4=table4.row_values(i)
        for item in _list4:
            wdig_list.append(item)

    #写配置文件的位置
    file = open('./opcdaserver.xml','w')

    #向配置文件中写入内容
    file.write('<?xml version="1.0"?>\n')
    file.write('<Gateway>\n')
    file.write('	<Setting LocalDrop="230" Redundancy="Simplex" PartnerDrop="0" Type="opcdas">\n')
    file.write('		<Startup Auto="True" />\n')
    file.write('	</Setting>\n')
    file.write('	<GatewayTagDef ConfigTimeSecs="1478671911" PrefixEnable="True">\n')
    #读点模拟量配置
    file.write('		<AnalogInput ItemCount="'+str(r_anacnt)+'">\n')
    for i in range(0,r_anacnt):
        file.write('			<Item Tag="'+str(rana_list[i])+'" Source="" Access="R" LowerLimit="0.000000" UpperLimit="100.000000" DataType="REAL" />\n')
    file.write('		</AnalogInput>'+'\n')

    #回写点模拟量配置
    file.write('		<AnalogOutput ItemCount="'+str(w_anacnt)+'">\n')
    for i in range(0,w_anacnt):
        file.write('			<Item Tag="'+str(wana_list[i])+'" Source="" Access="W" LowerLimit="0.000000" UpperLimit="100.000000" DataType="REAL">\n')
        file.write('				<TagDef Timeout="0" Desc="" Character="" AlmGrp="" Unit="" Format="" ExcDZone="0.000000" Share="FALSE" />\n')
        file.write('			</Item>\n')
    file.write('		</AnalogOutput>\n')

    #读点开关量配置
    file.write('		<Input ItemCount="'+str(r_digcnt)+'">\n')
    for i in range(0,r_digcnt):
        file.write('		    <Item Tag="'+str(rdig_list[i])+'" Source="" Access="R" />\n')
    file.write('		</Input>\n')

    #回写点开关量配置
    file.write('		<Output ItemCount="'+str(w_digcnt)+'">\n')
    for i in range(0,w_digcnt):
        file.write('		    <Item Tag="'+str(wdig_list[i])+'" Source="" Access="W">\n')
        file.write('		        <TagDef Timeout="0" Desc="" Character="" AlmGrp="" ZeroDesc="" OneDesc="" Share="FALSE" />\n')
        file.write('		    </Item>\n')
    file.write('		</Output>\n')
    file.write('	</GatewayTagDef>\n')
    file.write('</Gateway>')

# def Msg():
#     #创建顶层窗口、相关控件
#     top_layer = tkinter.Tk()
#     #窗口标题
#     top_layer.title('提示')
#     label = tkinter.Label(top_layer, text='配置文件生成完毕！',font=('宋体',14,'bold'))
#     label.pack()
#     button = tkinter.Button(top_layer, text='退出',command=top_layer.quit)
#     button.pack()
#
#     #设置窗口位置
#     top_layer.resizable(width=200, height=50)
#     screen_width = top_layer.winfo_screenwidth()
#     screen_hight = top_layer.winfo_screenheight()
#     print(screen_width)
#     print(screen_hight)
#     window_w = top_layer.winfo_reqwidth()
#     window_h = top_layer.winfo_reqheight()
#     print(window_w)
#     print(window_h)
#     x_local = (screen_width - window_w) / 2
#     y_local = (screen_hight - window_h) / 2
#     top_layer.geometry('%dx%d+%d+%d' %(window_w,window_h,x_local,y_local))
#
#     #禁止窗口最大、最小化、窗口拉伸
#     top_layer.maxsize(width=False, height=False)
#     top_layer.minsize(width=False, height=False)
#     top_layer.resizable(width=False, height=False)
#
#     #循环
#     top_layer.mainloop()

if __name__ == '__main__':
    ConfigFile()
    win32api.MessageBox(0,'配置文件生成完成！','提示信息')
