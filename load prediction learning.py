import xlrd
import xlwt
import datetime
import heapq
import math
#这里面是参数
k1=1                        #与星期相关的权重
k2=0.5                      #与间隔时间距离相关的权重
n_n=9                       #n的平方
epsilon=0.0000001           #微小量
Load_begin_col=3            #负荷的起始列，从0开始，目前为第4列
Day_measure_point=96        #日采样率点数 96点
Province_eg=['广东','广西','云南','贵州','海南']
#参数结束

def Get_96point_Theday(predate):#获取predate参数(datetime类)所表示的日期的那一行数据（即每日96点负荷）。返回一个二维list，[[A省当日96点数据]，[B省当日96点数据]...]

    Gdate=xlrd.xldate.xldate_from_date_tuple((predate.year,predate.month,predate.day),0)
    #print(Gdate)
    real_loadtheday=[]
    i=0
    for k in Load_data_sheet1.col(0):
        #print(k.value)
        if k.value!="日期":
            if abs(float(k.value)-Gdate)<0.00001:
                real_loadtheday.append( Load_data_sheet1.row(i))
            
        i+=1
    return tuple( real_loadtheday)
#def结束了



def Get_TheProvince_datarow_index(Province_input="none",real_loadpast=None,Province_index=-1):#获取指定的Province_input所述省份所在数据的序号
    pv_index=-1
    for pv in real_loadpast:
        if pv[Province_index].value==Province_input:
            pv_index=real_loadpast.index(pv)
            break
    return pv_index
#def结束了    

#程序正式开始了
Load_data=xlrd.open_workbook("2009年-test.xls")             #打开负荷数据工作簿
Week_rel=xlrd.open_workbook("星期对应系数表.xls")               #打开星期对应系数表工作簿
Load_data_sheet1=Load_data.sheet_by_index(0)                #打开负荷数据工作簿第一个工作表（要求负荷数据放在第一个工作表
Week_rel_sheet1=Week_rel.sheet_by_index(0)                  #打开星期对应系数表工作簿第一个工作表（要求星期系数放在第一个工作表
Load_data_sheet1_title=Load_data_sheet1.row_values(0)       #获取负荷数据的表头
Province_index=Load_data_sheet1_title.index("省区")         #找到表头中表示“省区”列的位置
Prediction_file=xlwt.Workbook()
Prediction_file_sheet=Prediction_file.add_sheet('预测值')   #创建预测值工作表
i=0
for temp_cell in Load_data_sheet1_title:
    Prediction_file_sheet.write(0,i,temp_cell)         #创建表头
    i+=1
Prediction_file.save('预测值.xls')
#print(Province_index)
i=0
Week_rel_sheet1_data=[] #将星期对应系数放置在该变量中（不包含表头）
while i < Week_rel_sheet1.nrows-2:
    Week_rel_sheet1_data.append(Week_rel_sheet1.row_values(i+2,1))
    #print(Week_rel_sheet1_data[i][1])
    i+=1

predateinput="2009-4-25"
Pretimeinput=input("输入当前时间点：（10:00）")
Pretimeinput="10:00"
pretime=Pretimeinput.split(':')
predate=predateinput.split('-')
predateout=datetime.datetime(year=int(predate[0]),month=int(predate[1]),day=int(predate[2]),hour=int(pretime[0]),minute=int(pretime[1]),second=0)
Theday_96point=Get_96point_Theday(predateout)
#print(Theday_96point)
Theday_weekday=predateout.isoweekday()
#print(Theday_weekday)
Pre_load_delta1=[]
Pre_load_delta2=[]
Diff_sign=[]
for i in range(0,15):
    d=predateout-datetime.timedelta(days=i)
    Pre_load_delta1.append(Week_rel_sheet1_data[d.isoweekday()-1][Theday_weekday-1])
    Pre_load_delta2.append(math.sqrt(i/n_n))
#print(Pre_load_delta1)
#map(Pre_load_delta1.index, heapq.nsmallest(3, Pre_load_delta1))
for i in range(0,15):
    if i<n_n:
        Diff_sign.append(Pre_load_delta1[i]*k1+Pre_load_delta2[i]*k2)
    else:
        Diff_sign.append(Pre_load_delta1[i]*k1+1*k2)
    
#print(Diff_sign)
temp=[]
Inf = 9999
for i in range(4):
    temp.append(Diff_sign.index(min(Diff_sign)))
    Diff_sign[Diff_sign.index(min(Diff_sign))]=Inf
temp.sort()
real_loadpast=[]
print("选取的相似日为:")
for i in temp:
    if i !=0:
        real_loadpast.append(Get_96point_Theday(predateout-datetime.timedelta(days=i)))
        print(predateout-datetime.timedelta(days=i))


write_row=1
for province in Province_eg:
    Province_input=province
    Prediction_load_init=[xlrd.xldate.xldate_from_date_tuple((predateout.year,predateout.month,predateout.day),0),'0',Province_input+'预测']
    Prediction_load_theday=[]
    Prediction_load_theday.extend(Prediction_load_init)
    avg_load=0
    for i in range (Day_measure_point):
        avg_load=0
        for j in range(3):
            pv_index=Get_TheProvince_datarow_index(Province_input=Province_input,real_loadpast=real_loadpast[j],Province_index=Province_index)
            if abs (pv_index+1)<epsilon:
                print("历史第",j+1,"日",province,"数据未找到，请重新输入")
            else:
                avg_load+=real_loadpast[j][pv_index][i+Load_begin_col].value
                #print(avg_load)
        Prediction_load_theday.append(avg_load/3)
    print(Prediction_load_theday)
    print(len(Prediction_load_theday))

    Time_offset=predateout.hour*4+int(predateout.minute/15)+Load_begin_col
    Prediction_offset=0
    for i in range(3):
        pv_index=Get_TheProvince_datarow_index(Province_input=Province_input,real_loadpast=Theday_96point,Province_index=Province_index)
        if abs (pv_index+1)<epsilon:
            print("第",i+1,"日",province,"数据未找到，请重新输入")
        else:
            Prediction_offset+=Theday_96point[pv_index][Time_offset-i].value-Prediction_load_theday[Time_offset-i]
    Prediction_offset=Prediction_offset/3
    #print(Prediction_offset)
    delta_load=Prediction_load_init
    for i in range(96):
        Prediction_load_theday[i+Load_begin_col]+=Prediction_offset
        delta_load.append( Prediction_load_theday[i+Load_begin_col]-Theday_96point[pv_index][i+Load_begin_col].value)
    print(Prediction_load_theday)
    print("---------------------------------------")
    print(delta_load)
    delta_load_temp=delta_load[Load_begin_col:]
    temp=[]
    Inf = 0
    for i in range(4):
        temp.append(delta_load_temp.index(max(delta_load_temp)))
        #print(max(delta_load_temp))
        delta_load_temp[delta_load_temp.index(max(delta_load_temp))]=Inf
    temp.sort()
    print(temp)

    i=0
    for temp_cell in Prediction_load_theday:
        Prediction_file_sheet.write(write_row,i,temp_cell)         #存入预测负荷值
        i+=1
    write_row+=1

    i=0
    for temp_cell in Theday_96point[pv_index]:
        Prediction_file_sheet.write(write_row,i,temp_cell.value)         #存入实际负荷值
        i+=1
    write_row+=1
    
Prediction_file.save('预测值.xls')