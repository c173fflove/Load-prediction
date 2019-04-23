import xlrd
import xlwt
from datetime import date
def Get_96point_Theday(predate):#获取predate参数(yyyy-mm-dd)所表示的日期的那一行数据（即每日96点负荷）。返回一个二维list，[[A省当日96点数据]，[B省当日96点数据]...]

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
Load_data=xlrd.open_workbook("2009年-test.xls")
Week_rel=xlrd.open_workbook("星期对应系数表.xls")
Load_data_sheet1=Load_data.sheet_by_index(0)
Week_rel_sheet1=Week_rel.sheet_by_index(0)
Load_data_sheet1_title=Load_data_sheet1.row(0)
i=0
Week_rel_sheet1_data=[] #将星期对应系数放置在该变量中（不包含表头）
while i < Week_rel_sheet1.nrows-2:
    Week_rel_sheet1_data.append(Week_rel_sheet1.row_values(i+2,1))
    #print(Week_rel_sheet1_data[i][1])
    i+=1

predateinput=input("Input your date(2009-4-10)")
predate=predateinput.split('-')
predateout=date(year=int(predate[0]),month=int(predate[1]),day=int(predate[2]))
Theday_96point=Get_96point_Theday(predateout)
print(Theday_96point)

