import xlrd
import xlwt
Load_data=xlrd.open_workbook("2009年.xls")
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
