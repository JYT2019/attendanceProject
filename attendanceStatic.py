import xlwt
from xlrd import open_workbook
from xlutils.copy import copy
import os
import time

def buildexcel(data):
    filename = "data.xls"
    #print(os.path.exists(filename))
    if not os.path.exists(filename):
        workbook = xlwt.Workbook(encoding="utf-8")
        sheet1 = workbook.add_sheet('Sheet 1', cell_overwrite_ok=True)
        
        sheet1.write(0, 0, "姓名")
        sheet1.write(0, 1, "性别")
        sheet1.write(0, 2, "年龄")
        sheet1.write(0, 3, "职位")
        sheet1.write(0, 4, "部门")
        sheet1.write(0, 5, "签名")
        sheet1.write(0, 6, "上班时间")
        sheet1.write(0, 7, "下班时间")

        workbook.save(filename)

    rexcel = open_workbook(filename)
    rows = rexcel.sheets()[0].nrows
    excel = copy(rexcel)
    table = excel.get_sheet(0)
    
    localtime = time.strftime("%Y-%m-%d %H:%M:%S",time.localtime(time.time()))
    
    data.append(localtime)
    data.append(localtime)
    
    for i, col in enumerate(data):
        table.write(rows,i,col)
    
    excel.save(filename)

#更新Excel表中指定元素
def updatexcel():
    strStr = ''
    filename = "data.xls"
    
    fileObj = xlrd.open_workbook(filename)
    sheetfirst = fileObj.sheets()[0]
    
    row_count = fileObj.sheets()[0].nrows
    col_count = fileObj.sheets()[0].ncols
    
    fileWb = copy(fileObj)
    sheet = fileWb.get_sheet(0)
    
    localtime = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time()))
    
    for element in range(row_count):
        if strStr.lower() in (str(sheetfirst.row_values(element))).lower():
            for col in range(col_count):
                if strStr == sheetfirst.row_values(element)[col]
                    sheet.write(element, col+7, localtime)
                    fileWb.save(filename)
