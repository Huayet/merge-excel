# -*- coding: utf-8 -*-
''' 
Created on 2016-08-03
@author：孙华琛
@description:测试提取文件名中的信息，测试读取excel信息
''' 
import xlrd
import xlwt
from datetime import date,datetime

def read_excel():
    file = u'预算完成情况综合表_湖南XXX有限公司201506.xls'
	#根据文件名获取公司名和时间
    filename = file.split('.')[0] #去除.xls
    company_and_time = filename.split('_')[1] #获取公司名和时间部分
    time    = company_and_time[-6:] #获取时间
    company = company_and_time[:-6] #获取公司名
    print u'公司名：',company
    print u'时间  ：',time
    # 打开文件
    workbook = xlrd.open_workbook(file)
    # 获取所有sheet
    print workbook.sheet_names() # [u'sheet...', u'sheet...']
    sheet_name = workbook.sheet_names()[0]

    # 根据sheet索引或者名称获取sheet内容
    sheet = workbook.sheet_by_index(0) # sheet索引从0开始
    #sheet = workbook.sheet_by_name('sheet')

    # sheet的名称，行数，列数
    print sheet.name,sheet.nrows,sheet.ncols
    # 获取整行和整列的值（数组）
    rows = sheet.row_values(3) # 获取第四行内容
    cols = sheet.col_values(2) # 获取第三列内容
    print rows
    print cols

    # 获取单元格内容
    print sheet.cell(1,0).value.encode('utf-8')
    print sheet.cell_value(1,0).encode('utf-8')
    print sheet.row(1)[0].value.encode('utf-8')
    
    # 获取单元格内容的数据类型
    print sheet.cell(1,0).ctype

if __name__ == '__main__':
    read_excel()