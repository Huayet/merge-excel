# coding=utf-8
# -*- coding: utf-8 -*-
''' 
Created on 2016-08-03
@author：孙华琛
@description: 把指定目录下所有excel文件，合并到一个表中，并添加“时间”和“公司名”两列
''' 
import os,sys
from xlutils.copy import copy
import xlrd as ExcelRead

#查看python默认编码
# print sys.getdefaultencoding()
#更改编码
reload(sys)   #必须要reload
sys.setdefaultencoding('utf-8')
# print sys.getdefaultencoding()

def write_append(old_file_name,new_file_name):

  #根据文件名获取公司名和时间
  filename = old_file_name.split('.')[0] #去除.xls
  company_and_time = filename.split('_')[1] #获取公司名和时间部分
  time    = company_and_time[-6:] #获取时间
  company = company_and_time[:-6] #获取公司名
  # ---------------------------------------
  #  ★  搞了一天这个编码问题！下面是解决问题遇到的经验
  #  ★  整体来说就是，需要先知道一个string的本身编码，然后再解码就变成了unicode，这时候才能给python使用！！
  #貌似time和company本身就是utf-8的字符了，不对 应该是GBK
  #先解码成unicode后才能编成其他码
  #print unicode("我是中文","utf-8")
  #类似u('xxx')给变量进行编码成unicode
  # company = eval("u'%s'" %(company))
  # ----------------------------------------
  # 把GBK解码成为unicode编码
  company  = company.decode("gbk")
  # print isinstance(company, unicode) #用来判断是否为unicode
  # print u'公司名：',company
  # print u'时间  ：',time

  #需要追加的内容测试
  #values = [u"利润总额", "2", "166533", "-2698" , "-0.0162"] 
  
  #打开源文件(需要复制的文件)
  r_xls = ExcelRead.open_workbook(old_file_name)
  #获取源文件第一个工作表的句柄
  r_sheet = r_xls.sheet_by_index(0)


  #获取目标excel的现在行数
  tar_xls = ExcelRead.open_workbook(new_file_name)
  #获取第一个工作表的句柄
  tar_sheet = tar_xls.sheet_by_index(0)
  #获取该工作表的行数
  rows = tar_sheet.nrows

  #拷贝 源工作表 的内容
  w_xls = copy(tar_xls)  #(r_xls)
  #选取第一个工作表
  sheet_write = w_xls.get_sheet(0)
  #根据实际情况，选定了6-29行的数据
  for j in range(5,29):
    #获取每一行 待追加的数据
    values = r_sheet.row_values(j)
    #遍历每一行value，追加到excel的最底端
    for i in range(0, len(values)):
        # print u'————————— 正在插入数据：——————————'
        # print values[i]
        # print u'————————— 本条数据成功插入：——————'
        sheet_write.write(rows, 0, time) #在坐标(rows,0)插入data内容
        sheet_write.write(rows, 1, company)
        sheet_write.write(rows, i+1, values[i])
    #每写入一行，Y轴向下移动一行
    rows = rows + 1
  #把文件进行保存操作
  w_xls.save(new_file_name)
  #w_xls.save(file_name + '.out' + os.path.splitext(file_name)[-1]); 


#======================================================================


# 获取指定路径下所有指定后缀的文件
# dir 指定路径
# ext 指定后缀，链表&不需要带点 或者不指定。例子：['xml', 'java']
def GetFileFromThisRootDir(dir,ext = None):
    allfiles = []
    needExtFilter = (ext != None)
    for root,dirs,files in os.walk(dir):
        for filespath in files:
            #只获取文件名，就不用连接路径了
            filepath = filespath
            #filepath = os.path.join(root, filespath)
            extension = os.path.splitext(filepath)[1][1:]
            if needExtFilter and extension in ext:
                allfiles.append(filepath)
            elif not needExtFilter:
                allfiles.append(filepath)
    return allfiles

#======================================================================

if __name__ == "__main__":
    #设置输出文件名
    result = "test.xls"
    #统计处理了多少文件
    count = 1
    #获取当前py工作目录下所以.xls文件的文件名
    for i in GetFileFromThisRootDir(os.getcwd(),['xls']):
        print '-------------------'
        print u'正在处理第',count,u'个文件 -> ',i,' <-'
        print '-------------------'
        count += 1
        #如果现在扫描到了输出文件，就跳过它
        if i == result:
            print u'跳过输出文件：',result
            continue
        #追加内容到输出文件
        write_append(i,result)