#!/usr/bin/python
# -*- coding:utf-8 -*-
#filename:text_process.py
"""此脚本用于处理每月发来的效果加强客户数据文件，按照相应字段要求整理形成导入文件。
   因xlrd还不支持excel07版，故用时须将原始文件另存为03版"""

from pyExcelerator import *  #写入excel数据
import os
import xlrd #读取excel数据

industries={'[1]':'Computer Products','[2]':'Auto Parts & Accessories',#行业、序号
            '[3]':'Agriculture & Food','[4]':'Apparel & Accessories',
            '[5]':'Arts & Crafts','[6]':'Bags, Cases & Boxes',
            '[7]':'Chemicals','[8]':'Construction & Decoration',
            '[9]':'Consumer Electronics','[10]':'Electrical & Electronics',
            '[11]':'Furniture & Furnishing','[12]':'Health & Medicine',
            '[13]':'Light Industry & Daily Use','[14]':'Lights & Lighting',
            '[15]':'Machinery','[16]':'Metallurgy, Mineral & Energy',
            '[17]':'Office Supplies','[18]':'Security & Protection',
            '[19]':'Service','[20]':'Sporting Goods & Recreation',
            '[21]':'Textile','[22]':'Tools & Hardware','[23]':'Toys',
            '[24]':'Transportation','[25]':'Manufacturing & Processing Machinery',
            '[26]':'Industrial Equipment & Components','[27]':'Instruments & Meters',}

file_name=u'刘_7月份客户效果加强名单(0~5封).xls' #待处理文件名
wb=xlrd.open_workbook(os.path.abspath('..')+'\\'+file_name.encode('gbk'))#打开Excel
sh=wb.sheet_by_index(0)  #获取第一张sheet对象
row_count=sh.nrows  #行数
col_count=sh.ncols  #列数
if col_count==25:   #简单的检查机制：看原文件是否为25个字段
    w=Workbook()    #创建一个excel文件，待写入数据
    ws=w.add_sheet('Sheet1')  #增加一张sheet
    table_headers=[u'公司英文名称',u'公司ID',u'主要产品',  #表头内容
                   u'英文行业','Showroom',u'推广服务结束时间（年月日）',
                   u'效果情况及个人分析意见',u'客服人员',]
    i=0
    for item in table_headers:  #写入表头
        ws.write(0,i,item)
        i+=1
    for i in range(1,row_count):   
        ws.write(i,0,sh.cell_value(i,2)) #写入公司名称
        ws.write(i,1,int(sh.cell_value(i,0))) #写入公司ID ID须为数值型
        ws.write(i,2,   #将原表格中几列的产品词合并写入到新文件的一列当中
                 sh.cell_value(i,4)+r', '+sh.cell_value(i,5)+r', '+
                 sh.cell_value(i,6)+r', '+sh.cell_value(i,7)+r', '+
                 sh.cell_value(i,8)+r', '+sh.cell_value(i,9)+r', '+
                 sh.cell_value(i,10))
        for k, v in industries.items():  #写入英文行业
           if sh.cell_value(i,20)==v:
                ws.write(i,3,k+sh.cell_value(i,20))
                break
        ws.write(i,4,'http://'+sh.cell_value(i,1)+'.en.made-in-china.com')#写入网址
        ws.write(i,5,sh.cell_value(i,12)) #写入推广服务结束时间
        ws.write(i,6,sh.cell_value(i,24)) #写入效果分析
        ws.write(i,7,sh.cell_value(i,17)) #写入客服人员
    w.save(os.path.abspath('..')+'\\import_file.xls') #保存文件到桌面
    print 'the work has done!'
else:
    print 'sorry! but the excel file got some errors!'

    


                     


