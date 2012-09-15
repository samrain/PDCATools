#!/usr/bin/env python
#-*- coding:utf-8 -
"""
@version: $Id$
@author: U{sam han<mailto:samrain.han@gmail.com>}
@license:GPL
@contact:samrainhan@gmail.com
@see:
"""
import xlrd
import sqlite3
import datetime
from datetime import date
from datetime import timedelta
import os
"""
    打开一个数据库
"""
conn = sqlite3.connect('/home/rain/下载/plantasktmp.sqlite')
cur = conn.cursor()
cur.execute('CREATE TABLE "checktrac" ("plan" VARCHAR, "missionname" VARCHAR, "enddate" DATETIME, "ET" INTEGER, "actor" VARCHAR, "record" VARCHAR,"result" INTEGER,"quality" INTEGER, "realrate" INTEGER, "realtime" INTEGER, "firstrate" INTEGER, "firsttime" INTEGER, "secondrate" INTEGER, "secondetime" INTEGER, "thirdrate" INTEGER, "thirdtime" INTEGER, "fourthrate" INTEGER, "fourthtime" INTEGER, "fifthrate" INTEGER, "fifthtime" INTEGER, "sixthrate" INTEGER, "sixthtime" INTEGER, "seventhrate" INTEGER, "seventhtime" INTEGER,"flag4delay" INTEGER,"flag4trac" INTEGER,"flag4check" INTEGER,"flag4qualty" INTEGER,"flag4finish" INTEGER)')
"""
    打开一个目录下所有xls文件
"""
dir_name = "/home/rain/下载/trac"
file_list = [f_name for f_name in os.listdir(dir_name) if f_name.endswith('xls')]
for f_in_name in file_list:
    if f_in_name.count('-1'):
        startdate = datetime.date(2012,7,30)
    elif f_in_name.count('-2'):
        startdate = datetime.date(2012,8,6)
    elif f_in_name.count('-3'):
        startdate = datetime.date(2012,8,13)
    elif f_in_name.count('-4'):
        startdate = datetime.date(2012,8,20)
    else:
        startdate = datetime.date(2012,8,27)
    print f_in_name,startdate
    xlsfile = xlrd.open_workbook(os.path.join(dir_name,f_in_name))
    """
        根据sheet位置，打开XLS文件中某个Sheet
    """
    worksheet = xlsfile.sheet_by_index(0)
    """
        取得每个Cell的值
    """
    num_rows = worksheet.nrows - 1
    num_cells = worksheet.ncols - 1
    """
        从第3行开始取数据
    """
    curr_row = 2
    """
        循环取数据
    """
    taskinfo = []
    while curr_row < num_rows:
        curr_row += 1
        row = worksheet.row(curr_row)
        curr_cell = -1
        rowinfo = []
        flag4finish = 0#默认没有完成
        flag4delay = 1#默认没有逾期
        flag4trac = 1#默认尚未开展跟踪
        flag4check = 1#默认没有检查
        flag4qualty = 1 #默认质量分评定正确
        while curr_cell < num_cells:
            curr_cell += 1
#            cell_type = worksheet.cell_type(curr_row, curr_cell)
            cell_value = worksheet.cell_value(curr_row, curr_cell)
            if curr_cell in (10,12,14,16,18,20,22) and cell_value == 100:#表示完成
                print rowinfo[9]
                flag4finish = 1
                if rowinfo[2] < date.isoformat(startdate+timedelta(curr_cell-10)/2):#逾期
#                    print curr_cell,rowinfo[1].encode('utf-8'),rowinfo[2],date.isoformat(startdate+timedelta(curr_cell-10)/2)
                    flag4delay = 0
                if rowinfo[5] == '':#没有记录位置
                    flag4trac = 0
                elif rowinfo[7] == '' and rowinfo[5] not in('N/A','NA'): #如果有质量记录时没有打质量分
                    flag4qualty = 0
                if rowinfo[6] == '': #没有打效果分
                    flag4check = 0
            rowinfo.append(cell_value)
        rowinfo.append(flag4delay)
        rowinfo.append(flag4trac)
        rowinfo.append(flag4check)
        rowinfo.append(flag4qualty)
        rowinfo.append(flag4finish)
        taskinfo.append(rowinfo)
    cur.executemany('insert into checktrac values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',taskinfo)

cur.close()
conn.commit()
conn.close()
