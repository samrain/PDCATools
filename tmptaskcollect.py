#!/usr/bin/env python
#-*- coding:utf-8 -*-
"""
@version: $Id$
@author: U{sam han<mailto:samrain.han@gmail.com>}
@license:
@contact:
@see:
"""

import os
import xlrd
import sqlite3

def readxlsfile():
    """
        @param p:
        @type v:
        @return:
        @rtype v:
    """
    dir_name = "/home/rain/下载/weekreport"
    tmptask = []
    file_list = [f_name for f_name in os.listdir(dir_name) if f_name.endswith('xls')]
    for f_in_name in file_list:
        print f_in_name
        actor = (f_in_name.rsplit('(')[2]).split(')')[0]
        xlsfile = xlrd.open_workbook(os.path.join(dir_name,f_in_name))
        """
            根据sheet名称，打开XLS文件中某个Sheet
        """
        worksheet = xlsfile.sheet_by_name(u'临时任务')
        
        """
            取得每个Cell的值
        """
        num_rows = worksheet.nrows-1
        num_cells = worksheet.ncols-1
        """
            从第3行开始取数据
        """
        curr_row = 1
        """
            循环取数据
        """
        while curr_row < num_rows:
            curr_row += 1
            rowinfo = []
            rowinfo.append(actor.decode('utf-8'))
            row = worksheet.row(curr_row)
            curr_cell = -1
            if worksheet.cell_value(curr_row, curr_cell+1):
                while curr_cell < num_cells:
                    curr_cell += 1
                    rowinfo.append(worksheet.cell_value(curr_row, curr_cell))
                tmptask.append(rowinfo)
    return tmptask

conn = sqlite3.connect('/home/rain/下载/tmptask1207.sqlite')
cur = conn.cursor()
cur.execute('CREATE TABLE "tmptask" ("actor" VARCHAR, "missionname" VARCHAR, “assigner” VARCHAR, "beneficiary" VARCHAR, "enddate" DATETIME, "ET" INTEGER, "record" VARCHAR, "realrate" INTEGER, "realtime" INTEGER, "firstrate" INTEGER, "firsttime" INTEGER, "secondrate" INTEGER, "secondetime" INTEGER, "thirdrate" INTEGER, "thirdtime" INTEGER, "fourthrate" INTEGER, "fourthtime" INTEGER, "fifthrate" INTEGER, "fifthtime" INTEGER, "sixthrate" INTEGER, "sixthtime" INTEGER, "seventhrate" INTEGER, "seventhtime" INTEGER)')
cur.executemany('insert into tmptask values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',readxlsfile())
conn.commit()
cur.close()
conn.close()

#weekreportdir = '/home/rain/下载'
#sheet2list = readxlsfile()
#book = xlwt.Workbook()
#"""
#    写sheet2=临时任务
#"""
#sheet2 = book.add_sheet(u"临时任务", cell_overwrite_ok=True)
#title2 = []#标题
#title2.append([u'执行人',u'任务名称',u'分配人',u'被执行人/受益人',u'截至日期',u'预计工期(分钟)',u'记录位置/撤销或推迟原因',u'实际工期',u'',u'一',u'',u'二',u'',u'三',u'',u'四',u'',u'五',u'',u'六',u'',u'七',u''])
#title2.append([u'',u'',u'',u'',u'',u'',u'',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时'])
#tnum = 0
#for t in title2:#填入标题
#    trow = 0
#    for tt in t:
#        sheet2.write(tnum,trow,tt)
#        trow += 1
#    tnum +=1
#r = 1
#for task in sheet2list:
#    tr = -1
#    r += 1
#    for taskinfo in task:
#        tr+=1
#        sheet2.write(r,tr,taskinfo)

#"""
#    保存XLS文件
#"""
#filename = '临时任务.xls'
#book.save(os.path.join(weekreportdir,filename))
