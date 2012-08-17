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
import os
import xlwt

def writetractable(tasklist,checker):
    """
        制作任务跟踪表
        @param p:
        @type v:
        @return:
        @rtype v:
    """
    book = xlwt.Workbook()
    sheet1 = book.add_sheet(u"计划任务", cell_overwrite_ok=True)
    title = []#标题
    title.append([u'计划名称',u'任务名称',u'截至日期',u'预计工期(分钟)',u'执行人',u'记录位置/撤销或推迟原因',u'效果',u'质量',u'实际工期',u'',u'一',u'',u'二',u'',u'三',u'',u'四',u'',u'五',u'',u'六',u'',u'七',u''])
    title.append([u'',u'',u'',u'',u'',u'',u'',u'',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时'])
    tnum = 0
    for t in title:#填入标题
        trow = 0
        for tt in t:
            sheet1.write(tnum,trow,tt)
            trow += 1
        tnum +=1

    """
        S1：计算当前最新的进度
        S2：计算实际工期合计
    """
    S1 = u'MAX(K?,M?,O?,Q?,S?,U?,W?)=100'
    S2 = u'L?+N?+P?+R?+T?+V?+X?'
    """
        逐格填入数据
    """
    r = 2
    for task in tasklist:
        tr = 0
        for taskinfo in task:
            sheet1.write(r,tr,taskinfo)
            tr+=1
        sheet1.write(r,tr-1,'')
        sheet1.write(r,8,xlwt.Formula(S1.replace('?', str(r+1),7)))#实际进度
        sheet1.write(r,9,xlwt.Formula(S2.replace('?', str(r+1),7)))#实际工时
        r+=1
    """
        保存XLS文件
    """

    filename = u'【TG-IT(1208-3)】任务执行跟踪表('+checker + u').xls'
    book.save(os.path.join(u"/home/rain/下载/",filename))

def readxlsfile(infolist,flag):
    """
        @param p:
        @type v:
        @return:
        @rtype v:
    """
    dir_name = "/home/rain/下载/"+flag
    file_list = [f_name for f_name in os.listdir(dir_name) if f_name.endswith('xls')]
    for f_in_name in file_list:
        actor = (f_in_name.rsplit('(')[2]).split(')')[0]
        print flag,actor
        xlsfile = xlrd.open_workbook(os.path.join(dir_name,f_in_name))
        """
            根据sheet名称，打开XLS文件中某个Sheet
        """
        worksheet = xlsfile.sheet_by_name(u'计划任务')
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
            row = worksheet.row(curr_row)
            curr_cell = -1
            while curr_cell < num_cells:
                curr_cell += 1
                rowinfo.append(worksheet.cell_value(curr_row, curr_cell))
            rowinfo.append(actor.decode('utf-8'))
            infolist.append(rowinfo)

weekreportlist = []
traclist = []
readxlsfile(weekreportlist,'weekreport')
readxlsfile(traclist,'trac')

"""
    打开一个数据库
"""
conn = sqlite3.connect(":memory:")#数据库放在内存中
#conn = sqlite3.connect("/home/rain/下载/autotrac.sqlite")#数据库存放在文件中
cur = conn.cursor()
"""
    用周报数据来更新跟踪表
"""
for trac in traclist:
    for weekreport in weekreportlist:
        if weekreport[1] == trac[1] and weekreport[22] == trac[4]:
            iiii = -1
            trac[5] = weekreport[5]
#            print weekreport[22].encode('utf-8')
            while iiii<15:
                iiii+=1
                trac[iiii+8] = weekreport[iiii+6]
            break
"""
    插入任务跟踪表数据
"""

#cur.execute('drop table "tracreport"')
cur.execute('CREATE TABLE "tracreport" ("plan" VARCHAR, "missionname" VARCHAR, "enddate" DATETIME, "ET" INTEGER, "actor" VARCHAR, "record" VARCHAR,"result" INTEGER,"quality" INTEGER, "realrate" INTEGER, "realtime" INTEGER, "firstrate" INTEGER, "firsttime" INTEGER, "secondrate" INTEGER, "secondetime" INTEGER, "thirdrate" INTEGER, "thirdtime" INTEGER, "fourthrate" INTEGER, "fourthtime" INTEGER, "fifthrate" INTEGER, "fifthtime" INTEGER, "sixthrate" INTEGER, "sixthtime" INTEGER, "seventhrate" INTEGER, "seventhtime" INTEGER,"checker" VARCHAR)')
cur.executemany('insert into tracreport values(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',traclist)

"""
    将周报中的任务完成情况导入到跟踪表中
"""
cur.execute('select checker from tracreport group by checker')
rows = cur.fetchall()
for row in rows:
    cur.execute('select * from tracreport where checker = ?',row)
    writetractable(cur.fetchall(),row[0])

cur.close()
#conn.commit()
conn.close()
