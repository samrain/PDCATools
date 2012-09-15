#!/usr/bin/env python
#-*- coding:utf-8 -
"""从任务分配表生成周报和跟踪表
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

def writeweeklyreport(tasklist):
    """
        制作个人周报
        @param p:
        @type v:
        @return:
        @rtype v:
    """
    book = xlwt.Workbook()
    """
        写sheet1=计划任务
    """
    sheet1 = book.add_sheet(u"计划任务", cell_overwrite_ok=True)
    title1 = []#标题
    title1.append([u'计划名称',u'任务名称',u'截至日期',u'预计工期(分钟)',u'检查人',u'记录位置/撤销或推迟原因',u'实际工期',u'',u'一',u'',u'二',u'',u'三',u'',u'四',u'',u'五',u'',u'六',u'',u'七',u''])
    title1.append([u'',u'',u'',u'',u'',u'',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时'])
    tnum = 0
    for t in title1:#填入标题
        trow = 0
        for tt in t:
            sheet1.write(tnum,trow,tt)
            trow += 1
        tnum +=1
    """
        调整标题格式
    """
#    sheet1.write_merge(r1=0, c1=0, r2=1, c2=0, label=u'计划名称') 
    r = 2
    """
        S1：计算当前最新的进度
        S2：计算实际工期合计
    """
    S1 = u'MAX(I?,K?,M?,O?,Q?,S?,U?)=100'
    S2 = u'J?+L?+N?+P?+R?+T?+V?'
    """
        逐格填入数据
    """
    for task in tasklist:
        sheet1.write(r,0,task[0])#计划名称
        sheet1.write(r,1,task[1])#任务名称
        sheet1.write(r,2,task[3])#截至日期
        sheet1.write(r,3,task[5])#预计工期
        sheet1.write(r,4,task[4])#执行人
        sheet1.write(r,6,xlwt.Formula(S1.replace('?', str(r+1),7)))#实际进度
        sheet1.write(r,7,xlwt.Formula(S2.replace('?', str(r+1),7)))#实际工时
        r+=1
    
    """
        写sheet2=临时任务
    """
    sheet2 = book.add_sheet(u"临时任务", cell_overwrite_ok=True)
    title2 = []#标题
    title2.append([u'任务名称',u'分配人',u'被执行人/受益人',u'截至日期',u'预计工期(分钟)',u'记录位置/撤销或推迟原因',u'实际工期',u'',u'一',u'',u'二',u'',u'三',u'',u'四',u'',u'五',u'',u'六',u'',u'七',u''])
    title2.append([u'',u'',u'',u'',u'',u'',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时',u'进度',u'工时'])
    tnum = 0
    for t in title2:#填入标题
        trow = 0
        for tt in t:
            sheet2.write(tnum,trow,tt)
            trow += 1
        tnum +=1
    """
        S1：计算当前最新的进度
        S2：计算实际工期合计
    """
    tempi=2
    while tempi < 25:#填入实际工期计算公式
        sheet2.write(tempi,6,xlwt.Formula(S1.replace('?', str(tempi+1),7)))#实际进度
        sheet2.write(tempi,7,xlwt.Formula(S2.replace('?', str(tempi+1),7)))#实际工时
        tempi+=1

    
    """
        保存XLS文件
    """
    filename = '【TG-IT(1208-5)】个人周工作报告(' + task[2].encode('utf-8') + ').xls'
    book.save(os.path.join(weekreportdir,filename))

def writetractable(tasklist):
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
        调整标题格式
    """
#    sheet1.write_merge(r1=0, c1=0, r2=1, c2=0, label=u'计划名称') 
    r = 2
    """
        S1：计算当前最新的进度
        S2：计算实际工期合计
    """
    S1 = u'MAX(K?,M?,O?,Q?,S?,U?,W?)=100'
    S2 = u'L?+N?+P?+R?+T?+V?+X?'
    """
        逐格填入数据
    """
    for task in tasklist:
        sheet1.write(r,0,task[0])#计划名称
        sheet1.write(r,1,task[1])#任务名称
        sheet1.write(r,2,task[3])#截至日期
        sheet1.write(r,3,task[5])#预计工期
        sheet1.write(r,4,task[2])#执行人
        sheet1.write(r,8,xlwt.Formula(S1.replace('?', str(r+1),7)))#实际进度
        sheet1.write(r,9,xlwt.Formula(S2.replace('?', str(r+1),7)))#实际工时
        r+=1
    """
        保存XLS文件
    """
    filename = '【TG-IT(1208-5)】任务执行跟踪表(' + task[4].encode('utf-8') + ').xls'
    book.save(os.path.join(tracdir,filename))


"""
    打开一个目录下所有xls文件
"""
dir_name = "/home/rain/下载/plan"
tracdir = '/home/rain/下载/trac'
weekreportdir = '/home/rain/下载/weekreport'
file_list = [f_name for f_name in os.listdir(dir_name) if f_name.endswith('plan.xls')]

tasklist = []

for f_in_name in file_list:
    xlsfile = xlrd.open_workbook(os.path.join(dir_name,f_in_name))
    """
        根据sheet名称，打开XLS文件中某个Sheet
    """
    worksheet = xlsfile.sheet_by_name('1')
    """
        取得每个Cell的值
    """
    num_rows = worksheet.nrows-1
    num_cells = worksheet.ncols-1
    """
        从第2行开始取数据
    """
    curr_row = 0
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
        tasklist.append(rowinfo)
"""
    打开一个数据库
"""
#print tasklist
conn = sqlite3.connect(":memory:")#数据库放在内存中
#conn = sqlite3.connect("temp.sqlite")
cur = conn.cursor()

"""
    插入计划任务数据
"""
cur.execute('CREATE TABLE "plan" ("plan" VARCHAR, "taskname" VARCHAR, "name" VARCHAR, "enddate" DATETIME, "checker" VARCHAR, "ET" INTEGER)')
cur.executemany('insert into plan values(?,?,?,?,?,?)',tasklist)

"""
    制作周报
"""
cur.execute('select name from plan group by name')
rows = cur.fetchall()
for row in rows:#遍历执行人集合，以执行人为条件查询他的计划任务信息
#    print row[0].encode('utf-8')
    cur.execute('select * from plan where name = ?',[row[0]])
    writeweeklyreport(cur.fetchall())

"""
    制作跟踪表
"""
cur.execute('select checker from plan group by checker')
rows = cur.fetchall()
for row in rows:
    print row[0].encode('utf-8')
    cur.execute('select * from plan where checker = ?',[row[0]])
    writetractable(cur.fetchall())


cur.close()
conn.commit()
conn.close()


