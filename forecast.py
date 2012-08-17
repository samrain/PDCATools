#!/usr/bin/env python
#-*- coding:utf-8 -*-
"""用任务分配表预测下周工作
@version: $Id$
@author: U{sam han<mailto:samrain.han@gmail.com>}
@license:GPL
@contact:samrainhan@gmail.com
@see:
"""

    cur.execute('INSERT INTO tasknow SELECT a.plan,a.name taskname,c.name,date(a.startdate,a.strdur) enddate,(select resources.name from resources where resources.id = a.pmid) checker FROM taskfromgantproj a,allocations b,resources c where a.id = b.taskid and b.resourceid = c.id and a.startdate >= ? and date(a.startdate,a.strdur)<= ?',['2012-06-18','2012-06-22'])

"""
    统计下周任务分配情况
"""
book1 = xlwt.Workbook()
sheet1 = book1.add_sheet(u"下周任务总览")
"""
    How many group by计划
"""
cur.execute('select plan,count(1) tasknum from tasknow group by plan')
r=0
sheet1.write(0,0,u'计划')
sheet1.write(0,1,u'任务数')
for row in cur.fetchall():
    r+=1
    sheet1.write(r,0,row[0])
    sheet1.write(r,1,row[1])
"""
    How many group by执行人
"""
cur.execute('select name,count(1) tasknum from tasknow group by name')
r+=2
sheet1.write(r,0,u'执行人')
sheet1.write(r,1,u'任务数')
for row in cur.fetchall():
    r+=1
    sheet1.write(r,0,row[0])
    sheet1.write(r,1,row[1])
"""
    How many group by截止日期
"""
cur.execute('select enddate,count(1) tasknum from tasknow group by enddate')
r+=2
sheet1.write(r,0,u'截止日期')
sheet1.write(r,1,u'任务数')
for row in cur.fetchall():
    r+=1
    sheet1.write(r,0,row[0])
    sheet1.write(r,1,row[1])

book1.save(os.path.join(outputdir,'下周任务总览.xls'))
