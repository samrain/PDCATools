#!/usr/bin/env python
#-*- coding:utf-8 -*-
"""从Ganttproject文件转换为任务分配表
@version: $Id$
@author: U{sam han<mailto:samrain.han@gmail.com>}
@license:GPL
@contact:samrainhan@gmail.com
@see:
"""


import xml.etree.cElementTree as ET
import sqlite3
import os
from xlwt import *
import xlwt
import sys
import tool4gan

"""
    打开一个目录下所有ganttproject文件
"""
dir_name = "/home/rain/下载/gan"
outputdir= "/home/rain/下载/plan"
file_list = [f_name for f_name in os.listdir(dir_name) if f_name.endswith('gan')]

"""
    打开一个数据库
"""
conn = sqlite3.connect(":memory:")
cur = conn.cursor()
"""
    生成4张表
    计划任务表  taskfromgantproj
    资源表  resources
    资源分配表  allocations
    当前任务表   tasknow
"""
cur.execute('CREATE TABLE "taskfromgantproj" ("id" INTEGER, "name" VARCHAR, "startdate" DATETIME, "duration" INTEGER,"plan" VARCHAR,"strdur" VARCHAR,"pmid" INTEGER)')
cur.execute('CREATE TABLE "resources" ("id" INTEGER, "name" VARCHAR)')
cur.execute('CREATE TABLE "allocations" ("taskid" INTEGER, "resourceid" INTEGER)')
cur.execute('CREATE TABLE "tasknow" ("plan" VARCHAR, "taskname" VARCHAR, "name" VARCHAR, "enddate" DATETIME, "checker" VARCHAR)')

for f_in_name in file_list:
    tree = ET.ElementTree(file=os.path.join(dir_name,f_in_name))
#    print tree.getroot().tag,tree.getroot().attrib
    projectname = tree.getroot().attrib['name']
    tasks = tree.getroot()[4]
    resources = tree.getroot()[5]
    allocations = tree.getroot()[6]
    tasklist = []
    listresources = []
    listallocations = []
    for subelement in resources.getchildren():
        """
            得到资源信息
        """
    #    print subelement.attrib['id'],subelement.attrib['name'].encode('utf-8')
        if subelement.attrib['function'] == 'Default:1': #得到项目经理的资源id
            pmid = subelement.attrib['id']        
        listresources.append([subelement.attrib['id'],subelement.attrib['name']])

    for subelement in allocations.getchildren():
        """
            得到任务和资源的关联关系
        """
        listallocations.append([subelement.attrib['task-id'],subelement.attrib['resource-id']])
    """
        得到任务信息
    """
#    printelement(tasks,'',projectname,pmid,tasklist)
    tool4gan.getelementlist(tasks,'',projectname,pmid,tasklist)
    
    """
        插入数据库,将以上得到的信息插入到各自表中
    """
    cur.execute('delete from taskfromgantproj')
    cur.execute('delete from resources')
    cur.execute('delete from allocations')
    
    cur.executemany('insert into taskfromgantproj values(?,?,?,?,?,?,?)',tasklist)
    cur.executemany('insert into resources values(?,?)',listresources)
    cur.executemany('insert into allocations values(?,?)',listallocations)

    """
        导出计划任务分配表
    """
    cur.execute('SELECT a.plan,a.name taskname,c.name,date(a.startdate,a.strdur) enddate,(select resources.name from resources where resources.id = a.pmid) checker FROM taskfromgantproj a,allocations b,resources c where a.id = b.taskid and b.resourceid = c.id and a.startdate >= ? and date(a.startdate,a.strdur)<= ? order by date(a.startdate,a.strdur),a.name',['2012-08-20','2012-08-24'])
    rows = cur.fetchall()
    r = 1
    book = xlwt.Workbook()
    sheet1 = book.add_sheet("1")
    for row in rows:
        plan = row[0]
        sheet1.write(r,0,plan)
        sheet1.write(r,1,row[1])
        sheet1.write(r,2,row[2])
        sheet1.write(r,3,row[3])
        sheet1.write(r,4,row[4])
        r+=1
    sheet1.write(0,0,u'计划名称')
    sheet1.write(0,1,u'任务名称')
    sheet1.write(0,2,u'执行人')
    sheet1.write(0,3,u'截止时间')
    sheet1.write(0,4,u'检查人')
    sheet1.write(0,5,u'预计工期')

    book.save(os.path.join(outputdir,projectname.encode('utf-8'))+'plan.xls')
cur.close()
#conn.commit()
conn.close()
