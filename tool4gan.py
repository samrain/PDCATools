#!/usr/bin/env python
#-*- coding:utf-8 -*-
"""Ganttproject文件工具
@version: @Id$
@author: U{sam han<mailto:samrain.han@gmail.com>}
@license:GPL
@contact:samrainhan@gmail.com
@see:
"""

import xml.etree.cElementTree as ET


class config():
    """
        @param p:
        @type v:
        @return:
        @rtype v:
    """
    tag4task = 'task'  # 取tag为task的数据
    attribtuple = ('name', 'id', 'start', 'duration', 'function', 'task-id', 'resource-id')  # 分别是任务名称、任务编号、起始日期和预计工期、角色
    symbel = ('+', ' day', '.', '', 'Default:1')  # 前2个都是用来计算截止日期的预计工期字符串；.是用来连接父亲和儿子的任务名称的；空格是用来连接一般字符串的;项目经理的角色ID
    pass


def getelementlist(element, parentname, projectname, pmid, tasklist):
    """
        递归方法得到ganttproject文件中Task任务树，并将儿子名称改为父亲的名称+儿子名称
        @param element:gant文件中信息制作成Tree后的对象
               parentname:父亲节点的名称
               projectname:项目名称
               pmid:项目经理ID
               tasklist:任务信息列表
        @type element:element
              parentname:str
              projectname:str
              pmid:int
              tasklist:list
        @return none
        @rtype none
    """
    for subelement in element.getchildren():
        if subelement.tag == config.tag4task:
            name = config.symbel[3].join([parentname, subelement.attrib[config.attribtuple[0]]])
            subtasklist=(subelement.attrib[config.attribtuple[1]], name, subelement.attrib[config.attribtuple[2]], subelement.attrib[config.attribtuple[3]], projectname, config.symbel[3].join([config.symbel[0], str(int(subelement.attrib[config.attribtuple[3]])-1), config.symbel[1]]), pmid)
            tasklist.append(subtasklist)
            getelementlist(subelement, config.symbel[3].join([name, config.symbel[2]]), projectname, pmid, tasklist)
    pass


def getresource(element, pmid):
    """执行人信息
        @param element:resource的element对象,pmid:项目经理的执行人id
        @type element:element,pmid:str
        @return listresources:执行人信息列表[id,name]
        @rtype listresources:list
    """
    listresources = []
    for subelement in element.getchildren():
        if subelement.attrib[config.attribtuple[4]] == config.symbel[4]: #得到项目经理的资源id
            pmid = subelement.attrib[config.attribtuple[1]]
        listresources.append([subelement.attrib[config.attribtuple[1]], subelement.attrib[config.attribtuple[0]]])
    return listresources
    pass


def getallocation(element):
    """执行人和任务关联信息
        @param element:执行人和任务关联信息的element对象
        @type element:element
        @return listallocation:执行人和任务关联信息列表[task-id, resource-id]
        @rtype listallocation:list
    """
    listallocations = []
    for subelement in element.getchildren():
        listallocations.append([subelement.attrib[config.attribtuple[5]],subelement.attrib[config.attribtuple[6]]])
    return listallocations
    pass


def readgan(filename):
    """
        @param filename:gan含全路径的文件名
        @type filename:str
        @return:
        @rtype v:
    """
    tree = ET.ElementTree(file=filename)
    pmid = ''
    gandict = {}
    gandict['resource'] = getresource(tree.getroot()[5], pmid)
    gandict['allocation'] = getallocation(tree.getroot()[6])
    # 得到任务信息
    tasklist = []
    getelementlist(tree.getroot()[4],config.symbel[3],tree.getroot().attrib[config.attribtuple[0]],pmid,tasklist)
    gandict['task'] = tasklist
    return gandict
    pass


if __name__=="__main__":
    print readgan('/home/rain/下载/gan/【TG-IT(12GWZN)】计划任务分配表(岗位职能).gan')
