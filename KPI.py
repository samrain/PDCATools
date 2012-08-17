#!/usr/bin/env python
#-*- coding:utf-8 -*-

import sqlite3


def calctask(self):
    pass

def calcperson(self):
    pass

def calcteam(self):
    pass

def calcbehavior():
    conn = sqlite3.connect('/home/rain/下载/behavior.sqlite')
    cur = conn.cursor()
    behaviorpoint = {}
    behaviortype={u'提示':1,
                  u'反馈':2,
                  u'分享':3,
                  u'公告':0}
    bhlist = cur.execute('select user,label,count(1) from wiki group by user,label')
    for bh in bhlist:
#        print user.encode('utf-8'),label.encode('utf-8'),bhcount
        if behaviorpoint.has_key(bh[0]):
            behaviorpoint[bh[0]] = behaviorpoint[bh[0]] + behaviortype[bh[1]] * bh[2]
        else:
            behaviorpoint[bh[0]] = behaviortype[bh[1]] * bh[2]
    for name, point in behaviorpoint.items():
        print 'user %s has %s' % (name.encode('utf-8'), point)
    pass

calcbehavior()
