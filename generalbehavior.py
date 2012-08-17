#!/usr/bin/env python
#-*- coding:utf-8 -*-
"""收集日常行为的数据
@version: $Id$
@author: U{sam han<mailto:samrain.han@gmail.com>}
@license:
@contact:
@see:
"""

from bs4 import BeautifulSoup
import urllib
import sqlite3


u = urllib.urlopen('http://kb.miao-shi.net/display/ITKB/2012/07')
soup = BeautifulSoup(u.read())
keylist = (("a","confluence-userlink"),("a","blogHeading"),("a","label"))
wikilist = []
for key in keylist:
    wikicont = []
    for link in soup.find_all(key[0],key[1]):
        wikicont.append(link.get_text())
#        print link.get_text().encode('utf-8')
#    print '--------'
    wikilist.append(wikicont)
if len(wikilist[0]) == len(wikilist[1]) and len(wikilist[0]) == len(wikilist[2]):
    print len(wikilist[0])
    a = zip(wikilist[0],wikilist[1],wikilist[2])
#    for b in a:
#        print b[0].encode('utf-8'),b[1].encode('utf-8'),b[2].encode('utf-8')
    conn = sqlite3.connect("/home/rain/下载/behavior1207.sqlite")
    cur = conn.cursor()
    cur.execute("create table 'wiki' ('user' VARCHAR,'title' VARCHAR, 'label' VARCHAR)")
    cur.executemany("insert into 'wiki' values (?,?,?)",a)
    conn.commit()
    cur.close()
    conn.close()
else:
    print len(wikilist[0]),len(wikilist[1]),len(wikilist[2])
