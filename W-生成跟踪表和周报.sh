#!/bin/sh
rm -f ~/下载/plan/*plan.xls
rm -f ~/下载/trac/*.xls
rm -f ~/下载/weekreport/*.xls


#!/bin/sh
ftp -v -n <<!
open msdocs.miao-shi.net
user hany ********
binary
cd /Alfresco/Sites/TSIT/documentLibrary/A工作计划/01任务分配表
lcd ~/下载/plan/
prompt
mget *
close
bye
!

python ~/repositories/PDCATools/plan2task.py
nautilus ~/下载/

