#!/bin/sh
rm -f ~/下载/gan/*.gan
rm -f ~/下载/plan/*.xls

#!/bin/sh
ftp -v -n <<!
open msdocs.miao-shi.net
user hany *********
binary
cd /Alfresco/Sites/TSIT/documentLibrary/A工作计划/00进行中
lcd ~/下载/gan/
prompt
mget *.gan
close
bye
!


python ~/repositories/PDCATools/gant2plan.py

nautilus ~/下载/plan

