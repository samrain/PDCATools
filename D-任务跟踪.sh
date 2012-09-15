#!/bin/sh
rm -f ~/下载/weekreport/*.xls
rm -f ~/下载/trac/*.xls

ftp -v -n <<!
open msdocs.miao-shi.net
user hany ********
binary
cd /Alfresco/Sites/TSIT/documentLibrary/E个人周报/1208/1208-5
lcd ~/下载/weekreport/
prompt
mget *

cd /Alfresco/Sites/TSIT/documentLibrary/C进度跟踪/1208/1208-5
lcd ~/下载/trac/
mget *

close
bye
!

python ~/repositories/PDCATools/collect.py
nautilus ~/下载

