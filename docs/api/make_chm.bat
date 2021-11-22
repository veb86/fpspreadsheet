set chm_cmd="C:\Program Files (x86)\HTML Help Workshop\hhc.exe"
::set chm_cmd=chmcmd.exe
%chm_cmd% output\fpspreadsheet.hhp
copy /Y output\fpspreadsheet.chm ..\fpspreadsheet-api.chm

