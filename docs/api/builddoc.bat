set pasdoc_cmd=pasdoc.exe
set hhc_cmd="C:\Program Files (x86)\HTML Help Workshop\hhc.exe"

%pasdoc_cmd% @options.txt --format=htmlhelp --output=output --name=fpspreadsheet --source=source-files.txt
::%hhc_cmd% output\fpspreadsheet.hhc