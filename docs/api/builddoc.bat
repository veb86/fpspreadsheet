set pasdoc_cmd=pasdoc.exe
if not exist output md output
%pasdoc_cmd% @options.txt --format=htmlhelp --output=output --name=fpspreadsheet --source=source-files.txt
