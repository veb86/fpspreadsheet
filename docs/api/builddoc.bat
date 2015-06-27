@echo off

set DOX_CMD="D:\Programme\Doc-O-Matic 7 Express\domexpress.exe"

if not exist output mkdir output

rem Prepare files...
path=%PROGRAMFILES%;%PATH%
pushd .
cd ..\..
ren fps.inc ---fps.inc
ren fpspreadsheetctrls.lrs ---fpspreadsheetctrls.lrs
popd

rem Extract help topics and create chm files...
%DOX_CMD% -config "HTML Help" fpspreadsheet.dox-express > doc-o-matic.txt

rem Clean up
pushd .
cd ..\..
chdir
ren ---fps.inc fps.inc
ren ---fpspreadsheetctrls.lrs fpspreadsheetctrls.lrs
popd
if exist output\fpspreadsheet.chm copy output\fpspreadsheet.chm ..\fpspreadsheet-api.chm /y
