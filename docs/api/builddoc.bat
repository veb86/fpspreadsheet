@echo off

set DOX_CMD="C:\Program Files (x86)\Doc-O-Matic 7 Express\domexpress.exe"
if not exist %DOX_CMD% goto :dox_error

rem *** Prepare files ***
if not exist output mkdir output
pushd .
cd ..\..

if not exist fps.inc goto :next1
ren fps.inc ---fps.inc

:next1
if not exist fpspreadsheetctrls.lrs goto :next2
ren fpspreadsheetctrls.lrs ---fpspreadsheetctrls.lrs

:next2
popd

rem Extract help topics and create chm files...
echo Running %DOX_CMD% -config "HTML Help" fpspreadsheet.dox-express 
%DOX_CMD% -config "HTML Help" fpspreadsheet.dox-express > doc-o-matic.txt

rem *** Clean up ***
pushd .

cd ..\..
chdir
if not exist ---fps.inc goto :next3
ren ---fps.inc fps.inc

:next3
if not exist ---fpspreadsheetctrls.lrs goto :next4
ren ---fpspreadsheetctrls.lrs fpspreadsheetctrls.lrs

:next4
popd

if exist output\fpspreadsheet.chm copy output\fpspreadsheet.chm ..\fpspreadsheet-api.chm /y

goto :end

:dox_error
echo Doc-O-Matic program not found. Check the script.

:end
