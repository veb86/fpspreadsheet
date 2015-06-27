echo off

rem set FMT=html
set FMT=chm
echo Downloading wiki...
wikiget --page=FPSpreadsheet --page=FPSpreadsheet:_List_of_formulas --page=RPN_Formulas_in_FPSpreadsheet
wikiget --page=FPSpreadsheet_tutorial:_Writing_a_mini_spreadsheet_application
wikiget --page=TsWorksheetGrid --page=TsWorksheetChartSource

echo.
echo Converting wiki to chm...
wikiconvert --format=%FMT% --css=css/wiki.css --root="FPSpreadsheet wiki pages" --title="FPSpreadsheet wiki pages (offline version, created %DATE%)" --chm="..\fpspreadsheet-wiki.chm" wikixml/FPSpreadsheet.s00.xml wikixml/FPSpreadsheet=3A_List_of_formulas*.xml wikixml/RPN_Formulas*.xml wikixml/FPSpreadsheet_tutorial=3A_Writing*.xml wikixml/TsWorksheet*.xml
set FMT=