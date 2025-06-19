echo off

rem set FMT=html
set FMT=chm
echo Downloading wiki...

wikiget --page=FPSpreadsheet --page=FPSpreadsheet:_Examples --page=FPSpreadsheet:_List_of_formulas --page=RPN_Formulas_in_FPSpreadsheet 
wikiget --page=FPSpreadsheet:_Chart_Tutorial
wikiget --page=FPSpreadsheet_tutorial:_Writing_a_mini_spreadsheet_application
wikiget --page=TsWorksheetGrid --page=TsWorksheetChartSource

echo.
echo Converting wiki to chm...

wikiconvert --format=chm --css=css/wiki.css --root="FPSpreadsheet wiki pages" --title="FPSpreadsheet wiki pages (offline version, created %DATE%)" --chm="..\fpspreadsheet-wiki.chm" wikixml/*.xml

wikiconvert --format=html --css=css/wiki.css --root="FPSpreadsheet wiki pages" --title="FPSpreadsheet wiki pages (offline version, created %DATE%)" --outputdir=wikihtml wikixml/*.xml


set FMT=