echo off

:: Download wkhtmltopdf from https://wkhtmltopdf.org/ and specify here the path to its binary
set html2pdf=d:\programme\wkhtmltox\bin\wkhtmltopdf.exe

echo Downloading wiki...
wikiget --page=FPSpreadsheet --page=FPSpreadsheet:_Examples --page=FPSpreadsheet:_List_of_formulas --page=RPN_Formulas_in_FPSpreadsheet 
wikiget --page=FPSpreadsheet:_Chart_Tutorial
wikiget --page=FPSpreadsheet_tutorial:_Writing_a_mini_spreadsheet_application
wikiget --page=TsWorksheetGrid --page=TsWorksheetChartSource

echo.
echo Converting wiki to html...
wikiconvert --format=html --css=css/wiki.css --root="FPSpreadsheet wiki pages" --title="FPSpreadsheet wiki pages (offline version, created %DATE%)" --outputdir=wikihtml wikixml/*.xml

echo Converting html to pdf...
cd wikihtml
set margins=-L 25 -T 25 -R 25 -B 25
set f1=FPSpreadsheet.s00.html
set f2=FPSpreadsheet=3A_List_of_formulas.s03000.html
set f3=FPSpreadsheet=3A_Examples.s0300.html
set f4=FPSpreadsheet_tutorial=3A_Writing_a_mini_spreadsheet_application.s000c0000000.html
set f5=FPSpreadsheet=3A_Chart_Tutorial.s03100.html
set f6=TsWorksheetGrid.k08.html
set f7=TsWorksheetChartSource.k0880.html
set f8=RPN_Formulas_in_FPSpreadsheet.u03g00.html
%html2pdf% %margins% --enable-local-file-access %f1% %f2% %f3% %f4% %f5% %f6% %f7% %f8% ../fpspreadsheet-wiki.pdf
cd ..

move fpspreadsheet-wiki.pdf ..
