This folder contains the palette icons of the visual spreadsheet components.
The icons with appended _150 and _200 are magnifited with respect to the
icons without appended number by factors 150% and 200%, respectively; they are 
used for screens at higher resolutions.

The icons are created from the gimp source files fpspreadsheetctrls.xcf,
fpspreadsheetctrls_150.xcf, and fpspreadsheet_200.xcf; these files contain
all icons in individual layers.

The basic icons are taken from the Lazarus images folder. The "spreadsheet"
overlay in the lower-right corner is the file "table.png" of the FatCow icon set
(http://www.fatcow.com/free-icons, license Creative Commons Attribution 3.0).

The Lazarus resource file is created by executing the make_lrs.bat batch file
(Linux script to be created accordingly). The script requires the program lazres
to be in the same directory - compile (lazarus)/tools/lazres to create this
binary (or change the script with the path to lazres on your system).