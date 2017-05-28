This folder contains the palette icons of the visual spreadsheet components.

fpspreadsheetctrls.spp is the source file for the palette icons. 
It contains the various images as layers. Use the software "PhotoPlus" 
(www.serif.com) to open and edit; the free starter edition is sufficient.

The basic icons are taken from the Lazarus images folder. The Excel overlay is
self-drawn according to an old Excel version.

In addition, there's also the cursor for drag-and-drop copy mode.

The Lazarus resource file is created by executing the make_lrs.bat batch file
(Linux script to be created accordingly). The script requires lazres.exe to
be in the same directory - compile (lazarus)/tools/lazres to create this
binary.