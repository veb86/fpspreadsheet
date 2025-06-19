--------------------------------------------------------------------------------
  Contents of this folder
--------------------------------------------------------------------------------

- fpsspreadsheet-wiki.chm is a snapshot of the wiki articles covering 
  fpspreadsheet.

- fpspreadsheet-api.chm contains the documentation of the most important
  procedures and functions of the library. It is extracted from the embedded
  comments in the source files.
  
  
--------------------------------------------------------------------------------
  How to create fpspreadsheet-wiki
--------------------------------------------------------------------------------
- Navigate to the folder components/wiki of the Lazarus installation.

- Compile the package "lazwiki" in the equally-named subfolder.

- Compile the programs "wikiget" and "wikiconvert".

- Add them to the system's search path (or copy the two executables to the
  folder docs/wiki of the fpspreadsheet installation.

- You also have to copy the OpenSSL DLLs libeay32.dll and ssleay32.dll of the
  correct bitness to the folder docs/wiki of the fpspreadsheed installation.
  
- Run the script "make_chms.bat" (no Linux script at the moment, but it should
  be easy to write one...)

- This script downloads the current fpspreadsheet wiki articles and creates
  a chm help file. 
  
- If you want to create a PDF file from the wiki articles run the script
  "make_pdf.bat". Note that it requires the wkhtmltopdf utility:
  download from https://wkhtmltopdf.org/ and specify the path to its binary
  at the top of "make_pdf.bat".
  
  
--------------------------------------------------------------------------------
  How to create fpspreadsheet-api
--------------------------------------------------------------------------------
- Download the program "PasDoc" from the site https://pasdoc.github.io/
  This is a source code documentation tool which extracts documentation from 
  the comments in the source code.
  
- Unzip the download and copy the file "pasdoc.exe" into the "docs/api" subfolder
  of the FPSpreadsheet installation.

- If not yet done already, compile the program "chmcmd" in folder
  "packages\chm\src\" of the FPC installation and copy the binary into the 
  folder "docs/api" of the FPSpreadsheet installation.
  
- Run the batch file "builddoc.bat" which extracts the documentation code from
  the sources and creates html files. (If you're not on Windows you must write a
  corresponding shell script for this task.)
  
- Run the batch file "make_chm.bat" which compiles a chm file from the html
  files created in the previous step. There may be some error messages which
  can be ignored. (If you are not on Windows you must write a corresponding
  shell script for this task.)
  
- Alternatively to chmcmd you can also use the Microsoft Help Workshop on
  Windows: Download the file "htmlhelp.exe" from
  https://learn.microsoft.com/en-us/previous-versions/windows/desktop/htmlhelp/microsoft-html-help-downloads
  and install it. In the "make_chm.bat" file activate the corresponding
  instruction (using hhc.exe) and comment out the chmcmd instruction.
--------------------------------------------------------------------------------