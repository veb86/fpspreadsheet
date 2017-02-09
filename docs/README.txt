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

- Run the script "make_docs.bat" (no Linux script at the moment)

- This script downloads the current fpspreadsheet wiki articles and creates
  a chm help file. 
  
  
--------------------------------------------------------------------------------
  How to create fpspreadsheet-api (Windows-only)
--------------------------------------------------------------------------------
- Download the program "Doc-o-matic Express" from the site
  http://www.doc-o-matic.com/download.shtml

- This is a source code documentation and help authoring tool which constructs
  a chm help file from comments embedded into the source code.

- Install the program

- Change the variable DOX_CMD in the batch file "builddoc.bat" to point to the 
  correct folder.

- Run the batch file "builddoc.bat" which extracts the documentation code from
  the sources and creates the chm file. It may take some time...
  
  
--------------------------------------------------------------------------------