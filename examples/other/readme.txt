This folder contains various demo applications:

- cell_formats/formatting/demo_write_formatting: shows some simple cell formatting

- colors/demo_write_colors: shows the colors of the Excel8 color palette
  
- conditional_formatting/demo_conditional_formatting: demonstrates how
  a cell range can be formatted depending on the cell content ("conditional
  formatting").

- expression_parser_demo_expression_parser: shows how the formula engine of
  FPSpreadsheet is used.
  
- frozen_rowscols/demo_frozen_cols_rows: shows how the first rows and columns
  can be "frozen", i.e. prevented from scrolling.
  
- header_footer_images/demo_write_headerfooter_images: adds images to 
  worksheet headers and footers
  
- images/demo_write_images: shows how to create workbooks/worksheets with
  embedded images.

- pagelayout/demo_pagelayout: show a few cases how the page layout can be
  configured for printing and print preview. After loading the created file into
  Excel or Calc switch to the print preview to see the effect of the selected
  parameters.
  
- protection/demo_protection: demonstrates cell and sheet protection
  supported by FPSpreadsheet.

- recursive_calculation/demo_recursive_calc: demonstrates recursive 
  calculation of formulas. All formulas in this demo depend on the result 
  of the formula in the next cell, except for the last cell. When the formula 
  in the first cell is calculated recursive calculation of the other cells 
  is requested.
  
- richtext/demo_richtext_utf8: shows working with rich text formatting in cells

- rpn_formulas/demo_write_formula: shows some rpn formulas

- searching/demo_search:  demonstrates how specific cell content can be 
  searched within a worksheet
  
- sorting/demo_sorting: shows how cell ranges can be sorted within FPSpreadsheet

- user_defined_formulas/demo_formula_func: shows how a user-provided function 
  can be registered in fpspreadsheet for usage in rpn formulas. The example 
  covers some financial functions.
  
- defined_names: shows how "defined names" (named cells) can be used in
  FPSpreadsheet. Also in formulas.
  
- virtual_mode/demo_virtualmode_writing: demonstrates how the virtual mode 
  of the workbook can be used to create huge spreadsheet files.
  
- virtual_mode/demo_virtualmode_reading: demonstrates how the virtual mode 
  of the workbook can be used to read huge spreadsheet files. Requires the 
  file written by demo_virtualmode_writing.
  
Users of Lazarus 2.1+ can compile all demo projects with a single click by using
the other_demos project group.

  