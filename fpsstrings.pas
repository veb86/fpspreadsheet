{ Translatable strings for fpspreadsheet }

unit fpsStrings;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}


interface

resourcestring
  rsDefaultSheetName = 'Sheet%d';

  // File formats
  rsFileFormatExcel2 = 'Excel 2.1';
  rsFileFormatExcel5 = 'Excel 5';
  rsFileFormatExcel8 = 'Excel 97-2003';
  rsFileFormatExcelXML = 'Excel XP/2003 XML';
  rsFileFormatExcelXLSX = 'Excel 2007+ XML';
  rsFileFormatOpenDocument = 'OpenDocument';
  rsFileFormatSYLK = 'SYLK';
  rsFileFormatCSV = 'CSV';
  rsFileFormatHTML = 'HTML';
  rsFileFormatWikiTablePipes = 'WikiTable (Pipes)';
  rsFileFormatWikiTableWikiMedia = 'WikiTable (WikiMedia)';
  rsFiles = 'files';
  rsAllSpreadsheetFiles = 'All spreadsheet files';
  rsAllExcelFiles = 'All Excel files';

  // Reading / writing
  rsCannotReadFile = 'Cannot read file "%s".' + LineEnding + 'Wrong, unknown or defective file format?';
  rsUnsupportedReadFormat = 'Tried to read a spreadsheet using an unsupported format';
  rsUnsupportedWriteFormat = 'Tried to write a spreadsheet using an unsupported format';
  rsNoValidSpreadsheetFile = '"%s" is not a valid spreadsheet file';
  rsUnknownSpreadsheetFormat = 'unknown format';
  rsReaderNotFound = 'Reader not found for file "%s"';

  rsExportFileIsRequired = 'Export file name is required';
  rsFPSExportDescription = 'Spreadsheet file';
  rsMultipleSheetsOnlyWithRestorePosition = 'Export to multiple sheets is possible '+
    'only if position is restored.';
  rsMaxRowsExceeded = 'This workbook contains %d rows, but the selected ' +
    'file format does not support more than %d rows.';
  rsMaxColsExceeded = 'This workbook contains %d columns, but the selected ' +
    'file format does not support more than %d columns.';
  rsTooManyPaletteColors = 'This workbook contains more colors (%d) than ' +
    'supported by the file format (%d). The additional colors are replaced by '+
    'the best-matching palette colors.';
  rsInvalidExtension = 'Attempting to save a spreadsheet by extension, ' +
    'but the extension %s is not valid.';
  rsInvalidFontIndex = 'Invalid font index';
  rsInvalidNumberFormat = 'Trying to use an incompatible number format.';
  rsInvalidDateTimeFormat = 'Trying to use an incompatible date/time format.';
  rsNoValidNumberFormatString = 'No valid number format string.';
  rsIsNoValidNumberFormatString = '%s is not a valid number format string.';
  rsNoValidCellAddress = '"%s" is not a valid cell address.';
  rsNoValidCellRangeAddress = '"%s" is not a valid cell range address.';
  rsNoValidCellRangeOrCellAddress = '"%s" is not a valid cell or cell range address.';
  rsSpecifyNumberOfParams = 'Specify number of parameters for function %s';
  rsIncorrectParamCount = 'Funtion %s requires at least %d and at most %d parameters.';
  rsCircularReference = 'Circular reference found when calculating worksheet formulas';
  rsFileNotFound = 'File "%s" not found.';
  rsFileAlreadyExists = 'File "%s" already exists.';
  rsWorksheetNotFound = 'Worksheet "%s" not found.';
  rsWorksheetNotFound1 = 'Worksheet not found.';
  rsInvalidWorksheetName = '"%s" is not a valid worksheet name.';
  rsDefectiveInternalStructure = 'Defective internal structure of %s file.';
  rsFileStructureError = 'File structure error in %s record, position %d.';
  rsUnknownDataType = 'Unknown data type.';
  rsUnknownErrorType = 'Unknown error type.';
  rsTruncateTooLongCellText = 'Text value exceeds %d character limit in cell %s '+
    'and has been truncated.';
  rsColumnStyleNotFound = 'Column style not found.';
  rsRowStyleNotFound = 'Row style not found.';
  rsInvalidCharacterInCell = 'Invalid character(s) in cell %s.';
  rsInvalidCharacterInCellComment = 'Invalid character(s) in cell comment "%s".';
  rsUTF8TextExpectedButANSIFoundInCell = 'Expected UTF8 text but probably ANSI '+
    'text found in cell %s.';
  rsIndexInSSTOutOfRange = 'Index %d in SST out of range (0-%d).';
  rsAmbiguousDecThouSeparator = 'Assuming usage of decimal separator in "%s".';
  rsCodePageNotSupported = 'Code page "%s" is not supported. Using "cp1252" (Latin 1) instead.';
  rsFormulaNotSupported = 'The formula in cell %s is not supported by this file format: %s';

  rsNoValidHyperlinkInternal = 'The hyperlink "%s" is not a valid cell address.';
  rsNoValidHyperlinkURI = 'The hyperlink "%s" is not a valid URI.';
  rsLocalFileHyperlinkAbs = 'The hyperlink "%s" points to a local file. ' +
    'In case of an absolute path the protocol "file:" must be specified.';
  rsEmptyHyperlink = 'The hyperlink is not specified.';
  rsODSHyperlinksOfTextCellsOnly = 'Cell %s: OpenDocument supports hyperlinks for text cells only.';
  rsStdHyperlinkTooltip = 'Hold the left mouse button down for a short time to activate the hyperlink.';

  rsCannotSortMerged = 'The cell range cannot be sorted because it contains merged cells.';

  rsDifferentSheetPrintRange = 'Print range "%s" requires a different worksheet.';

  // Colors
  rsAqua = 'aqua';
  rsBeige = 'beige';
  rsBlack = 'black';
  rsBlue = 'blue';
  rsBlueGray = 'blue gray';
  rsBrown = 'brown';
  rsCoral = 'coral';
  rsCyan = 'cyan';
  rsDarkBlue = 'dark blue';
  rsDarkGreen = 'dark green';
  rsDarkPurple = 'dark purple';
  rsDarkRed = 'dark red';
  rsDarkTeal = 'dark teal';
  rsGold = 'gold';
  rsGray = 'gray';
  rsGray10pct = '10% gray';
  rsGray20pct = '20% gray';
  rsGray25pct = '25% gray';
  rsGray40pct = '40% gray';
  rsGray50pct = '50% gray';
  rsGray80pct = '80% gray';
  rsGreen = 'green';
  rsIceBlue = 'ice blue';
  rsIndigo = 'indigo';
  rsIvory = 'ivory';
  rsLavander = 'lavander';
  rsLightBlue = 'light blue';
  rsLightGreen = 'light green';
  rsLightOrange = 'light orange';
  rsLightTurquoise = 'light turquoise';
  rsLightYellow = 'light yellow';
  rsLime = 'lime';
  rsMagenta = 'magenta';
  rsNavy = 'navy';
  rsOceanBlue = 'ocean blue';
  rsOlive = 'olive';
  rsOliveGreen = 'olive green';
  rsOrange = 'orange';
  rsPaleBlue = 'pale blue';
  rsPeriwinkle = 'periwinkle';
  rsPink = 'pink';
  rsPlum = 'plum';
  rsPurple = 'purple';
  rsRed = 'red';
  rsRose = 'rose';
  rsSeaGreen = 'sea green';
  rsSilver = 'silver';
  rsSkyBlue = 'sky blue';
  rsTan = 'tan';
  rsTeal = 'teal';
  rsVeryDarkGreen = 'very dark green';
  rsViolet = 'violet';
  rsWheat = 'wheat';
  rsWhite = 'white';
  rsYellow = 'yellow';

  rsNotDefined = 'not defined';
  rsTransparent = 'transparent';
  rsPaletteIndex = 'Palette index %d';

  rsTRUE = 'TRUE';               // wp: Do we really want to translate these strings?
  rsFALSE = 'FALSE';
  rsErrEmptyIntersection = '#NULL!';
  rsErrDivideByZero = '#DIV/0!';
  rsErrWrongType = '#VALUE!';
  rsErrIllegalRef = '#REF!';
  rsErrWrongName = '#NAME?';
  rsErrOverflow = '#NUM!';
  rsErrArgError = '#N/A';
  rsErrFormulaNotSupported = '<FORMULA?>';

  rsOperationExceedsColCount = 'This operation exceeds the range of defined grid columns.';
  rsOperationExceedsRowCount = 'This operation exceeds the range of defined grid rows.';

implementation

end.
