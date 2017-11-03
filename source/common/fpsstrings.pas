{ Translatable strings for fpspreadsheet }

unit fpsStrings;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

interface

resourcestring
  // Files & file formats
  rsAllSpreadsheetFiles = 'All spreadsheet files';
  rsAllExcelFiles = 'All Excel files';
  rsCannotReadFile = 'Cannot read file "%s". Wrong, unknown or defective file format?';
  rsDefectiveInternalFileStructure = 'Defective internal structure of %s file.';
  rsFileAlreadyExists = 'File "%s" already exists.';
  rsFileFormatNotSupported = 'File format of "%s" not supported.';
  rsFileNotFound = 'File "%s" not found.';
  rsFiles = 'files';
  rsFileStructureError = 'File structure error in %s record, position %d.';
  rsIndexInSSTOutOfRange = 'Index %d in SST out of range (0-%d).';
  rsInvalidExtension = 'Attempting to save a spreadsheet by extension, ' +
    'but the extension %s is not valid.';
  rsInvalidSpreadsheetFile = '"%s" is not a valid spreadsheet file';
  rsReaderNotFound = 'Reader not found for file "%s"';
  rsUnsupportedReadFormat = 'Tried to read a spreadsheet using an unsupported format';
  rsUnsupportedWriteFormat = 'Tried to write a spreadsheet using an unsupported format';

  // File format limitations
  rsMaxRowsExceeded = 'This workbook contains %d rows, but the selected ' +
    'file format does not support more than %d rows.';
  rsMaxColsExceeded = 'This workbook contains %d columns, but the selected ' +
    'file format does not support more than %d columns.';
  rsTooManyPaletteColors = 'This workbook contains more colors (%d) than ' +
    'supported by the file format (%d). The additional colors are replaced by '+
    'the best-matching palette colors.';
  rsTruncateTooLongCellText = 'Text value exceeds the %d character limit in ' +
    'cell %s and has been truncated.';
  rsTruncateTooLongText = 'Text value exceeds the %d character limit ' +
    'and has been truncated.';
  rsWriteError_WorksheetNameTooLong = 'File cannot be written because ' +
    'the name of worksheet "%0:s" is too long (max %1:d characters).';

  // Cells
  rsInvalidCharacterInCell = 'Invalid character(s) in cell %s.';
  rsNoValidCellAddress = '"%s" is not a valid cell address.';
  rsNoValidCellRangeAddress = '"%s" is not a valid cell range address.';
  rsNoValidCellRangeOrCellAddress = '"%s" is not a valid cell or cell range address.';
  rsUTF8TextExpectedButANSIFoundInCell = 'Expected UTF8 text, '+
    'but probably ANSI text found in cell %s.';

  // Code page
  rsCodePageNotSupported = 'Code page "%s" is not supported. Using "cp1252" (Latin 1) instead.';

  // Colors
  // EGA palette
  rsBlack = 'black';
  rsWhite = 'white';
  rsRed = 'red';
  rsGreen = 'green';
  rsBlue = 'blue';
  rsYellow = 'yellow';
  rsMagenta = 'magenta';
  rsCyan = 'cyan';
  rsDarkRed = 'dark red';
  rsDarkGreen = 'dark green';
  rsDarkBlue = 'dark blue';
  rsOlive = 'olive';
  rsPurple = 'purple';
  rsTeal = 'teal';
  rsSilver = 'silver';
  rsGray = 'gray';

  // Special colors
  rsNotDefined = 'not defined';
  rsTransparent = 'transparent';
  rsPaletteIndex = 'Palette index %d';

  // Columns
  rsColumnStyleNotFound = 'Column style not found.';

  // Comments
  rsInvalidCharacterInCellComment = 'Invalid character(s) in cell comment "%s".';

  // Expression parser
  // These strings are mostly taken or adapted from fpexprpars
  rsBadQuotes = 'Unterminated string';
  rsCircularReference = 'Circular reference found when calculating worksheet '+
    'formula in cell %s';
  rsCommaExpected =  'Expected comma (,) at position %d, but got %s';
  rsDuplicateIdentifier = 'An identifier with name "%s" already exists.';
  rsErrorInExpression = 'Cannot evaluate: error in expression';
  rsExpressionEmpty = 'Cannot evaluate: empty expression';
  rsInvalidArgumentCount = 'Invalid argument count for function %s';
  rsInvalidFloat = '%s is not a valid floating-point value';
  rsInvalidNumber = 'Invalid numerical value : %s';
  rsInvalidNumberChar = 'Unexpected character in number : %s';
  rsInvalidResultCharacter = '"%s" is not a valid return type indicator';
  rsInvalidResultType = 'Invalid result type: %s';
  rsLeftBracketExpected = 'Expected left bracket at position %d, but got %s';
  rsNoOperand = 'No operand for unary operation %s';
  rsNoPercentOperation = 'Cannot perform percent operation on expression ' +
    'of type %s: %s';
  rsNoVariable = 'Identifier %s is not a variable';
  rsRightBracketExpected = 'Expected right bracket at position %d, but got %s';
  rsUnexpectedEndOfExpression = 'Unexpected end of expression';
  rsUnknownCharacter = 'Unknown character at pos %d: "%s"';
  rsUnknownComparison = 'Internal error: Unknown comparison';
  rsUnknownDelimiter = 'Unknown delimiter character: "%s"';
  rsUnknownIdentifier = 'Unknown identifier: %s';
  rsUnknownTokenAtPos = 'Unknown token at pos %d : %s';
  rsUnterminatedExpression = 'Badly terminated expression. Found token at '+
    'position %d : %s';

  { -- currently not used:
  SErrNoLeftOperand = 'No left operand for binary operation %s';
  SErrNoRightOperand = 'No left operand for binary operation %s';
  SErrNoNegation = 'Cannot negate expression of type %s: %s';
  SErrNoUPlus = 'Cannot perform unary plus operation on type %s: %s';
  SErrTypesDoNotMatch = 'Type mismatch: %s<>%s for expressions "%s" and "%s".';
  SErrNoNodeToCheck = 'Internal error: No node to check !';
  SInvalidNodeType = 'Node type (%s) not in allowed types (%s) for expression: %s';
  SErrNoNOTOperation = 'Cannot perform NOT operation on expression of type %s: %s';
  }

  // Format
  rsAmbiguousDecThouSeparator = 'Assuming usage of decimal separator in "%s".';
  rsInvalidDateTimeFormat = 'Trying to use an incompatible date/time format.';
  rsInvalidFontIndex = 'Invalid font index';
  rsInvalidNumberFormat = 'Trying to use an incompatible number format.';
  rsNoValidNumberFormatString = 'No valid number format string.';

  // Formulas
  rsFormulaNotSupported = 'The formula in cell %s is not supported by this file format: %s';
  rsUnknownDataType = 'Unknown data type.';
  rsUnknownErrorType = 'Unknown error type.';

  // Hyperlinks
  rsEmptyHyperlink = 'The hyperlink is not specified.';
  rsLocalFileHyperlinkAbs = 'The hyperlink "%s" points to a local file. ' +
    'In case of an absolute path the protocol "file:" must be specified.';
  rsNoValidHyperlinkInternal = 'The hyperlink "%s" is not a valid cell address.';
  rsNoValidHyperlinkURI = 'The hyperlink "%s" is not a valid URI.';
  rsODSHyperlinksOfTextCellsOnly = 'Cell %s: OpenDocument supports hyperlinks '+
    'for text cells only.';
  rsStdHyperlinkTooltip = 'Hold the left mouse button down for a short time '+
    'to activate the hyperlink.';

  // Images
  rsImageFormatNotSupported = 'Image format not supported.';

  // PageLayout
  rsDifferentSheetPrintRange = 'Print range "%s" requires a different worksheet.';
  rsFooter = 'Footer';
  rsHeader = 'Header';
  rsIncorrectPositionOfImageInHeaderFooter = 'Incorrect position of %%G code in %s';
  rsOnlyOneHeaderFooterImageAllowed = 'Only one image per %s section allowed.';

  // Rows
  rsRowStyleNotFound = 'Row style not found.';

  // Sorting
  rsCannotSortMerged = 'The cell range cannot be sorted because it contains merged cells.';

  // Worksheets
  rsDefaultSheetName = 'Sheet%d';
  rsInvalidWorksheetName = '"%s" is not a valid worksheet name.';
  rsWorksheetNotFound = 'Worksheet "%s" not found.';
  rsWorksheetNotFound1 = 'Worksheet not found.';

  // WorksheetGrid
  rsOperationExceedsColCount = 'This operation at index %d exceeds the range of defined grid columns (%d).';
  rsOperationExceedsRowCount = 'This operation at index %d exceeds the range of defined grid rows (%d).';

  // Export
  rsExportFileIsRequired = 'Export file name is required.';
  rsFPSExportDescription = 'Spreadsheet file';
  rsMultipleSheetsOnlyWithRestorePosition = 'Export to multiple sheets is possible '+
    'only if position is restored.';

  // Protection
  rsPasswordRemoved_BIFF2 = 'Password removed (BIFF2 requires matching workbook '+
    'and worksheet passwords)';
  rsPasswordRemoved_NotValid = 'Password removed (Not valid).';
  rsPasswordRemoved_Excel = 'Password removed (Hashing algorithm not compatible with Excel)';

const
  // Color names which do not have to be translated. They will be removed.
  rsAqua = 'aqua' deprecated;
  rsBeige = 'beige' deprecated;
  rsBlueGray = 'blue gray' deprecated;
  rsBrown = 'brown' deprecated;
  rsCoral = 'coral' deprecated;
  rsDarkPurple = 'dark purple' deprecated;
  rsDarkTeal = 'dark teal' deprecated;
  rsFuchsia = 'fuchsia' deprecated;
  rsGold = 'gold' deprecated;
  rsGray10pct = '10% gray' deprecated;
  rsGray20pct = '20% gray' deprecated;
  rsGray25pct = '25% gray' deprecated;
  rsGray40pct = '40% gray' deprecated;
  rsGray50pct = '50% gray' deprecated;
  rsGray80pct = '80% gray' deprecated;
  rsIceBlue = 'ice blue' deprecated;
  rsIndigo = 'indigo' deprecated;
  rsIvory = 'ivory' deprecated;
  rsLavander = 'lavander' deprecated;
  rsLightBlue = 'light blue' deprecated;
  rsLightGreen = 'light green' deprecated;
  rsLightOrange = 'light orange' deprecated;
  rsLightTurquoise = 'light turquoise' deprecated;
  rsLightYellow = 'light yellow' deprecated;
  rsLime = 'lime' deprecated;
  rsMaroon = 'maroon' deprecated;
  rsNavy = 'navy' deprecated;
  rsOceanBlue = 'ocean blue' deprecated;
  rsOliveGreen = 'olive green' deprecated;
  rsOrange = 'orange' deprecated;
  rsPaleBlue = 'pale blue' deprecated;
  rsPeriwinkle = 'periwinkle' deprecated;
  rsPink = 'pink' deprecated;
  rsPlum = 'plum' deprecated;
  rsRose = 'rose' deprecated;
  rsSeaGreen = 'sea green' deprecated;
  rsSkyBlue = 'sky blue' deprecated;
  rsTan = 'tan' deprecated;
  rsVeryDarkGreen = 'very dark green' deprecated;
  rsViolet = 'violet' deprecated;
  rsWheat = 'wheat' deprecated;

implementation

end.
