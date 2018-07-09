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
  rsRightSquareBracketExpected = 'Expected right square bracket at positon %d, but got %s';
  rsUnexpectedEndOfExpression = 'Unexpected end of expression';
  rsUnknownCharacter = 'Unknown character at pos %d: "%s"';
  rsUnknownComparison = 'Internal error: Unknown comparison';
  rsUnknownDelimiter = 'Unknown delimiter character: "%s"';
  rsUnknownIdentifier = 'Unknown identifier: %s';
  rsUnknownTokenAtPos = 'Unknown token at pos %d : %s';
  rsUnterminatedExpression = 'Badly terminated expression. Found token at '+
    'position %d : %s';
  rsIllegalODSCellRange = 'Illegal structure of an OpenDocument cell range.';

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
  rsDuplicateWorksheetName = 'Duplicate worksheet "%s".';
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

{ Hints and caption for fpspreadsheet actions }
  rsAddCaption = 'Add...';
  rsDeleteCaption = 'Delete';
  rsRenameCaption = 'Rename...';
  rsZoomCaption = 'Zoom';
  rsAddWorksheetHint = 'Add worksheet';
  rsDeleteWorksheetHint = 'Delete worksheet';
  rsRenameWorksheetHint = 'Rename worksheet';
  rsZoomWorksheetHint = 'Zoom worksheet';
  rsClearFormat = 'Clear format';
  rsFontStyle_Bold = 'Bold';
  rsFontStyle_Italic = 'Italic';
  rsFontStyle_Underlined = 'Underlined';
  rsFontStyle_StrikeThrough = 'Strike-through';
  rsHorAlignment_Default = 'Default horizontal alignment';
  rsHorAlignment_Left = 'Left-justified text';
  rsHorAlignment_Center = 'Horizontally centered text';
  rsHorAlignment_Right = 'Right-justified text';
  rsVertAlignment_Default = 'Default vertical alignment';
  rsVertAlignment_Top = 'Top-aligned text';
  rsVertAlignment_Center = 'Vertically centered text';
  rsVertAlignment_Bottom = 'Bottom-aligned text';
  rsTextRotation_Hor = 'Horizontal';
  rsTextRotation_Vert_CW = 'Vertical (90° clockwise)';
  rsTextRotation_Vert_CCW = 'Vertical (90° counter-clockwise)';
  rsTextRotation_Vert_Stacked = 'Vertically stacked';
  rsWordwrap = 'Word-wrap';
  rsNumberFormatCaption_General = 'General';
  rsNumberFormatCaption_Fixed = 'Fixed';
  rsNumberFormatCaption_FixedTh = 'Fixed w/thousand separator';
  rsNumberFormatCaption_Exp = 'Exponential';
  rsNumberFormatCaption_Percentage = 'Percent';
  rsNumberFormatCaption_Fraction = 'Fraction';
  rsNumberFormatCaption_Currency = 'Currency';
  rsNumberFormatCaption_CurrencyRed = 'Currency (red)';
  rsNumberFormatCaption_ShortDateTime = 'Date and time';
  rsNumberFormatCaption_ShortDate = 'Short date';
  rsNumberFormatCaption_LongDate = 'Long date';
  rsNumberFormatCaption_ShortTime = 'Short time';
  rsNumberFormatCaption_LongTime = 'Long time';
  rsNumberFormatCaption_ShortTimeAM = 'Short time AM/PM';
  rsNumberFormatCaption_LongTimeAM = 'Long time AM/PM';
  rsNumberFormatCaption_DayMonth = 'Day and month';
  rsNumberFormatCaption_MonthYear = 'Month and year';
  rsNumberFormatCaption_TimeInterval = 'Time interval';
  rsNumberFormatCaption_Text = 'Text';
  rsNumberFormatCaption_Custom = 'Custom...';
  rsNumberFormatHint_General = 'General number format';
  rsNumberFormatHint_Fixed = 'Fixed number of decimal places';
  rsNumberFormatHint_FixedTh = 'Fixed number of decimal places, thousand separator';
  rsNumberFormatHint_Exp = 'Exponential (scientific) format';
  rsNumberFormatHint_Percentage = 'Percent';
  rsNumberFormatHint_Fraction = 'Fraction';
  rsNumberFormatHint_Currency = 'Currency';
  rsNumberFormatHint_CurrencyRed = 'Currency (negative values in red)';
  rsNumberFormatHint_ShortDateTime = 'Short date and time';
  rsNumberFormatHint_ShortDate = 'Short date';
  rsNumberFormatHint_LongDate = 'Long date';
  rsNumberFormatHint_ShortTime = 'Short time';
  rsNumberFormatHint_LongTime = 'Long time';
  rsNumberFormatHint_ShortTimeAM = 'Short time with AM/PM';
  rsNumberFormatHint_LongTimeAM = 'Long time with AM/PM';
  rsNumberFormatHint_DayMonth = 'Day and month';
  rsNumberFormatHint_MonthYear = 'Month and year';
  rsNumberFormatHint_TimeInterval = 'Time interval';
  rsNumberFormatHint_Text = 'Number as text';
  rsNumberFormatHint_Custom = 'Custom number format';
  rsMoreDecimals = 'More decimals';
  rsLessDecimals = 'Less decimals';
  rsCommentCaption_New = 'New comment...';
  rsCommentCaption_Edit = 'Edit comment...';
  rsCommentCaption_Delete = 'Delete comment';
  rsCommentHint_New = 'New comment';
  rsCommentHint_Edit = 'Edit comment';
  rsCommentHint_Delete = 'Delete comment';
  rsHyperlinkCaption_New = 'New hyperlink...';
  rsHyperlinkCaption_Edit = 'Edit hyperlink...';
  rsHyperlinkCaption_Delete = 'Delete hyperlink';
  rsHyperlinkHint_New = 'New hyperlink';
  rsHyperlinkHint_Edit = 'Edit hyperlink';
  rsHyperlinkHint_Delete = 'Delete hyperlink';
  rsMergeUnmerge = 'Merge/unmerge';
  rsCellFontCaption = 'Cell font...';
  rsCellFontHint = 'Cell font';
  rsBackgroundColorCaption = 'Background color...';
  rsBackgroundColorHint = 'Background color';
  rsCellBorder = 'Cell border';
  rsNoCellBorders = 'No borders';
  rsBorderTop_Menu = 'Top';
  rsBorderTop_Hint = 'Top border';
  rsBorderTopFmt_Menu = 'Top (%s)';
  rsBorderTopFmt_Hint = 'Top border (%s)';
  rsBorderBottom_Menu = 'Bottom';
  rsBorderBottom_Hint = 'Bottom border';
  rsBorderBottomFmt_Menu = 'Bottom (%s)';
  rsBorderBottomFmt_Hint = 'Bottom border (%s)';
  rsBorderLeft_Menu = 'Left';
  rsBorderLeft_Hint = 'Left border';
  rsBorderLeftFmt_Menu = 'Left (%s)';
  rsBorderLeftFmt_Hint = 'Left border (%s)';
  rsBorderRight_Menu = 'Right';
  rsBorderRight_Hint = 'Right border';
  rsBorderRightFmt_Menu = 'Right (%s)';
  rsBorderRightFmt_Hint = 'Right border (%s)';
  rsBorderTopBottomFmt_Menu = 'Top && %s bottom';
  rsBorderTopBottomFmt_Hint = 'Top && %s bottom border lines';
  rsBorderInnerHor_Menu = 'Inner horizontal';
  rsBorderInnerHor_Hint = 'Inner horizontal border';
  rsBorderInnerHorFmt_Menu = 'Inner horizontal (%s)';
  rsBorderinnerHorFmt_Hint = 'Inner horizontal lines (%s)';
  rsBorderInnerVert_Menu = 'Inner vertical';
  rsBorderInnerVert_Hint = 'Inner vertical border';
  rsBorderInnerVertFmt_Menu = 'Inner vertical (%s)';
  rsBorderInnerVertFmt_Hint = 'Inner vertical border (%s)';
  rsBorderAllHor_Menu = 'All horizontal';
  rsBorderAllHor_Hint = 'All horizontal lines';
  rsBorderAllHorFmt_Menu = 'All horizontal (%s)';
  rsBorderAllHorFmt_Hint = 'All horizontal lines (%s)';
  rsBorderAllVert_Menu = 'All vertical';
  rsBorderAllVert_Hint = 'All vertical lines';
  rsBorderAllVertFmt_Menu = 'All vertical (%s)';
  rsBorderAllVertFmt_Hint = 'All vertical lines (%s)';
  rsBorderAllInner_Menu = 'All inner';
  rsBorderAllInner_Hint = 'All inner lines';
  rsBorderAllInnerFmt_Menu = 'All inner (%s)';
  rsBorderAllInnerFmt_Hint = 'All inner lines (%s)';
  rsBorderAllOuter_Menu = 'All outer';
  rsBorderAllOuter_Hint = 'All outer lines';
  rsBorderAllOuterFmt_Menu = 'All outer (%s)';
  rsBorderAllOuterFmt_Hint = 'All outer lines (%s)';
  rsBorderAll_Menu = 'All';
  rsBorderAll_Hint = 'All lines';
  rsBorderAllFmt_Menu = 'All (%s)';
  rsBorderAllFmt_Hint = 'All lines (%s)';
  rsBorderDiagUp = 'Upward diagonal';
  rsBorderDiagUpFmt = 'Upward diagonal (%s)';
  rsBorderDiagDown = 'Downward diagonal';
  rsBorderDiagDownFmt = 'Downward diagonal (%s)';
  rsThin = 'thin';
  rsMedium = 'thick';
  rsDashed = 'dashed';
  rsDotted = 'dotted';
  rsThick = 'very thick';
  rsDouble = 'double';
  rsHair = 'hair';
  rsMediumDash = 'thick dash';
  rsDashDot = 'dash-dot';
  rsMediumDashDot = 'thick dash-dot';
  rsDashDotDot = 'dash-dot-dot';
  rsMediumDashDotDot = 'thick dash-dot-dot';
  rsSlantDashDot = 'slanted dash-dot';


implementation

end.
