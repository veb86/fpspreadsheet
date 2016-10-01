{@@ ----------------------------------------------------------------------------
  Unit fpsTypes collects the most <b>fundamental declarations</b> used
  throughout the fpspreadsheet library.

  AUTHORS: Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
            distribution, for details about the license.
-------------------------------------------------------------------------------}
unit fpsTypes;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpimage;

type
  {@@ Built-in file formats of fpspreadsheet }
  TsSpreadsheetFormat = (sfExcel2, sfExcel5, sfExcel8, sfExcelXML, sfOOXML,
    sfOpenDocument, sfCSV, sfHTML, sfWikiTable_Pipes, sfWikiTable_WikiMedia,
    sfUser);   // Use this for user-defined readers/writers

  {@@ Numerical identifier for file formats, built-in and user-provided }
  TsSpreadFormatID = integer;

  {@@ Array of file format identifiers }
  TsSpreadFormatIDArray = array of TsSpreadFormatID;

const
  {@@ Format identifier of an undefined, unknown, etc. file format. }
  sfidUnknown = -1;
  { Each unit implementing a reader/writer will define an sfidXXXX value as a
    numerical identifer of the file format. In case of the built-in formats,
    the identifier is equal to the ord of the TsSpreadsheetFormat value. }

type
  {@@ Flag set during reading or writing of a workbook }
  TsReadWriteFlag = (rwfNormal, rwfRead, rwfWrite);

  {@@ Record collection limitations of a particular file format }
  TsSpreadsheetFormatLimitations = record
    MaxRowCount: Cardinal;
    MaxColCount: Cardinal;
    MaxPaletteSize: Integer;
    MaxSheetNameLength: Integer;
  end;

const
  {@@ Explanatory name of sfBiff2 file format }
  STR_FILEFORMAT_EXCEL_2 = 'Excel 2.1';
  {@@ Explanatory name of sfBiff5 file format }
  STR_FILEFORMAT_EXCEL_5 = 'Excel 5';
  {@@ Explanatory name of sfBiff8 file format }
  STR_FILEFORMAT_EXCEL_8 = 'Excel 97-2003';
  {@@ Explanatory name of sfExcelXML file format }
  STR_FILEFORMAT_EXCEL_XML = 'Excel XP/2003 XML';
  {@@ Explanatory name of sfOOXLM file format }
  STR_FILEFORMAT_EXCEL_XLSX = 'Excel 2007+ XML';
  {@@ Explanatory name of sfOpenDocument file format }
  STR_FILEFORMAT_OPENDOCUMENT = 'OpenDocument';
  {@@ Explanatory name of sfCSV file format }
  STR_FILEFORMAT_CSV = 'CSV';
  {@@ Explanatory name of sfHTML file format }
  STR_FILEFORMAT_HTML = 'HTML';
  {@@ Explanatory name of sfWikiTablePipes file format }
  STR_FILEFORMAT_WIKITABLE_PIPES = 'WikiTable (Pipes)';
  {@@ Explanatory name of sfWikiTableWikiMedia file format }
  STR_FILEFORMAT_WIKITABLE_WIKIMEDIA = 'WikiTable (WikiMedia)';

  {@@ Default binary <b>Excel</b> file extension (<= Excel 97) }
  STR_EXCEL_EXTENSION = '.xls';
  {@@ Default xml <b>Excel</v> file extension (Excel XP, 2003) }
  STR_XML_EXCEL_EXTENSION = '.xml';
  {@@ Default xml <b>Excel</b> file extension (>= Excel 2007) }
  STR_OOXML_EXCEL_EXTENSION = '.xlsx';
  {@@ Default <b>OpenDocument</b> spreadsheet file extension }
  STR_OPENDOCUMENT_CALC_EXTENSION = '.ods';
  {@@ Default extension of <b>comma-separated-values</b> file }
  STR_COMMA_SEPARATED_EXTENSION = '.csv';
  {@@ Default extension for <b>HTML</b> files }
  STR_HTML_EXTENSION = '.html';
  {@@ Default extension of <b>wikitable files</b> in <b>pipes</b> format}
  STR_WIKITABLE_PIPES_EXTENSION = '.wikitable_pipes';
  {@@ Default extension of <b>wikitable files</b> in <b>wikimedia</b> format }
  STR_WIKITABLE_WIKIMEDIA_EXTENSION = '.wikitable_wikimedia';

  {@@ String for boolean value TRUE }
  STR_TRUE = 'TRUE';
  {@@ String for boolean value FALSE }
  STR_FALSE = 'FALSE';

  {@@ Error values }
  STR_ERR_EMPTY_INTERSECTION = '#NULL!';
  STR_ERR_DIVIDE_BY_ZERO = '#DIV/0!';
  STR_ERR_WRONG_TYPE = '#VALUE!';
  STR_ERR_ILLEGAL_REF = '#REF!';
  STR_ERR_WRONG_NAME = '#NAME?';
  STR_ERR_OVERFLOW = '#NUM!';
  STR_ERR_ARG_ERROR = '#N/A';
  // No Excel errors
  STR_ERR_FORMULA_NOT_SUPPORTED= '<FMLA?>';
  STR_ERR_UNKNOWN = '#UNKNWON!';

  {@@ Maximum count of worksheet columns}
  MAX_COL_COUNT = 65535;

  {@@ Unassigned row/col index }
  UNASSIGNED_ROW_COL_INDEX = $FFFFFFFF;

  {@@ Name of the default font}
  DEFAULT_FONTNAME = 'Arial';
  {@@ Size of the default font}
  DEFAULT_FONTSIZE = 10;
  {@@ Index of the default font in workbook's font list }
  DEFAULT_FONTINDEX = 0;
  {@@ Index of the hyperlink font in workbook's font list }
  HYPERLINK_FONTINDEX = 1;
  {@@ Index of bold default font in workbook's font list }
  BOLD_FONTINDEX = 2;
  {@@ Index of italic default font in workbook's font list - not used directly }
  ITALIC_FONTINDEX = 3;

  {@@ Line ending character in cell texts with fixed line break. Using a
      unique value simplifies many things... }
  FPS_LINE_ENDING = #10;


type
  {@@ Units for size dimensions }
  TsSizeUnits = (suChars, suLines, suMillimeters, suCentimeters, suPoints, suInches);

const
  {@@ Unit names }
  SizeUnitNames: array[TsSizeUnits] of string = (
    'chars', 'lines', 'mm', 'cm', 'pt', 'in');

  {@@ Takes account of effect of cell margins on row height by adding this
      value to the nominal row height. Note that this is an empirical value
      and may be wrong. }
  ROW_HEIGHT_CORRECTION = 0.3;

  {@@ Ratio of the width of the "0" character to the font size.
    Empirical value to match Excel and LibreOffice column withs.
    Needed because Excel defines colum width in terms of count of the "0"
    character. }
  ZERO_WIDTH_FACTOR = 351/640;


type
  {@@ Tokens to identify the <b>elements in an expanded formula</b>.

   NOTE: When adding or rearranging items
   * make sure that the subtypes TOperandTokens and TBasicOperationTokens
     are complete
   * make sure to keep the table "TokenIDs" in unit xlscommon in sync
  }
  TFEKind = (
    { Basic operands }
    fekCell, fekCellRef, fekCellRange, fekCellOffset,
    fekCellRef3d, fekCellRange3d,
    fekNum, fekInteger, fekString, fekBool, fekErr, fekMissingArg,
    { Basic operations }
    fekAdd, fekSub, fekMul, fekDiv, fekPercent, fekPower, fekUMinus, fekUPlus,
    fekConcat,  // string concatenation
    fekEqual, fekGreater, fekGreaterEqual, fekLess, fekLessEqual, fekNotEqual,
    fekList,    // List operator
    fekParen,   // show parenthesis around expression node  -- don't add anything after fekParen!
    { Functions - they are identified by their name }
    fekFunc
  );

  {@@ These tokens identify operands in RPN formulas. }
  TOperandTokens = fekCell..fekMissingArg;

  {@@ These tokens identify basic operations in RPN formulas. }
  TBasicOperationTokens = fekAdd..fekParen;

type
  {@@ Flags to mark the address or a cell or a range of cells to be <b>absolute</b>
      or <b>relative</b>. They are used in the set TsRelFlags. }
  TsRelFlag = (rfRelRow, rfRelCol, rfRelRow2, rfRelCol2);

  {@@ Flags to mark the address of a cell or a range of cells to be <b>absolute</b>
      or <b>relative</b>. It is a set consisting of TsRelFlag elements. }
  TsRelFlags = set of TsRelFlag;

const
  {@@ Abbreviation of all-relative cell reference flags }
  rfAllRel = [rfRelRow, rfRelCol, rfRelRow2, rfRelCol2];

  {@@ Separator between worksheet name and cell (range) reference in an address }
  SHEETSEPARATOR = '!';

type
  {@@ Elements of an expanded formula.
    Note: If ElementKind is fekCellOffset, "Row" and "Col" have to be cast to signed integers! }
  TsFormulaElement = record
    ElementKind: TFEKind;
    Row, Row2: Cardinal;    // zero-based
    Col, Col2: Cardinal;    // zero-based
    Sheet, Sheet2: Integer; // zero-based
    SheetNames: String;     // both sheet names separated by a TAB character (intermediate use only)
    DoubleValue: double;
    IntValue: Int64;
    StringValue: String;
    RelFlags: TsRelFlags;   // info on relative/absolute addresses
    FuncName: String;
    ParamsNum: Byte;
  end;

  {@@ RPN formula. Similar to the expanded formula, but in RPN notation.
      Simplifies the task of format writers which need RPN }
  TsRPNFormula = array of TsFormulaElement;

  {@@ Formula dialect }
  TsFormulaDialect = (fdExcelA1, fdExcelR1C1, fdOpenDocument);

  {@@ Describes the <b>type of content</b> in a cell of a TsWorksheet }
  TCellContentType = (cctEmpty, cctFormula, cctNumber, cctUTF8String,
    cctDateTime, cctBool, cctError);

  {@@ The record TsComment describes a comment attached to a cell.
     @param   Row        (0-based) row index of the cell
     @param   Col        (0-based) column index of the cell
     @param   Text       Comment text }
  TsComment = record
    Row, Col: Cardinal;
    Text: String;
  end;

  {@@ Pointer to a TsComment record }
  PsComment = ^TsComment;

  {@@ The record TsHyperlink contains info on a hyperlink in a cell
    @param   Row          Row index of the cell containing the hyperlink
    @param   Col          Column index of the cell containing the hyperlink
    @param   Target       Target of hyperlink: URI of file, web link, mail; or:
                          internal link (# followed by cell address)
    @param   Note         Text displayed as a popup hint by Excel }
  TsHyperlink = record
    Row, Col: Cardinal;
    Target: String;
    Tooltip: String;
  end;

  {@@ Pointer to a TsHyperlink record }
  PsHyperlink = ^TsHyperlink;

  {@@ Callback function, e.g. for iterating the internal AVL trees of the workbook/sheet}
  TsCallback = procedure (data, arg: Pointer) of object;

  {@@ Error code values }
  TsErrorValue = (
    errOK,                 // no error
    errEmptyIntersection,  // #NULL!
    errDivideByZero,       // #DIV/0!
    errWrongType,          // #VALUE!
    errIllegalRef,         // #REF!
    errWrongName,          // #NAME?
    errOverflow,           // #NUM!
    errArgError,           // #N/A
    // --- no Excel errors --
    errFormulaNotSupported
  );

  {@@ List of possible formatting fields }
  TsUsedFormattingField = (uffTextRotation, uffFont, uffBorder, uffBackground,
    uffNumberFormat, uffWordWrap, uffHorAlign, uffVertAlign, uffBiDi
  );
  { NOTE: "uffBackgroundColor" of older versions replaced by "uffBackground" }

  {@@ Describes which formatting fields are active }
  TsUsedFormattingFields = set of TsUsedFormattingField;

const
  {@@ Codes for curreny format according to FormatSettings.CurrencyFormat:
      "C" = currency symbol, "V" = currency value, "S" = space character
      For the negative value formats, we use also:
      "B" = bracket, "M" = Minus

      The order of these characters represents the order of these items.

      Example: 1000 dollars  --> "$1000"  for pCV,   or "1000 $"  for pVsC
              -1000 dollars --> "($1000)" for nbCVb, or "-$ 1000" for nMCSV

      Assignment taken from "sysstr.inc" }
  pcfDefault = -1;   // use value from Worksheet.FormatSettings.CurrencyFormat
  pcfCV      = 0;    // $1000
  pcfVC      = 1;    // 1000$
  pcfCSV     = 2;    // $ 1000
  pcfVSC     = 3;    // 1000 $

  ncfDefault = -1;   // use value from Worksheet.FormatSettings.NegCurrFormat
  ncfBCVB    = 0;    // ($1000)
  ncfMCV     = 1;    // -$1000
  ncfCMV     = 2;    // $-1000
  ncfCVM     = 3;    // $1000-
  ncfBVCB    = 4;    // (1000$)
  ncfMVC     = 5;    // -1000$
  ncfVMC     = 6;    // 1000-$
  ncfVCM     = 7;    // 1000$-
  ncfMVSC    = 8;    // -1000 $
  ncfMCSV    = 9;    // -$ 1000
  ncfVSCM    = 10;   // 1000 $-
  ncfCSVM    = 11;   // $ 1000-
  ncfCSMV    = 12;   // $ -1000
  ncfVMSC    = 13;   // 1000- $
  ncfBCSVB   = 14;   // ($ 1000)
  ncfBVSCB   = 15;   // (1000 $)

type
  {@@ Text rotation formatting. The text is rotated relative to the standard
      orientation, which is from left to right horizontal:
      <pre>
       --->
       ABC </pre>

      So 90 degrees clockwise means that the text will be:
      <pre>
       |  A
       |  B
       v  C </pre>

      And 90 degree counter clockwise will be:
      <pre>
       ^  C
       |  B
       |  A</pre>

      Due to limitations of the text mode the characters are not rotated here.
      There is, however, also a "stacked" variant which looks exactly like
      the 90-degrees-clockwise case.
  }
  TsTextRotation = (trHorizontal, rt90DegreeClockwiseRotation,
    rt90DegreeCounterClockwiseRotation, rtStacked);

  {@@ Indicates horizontal text alignment in cells }
  TsHorAlignment = (haDefault, haLeft, haCenter, haRight);

  {@@ Indicates vertical text alignment in cells }
  TsVertAlignment = (vaDefault, vaTop, vaCenter, vaBottom);

  {@@ Colors in fpspreadsheet are given as rgb values in little-endian notation
    (i.e. "r" is the low-value byte). The highest-value byte, if not zero,
    indicates special colors. }
  TsColor = DWord;

const
  {@@ These are some basic rgb color volues. FPSpreadsheet will support
    built-in color constants only for the EGA palette.
  }
  {@@ rgb value of <b>black</b> color, BIFF2 palette index 0, BIFF8 index 8}
  scBlack = $00000000;
  {@@ rgb value of <b>white</b> color, BIFF2 palette index 1, BIFF8 index 9 }
  scWhite = $00FFFFFF;
  {@@ rgb value of <b>red</b> color, BIFF2 palette index 2, BIFF8 index 10 }
  scRed = $000000FF;
  {@@ rgb value of <b>green</b> color, BIFF2 palette index 3, BIFF8 index 11 }
  scGreen = $0000FF00;
  {@@ rgb value of <b>blue</b> color, BIFF2 palette index 4, BIFF8 indexes 12 and 39}
  scBlue = $00FF0000;
  {@@ rgb value of <b>yellow</b> color, BIFF2 palette index 5, BIFF8 indexes 13 and 34}
  scYellow = $0000FFFF;
  {@@ rgb value of <b>magenta</b> color, BIFF2 palette index 6, BIFF8 index 14 and 33}
  scMagenta = $00FF00FF;
  {@@ rgb value of <b>cyan</b> color, BIFF2 palette index 7, BIFF8 indexes 15}
  scCyan = $00FFFF00;
  {@@ rgb value of <b>dark red</b> color, BIFF8 indexes 16 and 35}
  scDarkRed = $00000080;
  {@@ rgb value of <b>dark green</b> color, BIFF8 index 17 }
  scDarkGreen = $00008000;
  {@@ rgb value of <b>dark blue</b> color }
  scDarkBlue = $00800000;
  {@@ rgb value of <b>olive</b> color }
  scOlive = $00008080;
  {@@ rgb value of <b>purple</b> color, BIFF8 palette indexes 20 and 36 }
  scPurple = $00800080;
  {@@ rgb value of <b>teal</b> color, BIFF8 palette index 21 and 38 }
  scTeal = $00808000;
  {@@ rgb value of <b>silver</b> color }
  scSilver = $00C0C0C0;
  {@@ rgb value of <b>grey</b> color }
  scGray = $00808080;
  {@@ rgb value of <b>gray</b> color }
  scGrey = scGray;       // redefine to allow different spelling

  {@@ Identifier for not-defined color }
  scNotDefined = $40000000;
  {@@ Identifier for transparent color }
  scTransparent = $20000000;
  {@@ Identifier for palette index encoded into the TsColor }
  scPaletteIndexMask = $80000000;
  {@@ Mask for the rgb components contained in the TsColor }
  scRGBMask = $00FFFFFF;

  // aliases for LCL colors, deprecated
  scAqua = scCyan deprecated;
  scFuchsia = scMagenta deprecated;
  scLime = scGreen deprecated;
  scMaroon = scDarkRed deprecated;
  scNavy = scDarkBlue deprecated;

  { These color constants are deprecated, they will be removed in the long term }
  scPink = $00FE00FE deprecated;
  scTurquoise = scCyan deprecated;
  scGray25pct = scSilver deprecated;
  scGray50pct = scGray deprecated;
  scGray10pct = $00E6E6E6 deprecated;
  scGrey10pct = scGray10pct{%H-} deprecated;
  scGray20pct = $00CCCCCC deprecated;
  scGrey20pct = scGray20pct{%H-} deprecated;
  scPeriwinkle = $00FF9999 deprecated;
  scPlum = $00663399 deprecated;
  scIvory = $00CCFFFF deprecated;
  scLightTurquoise = $00FFFFCC deprecated;
  scDarkPurple = $00660066 deprecated;
  scCoral = $008080FF deprecated;
  scOceanBlue = $00CC6600 deprecated;
  scIceBlue = $00FFCCCC deprecated;
  scSkyBlue = $00FFCC00 deprecated;
  scLightGreen = $00CCFFCC deprecated;
  scLightYellow = $0099FFFF deprecated;
  scPaleBlue = $00FFCC99 deprecated;
  scRose = $00CC99FF deprecated;
  scLavander = $00FF99CC deprecated;
  scTan = $0099CCFF deprecated;
  scLightBlue = $00FF6633 deprecated;
  scGold = $0000CCFF deprecated;
  scLightOrange = $000099FF deprecated;
  scOrange = $000066FF deprecated;
  scBlueGray = $00996666 deprecated;
  scBlueGrey = scBlueGray{%H-} deprecated;
  scGray40pct = $00969696 deprecated;
  scDarkTeal = $00663300 deprecated;
  scSeaGreen = $00669933 deprecated;
  scVeryDarkGreen = $00003300 deprecated;
  scOliveGreen = $00003333 deprecated;
  scBrown = $00003399 deprecated;
  scIndigo = $00993333 deprecated;
  scGray80pct = $00333333 deprecated;
//  scGrey80pct = scGray80pct deprecated;
  scDarkBrown = $002D52A0 deprecated;
  scBeige = $00DCF5F5  deprecated;
  scWheat = $00B3DEF5 deprecated;

type
  {@@ Font style (redefined to avoid usage of "Graphics" }
  TsFontStyle = (fssBold, fssItalic, fssStrikeOut, fssUnderline);

  {@@ Set of font styles }
  TsFontStyles = set of TsFontStyle;

  {@@ Font position (subscript or superscript) }
  TsFontPosition = (fpNormal, fpSuperscript, fpSubscript);  // Keep order for compatibility with xls!

  {@@ Font record used in fpspreadsheet. Contains the font name, the font size
      (in points), the font style, and the font color. }
  TsFont = class
    {@@ Name of the font face, such as 'Arial' or 'Times New Roman' }
    FontName: String;
    {@@ Size of the font in points }
    Size: Single;   // in "points"
    {@@ Font style, such as bold, italics etc. - see TsFontStyle}
    Style: TsFontStyles;
    {@@ Text color given as rgb value }
    Color: TsColor;
    {@@ Text position }
    Position: TsFontPosition;
    constructor Create(AFontName: String; ASize: Single; AStyle: TsFontStyles;
      AColor: TsColor; APosition: TsFontPosition); overload;
    procedure CopyOf(AFont: TsFont);
  end;

  {@@ Parameter describing formatting of an text range in cell text }
  TsRichTextParam = record
    FirstIndex: Integer;  // 1-based utf8 character index
    FontIndex: Integer;
    HyperlinkIndex: Integer;
    {
    FontIndex: Integer;
    StartIndex: Integer;  // zero-based
    EndIndex: Integer;    // zero-based, next character!
    }
  end;

  {@@ Parameters describing formatting of text ranges in cell text }
  TsRichTextParams = array of TsRichTextParam;

  {@@ Indicates the border for a cell. If included in the CellBorders set the
      corresponding border is drawn in the style defined by the CellBorderStyle. }
  TsCellBorder = (cbNorth, cbWest, cbEast, cbSouth, cbDiagUp, cbDiagDown);

  {@@ Indicates the border for a cell }
  TsCellBorders = set of TsCellBorder;

  {@@ Line style (for cell borders) }
  TsLineStyle = (lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair,
    lsMediumDash, lsDashDot, lsMediumDashDot, lsDashDotDot, lsMediumDashDotDot,
    lsSlantDashDot);

  {@@ The Cell border style reocrd contains the linestyle and color of a cell
      border. There is a CellBorderStyle for each border. }
  TsCellBorderStyle = record
    LineStyle: TsLineStyle;
    Color: TsColor;
  end;

  {@@ The cell border styles of each cell border are collected in this array. }
  TsCellBorderStyles = array[TsCellBorder] of TsCellBorderStyle;

  {@@ Border styles for each cell border used by default: a thin, black, solid line }
const
  DEFAULT_BORDERSTYLES: TsCellBorderStyles = (
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack),
    (LineStyle: lsThin; Color: scBlack)
  );

type
  {@@ Style of fill pattern for cell backgrounds }
  TsFillStyle = (fsNoFill, fsSolidFill,
    fsGray75, fsGray50, fsGray25, fsGray12, fsGray6,
    fsStripeHor, fsStripeVert, fsStripeDiagUp, fsStripeDiagDown,
    fsThinStripeHor, fsThinStripeVert, fsThinStripeDiagUp, fsThinStripeDiagDown,
    fsHatchDiag, fsThinHatchDiag, fsThickHatchDiag, fsThinHatchHor);

  {@@ Fill pattern record }
  TsFillPattern = record
    Style: TsFillStyle;  // pattern type
    FgColor: TsColor;    // pattern color
    BgColor: TsColor;    // background color
  end;

const
  {@@ Parameters for a non-filled cell background }
  EMPTY_FILL: TsFillPattern = (
    Style: fsNoFill;
    FgColor: scTransparent;
    BgColor: scTransparent;
  );

type
  {@@ Identifier for a compare operation }
  TsCompareOperation = (coNotUsed,
    coEqual, coNotEqual, coLess, coGreater, coLessEqual, coGreaterEqual
  );

  {@@ Builtin number formats. Only uses a subset of the default formats,
      enough to be able to read/write date/time values.
      nfCustom allows to apply a format string directly. }
  TsNumberFormat = (
    // general-purpose for all numbers
    nfGeneral,
    // numbers
    nfFixed, nfFixedTh, nfExp, nfPercentage, nfFraction,
    // currency
    nfCurrency, nfCurrencyRed,
    // dates and times
    nfShortDateTime, nfShortDate, nfLongDate, nfShortTime, nfLongTime,
    nfShortTimeAM, nfLongTimeAM, nfDayMonth, nfMonthYear, nfTimeInterval,
    // text
    nfText,
    // other (format string goes directly into the file)
    nfCustom);

  {@@ Cell calculation state }
  TsCalcState = (csNotCalculated, csCalculating, csCalculated);

  {@@ Cell flag }
  TsCellFlag = (cfCalculating, cfCalculated, cfHasComment, cfHyperlink, cfMerged);

  {@@ Set of cell flags }
  TsCellFlags = set of TsCellFlag;

  {@@ Record combining a cell's row and column indexes }
  TsCellCoord = record
    Row, Col: Cardinal;
  end;

  {@@ Record combining row and column corner indexes of a range of cells }
  TsCellRange = record
    Row1, Col1, Row2, Col2: Cardinal;
  end;
  PsCellRange = ^TsCellRange;

  {@@ Array with cell ranges }
  TsCellRangeArray = array of TsCellRange;

  {@@ Record combining sheet index and row/column corner indexes of a cell range }
  TsCellRange3d = record
    Row1, Col1, Row2, Col2: Cardinal;
    Sheet1, Sheet2: Integer;
  end;

  {@@ Array of 3d cell ranges }
  TsCellRange3dArray = array of TsCellRange3d;

  {@@ Record containing limiting indexes of column or row range }
  TsRowColRange = record
    FirstIndex, LastIndex: Cardinal;
  end;

  {@@ Options for sorting }
  TsSortOption = (ssoDescending, ssoCaseInsensitive);
  {@@ Set of options for sorting }
  TsSortOptions = set of TsSortOption;

  {@@ Sort priority }
  TsSortPriority = (spNumAlpha, spAlphaNum);   // spNumAlpha: Number < Text

  {@@ Sort key: sorted column or row index and sort direction }
  TsSortKey = record
    ColRowIndex: Integer;
    Options: TsSortOptions;
  end;

  {@@ Array of sort keys for multiple sorting criteria }
  TsSortKeys = array of TsSortKey;

  {@@ Complete set of sorting parameters
    @param SortByCols  If true sorting is top-down, otherwise left-right
    @param Priority    Determines whether numbers are before or after text.
    @param SortKeys    Array of sorting indexes and sorting directions }
  TsSortParams = record
    SortByCols: Boolean;
    Priority: TsSortPriority;
    Keys: TsSortKeys;
  end;

  {@@ Switch a cell from left-to-right to right-to-left orientation }
  TsBiDiMode = (bdDefault, bdLTR, bdRTL);

  {@@ Record containing all details for cell formatting }
  TsCellFormat = record
    Name: String;
    ID: Integer;
    UsedFormattingFields: TsUsedFormattingFields;
    FontIndex: Integer;
    TextRotation: TsTextRotation;
    HorAlignment: TsHorAlignment;
    VertAlignment: TsVertAlignment;
    Border: TsCellBorders;
    BorderStyles: TsCelLBorderStyles;
    Background: TsFillPattern;
    NumberFormatIndex: Integer;
    BiDiMode: TsBiDiMode;
    // next two are deprecated...
    NumberFormat: TsNumberFormat;
    NumberFormatStr: String;
  end;

  {@@ Pointer to a format record }
  PsCellFormat = ^TsCellFormat;

  {@@ Cell structure for TsWorksheet
      The cell record contains information on the location of the cell (row and
      column index), on the value contained (number, date, text, ...), on
      formatting, etc.

      Never suppose that all *Value fields are valid,
      only one of the ContentTypes is valid. For other fields
      use TWorksheet.ReadAsUTF8Text and similar methods

      @see ReadAsUTF8Text }
  TCell = record
    { Location of the cell }
    Row: Cardinal; // zero-based
    Col: Cardinal; // zero-based
    Worksheet: Pointer;   // Must be cast to TsWorksheet when used  (avoids circular unit reference)
    { Status flags }
    Flags: TsCellFlags;
    { Index of format record in the workbook's FCellFormatList }
    FormatIndex: Integer;
    { Cell content }
    UTF8StringValue: String;   // Strings cannot be part of a variant record
    RichTextParams: TsRichTextParams; // Formatting of individual text ranges
    FormulaValue: String;      // Formula for calculation of cell content
    case ContentType: TCellContentType of  // variant part must be at the end
      cctEmpty      : ();      // has no data at all
      cctFormula    : ();      // FormulaValue is outside the variant record
      cctNumber     : (Numbervalue: Double);
      cctUTF8String : ();      // UTF8StringValue is outside the variant record
      cctDateTime   : (DateTimeValue: TDateTime);
      cctBool       : (BoolValue: boolean);
      cctError      : (ErrorValue: TsErrorValue);
  end;

  {@@ Pointer to a TCell record }
  PCell = ^TCell;

  {@@ Embedded image }
  TsImage = record
    Row, Col: Cardinal;       // cell for top/left edge of the image (anchor)
    Index: Integer;           // index into the workbook's embedded streams list
    OffsetX, OffsetY: Double; // mm, relative to anchor
    ScaleX, ScaleY: Double;   // scaling factor of image
  end;
  PsImage = ^TsImage;

  {@@ Image embedded in header or footer}
  TsHeaderFooterImage = record
    Index: Integer;           // index into the workbook's embedded streams list
  end;

  {@@ Page orientation for printing }
  TsPageOrientation = (spoPortrait, spoLandscape);

  {@@ Options for the print layout records }
  TsPrintOption = (poPrintGridLines, poPrintHeaders, poPrintPagesByRows,
    poMonochrome, poDraftQuality, poPrintCellComments, poDefaultOrientation,
    poUseStartPageNumber, poCommentsAtEnd, poHorCentered, poVertCentered,
    poDifferentOddEven, poDifferentFirst, poFitPages);

  {@@ Set of options used by the page layout }
  TsPrintOptions = set of TsPrintOption;

  {@@ Headers and footers are divided into three parts: left, center and right }
  TsHeaderFooterSectionIndex = (hfsLeft, hfsCenter, hfsRight);

  {@@ Array with all possible images in a header or a footer }
  TsHeaderFooterImages = array[TsHeaderFooterSectionIndex] of TsHeaderFooterImage;

const
  {@@ Indexes to be used for the various headers and footers }
  HEADER_FOOTER_INDEX_FIRST   = 0;
  HEADER_FOOTER_INDEX_ODD     = 1;
  HEADER_FOOTER_INDEX_EVEN    = 2;
  HEADER_FOOTER_INDEX_ALL     = 1;


type
  {@@ Search option }
  TsSearchOption = (soCompareEntireCell, soMatchCase, soRegularExpr, soAlongRows,
    soBackward, soWrapDocument, soEntireDocument);

  {@@ A set of search options }
  TsSearchOptions = set of TsSearchOption;

  {@@ Defines which part of document is scanned }
  TsSearchWithin = (swWorkbook, swWorksheet, swColumn, swRow);

  {@@ Search parameters }
  TsSearchParams = record
    SearchText: String;
    Options: TsSearchOptions;
    Within: TsSearchWithin;
  end;

  {@@ Replace option }
  TsReplaceOption = (roReplaceEntirecell, roReplaceAll, roConfirm);

  {@@ A set of replace options }
  TsReplaceOptions = set of TsReplaceOption;

  {@@ Replace parameters }
  TsReplaceParams = record
    ReplaceText: String;
    Options: TsReplaceOptions;
  end;

  {@@ Identifier for a copy operation }
  TsCopyOperation = (coNone, coCopyFormat, coCopyValue, coCopyFormula, coCopyCell);

  {@@ Parameters for stream access }
  TsStreamParam = (spClipboard, spWindowsClipboardHTML);
  TsStreamParams = set of TsStreamParam;

  {@@ Types of row heights
    rhtDefault - default row height
    rhtAuto - automatically determined row height, depends on font size,
      text rotation, rich-text parameters, word-wrap
    rhtCustom - user-determined row height (dragging the row header borders in
      the grid }
  TsRowHeightType = (rhtDefault, rhtAuto, rhtCustom);

const
  RowHeightTypeNames: array[TsRowHeightType] of string = (
    'Default', 'Auto', 'Custom');


implementation

constructor TsFont.Create(AFontName: String; ASize: Single; AStyle: TsFontStyles;
  AColor: TsColor; APosition: TsFontPosition);
begin
  FontName := AFontName;
  Size := ASize;
  Style := AStyle;
  Color := AColor;
  Position := APosition;
end;

procedure TsFont.CopyOf(AFont: TsFont);
begin
  FontName := AFont.FontName;
  Size := AFont.Size;
  Style := AFont.Style;
  Color := AFont.Color;
  Position := AFont.Position;
end;

end.

