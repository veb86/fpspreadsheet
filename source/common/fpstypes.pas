{@@ ----------------------------------------------------------------------------
  Unit **fpsTypes** collects the most **fundamental declarations** used
  throughout the fpspreadsheet library. It is very likey that this unit must
  be added to the uses clause of the application.

  AUTHORS: Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.
-------------------------------------------------------------------------------}
unit fpsTypes;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

{$include fps.inc}

interface

uses
  Classes, SysUtils, fpimage;

{$IF FPC_FullVersion < 30000}
{@@ This string type is not re-encoded by FPC. It is a standard type of FPC 3.0+,
  its declaration must be repeated here in order to keep fpSpreadsheet usable by
  older FPC versions. }
type
  RawByteString = ansistring;
{$ENDIF}

type
  { Forward declarations }
  TsBasicWorksheet = class;
  TsBasicWorkbook = class;

  {@@ Built-in file formats of fpspreadsheet 
    @value sfExcel2    File format of Excel 2.1
    @value sfExcel5    File format of Excel 5
    @value sfExcel8    File format of Excel 97
    @value sfExcelXML  XML file format of Excel 2003
    @value sfOOXML     Default file format of Excel 2007 and later
    @value sfOpenDocument  File format of LibreOffice/OpenOffice Calc
    @value sfCSV       Comma-separated text file
    @value sfHTML      HTML file format
    @value sfWikiTable_Pipes  Wiki table file format
    @value sfWikiTable_WikiMedia  Wiki media file format 
    @value sfUser      User-defined file format. The user must provide the reader and writer classes. }
  TsSpreadsheetFormat = (sfExcel2, sfExcel5, sfExcel8, sfExcelXML, sfOOXML,
    sfOpenDocument, sfCSV, sfHTML, sfWikiTable_Pipes, sfWikiTable_WikiMedia,
    sfUser);   // Use this for user-defined readers/writers

  {@@ Numerical identifier for file formats, built-in and user-provided }
  TsSpreadFormatID = integer;

  {@@ Array of file format identifiers }
  TsSpreadFormatIDArray = array of TsSpreadFormatID;

const
  {@@ Format identifier of an undefined, unknown, etc. file format. 
  
    Each unit implementing a reader/writer will define an sfidXXXX value as a
    numerical identifer of the file format. In case of the built-in formats,
    the identifier is equal to the ord of the TsSpreadsheetFormat value. }
  sfidUnknown = -1;

type
  {@@ Flag set during reading or writing of a workbook 
    @value  rwfNormal   The workbook is in normal read/write state, i.e. it is currently neither being read nor being written.
    @value  rwfRead     The workbook is currently being read.
    @value  rwfWrite    The workbook is currently being written. }
  TsReadWriteFlag = (rwfNormal, rwfRead, rwfWrite);

  {@@ This record collect limitations of a particular file format.
    @member  MaxRowCount         Gives the maximum number of rows supported by this file format.
    @member  MaxColCount         Gives the maximum number of columns supported by this file format.
    @member  MaxPaletteSize      Gives the maximum count of color palette entries supported by this file format  
    @member  MaxSheetNameLength  Defines the maximum length of the worksheet name supported by this file format.
    @member  MaxCharsInTextCell  Defines the maximum length of the text in a cell supported by this file format.    }
  TsSpreadsheetFormatLimitations = record
    MaxRowCount: Cardinal;
    MaxColCount: Cardinal;
    MaxPaletteSize: Integer;
    MaxSheetNameLength: Integer;
    MaxCharsInTextCell: Integer;
  end;

  {@@ Line ending used in CSV files }
  TsCSVLineEnding = (leSystem, leCRLF, leCR, leLF);

  {@@ Parameters controlling reading from and writing to CSV files
    @member  SheetIndex     Index of the sheet to be written (write-only)
    @member  LineEnding     Specification for the line endings to be written (write-only)
    @member  Delimiter      Column delimiter (read/write)
    @member  QuoteChar      Character used for quoting text in special cases (read/write)
    @member  Encoding       String identifying the endoding of the file, such as 'utf8', 'cp1252' etc (read/write)
    @member  DetectContentType  Try to convert strings to their content type (read-only)
    @member  NumberFormat   If empty numbers are written like in worksheet, otherwise this format string is applied (write-only)
    @member  AutoDetectNumberFormat  Try to detect the decimal and thousand separator used in numbers (read-only)
    @member  TrueText       String for boolean @true (read/write)
    @member  FalseText      String for boolean @false (read/write)
    @member  FormatSettings Additional parameter for string conversion (read/write) }
  TsCSVParams = record   // W = writing, R = reading, RW = reading/writing
    SheetIndex: Integer;             // W: Index of the sheet to be written
    LineEnding: TsCSVLineEnding;     // W: Specification for line ending to be written
    Delimiter: Char;                 // RW: Column delimiter
    QuoteChar: Char;                 // RW: Character for quoting texts
    Encoding: String;                // RW: Encoding of file (code page, such as "utf8", "cp1252" etc)
    DetectContentType: Boolean;      // R: try to convert strings to content types
    NumberFormat: String;            // W: if empty write numbers like in sheet, otherwise use this format
    AutoDetectNumberFormat: Boolean; // R: automatically detects decimal/thousand separator used in numbers
    TrueText: String;                // RW: String for boolean TRUE
    FalseText: String;               // RW: String for boolean FALSE
    FormatSettings: TFormatSettings; // RW: add'l parameters for conversion
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

  {@@ Default binary _Excel_ file extension (<= Excel 97) }
  STR_EXCEL_EXTENSION = '.xls';
  {@@ Default xml _Excel_ file extension (Excel XP, 2003) }
  STR_XML_EXCEL_EXTENSION = '.xml';
  {@@ Default xml _Excel_ file extension (>= Excel 2007) }
  STR_OOXML_EXCEL_EXTENSION = '.xlsx';
  {@@ Default _OpenDocument_ spreadsheet file extension }
  STR_OPENDOCUMENT_CALC_EXTENSION = '.ods';
  {@@ Default extension of _comma-separated-values_ file }
  STR_COMMA_SEPARATED_EXTENSION = '.csv';
  {@@ Default extension for _HTML_ files }
  STR_HTML_EXTENSION = '.html';
  {@@ Default extension of _wikitable files_ in _pipes_ format}
  STR_WIKITABLE_PIPES_EXTENSION = '.wikitable_pipes';
  {@@ Default extension of _wikitable files_ in _wikimedia_ format }
  STR_WIKITABLE_WIKIMEDIA_EXTENSION = '.wikitable_wikimedia';

  {@@ String for boolean value @TRUE }
  STR_TRUE = 'TRUE';
  {@@ String for boolean value @FALSE }
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

var
  CSVParams: TsCSVParams = (
    SheetIndex: 0;
    LineEnding: leSystem;
    Delimiter: ';';
    QuoteChar: '"';
    Encoding: '';    // '' = auto-detect when reading, UTF8 when writing
    DetectContentType: true;
    NumberFormat: '';
    AutoDetectNumberFormat: true;
    TrueText: 'TRUE';
    FalseText: 'FALSE';
  {%H-});

type
  {@@ Units for size dimensions 
    @value  suChars         Horizontal size is given as the count of '0' characters in the default font (not very exact)
    @value  suLines         Vertical size is given as the count of lines measured by the height of the default font (not very exact)
    @value  suMillimeters   Size is given in millimeters
    @value  suCentimeters   Size is given in centimeters
    @value  suPoints        Size is given in points
    @value  suInches        Size is given in inches
    
  Adjustment of the @link(fpsutils.ScreenPixelsPerInch) is required for accurate
  values in case of high-resolution monitors. }
  TsSizeUnits = (suChars, suLines, suMillimeters, suCentimeters, suPoints, suInches);

const
  {@@ Names of the size units }
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
  {@@ Tokens to identify the @bold(elements in an expanded RPN formula).

   @note(When adding or rearranging items
   * make sure that the subtypes TOperandTokens and TBasicOperationTokens
     are complete
   * make sure to keep the table "TokenIDs" in unit xlscommon in sync)
  }
  TFEKind = (
    { Basic operands }
    fekCell, fekCellRef, fekCellRange, fekCellOffset,
    fekCell3d, fekCellRef3d, fekCellRange3d,
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
  {@@ Flags to mark the address or a cell or a range of cells to be @bold(absolute) or @bold(relative). They are used in the set @link(TsRelFlags). 
    @value  rfRelRow  Signals that the row reference is relative
    @value  rfRelCol  Signals that the column reference is relative
    @value  rfRelRow2 Signals that the reference to the last row of a cell block is relative
    @value  rfRelCol2 Signals that the reference to the last column of a cell block is relative}
  TsRelFlag = (rfRelRow, rfRelCol, rfRelRow2, rfRelCol2);

  {@@ Flags to mark the address of a cell or a range of cells to be @bold(absolute)
      or @bold(relative). It is a set consisting of @link(TsRelFlag) elements. }
  TsRelFlags = set of TsRelFlag;

const
  {@@ Abbreviation of all-relative cell reference flags (@seeAlso(TsRelFlag))}
  rfAllRel = [rfRelRow, rfRelCol, rfRelRow2, rfRelCol2];

  {@@ Separator between worksheet name and cell (range) reference in an address }
  SHEETSEPARATOR = '!';

type
  {@@ Flag indicating the calculation state of a formula.
    @value ffCalculating  The formula is currently being calculated.
    @value ffCalculated   The calculation of the formula is completed. }
  TsFormulaFlag = (ffCalculating, ffCalculated);
  
  {@@ Set of formula calculation state flags, @link(TsFormulaFlag) }
  TsFormulaFlags = set of TsFormulaFlag;

  {@@ Elements of an expanded RPN formula.
    @member  ElementKind  Identifies the type of the formula element, @seeAlso(TFEKind)
    @member  Row          Row index of the cell to which this formula refers (zero-based)
    @member  Row2         Row index of the last cell in the cell range to which this formula refers (zero-based)
    @member  Col          Column index of the cell to which this formula refers (zero-based)
    @member  Col2         Column index of the last cell in the cell range to which this formula refers (zero-based)
    @member  Sheet        Index of the worksheet to which this formula refers (zero-based)
    @member  Sheet2       Index of the last worksheet in the 3d cell range to which this formula refers (zero-based)
    @member  SheetNames   Tab-separated list of the worksheet names refered by this formula (intermediate use only).
    @member  DoubleValue  Floating point value assigned to this formula element
    @member  IntValue     Integer value assigned to this formula element
    @member  Stringvalue  String value assigned to this formula element
    @member  RelFlags     Information about relative/absolute cell addresses, @seeAlso(TsRelFlags)
    @member  FuncName     Name of the function called by this formula element
    @member  ParamsNum    Count of parameters needed by this formula element
    
    @note(If ElementKind is fekCellOffset, "Row" and "Col" have to be cast to signed integers!) }
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

  {@@ Formula dialect 
    @value  fdExcelA1       Default A1 syntax of Excel cell references: Cells are identified by column letters ('A', 'B', ...) and 1-based row numbers ('1', '2') 
    @value  fdExcelR1C1     R1C1 syntax of excel 
    @value  fdOpenDocument  Syntax of OpenOffice/LibreOffice Calc.
    @value  fdLocalized     The formula uses localized format settings in A1 syntax. }
  TsFormulaDialect = (fdExcelA1, fdExcelR1C1, fdOpenDocument, fdLocalized);

  {@@ Describes the **type of content** in a cell of a @link(TsWorksheet)
    @value  cctEmpty       The cell is empty, however, can carry a format.
    @value  cctFormula     The cell contains a formula which has not yet been calculated
    @value  cctNumber      The cell contains a number (integer or float)
    @value  cctUTF8String  The cell contains a string which is considered to be UTF8-encoded.
    @value  cctDateTime    The cell contains a date/time value.
    @value  cctBool        The cell contains a boolean value
    @value  cctError       The cell contains an error value

  @seeAlso TsErrorValue
  }
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
    @param   Target       Target of hyperlink: URI of file, web link, mail; or: internal link (# followed by cell address)
    @param   Tooltip      Text displayed as a popup hint by Excel }
  TsHyperlink = record
    Row, Col: Cardinal;
    Target: String;
    Tooltip: String;
  end;

  {@@ Pointer to a @link(TsHyperlink) record }
  PsHyperlink = ^TsHyperlink;

  {@@ Callback function, e.g. for iterating the internal AVL trees of the workbook/sheet}
  TsCallback = procedure (data, arg: Pointer) of object;

  {@@ Error code values 
    @value  errOK                  ok, no error
    @value  errEmptyIntersection   A space was used in formulas that reference multiple ranges; a comma separates range references (#NULL!)
    @value  errDivideByZero        Trying to divide by zero (#DIV/0!)
    @value  errWrongType           The wrong type of operand or function argument is used (#VALUE!)
    @value  errIllegalRef          A reference is invalid (#REF!)
    @value  errWrongName           Text in the formula is not recognized (#NAME?)
    @value  errOverflow            A formula has invalid numeric data for the type of operation (#NUM!)
    @value  errArgError            A formula or a function inside a formula cannot find the referenced data (#N/A)
    @value  errFormulaNotSupported This formula is not suppored by fpSpreadsheet (error code not used by Excel and Calc)   }
  TsErrorValue = (
    errOK,                 // no error
    errEmptyIntersection,  // #NULL!
    errDivideByZero,       // #DIV/0!
    errWrongType,          // #VALUE!
    errIllegalRef,         // #REF!
    errWrongName,          // #NAME?
    errOverflow,           // #NUM!
    errArgError,           // #N/A  ( = #NV in German )
    errFormulaNotSupported // No excel error
  );

  {@@ List of possible formatting fields 
    @value uffTextRotation  The cell format supports text rotation.
    @value uffFont          The cell format supports using font different from default. 
    @value uffBorder        The cell format supports decorating the cell with a border. 
    @value uffBackground    The cell format supports a dedicated background color and fill style.
    @value uffNumberFormat  The cell format supports individual number formats of numeric cell values.
    @value uffWordwrap      The cell format supports wrapping of long cell text into new lines.
    @value uffHorAlign      The cell format supports horizontal text alignment.
    @value uffVertAlign     The cell format supports vertical text alignment
    @value uffBiDi          The cell format supports right-to-left text display.
    @value uffProtection    The cell format supports locking of cells. }
  TsUsedFormattingField = (uffTextRotation, uffFont, uffBorder, uffBackground,
    uffNumberFormat, uffWordWrap, uffHorAlign, uffVertAlign, uffBiDi,
    uffProtection
  );
  { NOTE: "uffBackgroundColor" of older versions replaced by "uffBackground" }

  {@@ Describes which formatting fields are active (see @link(TsUsedFormattingField)). }
  TsUsedFormattingFields = set of TsUsedFormattingField;

  {$IFDEF NO_RAWBYTESTRING}
  RawByteString = ansistring;
  {$ENDIF}

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
      @preformatted(
       --->
       ABC)

      So 90 degrees clockwise means that the text will be:
      @preformatted(
       |  A
       |  B
       v  C)

      And 90 degree counter clockwise will be:
      @preformatted(
       ^  C
       |  B
       |  A)

      Due to limitations of the text mode the characters are not rotated here.
      There is, however, also a "stacked" variant which looks exactly like
      the 90-degrees-clockwise case.
      
   @value trHorizontal  Text is written horizontally as usual.
   @value rt90DegreeClockwiseRotation  Text is written vertically from top to bottom
   @value rt90DegreeCounterClockwiseRotation  Text is written vertically from bottom to top
   @value rtStacked  Text is written vertically from top to bottom, but with horizontal chararacters.
  }
  TsTextRotation = (trHorizontal, rt90DegreeClockwiseRotation,
    rt90DegreeCounterClockwiseRotation, rtStacked);

  {@@ Indicates horizontal text alignment in cells 
    @value  haDefault  Default text alignment (left for alphanumeric, right for numbers and dates)
    @value  haLeft     Left-aligned cell text
    @value  haCenter   Centered cell text
    @value  haRight    Right-aligned cell text }
  TsHorAlignment = (haDefault, haLeft, haCenter, haRight);

  {@@ Indicates vertical text alignment in cells 
    @value  vaDefault  Default vertical alignment (bottom)
    @value  vaTop      Top-aligned cell text
    @value  vaCenter   Vertically centered cell text
    @value  vaBottom   Bottom-aligned cell text }
  TsVertAlignment = (vaDefault, vaTop, vaCenter, vaBottom);

  {@@ Colors in fpspreadsheet are given as rgb values in little-endian notation
    (i.e. "r" is the low-value byte). The highest-value byte, if not zero,
    indicates special colors. 
    
    @note(This byte order in TsColor is opposite to that in HTML colors.) }
  TsColor = DWord;

const
  {@@ These are some basic rgb color volues. FPSpreadsheet will support
    only those built-in color constants originating in the EGA palette.
  }
  {@@ rgb value of @bold(black) color, BIFF2 palette index 0, BIFF8 index 8}
  scBlack = $00000000;
  {@@ rgb value of @bold(white) color, BIFF2 palette index 1, BIFF8 index 9 }
  scWhite = $00FFFFFF;
  {@@ rgb value of @bold(red) color, BIFF2 palette index 2, BIFF8 index 10 }
  scRed = $000000FF;
  {@@ rgb value of @bold(green) color, BIFF2 palette index 3, BIFF8 index 11 }
  scGreen = $0000FF00;
  {@@ rgb value of @bold(blue) color, BIFF2 palette index 4, BIFF8 indexes 12 and 39}
  scBlue = $00FF0000;
  {@@ rgb value of @bold(yellow) color, BIFF2 palette index 5, BIFF8 indexes 13 and 34}
  scYellow = $0000FFFF;
  {@@ rgb value of @bold(magenta) color, BIFF2 palette index 6, BIFF8 index 14 and 33}
  scMagenta = $00FF00FF;
  {@@ rgb value of @bold(cyan) color, BIFF2 palette index 7, BIFF8 indexes 15}
  scCyan = $00FFFF00;
  {@@ rgb value of @bold(dark red) color, BIFF8 indexes 16 and 35}
  scDarkRed = $00000080;
  {@@ rgb value of @bold(dark green) color, BIFF8 index 17 }
  scDarkGreen = $00008000;
  {@@ rgb value of @bold(dark blue) color }
  scDarkBlue = $00800000;
  {@@ rgb value of @bold(olive) color }
  scOlive = $00008080;
  {@@ rgb value of @bold(purple) color, BIFF8 palette indexes 20 and 36 }
  scPurple = $00800080;
  {@@ rgb value of @bold(teal) color, BIFF8 palette index 21 and 38 }
  scTeal = $00808000;
  {@@ rgb value of @bold(silver) color }
  scSilver = $00C0C0C0;
  {@@ rgb value of @bold(grey) color }
  scGray = $00808080;
  {@@ rgb value of @bold(gray) color }
  scGrey = scGray;       // redefine to allow different spelling

  {@@ Identifier for not-defined color }
  scNotDefined = $40000000;
  {@@ Identifier for @bold(transparent) color }
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
  {@@ Font style (redefined to avoid usage of graphics unit) 
    @value fssBold       Bold text
    @value fssItalic     Italic text
    @value fssStrikeOut  Strike-through text (there is only a single strike-out line)
    @value fssUnderLine  Underlined text (there is only a single underline) }
  TsFontStyle = (fssBold, fssItalic, fssStrikeOut, fssUnderline);

  {@@ Set of font styles }
  TsFontStyles = set of TsFontStyle;

  {@@ Font position (subscript or superscript) 
    @value fpNormal      Normals character position
    @value fpSuperscript Superscripted characters, like in @code('10²') 
    @value fpSubscript   Subscripted characters}
  TsFontPosition = (fpNormal, fpSuperscript, fpSubscript);  // Keep order for compatibility with xls!

  {@@ Font record used in fpspreadsheet. Contains the font name, the font size
      (in points), the font style, and the font color. }
  TsFont = class
    {@@ Name of the font face, such as 'Arial' or 'Times New Roman' }
    FontName: String;
    {@@ Size of the font, in points }
    Size: Single;   // in "points"
    {@@ Font style, such as bold, italics etc. - see @link(TsFontStyle)}
    Style: TsFontStyles;
    {@@ Text color given as rgb value }
    Color: TsColor;
    {@@ Text position }
    Position: TsFontPosition;
    constructor Create(AFontName: String; ASize: Single; AStyle: TsFontStyles;
      AColor: TsColor; APosition: TsFontPosition); overload;
    procedure CopyOf(AFont: TsFont);
  end;

  {@@ Array of font records }
  TsFontArray = array of TsFont;

  {@@ Parameter describing formatting of an text range in cell text 
  
    @member  FirstIndex      One-based index of the utf8 code-point ("character")
                             at which the FontIndex should be used.
    @member  FontIndex       Index of the font in the workbook's font list to
                             be used for painting the characters following the
                             FirstIndex.
    @member  HyperlinkIndex  Index of the hyperlink assigned to the text following
                             the character at FirstIndex.
    }
  TsRichTextParam = record
    FirstIndex: Integer; 
    FontIndex: Integer;
    HyperlinkIndex: Integer;
  end;

  {@@ Parameters describing formatting of text ranges in cell text }
  TsRichTextParams = array of TsRichTextParam;

  {@@ Indicates the border for a cell. If included in the CellBorders set 
      the corresponding border is drawn in the style defined by 
      the CellBorderStyle. 
      
     @value cbNorth    Horizontal border line at the top of the cell
     @value cbWest     Vertical border line at the left of the cell
     @value cbEast     Vertical border line at the right of the cell
     @value cbSouth    Horizontal border line at the bottom of the cell
     @value cbDiagUp   Diagonal border line running from the bottom-left to the top-right corner of the cell (/)
     @value cbDiagDown Diagonal border line running from the top-left to the bottom-right corner of the cell (\)  
     
     @seeAlso(TsCellBorderStyle)}
  TsCellBorder = (cbNorth, cbWest, cbEast, cbSouth, cbDiagUp, cbDiagDown);

  {@@ Indicates the border for a cell }
  TsCellBorders = set of TsCellBorder;

  {@@ Line style (for cell borders) 
    @value lsThin             Thin solid line
    @value lsMedium           Medium thick solid line
    @value lsDashed           Dashed line, thin
    @value lsDotted           Dotted line, thin
    @value lsThick            Very thick solid line
    @value lsDouble           Double line (solid)
    @value lsHair             Very thin line
    @value lsMediumDash       Dashed line, medium thickness
    @value lsDashDot          Dash-dot line, thin
    @value lsMediumDashDot    Dash-dot line, medium thickness
    @value lsDashDotDot       Dash-dot-dot line, thin
    @value lsMediumDashDotDot Dash-dot-dot line, medium thickness
    @value lsSlantDashDot     Dash-dot line, slanted segment ends    }
  TsLineStyle = (lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair,
    lsMediumDash, lsDashDot, lsMediumDashDot, lsDashDotDot, lsMediumDashDotDot,
    lsSlantDashDot);

  {@@ The Cell border style reocrd contains the linestyle and color of a cell
      border. There is a CellBorderStyle for each border. 
      
    @member  LineStyle   LineStyle to be used for this cell border. See @link(TsLineStyle).
    @member  Color       Color to be used for this  cell border. See @link(TsColor).
   }
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

  {@@ Border style to be used for "no border"}
  NO_CELL_BORDER: TsCellBorderStyle = (LineStyle: lsThin; Color: scNotDefined);

  {@@ Default border style in which all borders are used. }
  ALL_BORDERS: TsCellBorders = [cbNorth, cbEast, cbSouth, cbWest];

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
    BgColor: TsColor;    // background color (undefined when Style=fsSolidFill)
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

  {@@ Cell calculation state 
  
    @value  csNotCalculated   The cell formula has not yet been calculated.
    @value  csCalculating     The flag indicates that the cell is currently 
                              being calculated.
    @value  csCalculated      The cell formula has been calculated, and the 
                              result is stored in the cell's data fields. }
  TsCalcState = (csNotCalculated, csCalculating, csCalculated);

  {@@ Cell flag providing particular information about the state of a cell
    @value cfHasComment    The cell has a comment record
    @value cfHyperlink     The cell has a hyperlink record
    @value cfMerged        The cell belongs to a block of merged cells
    @value cfHasFormula    The cell has a formula.
    @value cf3dFormula     The cell formula links to other worksheets. }
  TsCellFlag = (cfHasComment, cfHyperlink, cfMerged, cfHasFormula, cf3dFormula);

  {@@ Set of cell flags }
  TsCellFlags = set of TsCellFlag;

  {@@ Record combining a cell's row and column indexes 
    @member   Row   Row index of the cell (0-based)
    @member   Col   Column index of the cell (0-based) }
  TsCellCoord = record
    Row, Col: Cardinal;
  end;

  {@@ Record combining row and column corner indexes of a range of cells 
    @member  Row1    The index of the top row of the cell block (0-based)
    @member  Col1    The index of the left column of the cell block (0-based)
    @member  Row2    The index of the bottom row of the cell block (0-based)
    @member  Col2    The index of the right column of the cell block (0-based)  }
  TsCellRange = record
    Row1, Col1, Row2, Col2: Cardinal;
  end;
  {@@ Pointer to a @link(TsCellRange) record }
  PsCellRange = ^TsCellRange;

  {@@ Array with cell ranges }
  TsCellRangeArray = array of TsCellRange;

  {@@ Record combining sheet index and row/column corner indexes of a 3d cell range 
    @member  Row1    The index of the top row of the 3d cell block (0-based)
    @member  Col1    The index of the left column of the 3d cell block (0-based)
    @member  Row2    The index of the bottom row of the 3d cell block (0-based)
    @member  Col2    The index of the right column of the 3d cell block (0-based)
    @member  Sheet1  The index of the first sheet of the 3d cell block (0-based)
    @member  Sheet2  The index of the last sheet of the 3d cell block (0-based)}
  TsCellRange3d = record
    Row1, Col1, Row2, Col2: Cardinal;
    Sheet1, Sheet2: Integer;
  end;

  {@@ Array of 3d cell ranges }
  TsCellRange3dArray = array of TsCellRange3d;

  {@@ Record containing limiting indexes of column or row range 
    @member  FirstIndex   Index of the first column/row of a range of columns/rows 
    @member  LastIndex    Index of the last column/row of a range of columns/rows }
  TsRowColRange = record
    FirstIndex, LastIndex: Cardinal;
  end;

  {@@ Options for sorting 
    @value  ssoDescending       Sort values in descending order
    @value  ssoCaseInsensitive  Ignore character case for sorting
    @value  ssoAlphaBeforeNum   Sort alphanumeric characters to be before 
                                numeric characters}
  TsSortOption = (ssoDescending, ssoCaseInsensitive, ssoAlphaBeforeNum);
  {@@ Set of options for sorting }
  TsSortOptions = set of TsSortOption;

  {@@ Sort priority 
    @value  spNumAlpha  Numbers are considered to be "smaller" than Alpha-Text, 
                        i.e. will be put before text
    @value  spAlphaNum  Numbers are considered to be "larger" than Alpha-Text,
                        i.e. will be sorted after text. }
  TsSortPriority = (spNumAlpha, spAlphaNum);   // spNumAlpha: Number < Text

  {@@ Sort key: parameters for sorting
    @member  ColRowIndex    Index of the sorted column or row
    @member  Options        Options used for sorting)
    
    @seeAlso TsSortOption }
  TsSortKey = record
    ColRowIndex: Integer;
    Options: TsSortOptions;
  end;

  {@@ Array of sort keys for multiple sorting criteria }
  TsSortKeys = array of TsSortKey;

  {@@ Complete set of sorting parameters
    @member SortByCols  If true sorting is top-down, otherwise left-right
    @member Priority    Determines whether numbers are before or after text.
    @member Keys        Array of sorting col/row indexes and sorting directions }
  TsSortParams = record
    SortByCols: Boolean;
    Priority: TsSortPriority;
    Keys: TsSortKeys;
  end;

  {@@ Switch a cell from left-to-right to right-to-left orientation }
  TsBiDiMode = (bdDefault, bdLTR, bdRTL);

  {@@ Algorithm used for encryption/decryption }
  TsCryptoAlgorithm = (caUnknown,
    caExcel,    // Excel <= 2010
    caMD2, caMD4, caMD5, caRIPEMD128, caRIPEMD160,
    caSHA1, caSHA256, caSHA384, caSHA512,
    caWHIRLPOOL
    );

  {@@ Record collection information for encryption/decryption }
  TsCryptoInfo = record
    PasswordHash: String;
    Algorithm: TsCryptoAlgorithm;
    SaltValue: string;
    SpinCount: Integer;
  end;

  {@@ Workbook protection options }
  TsWorkbookProtection = (bpLockRevision, bpLockStructure, bpLockWindows);
  TsWorkbookProtections = set of TsWorkbookProtection;

  {@@ Worksheet protection options. All used items are locked. }
  TsWorksheetProtection = (
    spFormatCells, spFormatColumns, spFormatRows,
    spDeleteColumns, spDeleteRows,
    spInsertColumns, spInsertRows, spInsertHyperlinks,
    spCells, spSort, spObjects,
    spSelectLockedCells, spSelectUnlockedCells
    {spPivotTables, spScenarios }
  );
  TsWorksheetProtections = set of TsWorksheetProtection;

  {@@ Cell protection options }
  TsCellProtection = (cpLockCell, cpHideFormulas);
  TsCellProtections = set of TsCellProtection;

const     // all this actions are FORBIDDEN is included and ALLOWED of excluded!
  ALL_SHEET_PROTECTIONS = [spFormatCells, spFormatColumns, spFormatRows,
    spDeleteColumns, spDeleteRows, spInsertColumns, spInsertRows, spInsertHyperlinks,
    spCells, spSort, spObjects, spSelectLockedCells, spSelectUnlockedCells
    {spPivotTables, spScenarios} ];

  DEFAULT_SHEET_PROTECTION = ALL_SHEET_PROTECTIONS - [spSelectLockedCells, spSelectUnlockedcells];

  DEFAULT_CELL_PROTECTION = [cpLockCell];

type
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
    BorderStyles: TsCellBorderStyles;
    Background: TsFillPattern;
    NumberFormatIndex: Integer;
    BiDiMode: TsBiDiMode;
    Protection: TsCellProtections;
    // next two are deprecated...
    NumberFormat: TsNumberFormat;
    NumberFormatStr: String;
    procedure SetBackground(AFillStyle: TsFillStyle; AFgColor, ABgColor: TsColor);
    procedure SetBackgroundColor(AColor: TsColor);
    procedure SetBorders(ABorders: TsCellBorders;
      const AColor: TsColor = scBlack; const ALineStyle: TsLineStyle = lsThin);
    procedure SetFont(AFontIndex: Integer);
    procedure SetHorAlignment(AHorAlign: TsHorAlignment);
    procedure SetNumberFormat(AIndex: Integer);
    procedure SetTextRotation(ARotation: TsTextRotation);
    procedure SetVertAlignment(AVertAlign: TsVertAlignment);
  end;
  {@@ Pointer to a @link(TsCellFormat) record }
  PsCellFormat = ^TsCellFormat;

  {@@ Cell structure for TsWorksheet
      The cell record contains information on the location of the cell (row and
      column index), on the value contained (number, date, text, ...), on
      formatting, etc.

      Never suppose that all *Value fields are valid,
      only one of the ContentTypes is valid. For other fields
      use @link(TsWorksheet.ReadAsText) and similar methods

      @member  Row            Row index of the cell (zero-based)
      @member  Col            Column index of the cell (zero-based)
      @member  Worksheet      Worksheet to which this cell belongs. In order to avoid circular unit references the worksheet is declared as @link(TsBasicWorksheet) here; usually it must be cast to TsWorksheet when used.
      @member  Flags          Status flags for this cell (see @link(TsCellFlags))
      @member  FormatIndex    Index to the @link(TsCellFormat) record to be applied to this cell. The format records are collected in the workbook's CellFormatList.
      @member  ConditionalFormatIndex  Array of indexes to the worksheet's  ConditionalalFormats list needed for conditional formatting
      @member  UTF8StringValue  String to be displayed in the cell if ContentType is cctUTF8String
      @member  RichTextParams Descriptions to be used for individual text formatting of parts of the cell text
      @member  ContentType    Type of the data stored in the cell. See @link(TsCellContentType).
      @member  NumberValue    Floating point value assigned to the cell. It is valid only when ContentType is cctNumber.
      @member  DateTimeValue  Date/time value assigned to the cell. It is valid only when ContentType is cctDateTime.
      @member  BoolValue      Boolean value assigned to the cell. It is valid only when ContentType is cctBoole.
      @member  ErrorValue     Error value assigned to the cell. The value is valid only when ContentType is cctError. .}
  TCell = record
    Row: Cardinal; // zero-based
    Col: Cardinal; // zero-based
    Worksheet: TsBasicWorksheet;   // Must be cast to TsWorksheet when used  (avoids circular unit reference)
    Flags: TsCellFlags;
    FormatIndex: Integer;
    ConditionalFormatIndex: array of Integer;
    
    // Cell content 
    UTF8StringValue: String;   // Strings cannot be part of a variant record
    RichTextParams: TsRichTextParams; // Formatting of individual text ranges
//    FormulaValue: String;      // Formula for calculation of cell content
    case ContentType: TCellContentType of  // variant part must be at the end
      cctEmpty      : ();      // has no data at all
//      cctFormula    : ();      // FormulaValue is outside the variant record
      cctNumber     : (Numbervalue: Double);
      cctUTF8String : ();      // UTF8StringValue is outside the variant record
      cctDateTime   : (DateTimeValue: TDateTime);
      cctBool       : (BoolValue: boolean);
      cctError      : (ErrorValue: TsErrorValue);
  end;
  {@@ Pointer to a @link(TCell) record }
  PCell = ^TCell;

  {@@ Types of row heights
    @value rhtDefault  Default row height
    @value rhtAuto     Automatically determined row height, depends on font size, text rotation, rich-text parameters, word-wrap
    @value rhtCustom   User-determined row height (dragging the row header borders in the grid, or changed by code) }
  TsRowHeightType = (rhtDefault, rhtCustom, rhtAuto);

  {@@ Types of column widths
    @value cwtDefault  Default column width
    @value cwtCustom   Userdefined column width (dragging the column header border in the grid, or by changed by code) }
  TsColWidthtype = (cwtDefault, cwtCustom);

  {@@ Column or row options
    @value croHidden     Column or row is hidden
    @value croPageBreak  Enforces a pagebreak before this column/row during printing }
  TsColRowOption = (croHidden, croPageBreak);
  TsColRowOptions = set of TsColRowOption;

  {@@ The record TRow contains information about a spreadsheet row:
    @member  Row            The index of the row (beginning with 0)
    @member  Height         The height of the row (expressed in the units defined by the workbook)
    @member  RowHeightType  Specifies whether the row has default, custom, or automatic height
    @member  FormatIndex    Row default format, index into the workbook's FCellFormatList
    @member  Options        See @link(TsColRowOption)
    
    Only rows with non-default height or non-default format or non-default
    Options have a row record. }
  TRow = record
    Row: Cardinal;
    Height: Single;
    RowHeightType: TsRowHeightType;
    FormatIndex: Integer;
    Options: TsColRowOptions;
  end;
  {@@ Pointer to a @link(TRow) record }
  PRow = ^TRow;

  {@@ The record TCol contains information about a spreadsheet column:
   @member Col          The index of the column (beginning with 0)
   @member Width        The width of the column (expressed in the units defined in the workbook)
   @member ColWidthType Specifies whether the column has default or custom width
   @member FormatIndex  Column default format, index into the workbook's FCellFormatlist
   @member Options      See @link(TsColRowOptions)
   
   Only columns with non-default width or non-default format or non-default
   Options have a column record. }
  TCol = record
    Col: Cardinal;
    Width: Single;
    ColWidthType: TsColWidthType;
    FormatIndex: Integer;
    Options: TsColRowOptions;
  end;
  {@@ Pointer to a @link(TCol) record }
  PCol = ^TCol;

  {@@ Embedded image 
    @member  Row                Row index of the cell at which the top sie of the image is anchored
    @member  Index              Index of the image in the workbook's embedded streams list.
    @member  Col                Column index of cell at which the left side of the image is anchored
    @member  OffsetX            Horizontal displacement of the image relative to the top/left corner of the anchor cell (in millimeters)
    @member  OffsetY            Vertical displacement of the image relative to the top/left corner of the anchor cell (in millimeters)
    @member  ScaleX             Horizontal scaling factor of the image
    @member  ScaleY             Vertical scaling factor of the image
    @member  Picture            Used internally by TPicture to display the image in the worksheet grid
    @member  HyperlinkTarget    Hyperlink assigned to the image
    @member  HyperlinkToolTip   Tooltip for the hyperlink of the image }
  TsImage = record
    Row, Col: Cardinal;       
    Index: Integer;           
    OffsetX, OffsetY: Double; 
    ScaleX, ScaleY: Double;   
    Picture: TObject;         
    HyperlinkTarget: String;  
    HyperlinkToolTip: String; 
  end;
  {@@ Pointer to a @link(TsImage) record}
  PsImage = ^TsImage;

  {@@ Image embedded in header or footer
    @member  Index   Index of the image in the workbook's embedded streams list }
  TsHeaderFooterImage = record
    Index: Integer;      
  end;

  {@@ Page orientation for printing 
    @value  spoPortrait   Printed page is in portrait orientation
    @value  spoLandscape  Printed page is in landscape orientation }
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

  {@@ Search option }
  TsSearchOption = (soCompareEntireCell, soMatchCase, soRegularExpr, soAlongRows,
    soBackward, soWrapDocument, soEntireDocument, soSearchInComment);

  {@@ A set of search options }
  TsSearchOptions = set of TsSearchOption;

  {@@ Defines which part of document is scanned }
  TsSearchWithin = (swWorkbook, swWorksheet, swColumn, swRow, swColumns, swRows);

  {@@ Search parameters }
  TsSearchParams = record
    SearchText: String;
    Options: TsSearchOptions;
    Within: TsSearchWithin;
    ColsRows: String;
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

  {@@ Worksheet user interface options:
    @value soShowGridLines    Show or hide the grid lines in the spreadsheet
    @value soShowHeaders      Show or hide the column or row headers of the spreadsheet
    @value soHasFrozenPanes   If set a number of rows and columns of the spreadsheet is fixed and does not scroll. The number is defined by LeftPaneWidth and TopPaneHeight.
    @value soHidden           Worksheet is hidden.
    @value soProtected        Worksheet is protected
    @value soPanesProtection  Panes are locked due to workbook protection
    @value soAutoDetectCellType  Auomatically detect type of cell content}
  TsSheetOption = (soShowGridLines, soShowHeaders, soHasFrozenPanes, soHidden,
    soProtected, soPanesProtection, soAutoDetectCellType);

  {@@ Set of user interface options (
    @seeAlso(TsSheetOption) }
  TsSheetOptions = set of TsSheetOption;

  {@@ Option flags for the workbook
    @value  boVirtualMode      If in virtual mode date are not taken from cells when a spreadsheet is written to file, but are provided by means of the event OnWriteCellData. Similarly, when data are read they are not added  as cells but passed the the event OnReadCellData;
    @value  boBufStream        When this option is set a buffered stream is used for writing (a memory stream swapping to disk) or reading (a file stream pre-reading chunks of data to memory)
    @value  boFileStream       Uses file streams and temporary files during reading and writing. Lowest memory consumptions, but slow.
    @value  boAutoCalc         Automatically recalculate formulas whenever a cell value changes, in particular when file is loaded.
    @value  boCalcBeforeSaving Calculates formulas before saving the file. Otherwise there are no results when the file is loaded back by fpspreadsheet.
    @value  boReadFormulas     Allows to turn off reading of rpn formulas; this is a precaution since formulas not correctly implemented by fpspreadsheet could crash the reading operation.
    @value boWriteZoomfactor   Instructs the writer to write the current zoom factors of the worksheets to file.
    @value boAbortReadOnFormulaError Aborts reading if a formula error is encountered
    @value boIgnoreFormulas    Formulas are not checked and not calculated. Cannot be used for biff formats. }
  TsWorkbookOption = (boVirtualMode, boBufStream, boFileStream,
    boAutoCalc, boCalcBeforeSaving, boReadFormulas, boWriteZoomFactor,
    boAbortReadOnFormulaError, boIgnoreFormulas);

  {@@ Set of option flags for the workbook (see @link(TsWorkbookOption)}
  TsWorkbookOptions = set of TsWorkbookOption;

  {@@ Workbook metadata }
  TsMetaData = class
  private
    FDateCreated: TDateTime;
    FDateLastModified: TDateTime;
    FLastModifiedBy: String;
    FTitle: String;
    FSubject: String;
    FAuthors: TStrings;
    FComments: TStrings;
    FKeywords: TStrings;
    FCustom: TStrings;
    function GetCreatedBy: String;
    procedure SetCreatedBy(AValue: String);
  public
    constructor Create;
    destructor Destroy; override;
    function AddCustom(AName, AValue: String): Integer;
    procedure Clear;
    function IsEmpty: Boolean;
    property CreatedBy: String read GetCreatedBy write SetCreatedBy;
    property LastModifiedBy: String read FLastModifiedBy write FLastModifiedBy;
    property DateCreated: TDateTime read FDateCreated write FDateCreated;
    property DateLastModified: TDatetime read FDateLastModified write FDateLastModified;
    property Subject: String read FSubject write FSubject;
    property Title: String read FTitle write FTitle;
    property Authors: TStrings read FAuthors write FAuthors;
    property Comments: TStrings read FComments write FComments;
    property Custom: TStrings read FCustom write FCustom;
    property Keywords: TStrings read FKeywords write FKeywords;
  end;

  {@@ Basic worksheet class to avoid circular unit references. It has only those
    properties and methods which do not require any other unit than fpstypes. }
  TsBasicWorksheet = class
  protected
    FName: String;  // Name of the worksheet (displayed at the tab)
    FOptions: TsSheetOptions;
    FProtection: TsWorksheetProtections;
    procedure SetName(const AName: String); virtual; abstract;
  public
    constructor Create;
    function HasHyperlink(ACell: PCell): Boolean;
    function IsProtected: Boolean;
    {@@ Name of the sheet. In the popular spreadsheet applications this is
      displayed in the tab of the sheet. }
    property Name: string read FName write SetName;
    {@@ Parameters controlling visibility of grid lines and row/column headers,
      usage of frozen panes etc. See @link(TsSheetOption). }
    property  Options: TsSheetOptions read FOptions write FOptions;
    {@@ Worksheet protection options }
    property Protection: TsWorksheetProtections read FProtection write FProtection;
  end;

  {@@ Basic worksheet class to avoid circular unit references. It contains only
    those properties and methods which do not require any other unit than
    fpstypes. }
  TsBasicWorkbook = class
  private
    FLog: TStringList;
    function GetErrorMsg: String;
  protected
    FFileName: String;
    FFormatID: TsSpreadFormatID;
    FOptions: TsWorkbookOptions;
    FProtection: TsWorkbookProtections;
    FUnits: TsSizeUnits;  // Units for row heights and col widths
  public
    {@@ A copy of SysUtil's DefaultFormatSettings (converted to UTF8) to provide
      some kind of localization of some formatting strings.
      Can be modified before loading/writing files }
    FormatSettings: TFormatSettings;

    constructor Create;
    destructor Destroy; override;

    { Error messages }
    procedure AddErrorMsg(const AMsg: String); overload;
    procedure AddErrorMsg(const AMsg: String; const Args: array of const); overload;
    procedure ClearErrorList; inline;

    { Protection }
    function IsProtected: Boolean;

    {@@ Retrieves error messages collected during reading/writing }
    property ErrorMsg: String read GetErrorMsg;
    {@@ Identifies the file format which was detected when reading the file }
    property FileFormatID: TsSpreadFormatID read FFormatID;
    {@@ Filename of the saved workbook }
    property FileName: String read FFileName;
    {@@ Option flags for the workbook - see boXXXX declarations }
    property Options: TsWorkbookOptions read FOptions write FOptions;
    {@@ Workbook protection flags }
    property Protection: TsWorkbookProtections read FProtection write FProtection;
    {@@ Units of row heights and column widths }
    property Units: TsSizeUnits read FUnits;
  end;

  {@@ Ancestor of the fpSpreadsheet exceptions } 
  EFpSpreadsheet = class(Exception);
  {@@ Class of exceptions fired by the workbook reader }
  EFpSpreadsheetReader = class(EFpSpreadsheet);
  {@@ Class of exceptions fired for the workbook writer }
  EFpSpreadsheetWriter = class(EFpSpreadsheet);

const
  RowHeightTypeNames: array[TsRowHeightType] of string = (
    'Default', 'Custom', 'Auto');

  ColWidthTypeNames: array[TsColWidthType] of string = (
    'Default', 'Custom');

  { Indexes to be used for the various headers and footers }
  
  {@@ Index of the first header/footer to be used }
  HEADER_FOOTER_INDEX_FIRST   = 0;
  {@@ Index of the header/footer to be used for odd page numbers }
  HEADER_FOOTER_INDEX_ODD     = 1;
  {@@ Index of the header/footer to be used for even page numbers }
  HEADER_FOOTER_INDEX_EVEN    = 2;
  {@@ Index of the header/footer to be used for all pages }
  HEADER_FOOTER_INDEX_ALL     = 1;

procedure InitUTF8FormatSettings(out AFormatSettings: TFormatSettings);


implementation

{@@ ----------------------------------------------------------------------------
  Creates a localized FPC format settings record in which all strings are
  encoded as UTF8.
-------------------------------------------------------------------------------}
procedure InitUTF8FormatSettings(out AFormatSettings: TFormatSettings);
// remove when available in LazUtils
var
  i: Integer;
begin
  AFormatSettings := DefaultFormatSettings;
  AFormatSettings.CurrencyString := AnsiToUTF8(DefaultFormatSettings.CurrencyString);
  for i:=1 to 12 do begin
    AFormatSettings.LongMonthNames[i] := AnsiToUTF8(DefaultFormatSettings.LongMonthNames[i]);
    AFormatSettings.ShortMonthNames[i] := AnsiToUTF8(DefaultFormatSettings.ShortMonthNames[i]);
  end;
  for i:=1 to 7 do begin
    AFormatSettings.LongDayNames[i] := AnsiToUTF8(DefaultFormatSettings.LongDayNames[i]);
    AFormatSettings.ShortDayNames[i] := AnsiToUTF8(DefaultFormatSettings.ShortDayNames[i]);
  end;
end;

{ TsCellFormat }

procedure TsCellFormat.SetBackground(AFillStyle: TsFillStyle;
  AFgColor, ABgColor: TsColor);
begin
  UsedFormattingFields := UsedFormattingFields + [uffBackground];
  Background.FgColor := AFgColor;
  Background.BgColor := ABgColor;
  Background.Style := AFillStyle;
end;

procedure TsCellFormat.SetBackgroundColor(AColor: TsColor);
begin
  SetBackground(fsSolidFill, AColor, AColor);
end;

procedure TsCellFormat.SetBorders(ABorders: TsCellBorders;
  const AColor: TsColor = scBlack; const ALineStyle: TsLineStyle = lsThin);
var
  cb: TsCellBorder;
begin
  for cb in ABorders do
  begin
    if (AColor = scTransparent) or (AColor = scNotDefined) then
      Exclude(Border, cb)
    else
    begin
      Include(Border, cb);
      BorderStyles[cb].LineStyle := ALineStyle;
      BorderStyles[cb].Color := AColor;
    end;
  end;
  if Border = [] then
    UsedFormattingFields := UsedFormattingfields - [uffBorder]
  else
    UsedFormattingFields := UsedFormattingfields + [uffBorder];
end;

procedure TsCellFormat.SetFont(AFontIndex: Integer);
begin
  FontIndex := AFontIndex;
  UsedFormattingFields := UsedFormattingFields + [uffFont];
end;

procedure TsCellFormat.SetHorAlignment(AHorAlign: TsHorAlignment);
begin
  HorAlignment := AHorAlign;
  UsedFormattingFields := usedFormattingFields + [uffHorAlign];
end;

procedure TsCellFormat.SetNumberFormat(AIndex: Integer);
begin
  NumberFormatIndex := AIndex;
  UsedFormattingFields := UsedFormattingFields + [uffNumberFormat];
end;

procedure TsCellFormat.SetTextRotation(ARotation: TsTextRotation);
begin
  TextRotation := ARotation;
  UsedFormattingFields := UsedFormattingFields + [uffTextRotation];
end;

procedure TsCellFormat.SetVertAlignment(AVertAlign: TsVertAlignment);
begin
  VertAlignment := AVertAlign;
  UsedFormattingfields := UsedFormattingFields + [uffVertAlign];
end;


{ TsFont }

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


{-------------------------------------------------------------------------------
                                TsMetaData
-------------------------------------------------------------------------------}
constructor TsMetaData.Create;
begin
  inherited;
  FAuthors := TStringList.Create;
  FAuthors.StrictDelimiter := true;
  FAuthors.Delimiter := ';';
  FComments := TStringList.Create;
  FKeywords := TStringList.Create;
  FCustom := TStringList.Create;
end;

destructor TsMetaData.Destroy;
begin
  FAuthors.Free;
  FComments.Free;
  FKeywords.Free;
  FCustom.Free;
  inherited;
end;

procedure TsMetaData.Clear;
begin
  FTitle := '';
  FSubject := '';
  FLastModifiedBy := '';
  FDateCreated := 0;
  FDateLastModified := 0;
  FAuthors.Clear;
  FComments.Clear;
  FKeywords.Clear;
  FCustom.Clear;
end;

function TsMetaData.AddCustom(AName, AValue: String): Integer;
begin
  Result := FCustom.IndexOf(AName);
  if result > -1 then
    FCustom.ValueFromIndex[Result] := AValue
  else
    Result := FCustom.Add(AName + '=' + AValue);
end;

function TsMetaData.GetCreatedBy: String;
begin
  Result := FAuthors.DelimitedText;
end;

function TsMetaData.IsEmpty: Boolean;
begin
  Result := (FLastModifiedBy = '') and (FTitle = '') and (FSubject = '') and
    (FAuthors.Count = 0) and (FComments.Count = 0) and (FKeywords.Count = 0) and
    (FCustom.Count = 0) and (FDateCreated = 0) and (FDateLastModified = 0);
end;

{@@ ----------------------------------------------------------------------------
  Provide initial author. In case of multiple authors, separate the names by
  semicolons. 
-------------------------------------------------------------------------------}
procedure TsMetaData.SetCreatedBy(AValue: String);
begin
  FAuthors.DelimitedText := AValue;
end;


{-------------------------------------------------------------------------------
                              TsBasicWorksheet
-------------------------------------------------------------------------------}

constructor TsBasicWorksheet.Create;
begin
  inherited;
  FProtection := DEFAULT_SHEET_PROTECTION;
end;

{@@ ----------------------------------------------------------------------------
  Checks whether the specified cell contains a hyperlink
-------------------------------------------------------------------------------}
function TsBasicWorksheet.HasHyperlink(ACell: PCell): Boolean;
begin
  Result := (ACell <> nil) and (cfHyperlink in ACell^.Flags);
end;

{@@ ----------------------------------------------------------------------------
  Returns whether the worksheet is protected
-------------------------------------------------------------------------------}
function TsBasicWorksheet.IsProtected: Boolean;
begin
  Result := soProtected in FOptions;
end;


{-------------------------------------------------------------------------------
                             TsBasicWorkbook
-------------------------------------------------------------------------------}
constructor TsBasicWorkbook.Create;
begin
  inherited;
  InitUTF8FormatSettings(FormatSettings);
  FUnits := suMillimeters;              // Units for column width and row height
  FFormatID := sfidUnknown;
  FLog := TStringList.Create;
  FProtection := [];
end;

destructor TsBasicWorkbook.Destroy;
begin
  FLog.Free;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Adds a (simple) error message to an internal list

  @param   AMsg   Error text to be stored in the list
-------------------------------------------------------------------------------}
procedure TsBasicWorkbook.AddErrorMsg(const AMsg: String);
begin
  FLog.Add(AMsg);
end;

{@@ ----------------------------------------------------------------------------
  Adds an error message composed by means of format codes to an internal list

  @param   AMsg   Error text to be stored in the list
  @param   Args   Array of arguments to be used by the Format() function
-------------------------------------------------------------------------------}
procedure TsBasicWorkbook.AddErrorMsg(const AMsg: String;
  const Args: Array of const);
begin
  FLog.Add(Format(AMsg, Args));
end;

{@@ ----------------------------------------------------------------------------
  Clears the internal error message list
-------------------------------------------------------------------------------}
procedure TsBasicWorkbook.ClearErrorList;
begin
  FLog.Clear;
end;

{@@ ----------------------------------------------------------------------------
  Getter to retrieve the error messages collected during reading/writing
-------------------------------------------------------------------------------}
function TsBasicWorkbook.GetErrorMsg: String;
begin
  Result := FLog.Text;
end;

{@@ ----------------------------------------------------------------------------
  Returns whether the workbook is protected
-------------------------------------------------------------------------------}
function TsBasicWorkbook.IsProtected: Boolean;
begin
  Result := (FProtection <> []);
end;


end.

