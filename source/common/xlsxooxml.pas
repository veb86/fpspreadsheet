{
xlsxooxml.pas

Writes an OOXML (Office Open XML) document

An OOXML document is a compressed ZIP file with the following files inside:

[Content_Types].xml         -
_rels/.rels                 -
xl/_rels\workbook.xml.rels  -
xl/workbook.xml             - Global workbook data and list of worksheets
xl/styles.xml               -
xl/sharedStrings.xml        -
xl/worksheets\sheet1.xml    - Contents of each worksheet
...
xl/worksheets\sheetN.xml

Specifications obtained from:

http://openxmldeveloper.org/default.aspx

also:
http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx#BMworksheetworkbook

AUTHORS: Felipe Monteiro de Carvalho, Reinier Olislagers, Werner Pamler
}

unit xlsxooxml;

{$mode objfpc}{$H+}
{$include ../fps.inc}

{$WARN 6058 off : Call to subroutine "$1" marked as inline is not inlined}

interface

uses
  Classes, SysUtils,
  laz2_xmlread, laz2_DOM, avglvltree,
 {$ifdef FPS_PATCHED_ZIPPER}
  fpszipper,
 {$else}
  zipper,
 {$endif}
  fpsTypes, fpsUtils, fpsReaderWriter, fpsNumFormat, fpsPalette,
  fpsConditionalFormat,
 {$ifdef FPS_CHARTS}
  fpsChart,
 {$ENDIF}
  fpsxmlcommon, xlsCommon;
  
type
  TXlsxRelationship = class
    RelID: String;
    Target: String;
    Schema: String;
  end;

  TXlsxRelationshipList = class(TFPList)
  private
    function GetItem(AIndex: Integer): TXlsxRelationship;
    procedure SetItem(AIndex: Integer; AValue: TXlsxRelationship);
  public
    destructor Destroy; override;
    function Add(ARelID, ASchema, ATarget: String): TXlsxRelationship;
    procedure Clear;
    procedure Delete(AIndex: Integer);
    function FindTarget(ARelID: String): String;
    property Items[AIndex: Integer]: TXlsxRelationship read GetItem write SetItem; default;
  end;

  { TsSpreadOOXMLReader }

  TsSpreadOOXMLReader = class(TsSpreadXMLReader)
  private
   {$ifdef FPS_CHARTS}
    FChartReader: TsBasicSpreadChartReader;
   {$endif}
    FDateMode: TDateMode;
    FPointSeparatorSettings: TFormatSettings;
    FSharedStrings: TStringList;
    FSheetList: TFPList;
    FWorkbookRels: TFPList;
    FFillList: TFPList;
    FBorderList: TFPList;
    FDrawingToSheetRelList: TFPList;
    FEmbeddedObjList: TFPList;
    FHyperlinkList: TFPList;
    FSharedFormulaBaseList: TFPList;
    FPalette: TsPalette;
    FThemeColors: array of TsColor;
    FLastRow, FLastCol: Cardinal;
    FWrittenByFPS: Boolean;
    procedure ApplyCellFormatting(ACell: PCell; XfIndex: Integer);
    procedure ApplyHyperlinks(AWorksheet: TsBasicWorksheet);
  protected
    procedure ReadActiveSheet(ANode: TDOMNode; out ActiveSheetIndex: Integer);
    procedure ReadBorders(ANode: TDOMNode);
    function ReadBorderStyle(ANode: TDOMNode; out ABorderStyle: TsCellBorderStyle): Boolean;
    procedure ReadCell(ANode: TDOMNode; AWorksheet: TsBasicWorksheet;
      ARowIndex: Cardinal; var AColIndex: Cardinal);
    procedure ReadCellXfs(ANode: TDOMNode);
    procedure ReadCFAverage(ANode: TDOMNode; AWorksheet: TsBasicWorksheet;
      ARange: TsCellRange; AFormatIndex: Integer);
    procedure ReadCFCellFormat(ANode: TDOMNode; AWorksheet: TsBasicWorksheet;
      ARange: TsCellRange; AFormatIndex: Integer);
    procedure ReadCFColorRange(ANode: TDOMNode; AWorksheet: TsBasicWorksheet;
      ARange: TsCellRange);
    procedure ReadCFDataBars(ANode: TDOMNode; AWorksheet: TsBasicWorksheet;
      ARange: TsCellRange);
    procedure ReadCFDateFormat(ANode: TDOMNode; AWorksheet: TsBasicWorksheet;
      ARange: TsCellRange; AFormatIndex: Integer);
    procedure ReadCFExpression(ANode: TDOMNode; AWorksheet: TsBasicWorksheet;
      ARange: TsCellRange; AFormatIndex: Integer);
    procedure ReadCFIconSet(ANode: TDOMNode; AWorksheet: TsBasicWorksheet;
      ARange: TsCellRange);
    procedure ReadCFMisc(ANode: TDOMNode; AWorksheet: TsBasicWorksheet;
      ARange: TsCellRange; AFormatIndex: Integer);
    procedure ReadCFTop10(ANode: TDOMNode; AWorksheet: TsBasicWorksheet;
      ARange: TsCellRange; AFormatIndex: Integer);
    procedure ReadColRowBreaks(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    function  ReadColor(ANode: TDOMNode): TsColor;
    procedure ReadCols(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadComments(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadComments(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure ReadConditionalFormatting(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadDateMode(ANode: TDOMNode);
    procedure ReadDefinedNames(ANode: TDOMNode);
    procedure ReadDifferentialFormat(ANode: TDOMNode);
    procedure ReadDifferentialFormats(ANode: TDOMNode);
    procedure ReadDimension(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadDrawing(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadDrawingRels(ANode: TDOMNode; ASheet: TsBasicWorksheet);
    procedure ReadEmbeddedObjs(AStream: TStream);
    procedure ReadFileVersion(ANode: TDOMNode);
    procedure ReadFills(ANode: TDOMNode);
    function ReadFont(ANode: TDOMNode): Integer;
    procedure ReadFonts(ANode: TDOMNode);
    procedure ReadHeaderFooter(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadHyperlinks(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadMedia(AStream: TStream);
    procedure ReadMetaData(ANode: TDOMNode);
    procedure ReadMergedCells(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadNumFormats(ANode: TDOMNode);
    procedure ReadPageMargins(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadPageSetup(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadPalette(ANode: TDOMNode);
    procedure ReadPrintOptions(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadRow(ANode: TDOMNode; AWorksheet: TsBasicWorksheet; var ARowIndex: Cardinal);
    procedure ReadSharedStrings(ANode: TDOMNode);
    procedure ReadSheetFormatPr(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadSheetList(ANode: TDOMNode);
    procedure ReadSheetPr(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadSheetProtection(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadSheetRels(AStream: TStream);
    procedure ReadSheetRels(AStream: TStream; ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadSheetViews(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadThemeColors(ANode: TDOMNode);
    procedure ReadThemeElements(ANode: TDOMNode);
    procedure ReadVmlDrawing(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadWorkbookProtection(ANode: TDOMNode);
    procedure ReadWorksheet(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
  protected
    FFirstNumFormatIndexInFile: Integer;
    procedure AddBuiltinNumFormats; override;
    class function IsEncrypted(AStream: TStream): Boolean;
  public
    constructor Create(AWorkbook: TsBasicWorkbook); override;
    destructor Destroy; override;
    class function CheckFileFormat(AStream: TStream): Boolean; override;
    function CreateXMLStream: TStream;
    function NeedsPassword(AStream: TStream): Boolean; override;
    procedure ReadFromStream(AStream: TStream; APassword: String = '';
      AParams: TsStreamParams = []); override;
    procedure ReadRels(AStream: TStream; ARelsFile: String; ARelsList: TFPList);
  end;

  { TsSpreadOOXMLWriter }

  TsSpreadOOXMLWriter = class(TsCustomSpreadWriter)
  private
   {$IFDEF FPS_CHARTS}
    FChartWriter: TsBasicSpreadChartWriter;
   {$ENDIF}
    FFirstNumFormatIndexInFile: Integer;
    vmlDrawingCounter: Integer;
  protected
    FDateMode: TDateMode;
    FPointSeparatorSettings: TFormatSettings;
    FSharedStringsCount: Integer;
    FFillList: array of PsCellFormat;
    FBorderList: array of PsCellFormat;
    FDifferentialFormatIndexList: array of Integer;
    function GetActiveTab: String;
    procedure Get_rId(AWorksheet: TsBasicWorksheet;
      out AComment_rId, AFirstHyperlink_rId, ADrawing_rId, ADrawingHF_rId: Integer);
  protected
    procedure AddBuiltinNumFormats; override;
    procedure CreateStreams;
    procedure DestroyStreams;
    function  FindBorderInList(AFormat: PsCellFormat): Integer;
    function  FindFillInList(AFormat: PsCellFormat): Integer;
    function GetStyleIndex(ACell: PCell): Cardinal;
    procedure ListAllBorders;
    procedure ListAllDifferentialFormats;
    procedure ListAllFills;
    function  PrepareFormula(const AFormula: String): String;
    procedure ResetStreams;
    procedure WriteBorderList(AStream: TStream);
    procedure WriteCFCellRule(AStream: TStream; ARule: TsCFCellRule; ARange: TsCellRange; APriority: Integer);
    procedure WriteCFColorRangeRule(AStream: TStream; ARule: TsCFColorRangeRule; APriority: Integer);
    procedure WriteCFDataBarRule(AStream: TStream; ARule: TsCFDatabarRule; APriority: Integer);
    procedure WriteCFIconSetRule(AStream: TStream; ARule: TsCFIconSetRule; APriority: Integer);
    procedure WriteColBreaks(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteCols(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteComments(AWorksheet: TsBasicWorksheet);
    procedure WriteConditionalFormat(AStream: TStream; AFormat: TsConditionalFormat; var APriority: Integer);
    procedure WriteConditionalFormats(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteCustomMetaData(AStream: TStream);
    procedure WriteDefinedNames(AStream: TStream);
    procedure WriteDifferentialFormat(AStream: TStream; AFormat: PsCellFormat);
    procedure WriteDifferentialFormats(AStream: TStream);
    procedure WriteDimension(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteDrawings(AWorksheet: TsBasicWorksheet);
    procedure WriteDrawingRels(AWorksheet: TsBasicWorksheet);
//    procedure WriteDrawingsOfSheet(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteFillList(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFont: TsFont; UseInStyleNode: Boolean);
    procedure WriteFontList(AStream: TStream);
    procedure WriteHeaderFooter(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteHyperlinks(AStream: TStream; AWorksheet: TsBasicWorksheet; rId: Integer);
    procedure WriteMetadata(AStream: TStream);
    procedure WriteMergedCells(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteNumFormatList(AStream: TStream);
    procedure WritePalette(AStream: TStream);
    procedure WritePageMargins(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WritePageSetup(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WritePrintOptions(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteRowBreaks(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteSheetData(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteSheetFormatPr(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteSheetPr(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteSheetProtection(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteSheets(AStream: TStream);
    procedure WriteSheetViews(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteStyle(AStream: TStream; ANodeName: String; AFormat: PsCellFormat);
    procedure WriteStyleList(AStream: TStream; ANodeName: String);
    procedure WriteVMLDrawings(AWorksheet: TsBasicWorksheet);
    procedure WriteVMLDrawings_Comments(AWorksheet: TsBasicWorksheet);
    procedure WriteVMLDrawings_HeaderFooterImages(AWorksheet: TsBasicWorksheet);
    procedure WriteVMLDrawingRels(AWorksheet: TsBasicWorksheet);
    procedure WriteWorkbook(AStream: TStream);
    procedure WriteWorkbookProtection(AStream: TStream);
    procedure WriteWorkbookRels(AStream: TStream);
    procedure WriteWorksheet(AWorksheet: TsBasicWorksheet);
    procedure WriteWorksheetRels(AWorksheet: TsBasicWorksheet);
  protected
    { Streams with the contents of files }
    FSContentTypes: TStream;
    FSRelsRels: TStream;
    FSWorkbook: TStream;
    FSWorkbookRels: TStream;
    FSMetaData: TStream;
    FSCustomMetaData: TStream;
    FSStyles: TStream;
    FSSharedStrings: TStream;
    FSSharedStrings_complete: TStream;
    FSMedia: array of TStream;
    FSSheets: array of TStream;
    FSSheetRels: array of TStream;
    FSComments: array of TStream;
    FSDrawings: array of TStream;
    FSDrawingsRels: array of TStream;
    FSVmlDrawings: array of TStream;
    FSVmlDrawingsRels: array of TStream;
    FCurSheetNum: Integer;
  protected
    { Routines to write the files }
    procedure WriteContent;
    procedure WriteContentTypes;
    procedure WriteGlobalFiles;
    procedure WriteMedia(AZip: TZipper);
  protected
    { Record writing methods }
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: Boolean; ACell: PCell); override;
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TDateTime; ACell: PCell); override;
    procedure WriteError(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TsErrorValue; ACell: PCell); override;
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;

  public
    constructor Create(AWorkbook: TsBasicWorkbook); override;
    destructor Destroy; override;
    { General writing methods }
    procedure WriteStringToFile(AFileName, AString: string);
    procedure WriteToStream(AStream: TStream; AParams: TsStreamParams = []); override;
  end;

  TXlsxSettings = record
    DateMode: TDateMode;
  end;

var
  {@@ Default settings for reading/writing of XLSX files }
  XlsxSettings: TXlsxSettings = (
    DateMode: dm1900;
  );

  sfidOOXML: TsSpreadFormatID;

procedure InitOOXMLLimitations(out ALimitations: TsSpreadsheetFormatLimitations);
function MakeXLPath(AFileName: String): String;


implementation

uses
  variants, strutils, dateutils, math, lazutf8, LazFileUtils, uriparser, typinfo,
  {%H-}fpsPatches, fpSpreadsheet, fpsCrypto, fpsExprParser,
  fpsStrings, fpsStreams, fpsClasses,
  {$ifdef FPS_CHARTS}xlsxooxmlChart,{$endif}
  fpsImages;

const
  { OOXML general XML constants }
     XML_HEADER           = '<?xml version="1.0" encoding="utf-8" ?>';

  { OOXML Directory structure constants }
  // Note: directory separators are always / because the .xlsx is a zip file which
  // requires / instead of \, even on Windows; see 
  // http://www.pkware.com/documents/casestudies/APPNOTE.TXT
  // 4.4.17.1 All slashes MUST be forward slashes '/' as opposed to backward slashes '\'
     OOXML_PATH_TYPES              = '[Content_Types].xml';
{%H-}OOXML_PATH_RELS               = '_rels/';
     OOXML_PATH_RELS_RELS          = '_rels/.rels';
{%H-}OOXML_PATH_XL                 = 'xl/';
{%H-}OOXML_PATH_XL_RELS            = 'xl/_rels/';
     OOXML_PATH_XL_RELS_RELS       = 'xl/_rels/workbook.xml.rels';
     OOXML_PATH_XL_WORKBOOK        = 'xl/workbook.xml';
     OOXML_PATH_XL_STYLES          = 'xl/styles.xml';
     OOXML_PATH_XL_STRINGS         = 'xl/sharedStrings.xml';
     OOXML_PATH_XL_WORKSHEETS      = 'xl/worksheets/';
     OOXML_PATH_XL_WORKSHEETS_RELS = 'xl/worksheets/_rels/';
     OOXML_PATH_XL_CHARTS          = 'xl/charts/';
     OOXML_PATH_XL_CHARTS_RELS     = 'xl/charts/_rels/';
     OOXML_PATH_XL_DRAWINGS        = 'xl/drawings/';
     OOXML_PATH_XL_DRAWINGS_RELS   = 'xl/drawings/_rels/';
     OOXML_PATH_XL_THEME           = 'xl/theme/theme1.xml';
     OOXML_PATH_XL_MEDIA           = 'xl/media/';
     OOXML_PATH_DOCPROPS_CORE      = 'docProps/core.xml';
     OOXML_PATH_DOCPROPS_CUSTOM    = 'docProps/custom.xml';

     { OOXML schemas constants }
     SCHEMAS_TYPES        = 'http://schemas.openxmlformats.org/package/2006/content-types';
     SCHEMAS_RELS         = 'http://schemas.openxmlformats.org/package/2006/relationships';
     SCHEMAS_DOC_RELS     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
     SCHEMAS_DOCUMENT     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
     SCHEMAS_META_CORE    = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
     SCHEMAS_META_CUSTOM  = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties';
     SCHEMAS_WORKSHEET    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
     SCHEMAS_STYLES       = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
     SCHEMAS_STRINGS      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
     SCHEMAS_COMMENTS     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
     SCHEMAS_DRAWING      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
     SCHEMAS_VMLDRAWING   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
     SCHEMAS_HYPERLINK    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
     SCHEMAS_IMAGE        = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
     SCHEMAS_CHART        = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
     SCHEMAS_SPREADML     = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
     SCHEMAS_CORE         = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
     SCHEMAS_DRAWINGML_CHART= 'http://schemas.openxmlformats.org/drawingml/2006/chart';

     { OOXML mime types constants }
     MIME_XML             = 'application/xml';
     MIME_RELS            = 'application/vnd.openxmlformats-package.relationships+xml';
     MIME_OFFICEDOCUMENT  = 'application/vnd.openxmlformats-officedocument';
     MIME_CORE            = 'application/vnd.openxmlformats-package.core-properties+xml';
     MIME_CUSTOM          = 'application/vnd.openxmlformats-officedocument.custom-properties+xml';
     MIME_SPREADML        = MIME_OFFICEDOCUMENT + '.spreadsheetml';
     MIME_SHEET           = MIME_SPREADML + '.sheet.main+xml';
     MIME_WORKSHEET       = MIME_SPREADML + '.worksheet+xml';
     MIME_STYLES          = MIME_SPREADML + '.styles+xml';
     MIME_STRINGS         = MIME_SPREADML + '.sharedStrings+xml';
     MIME_COMMENTS        = MIME_SPREADML + '.comments+xml';
     MIME_DRAWING         = MIME_OFFICEDOCUMENT + '.drawing+xml';
     MIME_VMLDRAWING      = MIME_OFFICEDOCUMENT + '.vmlDrawing';

     LAST_PALETTE_INDEX   = 63;

     CFB_SIGNATURE        = $E11AB1A1E011CFD0;   // Compound File Binary Signature

     LE = LineEnding;

type
  TFillListData = class
    PatternType: String;
    FgColor: TsColor;
    BgColor: Tscolor;
  end;

  TBorderListData = class
    Borders: TsCellBorders;
    BorderStyles: TsCellBorderStyles;
  end;

  TDifferentialFormatData = class
    CellFormatIndex: Integer;
    UsedFormattingFields: TsUsedFormattingFields;
    Borders: TsCellBorders;
    BorderStyles: TsCellBorderStyles;
    FillPatternType: TsFillStyle;
    FillFgColor: TsColor;
    FillBgColor: TsColor;
    FontColor: TsColor;
    FontStyles: TsFontStyles;
    NumFormatStr: String;
  end;

  TDrawingToSheetRel = class
    DrawingFileName: String;
    Worksheet: TsBasicWorksheet;
  end;
  
  TEmbeddedObjData = class
    RelID: String;
    MediaName: String;
    FileName: String;
    ImgIndex: Integer;
    Worksheet: TsBasicWorksheet;
    IsChart: boolean;
    IsHeaderFooter: Boolean;
    // This part is for images or charts embedded in the worksheet
    XPos, YPos, Width, Height: Double;
    FromRow, FromCol, ToRow, ToCol: Cardinal;
    FromRowOffs, FromColOffs, ToRowOffs, ToColOffs: Double;
    // This part is for header/footer images.
    HFImgPosition: string; // 'LH', 'CH', 'RH', 'LF', 'CF', 'RF';
    HFImgWidth: Double;
    HFImgHeight: Double;
  end;
  
  THyperlinkListData = class
    ID: String;
    CellRef: String;
    Target: String;
    TextMark: String;
    Display: String;
    Tooltip: String;
    Worksheet: TsBasicWorksheet;
  end;

  TSharedFormulaData = class
    Worksheet: TsWorksheet;
    Row: Integer;
    Col: Integer;
    Formula: String;
  end;

  TSheetData = class
    Name: String;
    ID: String;
    RelID : String;
    Target : String;
    Hidden: Boolean;
    SheetRels: TFPList;
    SheetRelsFile: String;
    Drawing_File: String;
    DrawingRels: TXlsxRelationshipList;
    VmlDrawing_File: String;
    VmlDrawingRels: TXlsxRelationshipList;
    HyperlinkRels: TXlsxRelationshipList;
    Comments_File: String;
    constructor Create;
    destructor Destroy; override;
  end;
  
  TSharedObjData = class
    Picture: array of byte;
    RelId: String;
  end;

const
  PATTERN_TYPES: array [TsFillStyle] of string = (
    'none',            // fsNoFill
    'solid',           // fsSolidFill
    'darkGray',        // fsGray75
    'mediumGray',      // fsGray50
    'lightGray',       // fsGray25
    'gray125',         // fsGray12
    'gray0625',        // fsGray6,
    'darkHorizontal',  // fsStripeHor
    'darkVertical',    // fsStripeVert
    'darkUp',          // fsStripeDiagUp
    'darkDown',        // fsStripeDiagDown
    'lightHorizontal', // fsThinStripeHor
    'lightVertical',   // fsThinStripeVert
    'lightUp',         // fsThinStripeDiagUp
    'lightDown',       // fsThinStripeDiagDown
    'darkTrellis',     // fsHatchDiag
    'lightTrellis',    // fsHatchThinDiag
    'darkTellis',      // fsHatchThickDiag
    'lightGrid'        // fsHatchThinHor
    );

  LINESTYLE_TYPES: array[TsLineStyle] of String = (
     'thin',                   // lsThin
     'medium',                 // lsMedium
     'dashed',                 // lsDashed
     'dotted',                 // lsDotted
     'thick',                  // lsThick
     'double',                 // lsDouble
     'hair',                   // lsHair
     'mediumDashed',           // lsMediumDash
     'dashDot',                // lsDashDot
     'mediumDashDot',          // lsMediumDashDot
     'dashDotDot',             // lsDashDotDot
     'mediumDashDotDot',       // lsMediumDashDotDot
     'slantDashDot'            // lsSlantDashDot
  );

  CF_TYPE_NAMES: array[TsCFCondition] of String = (
    'cellIs', 'cellIs',                      // cfcEqual, cfcNotEqual,
    'cellIs', 'cellIs', 'cellIs', 'cellIs',  //  cfcGreaterThan, cfcLessThan, cfcGreaterEqual, cfcLessEqual,
    'cellIs', 'cellIs',                      // cfcBetween, cfcNotBetween,
    'aboveAverage', 'aboveAverage', 'aboveAverage', 'aboveAverage', // cfcAboveAverage, cfcBelowAverage, cfcAboveEqualAverage, cfcBelowEqualAverage,
    'top10', 'top10', 'top10', 'top10',      // cfcTop, cfcBottom, cfcTopPercent, cfcBottomPercent,
    'duplicateValues', 'uniqueValues',       // cfcDuplicate, cfcUnique,
    'beginsWith', 'endsWith',                // cfcBeginsWith, cfcEndsWith,
    'containsText', 'notContainsText',       // cfcContainsText, cfcNotContainsText,
    'containsErrors', 'notContainsErrors',   // cfcContainsErrors, cfcNotContainsErrors
    'timePeriod', 'timePeriod', 'timePeriod',// cfcYesterday, cfcToday, cfcTomorrow
    'timePeriod',                            // cfcLast7Days
    'timePeriod', 'timePeriod', 'timePeriod',// cfcLastWeek, cfcThisWeek, cfcNextWeek
    'timePeriod', 'timePeriod', 'timePeriod',// cfcLastMonth, cfcThisMonth, cfcNextMonth
    'expression', 'expression', 'expression',// cfcLastYear, cfcThisYear, cfcNextYear      // implemented as expression
    'expression'                             // cfcExpression
  );

  CF_OPERATOR_NAMES: array[TsCFCondition] of string = (
    'equal', 'notEqual', 'greaterThan', 'lessThan', 'greaterThanOrEqual', 'lessThanOrEqual',
    'between', 'notBetween',
    '', '', '', '',   // cfcAboveAverage, cfcBelowAverage, cfcAboveEqualAverage, cfcBelowEqualAverage,
    '', '', '', '',   // cfcTop, cfcBottom, cfcTopPercent, cfcBottomPercent,
    '', '',           // cfcDuplicate, cfcUnique,
    '', '', '', 'notContains', //cfcBeginsWith, cfcEndsWith, cfcContainsText, cfcNotContainsText,
    '', '',           // cfcContainsErrors, cfcNotContainsErrors
    'yesterday', 'today', 'tomorrow', 'last7Days', // cfcYesterday, cfcToday, cfcTomorrow, cfcLast7Days
    'lastWeek', 'thisWeek', 'nextWeek',            // cfcLastWeek, cfcThisWeek, cfcNextWeek
    'lastMonth', 'thisMonth', 'nextMonth',         // cfcLastMonth, cfcThisMonth, cfcNextMonth
    '', '', '',       // cfcLastYear, cfcThisYear, cfcNextYear
    ''                // cfcExpression
  );

  CF_YEAR_FORMULAS: array[cfcLastYear..cfcNextYear] of string = (
    'YEAR(%0:s)=YEAR(TODAY())-1',         // cfcLastYear
    'YEAR(%0:s)=YEAR(TODAY())',           // cfcThisYear
    'YEAR(%0:s)=YEAR(TODAY())+1'          // cfcNextYear
  );

  CF_ICON_SET: array[TsCFIconSet] of string = (
    '3Arrows', '3ArrowsGray','3Flags',  // is3Arrows, is3ArrowsGray, is3Flags
    '3TrafficLights2',                  // REPLACEMENT FOR is3TrafficLights1 which requires x14
    '3TrafficLights',                   // is3TrafficLights2
    '3Signs', '3Symbols', '3Symbols2',  // is3Signs, is3Symbols, is3Symbols2
    '3Signs', '3Signs', '3Signs', '3Signs', // REPLACEMENT FOR  is3Smilies, is3Stars, is3Triangles, is3ColorSmilies which need x14
    '4Arrows', '4ArrowsGray',           // is4Arrows, is4ArrowsGray
    '4RedToBlack', '4Rating',           // is4RedToBlack, is4Rating
    '4TrafficLights',                   // is4TrafficLights,
    '5Arrows', '5ArrowsGray',           // is5Arrows, is5ArrowsGray,
    '5Rating', '5Quarters',             // is5Rating, is5Quarters,
    '5Quarters'                         // REPLACEMENT FOR is5Boxes which needs x14
  );

function ColorToExcelStr(AColor: TsColor): String;
begin
  Result := ColorToHTMLColorStr(AColor, true);  // --> "00rrggbb"
  Result[1] := 'F';
  Result[2] := 'F';  // --> "FFrrggbb"
end;

{ The rels files store relative file paths (e.g. ../media/image1.png).
  This function makes sure that the file path begins with 'xl'. This filename
  can be used by the unzipper to extract the file from the xlsx (zip) archive. }
function MakeXLPath(AFileName: String): String;
begin
  Result := AFileName;
  if Length(AFileName) <= 3 then
    exit;
  
  if (Result[1] = '.') and (Result[2] = '.') then
  begin
    Result[1] := 'x';
    Result[2] := 'l';
  end else 
  if not ((Result[1] ='x') and (Result[2] = 'l') and (Result[3] = '/')) then
    Result := 'xl/' + AFileName;
end;
  
function RelsFileFor(AFileName: String): String;
var
  path: String;
  fn: String;
  p: Integer;
begin
  p := RPos('/', AFileName);
  path := Copy(AFileName, 1, p);
  fn := copy(AFileName, p+1, MaxInt);
  Result := path + '_rels/' + fn + '.rels';
end;

function StrToFillStyle(s: String): TsFillStyle;
var
  fs: TsFillStyle;
begin
  if s = '' then
  begin
    Result := fsSolidFill;
    exit;
  end;

  for fs in TsFillStyle do
    if PATTERN_TYPES[fs] = s then
    begin
      Result := fs;
      exit;
    end;
  Result := fsNoFill;
end;

function StrToLineStyle(s: String): TsLineStyle;
var
  ls: TsLineStyle;
begin
  if s = '' then
  begin
    Result := lsThin;
    exit;
  end;

  for ls in TsLineStyle do
    if LINESTYLE_TYPES[ls] = s then
    begin
      Result := ls;
      exit;
    end;
  Result := lsThin;
end;

function StrToCFValueKind(s: String): TsCFValueKind;
begin
  case s of
    'min', 'automin':
      Result := vkMin;
    'max', 'automax':
      Result := vkMax;
    'percent':
      Result := vkPercent;
    'percentile':
      Result := vkPercentile;
    'num':
      Result := vkValue;
    else
      Result := vkMin;
  end;
end;

function CFOperandToStr(v: Variant; const AFormatSettings: TFormatSettings): String;
const
  ERR = cardinal(-1);
var
  r, c: Cardinal;
  f: Double;
begin
  if VarIsFloat(v) then
  begin
    f := v;
    Result := Format('%g', [f], AFormatSettings);
  end else
  begin
    Result := VarToStr(v);
    if Result = '' then
      exit;
  end;

  if Result[1] = '=' then
    Delete(Result, 1, 1)
  else
  if ParseCellString(Result, r, c) and (r <> ERR) and (c <> ERR) then
    Result := GetCellString(r, c, [])
  else
  if VarIsStr(v) then
    Result := UTF8TextToXMLText(SafeQuoteStr(Result));
end;

procedure InitOOXMLLimitations(out ALimitations: TsSpreadsheetFormatLimitations);
begin
  // http://en.wikipedia.org/wiki/List_of_spreadsheet_software#Specifications
  ALimitations.MaxColCount := 16384;
  aLimitations.MaxRowCount := 1048576;
  ALimitations.MaxSheetNameLength := 31;
  // https://support.office.com/en-us/article/Excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3#ID0EBABAAA=Excel_2007
  ALimitations.MaxCharsInTextCell := 32767;
end;

function StrIsTrue(s: String): boolean;
begin
  Result := (s = '1') or (Lowercase(s) = 'true');
end;

function StrIsFalse(s: String): boolean;
begin
  Result := (s = '0') or (Lowercase(s) = 'false');
end;

function CF_ValueNode(AKind: TsCFValueKind; AValue: Double): String;
begin
  Result := '<cfvo';
  case AKind of
    vkMin        : Result := Result + ' type="min"';
    vkMax        : Result := Result + ' type="max"';
    vkPercent    : Result := Result + Format(' type="percent" val="%g"', [AValue]);
    vkPercentile : Result := Result + Format(' type="percentile" val="%g"', [AValue]);
    vkValue      : Result := Result + Format(' type="num" val="%g"', [AValue]);
  end;
  Result := Result + ' />';
end;

function CF_ColorNode(AColor: TsColor): String;
begin
  Result := Format('<color rgb="%s" />', [ColorToExcelStr(AColor)]);
end;


{ TXlsxRelationshipList }

destructor TXlsxRelationshipList.Destroy;
begin
  Clear;
  inherited;
end;

function TXlsxRelationshipList.Add(ARelID, ASchema, ATarget: String): TXlsxRelationship;
var
  rel: TXlsxRelationship;
  idx: Integer;
begin
  rel := TXlsxRelationship.Create;
  rel.RelID := ARelID;
  rel.Schema := ASchema;
  rel.Target := ATarget;
  idx := inherited Add(rel);
  Result := Items[idx];
end;
  
procedure TXlsxRelationshipList.Clear;
var
  j: Integer;
begin
  for j := 0 to Count-1 do Items[j].Free;
  inherited;
end;

procedure TXlsxRelationshipList.Delete(AIndex: Integer);
begin
  Items[AIndex].Free;
  inherited;
end;

function TXlsxRelationshipList.FindTarget(ARelID: String): String;
var
  i: Integer;
  rel: TXlsxRelationship;
begin
  for i := 0 to Count-1 do
  begin
    rel := Items[i];
    if (rel.RelID = ARelID) then
    begin
      Result := rel.Target;
      exit;
    end;
  end;
  Result := '';
end;
  
function TXlsxRelationshipList.GetItem(AIndex: Integer): TXlsxRelationship;
begin
  Result := TXlsxRelationship(inherited Items[AIndex]);
end;

procedure TXlsxRelationshipList.SetItem(AIndex: Integer; AValue: TXlsxRelationship);
begin
  inherited Items[AIndex] := AValue;
end;


{ TSheetData }

constructor TSheetData.Create;
begin
  inherited;
  SheetRels := TXlsxRelationshipList.Create;
  DrawingRels := TXlsxRelationshipList.Create;
  VmlDrawingRels := TXlsxRelationshipList.Create;
  HyperlinkRels := TXlsxRelationshipList.Create;
end;

destructor TSheetData.Destroy;
begin
  SheetRels.Free;
  DrawingRels.Free;
  VmlDrawingRels.Free;
  HyperlinkRels.Free;
  inherited;
end;

{------------------------------------------------------------------------------}
{                           TsSpreadOOXMLReader                                }
{------------------------------------------------------------------------------}

constructor TsSpreadOOXMLReader.Create(AWorkbook: TsBasicWorkbook);
begin
  inherited Create(AWorkbook);

 {$ifdef FPS_CHARTS}
  FChartReader := TsSpreadOOXMLChartReader.Create(self);
 {$endif}

  FDateMode := XlsxSettings.DateMode;

  FSharedStrings := TStringList.Create;
  FSheetList := TFPList.Create;
  FWorkbookRels := TXlsxRelationshipList.Create;
  FFillList := TFPList.Create;
  FBorderList := TFPList.Create;
  FHyperlinkList := TFPList.Create;
  FCellFormatList := TsCellFormatList.Create(true);
  FDifferentialFormatList := TFPList.Create;
  FDrawingToSheetRelList := TFPList.Create;
  FEmbeddedObjList := TFPList.Create;
  // Allow duplicates because xf indexes used in cell records cannot be found any more.
  FSharedFormulaBaseList := TFPList.Create;

  FPalette := TsPalette.Create;

  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';

  InitOOXMLLimitations(FLimitations);
end;

destructor TsSpreadOOXMLReader.Destroy;
var
  j: Integer;
begin
  for j := FFillList.Count-1 downto 0 do
    TObject(FFillList[j]).Free;
  FFillList.Free;

  for j := FBorderList.Count-1 downto 0 do
    TObject(FBorderList[j]).Free;
  FBorderList.Free;

  for j := FEmbeddedObjList.Count-1 downto 0 do
    TObject(FEmbeddedObjList[j]).Free;
  FEmbeddedObjList.Free;
  
  for j := FHyperlinkList.Count-1 downto 0 do
    TObject(FHyperlinkList[j]).Free;
  FHyperlinkList.Free;

  for j := FDifferentialFormatList.Count-1 downto 0 do
    TObject(FDifferentialFormatList[j]).Free;
  FDifferentialFormatList.Free;

  for j := FDrawingToSheetRelList.Count-1 downto 0 do
    TObject(FDrawingToSheetRelList[j]).Free;
  FDrawingToSheetRelList.Free;
  for j := FSheetList.Count-1 downto 0 do
    TObject(FSheetList[j]).Free;
  FSheetList.Free;

  FWorkbookRels.Free;

  for j := FSharedStrings.Count-1 downto 0 do
    FSharedStrings.Objects[j].Free;
  FSharedStrings.Free;

  for j := FSharedFormulaBaseList.Count-1 downto 0 do
    TObject(FSharedFormulaBaseList[j]).Free;
  FSharedFormulaBaseList.Free;

  // FCellFormatList, FNumFormatList and FFontList are destroyed by ancestor

  FPalette.Free;
  
 {$ifdef FPS_CHARTS}
  FChartReader.Free;
 {$endif}

  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Adds the built-in number formats to the NumFormatList.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLReader.AddBuiltinNumFormats;
begin
  FFirstNumFormatIndexInFile := 164;
  AddBuiltInBiffFormats(
    FNumFormatList, Workbook.FormatSettings, FFirstNumFormatIndexInFile-1
  );
end;

procedure TsSpreadOOXMLReader.ApplyCellFormatting(ACell: PCell; XFIndex: Integer);
var
  i: Integer;
  fmt: PsCellFormat;
begin
  if Assigned(ACell) then begin
    i := FCellFormatList.FindIndexOfID(XFIndex);
    fmt := FCellFormatList.Items[i];
    ACell^.FormatIndex := (FWorkbook as TsWorkbook).AddCellFormat(fmt^);
  end;
end;

procedure TsSpreadOOXMLReader.ApplyHyperlinks(AWorksheet: TsbasicWorksheet);
var
  i: Integer;
  hyperlinkData: THyperlinkListData;
  r1, c1, r2, c2, r, c: Cardinal;
  sheet: TsWorksheet;
begin
  sheet := AWorksheet as TsWorksheet;

  for i:=0 to FHyperlinkList.Count-1 do
  begin
    hyperlinkData := THyperlinkListData(FHyperlinkList.Items[i]);
    if hyperlinkData.Worksheet <> sheet then
      Continue;

    if pos(':', hyperlinkdata.CellRef) = 0 then
    begin
      ParseCellString(hyperlinkData.CellRef, r1, c1);
      r2 := r1;
      c2 := c1;
    end else
      ParseCellRangeString(hyperlinkData.CellRef, r1, c1, r2, c2);

    for r := r1 to r2 do
      for c := c1 to c2 do
        with hyperlinkData do
          if Target = '' then
            sheet.WriteHyperlink(r, c, '#'+TextMark, ToolTip)
          else
          if TextMark = '' then
            sheet.WriteHyperlink(r, c, Target, ToolTip)
          else
            sheet.WriteHyperlink(r, c, Target+'#'+TextMark, ToolTip);
  end;
end;

class function TsSpreadOOXMLReader.CheckFileFormat(AStream: TStream): Boolean;
begin
  Result := HasZipHeader(AStream);
end;

function TsSpreadOOXMLReader.CreateXMLStream: TStream;
begin
  if boFileStream in FWorkbook.Options then
    Result := TFileStream.Create(GetTempFileName, fmCreate)
  else
  if boBufStream in FWorkbook.Options then
    Result := TBufStream.Create(GetTempFileName, fmCreate)
  else
    Result := TMemoryStream.Create;
end;

{ Checks the file header for the signature of the decrypted file format. }
class function TsSpreadOOXMLReader.IsEncrypted(AStream: TStream): Boolean;
var
  p: Int64;
  buf: Cardinal;
begin
  p := AStream.Position;
  AStream.Position := 0;
  AStream.Read(buf, SizeOf(buf));
  Result := (buf = Cardinal(CFB_SIGNATURE));
  AStream.Position := p;
end;

{ Returns TRUE if the file is encrypted and requires a password. }
function TsSpreadOOXMLReader.NeedsPassword(AStream: TStream): Boolean;
begin
  Result := IsEncrypted(AStream);
end;

procedure TsSpreadOOXMLReader.ReadActiveSheet(ANode: TDOMNode;
  out ActiveSheetIndex: Integer);
var
  S: string;
  i: Integer;
  node: TDOMNode;
begin
  ActiveSheetIndex := -1;
  if ANode = nil then
    Exit;
  node := ANode.FindNode('workbookView');
  if node = nil then
    node := ANode.FindNode('x:workbookView');
  S := GetAttrValue(node, 'activeTab');
  if TryStrToInt(S, i) then
    ActiveSheetIndex := i;
end;

procedure TsSpreadOOXMLReader.ReadBorders(ANode: TDOMNode);
var
  borderNode: TDOMNode;
  edgeNode: TDOMNode;
  nodeName: String;
  borders: TsCellBorders;
  borderStyles: TsCellBorderStyles;
  borderData: TBorderListData;
  s: String;

begin
  if ANode = nil then
    exit;

  borderStyles := DEFAULT_BORDERSTYLES;
  borderNode := ANode.FirstChild;
  while Assigned(borderNode) do begin
    nodeName := borderNode.NodeName;
    if (nodeName = 'border') or (nodeName = 'x:border') then begin
      borders := [];
      s := GetAttrValue(borderNode, 'diagonalUp');
      if strIsTrue(s) then
        Include(borders, cbDiagUp);
      s := GetAttrValue(borderNode, 'diagonalDown');
      if StrIsTrue(s) then
        Include(borders, cbDiagDown);
      edgeNode := borderNode.FirstChild;
      while Assigned(edgeNode) do begin
        nodeName := edgeNode.NodeName;
        if (nodeName = 'left') or (nodeName = 'x:left') then begin
          if ReadBorderStyle(edgeNode, borderStyles[cbWest]) then
            Include(borders, cbWest);
        end
        else if (nodeName = 'right') or (nodeName = 'x:right') then begin
          if ReadBorderStyle(edgeNode, borderStyles[cbEast]) then
            Include(borders, cbEast);
        end
        else if (nodeName = 'top') or (nodeName = 'x:top') then begin
          if ReadBorderStyle(edgeNode, borderStyles[cbNorth]) then
            Include(borders, cbNorth);
        end
        else if (nodeName = 'bottom') or (nodename = 'x:bottom') then begin
          if ReadBorderStyle(edgeNode, borderStyles[cbSouth]) then
            Include(borders, cbSouth);
        end
        else if (nodeName = 'diagonal') or (nodeName = 'x:diagonal') then begin
          if ReadBorderStyle(edgeNode, borderStyles[cbDiagUp]) then
            borderStyles[cbDiagDown] := borderStyles[cbDiagUp];
        end;
        edgeNode := edgeNode.NextSibling;
      end;

      // add to border list
      borderData := TBorderListData.Create;
      borderData.Borders := borders;
      borderData.BorderStyles := borderStyles;
      FBorderList.Add(borderData);
    end;
    borderNode := borderNode.NextSibling;
  end;
end;

function TsSpreadOOXMLReader.ReadBorderStyle(ANode: TDOMNode;
  out ABorderStyle: TsCellBorderStyle): Boolean;
var
  s: String;
  colorNode: TDOMNode;
  nodeName: String;
begin
  Result := false;
  ABorderStyle.LineStyle := lsThin;
  ABorderStyle.Color := scBlack;

  s := GetAttrValue(ANode, 'style');
  if (s = '') or (s = 'none') then
    exit;

  ABorderStyle.LineStyle := StrToLineStyle(s);
            {
  if s = 'thin' then
    ABorderStyle.LineStyle := lsThin
  else if s = 'medium'then
    ABorderStyle.LineStyle := lsMedium
  else if s = 'thick' then
    ABorderStyle.LineStyle := lsThick
  else if s = 'dotted' then
    ABorderStyle.LineStyle := lsDotted
  else if s = 'dashed' then
    ABorderStyle.LineStyle := lsDashed
  else if s = 'double' then
    ABorderStyle.LineStyle := lsDouble
  else if s = 'hair' then
    ABorderStyle.LineStyle := lsHair
  else if s = 'dashDot' then
    ABorderStyle.LineStyle := lsDashDot
  else if s = 'dashDotDot' then
    ABorderStyle.LineStyle := lsDashDotDot
  else if s = 'mediumDashed' then
    ABorderStyle.LineStyle := lsMediumDash
  else if s = 'mediumDashDot' then
    ABorderStyle.LineSTyle := lsMediumDashDot
  else if s = 'mediumDashDotDot' then
    ABorderStyle.LineStyle := lsMediumDashDotDot
  else if s = 'slantDashDot' then
    ABorderStyle.LineStyle := lsSlantDashDot;
    }

  colorNode := ANode.FirstChild;
  while Assigned(colorNode) do begin
    nodeName := colorNode.NodeName;
    if (nodeName = 'color') or (nodename = 'x:color') then
      ABorderStyle.Color := ReadColor(colorNode);
    colorNode := colorNode.NextSibling;
  end;
  Result := true;
end;

procedure TsSpreadOOXMLReader.ReadCell(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARowIndex: Cardinal; var AColIndex: Cardinal);
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  addr, s, typeStr: String;
  rowIndex, colIndex: Cardinal;
  cell: PCell;
  lCell: TCell;
  sharedFormulabase: TSharedFormulaData;
  datanode, tnode, rnode: TDOMNode;
  dataStr: String;
  formulaStr: String;
  formula: PsFormula;
  nodeName: String;
  sstIndex: Integer;
  number: Double;
  fmt: TsCellFormat;
  numFmt: TsNumFormatParams = nil;
  ms: TMemoryStream;
  n: Integer;
begin
  if ANode = nil then
    exit;

  book := FWorkbook as TsWorkbook;
  sheet := AWorksheet as TsWorksheet;

  // get row and column address
  addr := GetAttrValue(ANode, 'r');       // cell address, like 'A1'
  if addr <> '' then
    ParseCellString(addr, rowIndex, colIndex)
  else
  begin
    rowIndex := ARowIndex;
    colIndex := AColIndex;
  end;
  AColIndex := colIndex + 1;

  // create cell
  if FIsVirtualMode then
  begin
    InitCell(FWorksheet, rowIndex, colIndex, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := sheet.AddCell(rowIndex, colIndex);

  // get style index
  s := GetAttrValue(ANode, 's');
  if s <> '' then begin
    ApplyCellFormatting(cell, StrToInt(s));
    fmt := book.GetCellFormat(cell^.FormatIndex);
  end else
  begin
    InitFormatRecord(fmt);
    cell^.FormatIndex := 0;
  end;

  // Get cell type
  typeStr := GetAttrValue(ANode, 't');

  // get number format parameters
  numFmt := book.GetNumberFormat(fmt.NumberFormatIndex);

  // get data
  datanode := ANode.FirstChild;
  dataStr := '';
  formulaStr := '';
  while Assigned(datanode) do
  begin
    nodeName := datanode.NodeName;
    if (nodeName = 'v') or (nodeName = 'x:v') then
      dataStr := GetNodeValue(datanode)
    else
    if (nodeName = 'is') or (nodeName = 'x:is') then
    begin
      tnode := datanode.FirstChild;
      while Assigned(tnode) do begin
        nodename := tnode.NodeName;
        if (nodename = 'r') and (typeStr = 'inlineStr') then begin
          rnode := tnode.FirstChild;
          while Assigned(rnode) do begin
            nodename := rnode.NodeName;
            if (nodename = 't') or (nodeName = 'x:t') then
              dataStr := dataStr + GetNodeValue(rnode);
            rnode := rnode.NextSibling;
          end;
        end else
        if (nodename = 't') or (nodeName = 'x:t') then
          dataStr := dataStr + GetNodeValue(tnode);
        tnode := tnode.NextSibling;
      end;
    end;

    if (boReadFormulas in book.Options) and ((nodeName = 'f') or (nodeName = 'x:f'))then
    begin
      // Formula to cell
      formulaStr := GetNodeValue(datanode);
      try
        s := GetAttrValue(datanode, 't');
        if s = 'shared' then
        begin
          // Shared formula
          s := GetAttrValue(datanode, 'ref');
          if (s <> '') then      // This defines the shared formula range
          begin
            sheet.WriteFormula(cell, formulaStr);
            // We store the shared formula base in the SharedFormulaBaseList.
            // The list index is identical with the 'si' attribute of the node.
            sharedformulabase := TSharedFormulaData.Create;
            sharedformulabase.Worksheet := sheet;
            sharedformulabase.Row := rowindex;
            sharedformulabase.Col := colindex;
            sharedformulabase.Formula := formulaStr;
            FSharedFormulaBaseList.Add(sharedformulabase);
            sheet.WriteFormula(rowindex, colindex, formulaStr);
          end else
          begin
            // Get index into the SharedFormulaBaseList...
            s := GetAttrValue(datanode, 'si');
            if s <> '' then
            begin
              sharedformulabase := TSharedFormulaData(FSharedFormulaBaseList[StrToInt(s)]);
              // ... and copy shared formula to destination cell
              formula := sharedFormulaBase.Worksheet.Formulas.FindFormula(
                sharedFormulabase.Row, sharedFormulaBase.Col);
              InitCell(FWorksheet, sharedformulabase.Row, sharedformulabase.Col, lCell);
              sheet.UseFormulaInCell(@lCell, formula);
              sheet.CopyFormula(@lCell, cell);
            end;
          end;
        end
        else
          // "Normal" formula
          sheet.WriteFormula(cell, formulaStr);
      except
        on E:EExprParser do begin
          FWorkbook.AddErrorMsg(E.Message);
          if (boAbortReadOnFormulaError in Workbook.Options) then raise;
        end;
        on E:ECalcEngine do begin
          FWorkbook.AddErrorMsg(E.Message);
          if (boAbortReadOnFormulaError in Workbook.Options) then raise;
        end;
      end;
    end;
    datanode := datanode.NextSibling;
  end;

  book.LockFormulas;  // Protect formulas from being deleted by the WriteXXXX calls
  try
    // get data type
    s := GetAttrValue(ANode, 't');   // "t" = data type
    if (s = '') and (dataStr = '') then
      sheet.WriteBlank(cell, true)     // true --> do not erase the formula!!!
    else
    if (s = '') or (s = 'n') then begin
      // Number or date/time, depending on format
      if TryStrToFloat(dataStr, number, FPointSeparatorSettings) then
      begin
        if IsDateTimeFormat(numFmt) then
        begin
          if not IsTimeIntervalFormat(numFmt) then
            // no correction of time origin for "time" or "time interval" format
            number := ConvertExcelDateTimeToDateTime(number, FDateMode);
          sheet.WriteDateTime(cell, number);
        end
        else if IsTextFormat(numFmt) then
          sheet.WriteText(cell, dataStr)
        else
          sheet.WriteNumber(cell, number);
      end else
        workbook.AddErrorMsg(
          'Error reading cell %s: Failure to convert "%s" to number.', [
          GetCellString(rowindex, colindex), dataStr
        ]);
    end
    else
    if s = 's' then begin
      // String from shared strings table
      if TryStrToInt(dataStr, sstIndex) then
      begin
        sheet.WriteText(cell, FSharedStrings[sstIndex]);
        // Read rich-text parameters from the stream stored in the Objects of the stringlist
        if FSharedStrings.Objects[sstIndex] <> nil then
        begin
          ms := TMemoryStream(FSharedStrings.Objects[sstIndex]);
          ms.Position := 0;
          n := ms.ReadWord;   // Count of array elements
          SetLength(cell^.RichTextParams, n);
          ms.ReadBuffer(cell^.RichTextParams[0], n*SizeOf(TsRichTextParam));
        end;
      end else
        workbook.AddErrorMsg(
          'Error readind cell %s: Failure to extract SST index from value "%s"', [
          GetCellString(rowindex, colindex), dataStr
        ]);
    end else
    if (s = 'str') or (s = 'inlineStr') then begin
      // literal string
      sheet.WriteText(cell, datastr);
    end else
    if s = 'b' then
      // boolean
      sheet.WriteBoolValue(cell, dataStr='1')
    else
    if s = 'e' then begin
      // error value
      if dataStr = '#NULL!' then
        sheet.WriteErrorValue(cell, errEmptyIntersection)
      else if dataStr = '#DIV/0!' then
        sheet.WriteErrorValue(cell, errDivideByZero)
      else if dataStr = '#VALUE!' then
        sheet.WriteErrorValue(cell, errWrongType)
      else if dataStr = '#REF!' then
        sheet.WriteErrorValue(cell, errIllegalRef)
      else if dataStr = '#NAME?' then
        sheet.WriteErrorValue(cell, errWrongName)
      else if dataStr = '#NUM!' then
        sheet.WriteErrorValue(cell, errOverflow)
      else if dataStr = '#N/A' then
        sheet.WriteErrorValue(cell, errArgError)
      else if dataStr = '' then
        // rare case...
        // see http://forum.lazarus.freepascal.org/index.php/topic,38726.0.html
        sheet.WriteBlank(cell)
      else
        raise EFPSpreadsheetReader.Create(rsUnknownErrorType);


    end else
      raise EFPSpreadsheetReader.Create(rsUnknownDataType);

    if FIsVirtualMode then
      book.OnReadCellData(book, rowIndex, colIndex, cell);

  finally
    book.UnlockFormulas;
  end;
end;

procedure TsSpreadOOXMLReader.ReadCellXfs(ANode: TDOMNode);
var
  node: TDOMNode;
  childNode: TDOMNode;
  nodeName: String;
  fmt: TsCellFormat;
  fs: TsFillStyle;
  s1, s2: String;
  numFmtIndex, fillIndex, borderIndex: Integer;
  numFmtStr: String;
  numFmtParams: TsNumFormatParams;
  fillData: TFillListData;
  borderData: TBorderListData;
  fnt: TsFont;
  cp: TsCellProtections;
  book: TsWorkbook;
begin
  if ANode = nil then
    exit;

  book := FWorkbook as TsWorkbook;

  node := ANode.FirstChild;
  while Assigned(node) do
  begin
    nodeName := node.NodeName;
    if (nodeName = 'xf') or (nodeName = 'x:xf') then
    begin
      InitFormatRecord(fmt);
      fmt.ID := FCellFormatList.Count;
      fmt.Name := '';

      // strange: sometimes the "apply*" are missing. Therefore, it may be better
      // to check against "<>0" instead of "=1"
      s1 := GetAttrValue(node, 'numFmtId');
      s2 := GetAttrValue(node, 'applyNumberFormat');
      if (s1 <> '') and not StrIsFalse(s2) then //(s2 <> '0') and (s2 <> 'false') then
      begin
        numFmtIndex := StrToInt(s1);
        numFmtStr := NumFormatList[numFmtIndex];
        if SameText(numFmtStr, 'General') then
          numFmtParams := nil
        else
        begin
          fmt.NumberFormatIndex := book.AddNumberFormat(numFmtStr);
          numFmtParams := book.GetNumberFormat(fmt.NumberFormatIndex);
        end;
        if numFmtParams <> nil then
        begin
          fmt.NumberFormat := numFmtParams.NumFormat;
          fmt.NumberFormatStr := numFmtStr;
          Include(fmt.UsedFormattingFields, uffNumberFormat);
        end else
        begin
          fmt.NumberFormat := nfGeneral;
          fmt.NumberFormatStr := '';
          Exclude(fmt.UsedFormattingFields, uffNumberFormat);
        end;
      end;

      s1 := GetAttrValue(node, 'fontId');
      s2 := Lowercase(GetAttrValue(node, 'applyFont'));
      if (s1 <> '') and not StrIsFalse(s2) then //(s2 <> '0') and (s2 <> 'false') then
      begin
        fnt := TsFont(FFontList.Items[StrToInt(s1)]);
        fmt.FontIndex := book.FindFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
        if fmt.FontIndex = -1 then
          fmt.FontIndex := book.AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
        if fmt.FontIndex > 0 then
          Include(fmt.UsedFormattingFields, uffFont);
      end;

      s1 := GetAttrValue(node, 'fillId');
      s2 := Lowercase(GetAttrValue(node, 'applyFill'));
      if (s1 <> '') and not StrIsFalse(s2) then //(s2 <> '0') and (s2 <> 'false') then
      begin
        fillIndex := StrToInt(s1);
        fillData := TFillListData(FFillList[fillIndex]);
        if (fillData <> nil) and (fillData.PatternType <> 'none') then begin
          fmt.Background.FgColor := fillData.FgColor;
          fmt.Background.BgColor := fillData.BgColor;
          for fs in TsFillStyle do
            if SameText(fillData.PatternType, PATTERN_TYPES[fs]) then
            begin
              fmt.Background.Style := fs;
              Include(fmt.UsedFormattingFields, uffBackground);
              break;
            end;
        end;
      end;

      s1 := GetAttrValue(node, 'borderId');
      s2 := Lowercase(GetAttrValue(node, 'applyBorder'));
      if (s1 <> '') and not StrIsFalse(s2) then //(s2 <> '0') and (s2 <> 'false') then
      begin
        borderIndex := StrToInt(s1);
        borderData := TBorderListData(FBorderList[borderIndex]);
        if (borderData <> nil) then
        begin
          fmt.BorderStyles := borderData.BorderStyles;
          fmt.Border := borderData.Borders;
        end;
      end;

      s2 := Lowercase(GetAttrValue(node, 'applyAlignment'));
      if StrIsTrue(s2) then  // if (s2 <> '0') and (s2 <> '') and (s2 <> 'false') then
      begin
        childNode := node.FirstChild;
        while Assigned(childNode) do begin
          nodeName := childNode.NodeName;
          if (nodeName = 'alignment') or (nodeName = 'x:alignment') then begin
            s1 := GetAttrValue(childNode, 'horizontal');
            case s1 of
              'left'       : fmt.HorAlignment := haLeft;
              'center'     : fmt.HorAlignment := haCenter;
              'right'      : fmt.HorAlignment := haRight;
              'justify'    : fmt.HorAlignment := haJustified;
              'distributed': fmt.HorAlignment := haDistributed;
              'fill'       : fmt.HorAlignment := haFilled;
              else           fmt.HorAlignment := haDefault;
            end;

            s1 := GetAttrValue(childNode, 'vertical');
            case s1 of
              'top': fmt.VertAlignment := vaTop;
              'center': fmt.VertAlignment := vaCenter;
              'bottom': fmt.VertAlignment := vaBottom;
              else fmt.VertAlignment := vaDefault;
            end;

            s1 := GetAttrValue(childNode, 'readingOrder');
            if (s1 = '1') or (s1 = '2') then
              fmt.BiDiMode := TsBiDiMode(StrToInt(s1));

            s1 := GetAttrValue(childNode, 'wrapText');
            if StrIsTrue(s1) then //(s1 <> '') and (s1 <> '0') and (s1 <> 'false') then
              Include(fmt.UsedFormattingFields, uffWordWrap);

            s1 := GetAttrValue(childNode, 'textRotation');
            if s1 = '90' then
              fmt.TextRotation := rt90DegreeCounterClockwiseRotation
            else
            if s1 = '180' then
              fmt.TextRotation := rt90DegreeClockwiseRotation
            else
            if s1 = '255' then
              fmt.TextRotation := rtStacked
            else
              fmt.TextRotation := trHorizontal;
          end;
          childNode := childNode.NextSibling;
        end;
      end;

      // protection
      s2 := Lowercase(GetAttrValue(node, 'applyProtection'));
      if StrIsTrue(s2) then //(s2 <> '0') and (s2 <> '') and (s2 <> 'false') then
      begin
        cp := [cpLockCell];
        childNode := node.FirstChild;
        while Assigned(childNode) do begin
          nodeName := childNode.NodeName;
          if (nodeName = 'protection') or (nodeName = 'x:protection') then
          begin
            s1 := GetAttrValue(childNode, 'locked');
            if StrIsFalse(s1) then
              Exclude(cp, cpLockCell);
            s1 := GetAttrValue(childNode, 'hidden');
            if StrIsTrue(s1) then
              Include(cp, cpHideFormulas);
          end;
          childNode := childNode.NextSibling;
        end;
        fmt.Protection := cp;
      end;

      if fmt.FontIndex > 0 then
        Include(fmt.UsedFormattingFields, uffFont);
      if fmt.Border  <> [] then
        Include(fmt.UsedFormattingFields, uffBorder);
      if fmt.HorAlignment <> haDefault then
        Include(fmt.UsedFormattingFields, uffHorAlign);
      if fmt.VertAlignment <> vaDefault then
        Include(fmt.UsedFormattingFields, uffVertAlign);
      if fmt.TextRotation <> trHorizontal then
        Include(fmt.UsedFormattingFields, uffTextRotation);
      if fmt.BiDiMode <> bdDefault then
        Include(fmt.UsedFormattingFields, uffBiDi);
      if fmt.Protection <> DEFAULT_CELL_PROTECTION then
        Include(fmt.UsedFormattingFields, uffProtection);
      FCellFormatList.Add(fmt);
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadCFAverage(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARange: TsCellRange; AFormatIndex: Integer);
var
  sStdDev, sEquAve, sAboveAve: String;
  condition: TsCFCondition;
  stdDev: Double;
  sheet: TsWorksheet;
begin
  sheet := TsWorksheet(AWorksheet);

  sEquAve := GetAttrValue(ANode, 'equalAverage');
  sAboveAve := GetAttrValue(ANode, 'aboveAverage');

  if (sEquAve='1') then
  begin
    if (sAboveAve='0') then
      condition := cfcBelowEqualAverage
    else
      condition := cfcAboveEqualAverage;
  end else
  begin
    if (sAboveAve = '0') then
      condition := cfcBelowAverage
    else
      condition := cfcAboveAverage;
  end;

  sStdDev := GetAttrValue(ANode, 'stdDev');
  if not TryStrToFloat(sStdDev, stdDev, FPointSeparatorSettings) then
    sStdDev := '';     // This will omit the "stdDev" attribute

  if sStdDev = '' then
    sheet.WriteConditionalCellFormat(ARange, condition, AFormatIndex)
  else
    sheet.WriteConditionalCellFormat(ARange, condition, stdDev, AFormatIndex);
end;


procedure TsSpreadOOXMLReader.ReadCFCellFormat(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARange: TsCellRange; AFormatIndex: Integer);
var
  nodeName: String;
  sOp: String;
  sFormula: Array of String = nil;
  cf: TsCFCondition;
  found: Boolean;
  i: Integer;
  n: Integer;
  x: Double;
  r, c: Cardinal;
  condition: TsCFCondition;
  values: array of Variant = nil;
  sheet: TsWorksheet;
begin
  sheet := TsWorksheet(AWorksheet);

  sOp := GetAttrValue(ANode, 'operator');
  found := false;
  for cf in TsCFCondition do
    if sOp = CF_OPERATOR_NAMES[cf] then
    begin
      found := true;
      condition := cf;
      break;
    end;
  if (not found) or not (condition in [cfcEqual..cfcNotBetween]) then
    exit;

  // Process cases related to (un)equality of cells
  SetLength(sFormula, 0);
  ANode := ANode.FirstChild;
  while (ANode <> nil) do
  begin
    nodeName := ANode.NodeName;
    if nodeName = 'formula' then
    begin
      SetLength(sFormula, Length(sFormula) + 1);
      sFormula[High(sFormula)] := GetNodeValue(ANode);
    end;
    ANode := ANode.NextSibling;
  end;

  SetLength(values, Length(sFormula));
  for i := 0 to High(sFormula) do begin
    values[i] := sFormula[i];
    if (sFormula[i] <> '') then begin
      if TryStrToInt(sFormula[i], n) then
        values[i] := n
      else if TryStrToFloat(sFormula[i], x, FPointSeparatorSettings) then
        values[i] := x
      else if sFormula[i][1] = '"' then
        values[i] := sFormula[i]
      else if ParseCellString(sFormula[i], r, c) then
        values[i] := sFormula[i]
      else
        values[i] := '=' + sFormula[i];
    end;
  end;

  case Length(values) of
    0: sheet.WriteConditionalCellFormat(ARange, condition, AFormatIndex);
    1: sheet.WriteConditionalCellFormat(ARange, condition, values[0], AFormatIndex);
    2: sheet.WriteConditionalCellFormat(ARange, condition, values[0], values[1], AFormatIndex);
  end;
end;

procedure TsSpreadOOXMLReader.ReadCFColorRange(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARange: TsCellRange);
{ <colorScale>
    <cfvo type="num" val="2" />
    <cfvo type="percentile" val="50" />
    <cfvo type="max" />
    <color rgb="FF63BE7B" />
    <color rgb="FFFFEB84" />
    <color rgb="FFF8696B" />
  </colorScale>
}
var
  sheet: TsWorksheet;
  nodeName, s: String;
  vk: array[0..2] of TsCFValueKind = (vkNone, vkNone, vkNone);
  v: array[0..2] of Double = (NaN, NaN, NaN);
  c: array[0..2] of TsColor = (scNotDefined, scNotDefined, scNotDefined);
  x: Double;
  iv, ic: Integer;
  clevels, vlevels: Integer;
begin
  iv := 0;
  ic := 0;
  ANode := ANode.FirstChild;
  if (ANode <> nil) and (ANode.NodeName = 'colorScale') then
    ANode := ANode.FirstChild;
  while (ANode <> nil) do
  begin
    nodeName := ANode.NodeName;
    if (nodeName = 'color') and (ic <= High(c)) then
    begin
      c[ic] := ReadColor(ANode);
      inc(ic);
    end;
    if (nodeName = 'cfvo') and (iv <= High(vk)) then
    begin
      s := GetAttrValue(ANode, 'type');
      vk[iv] := StrToCFValueKind(s);
      s := GetAttrValue(ANode, 'val');
      if TryStrToFloat(s, x, FPointSeparatorSettings) then
        v[iv] := x
      else
        v[iv] := 0.0;
      inc(iv);
    end;
    ANode := ANode.NextSibling;
  end;

  clevels := 0;
  for ic := 0 to High(c) do
  begin
    if c[ic] = scNotDefined then
      break;
    inc(cLevels);
  end;

  vlevels := 0;
  for iv := 0 to High(vk) do
  begin
    if vk[iv] = vkNone then
      break;
    inc(vlevels);
  end;

  // Not 100% sure, but I guess there must be as many colors as value kinds.
  if vlevels <> cLevels then
  begin
    FWorkbook.AddErrorMsg('ColorRange: colors-levels mismatch');
    exit;
  end;

  sheet := TsWorksheet(AWorksheet);
  case clevels of
    1: exit;
    2: sheet.WriteColorRange(ARange, c[0],vk[0],v[0], c[1],vk[1],v[1]);
    3: sheet.WriteColorRange(ARange, c[0],vk[0],v[0], c[1],vk[1],v[1], c[2],vk[2],v[2]);
  end;
end;

procedure TsSpreadOOXMLReader.ReadCFDataBars(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARange: TsCelLRange);
{ <dataBar>
    <cfvo type="num" val="2" />
    <cfvo type="percent" val="80" />
    <color rgb="FFFF0000" />
  </dataBar>   } // We do not support the x14 namespace extension ATM
var
  sheet: TsWorksheet;
  nodeName, s: String;
  vk: array[0..1] of TsCFValuekind;
  v: array[0..1] of Double;
  x: Double;
  idx: Integer;
  clr: TsColor;
begin
  idx := 0;
  ANode := ANode.FirstChild;
  if (ANode <> nil) and (ANode.NodeName = 'dataBar') then
    ANode := ANode.FirstChild;
  while ANode <> nil do
  begin
    nodeName := ANode.NodeName;
    if nodeName = 'color' then
      clr := ReadColor(ANode)
    else
      clr := scBlue;

    if (nodeName = 'cfvo') and (idx <= 1) then
    begin
      s := GetAttrValue(ANode, 'type');
      vk[idx] := StrToCFValueKind(s);

      s := GetAttrValue(ANode, 'val');
      if TryStrToFloat(s, x, FPointSeparatorSettings) then
        v[idx] := x
      else
        v[idx] := NaN;

      inc(idx);
    end;

    ANode := ANode.NextSibling;
  end;

  // Check for value attribute. If not existing, use simple min/max scaling and log error.
  if ((vk[0] in [vkPercent, vkPercentile, vkValue]) and IsNaN(v[0])) or
     ((vk[1] in [vkPercent, vkPercentile, vkValue]) and IsNaN(v[1])) then
  begin
    FWorkbook.AddErrorMsg('DataBars: value needed.');
    vk[0] := vkMin;
    vk[1] := vkMax;
  end;

  sheet := TsWorksheet(AWorksheet);
  sheet.WriteDataBars(ARange, clr, vk[0], v[0], vk[1], v[1]);
end;

procedure TsSpreadOOXMLReader.ReadCFDateFormat(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARange: TsCellRange; AFormatIndex: Integer);
var
  sheet: TsWorksheet;
  s: String;
  cond: TsCFCondition;
begin
  if ANode = nil then
    exit;

  sheet := TsWorksheet(AWorksheet);
  s := GetAttrValue(ANode, 'timePeriod');

  for cond in [cfcYesterday..cfcNextMonth] do
    if CF_OPERATOR_NAMES[cond] = s then
    begin
      sheet.WriteConditionalCellFormat(ARange, cond, AFormatIndex);
      exit;
    end;
end;

procedure TsSpreadOOXMLReader.ReadCFExpression(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARange: TsCellRange; AFormatIndex: Integer);
var
  sheet: TsWorksheet;
  nodeName: String;
  s, s1: String;
  c: TsCFCondition;
  firstCellInRange: String;
begin
  sheet := TsWorksheet(AWorksheet);
  firstCellInRange := GetCellString(ARange.Row1, ARange.Col1);

  ANode := ANode.FirstChild;
  while ANode <> nil do
  begin
    nodeName := ANode.NodeName;
    if nodeName = 'formula' then
    begin
      s := GetNodeValue(ANode);

      for c in [cfcLastYear..cfcNextYear] do
      begin
        s1 := Format(CF_YEAR_FORMULAS[c], [firstCellInRange]);
        if s1 = s then begin
          sheet.WriteConditionalCellFormat(ARange, c, AFormatIndex);
          exit;
        end;
      end;

      sheet.WriteConditionalCellFormat(ARange, cfcExpression, s, AFormatIndex);
      exit;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadCFIconSet(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARange: TsCellRange);
var
  sheet: TsWorksheet;
  s: String;
  sIconSet: String;
  iconSet: TsCFIconSet;
  found: Boolean;
  tmpIconSet: TsCFIconSet;
  nodeName: String;
  iv: Integer;
  vk: array of TsCFValueKind = nil;
  v: array of Double = nil;
  n: Integer;
  x: Double;
begin
  ANode := ANode.FirstChild;
  if (ANode <> nil) and (ANode.NodeName = 'iconSet') then
  begin
    sIconSet := GetAttrValue(ANode, 'iconSet');
    found := false;
    for tmpIconSet in TsCFIconSet do
      if 'is' + sIconSet = GetEnumName(typeInfo(TsCFIconSet), integer(tmpIconSet)) then
      begin
        iconSet := tmpIconSet;
        found := true;
        break;
      end;

    if (not found) or (sIconSet = '') then
      Exit;

    // Determine icon count from name of icon set
    n := GetCFIconCount(iconSet);
    if (n < 3) or (n > 5) then  // only 3, 4 or 5 icons allowed
      Exit;

    SetLength(v, n);
    SetLength(vk, n);
    iv := 0;
    ANode := ANode.FirstChild;
    while (ANode <> nil) do
    begin
      nodeName := ANode.NodeName;
      if (nodeName = 'cfvo') and (iv <= High(vk)) then
      begin
        s := GetAttrValue(ANode, 'type');
        vk[iv] := StrToCFValueKind(s);
        s := GetAttrValue(ANode, 'val');
        if TryStrToFloat(s, x, FPointSeparatorSettings) then
          v[iv] := x
        else
          v[iv] := 0.0;
        inc(iv);
        if iv >= n then
          break;
      end;
      ANode := ANode.NextSibling;
    end;

    sheet := TsWorksheet(AWorksheet);
    // Ignore the first value because it is always 0
    case n of
      3: sheet.WriteIconSet(ARange, iconSet, vk[1], v[1], vk[2], v[2]);
      4: sheet.WriteIconSet(ARange, iconSet, vk[1], v[1], vk[2], v[2], vk[3], v[3]);
      5: sheet.WriteIconSet(ARange, iconSet, vk[1], v[1], vk[2], v[2], vk[3], v[3], vk[4], v[4]);
    end;

    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadCFMisc(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARange: TsCellRange; AFormatIndex: Integer);
var
  sheet: TsWorksheet;
  s: String;
  cf, condition: TsCFCondition;
  found: Boolean;
begin
  sheet := TsWorksheet(AWorksheet);

  found := false;
  s := GetAttrValue(ANode, 'type');
  for cf in TsCFCondition do
    if CF_TYPE_NAMES[cf] = s then
    begin
      condition := cf;
      found := true;
      break;
    end;
  if not found then
    exit;

  if condition in [cfcBeginsWith..cfcNotContainsText] then
  begin
    s := GetAttrValue(ANode, 'text');
    sheet. WriteConditionalCellFormat(ARange, condition, s, AFormatIndex);
  end else
    sheet. WriteConditionalCellFormat(ARange, condition, AFormatIndex);
end;

procedure TsSpreadOOXMLReader.ReadCFTop10(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARange: TsCellRange; AFormatIndex: Integer);
var
  sheet: TsWorksheet;
  rank: Integer;
  s, sPercent, sBottom: String;
  condition: TsCFCondition;
begin
  sheet := TsWorksheet(AWorksheet);

  s := GetAttrValue(ANode, 'rank');
  if not TryStrToInt(s, rank) then
    rank := 10;

  sPercent := GetAttrValue(ANode, 'percent');
  sBottom := GetAttrValue(ANode, 'bottom');
  if (sBottom = '1') then
  begin
    if sPercent = '1' then
      condition := cfcBottomPercent
    else
      condition := cfcBottom;
  end else
  begin
    if sPercent = '1' then
      condition := cfcTopPercent
    else
      condition := cfcTop;
  end;

  sheet.WriteConditionalCellFormat(ARange, condition, rank, AFormatIndex);
end;

procedure TsSpreadOOXMLReader.ReadColRowBreaks(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  s: String;
  sheet: TsWorksheet absolute AWorksheet;
  node: TDOMNode;
  nodeName: String;
  n: Integer;
  isCol: Boolean;
begin
  if (ANode = nil) or (AWorksheet = nil) then     // just to make sure...
    exit;

  nodeName := ANode.NodeName;
  isCol := nodeName = 'colBreaks';
  node := ANode.FirstChild;
  while node <> nil do
  begin
    nodeName := node.NodeName;
    if nodeName = 'brk' then begin
      s := GetAttrValue(node, 'id');
      if (s <> '') and TryStrToInt(s, n) then
      begin
        s := GetAttrValue(node, 'man');
        if s = '1' then
        begin
          if isCol then
            sheet.AddPageBreakToCol(n)
          else
            sheet.AddPageBreakToRow(n);
        end;
      end;
    end;
    node := node.NextSibling;
  end;
end;

function TsSpreadOOXMLReader.ReadColor(ANode: TDOMNode): TsColor;
var
  s: String;
  rgb: TsColor;
  idx: Integer;
  tint: Double;
  n, i: Integer;
  nodename: String;
begin
  Result := scNotDefined;
  Assert(ANode <> nil);

  s := Lowercase(GetAttrValue(ANode, 'auto'));
  if (s = '1') or (s = 'true') then begin
    nodename := ANode.NodeName;
    if (nodeName = 'fgColor') or (nodename = 'x:fgColor') then
      Result := scBlack
    else
      Result := scTransparent;
    exit;
  end;

  s := GetAttrValue(ANode, 'rgb');
  if s <> '' then begin
    if s[1] = '#' then
      Result := HTMLColorStrToColor(s)
    else
      Result := HTMLColorStrToColor('#' + s);
    exit;
  end;

  s := GetAttrValue(ANode, 'indexed');
  if s <> '' then begin
    i := StrToInt(s);
    n := FPalette.Count;
    if (i <= LAST_PALETTE_INDEX) and (i < n) then
    begin
      Result := FPalette[i];
      exit;
    end;
    // System colors
    // taken from OpenOffice docs
    case i of
      $0040: Result := scBlack;  // Default border color
      $0041: Result := scWhite;  // Default background color
      $0043: Result := scGray;   // Dialog background color
      $004D: Result := scBlack;  // Text color, chart border lines
      $004E: Result := scGray;   // Background color for chart areas
      $004F: Result := scBlack;  // Automatic color for chart border lines
      $0050: Result := scBlack;  // ???
      $0051: Result := scBlack;  // ??
      $7FFF: Result := scBlack;  // Automatic text color
      else   Result := scBlack;
    end;
    exit;
  end;

  s := GetAttrValue(ANode, 'theme');
  if s <> '' then begin
    idx := StrToInt(s);
    if idx < Length(FThemeColors) then begin
      // For some reason the first two pairs of colors are interchanged in Excel!
      case idx of
        0: idx := 1;
        1: idx := 0;
        2: idx := 3;
        3: idx := 2;
      end;
      rgb := FThemeColors[idx];
      s := GetAttrValue(ANode, 'tint');
      if s <> '' then begin
        tint := StrToFloat(s, FPointSeparatorSettings);
        rgb := TintedColor(rgb, tint);
      end;
      Result := rgb;
      exit;
    end;
  end;

  Result := scBlack;
end;

procedure TsSpreadOOXMLReader.ReadCols(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
const
  EPS = 1e-3;
var
  colNode: TDOMNode;
  col, col1, col2: Cardinal;  // column indexes
  lCol: TCol;                 // column record
  w: Double;
  s: String;
  nodeName: String;
  idx: Integer;
  fmt: PsCellFormat;
  book: TsWorkbook;
  sheet: TsWorksheet;
begin
  if ANode = nil then
    exit;

  book := FWorkbook as TsWorkbook;
  sheet := AWorksheet as TsWorksheet;

  colNode := ANode.FirstChild;
  while Assigned(colNode) do begin
    nodeName := colNode.NodeName;
    if (nodename = 'col') or (nodeName = 'x:col') then
    begin
      s := GetAttrValue(colNode, 'min');
      if s <> '' then col1 := StrToInt(s)-1 else col1 := 0;

      s := GetAttrValue(colNode, 'max');
      if s <> '' then col2 := StrToInt(s)-1 else col2 := col1;

      s := GetAttrValue(colNode, 'customWidth');
      if StrIsTrue(s) then  //(s = '1') or (s = 'true') then
      begin
        s := GetAttrValue(colNode, 'width');
        if (s <> '') and TryStrToFloat(s, w, FPointSeparatorSettings) then
        begin
          if SameValue(w, sheet.ReadDefaultColWidth(suChars), EPS) then  // is this needed?
            lCol.ColWidthType := cwtDefault
          else
            lCol.ColWidthType := cwtCustom;
          lCol.Width := book.ConvertUnits(w, suChars, book.Units);
        end;
      end else begin
        lCol.ColWidthType := cwtDefault;
        lCol.Width := sheet.ReadDefaultColWidth(FWorkbook.Units);
      end;

      s := GetAttrValue(colNode, 'style');
      if s <> '' then begin
        idx := FCellFormatList.FindIndexOfID(StrToInt(s));
        fmt := FCellFormatList.Items[idx];
        lCol.FormatIndex := book.AddCellFormat(fmt^);
      end else
        lCol.FormatIndex := 0;

      lCol.Options := [];
      s := GetAttrValue(colNode, 'hidden');
      if StrIsTrue(s) then
        Include(lCol.Options, croHidden);

      if (lCol.ColWidthType = cwtCustom) or (lCol.FormatIndex > 0) or (lCol.Options <> []) then
        for col := col1 to Min(col2, FLastCol) do
          sheet.WriteColInfo(col, lCol);
    end;
    colNode := colNode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadComments(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  node, txtNode, rNode, rchild: TDOMNode;
  nodeName: String;
  cellAddr: String;
  s: String;
  r, c: Cardinal;
  comment: String;
begin
  comment := '';
  node := ANode.FirstChild;
  while node <> nil do
  begin
    nodeName := node.NodeName;
    cellAddr := GetAttrValue(node, 'ref');
    if cellAddr <> '' then
    begin
      comment := '';
      txtNode := node.FirstChild;
      while txtNode <> nil do
      begin
        rNode := txtnode.FirstChild;
        while rNode <> nil do
        begin
          nodeName := rnode.NodeName;
          rchild := rNode.FirstChild;
          while rchild <> nil do begin
            nodename := rchild.NodeName;
            if (nodename = 't') or (nodename = 'x:t') then begin
              s := GetNodeValue(rchild);
              if comment = '' then comment := s else comment := comment + s;
            end;
            rchild := rchild.NextSibling;
          end;
          rNode := rNode.NextSibling;
        end;
        if (comment <> '') and ParseCellString(cellAddr, r, c) then begin
          // Fix line endings  // #10 --> "LineEnding"
          comment := UTF8StringReplace(comment, #10, LineEnding, [rfReplaceAll]);
          (AWorksheet as TsWorksheet).WriteComment(r, c, comment);
        end;
        txtNode := txtNode.NextSibling;
      end;
    end;
    node := node.NextSibling;
  end;
end;

{ The comments are stored in separate "comments<n>.xml" files (n = 1, 2, ...)
  for each worksheet.
  The relationship which comment belongs to which sheet file is contained
  in the "sheet<n>.xml.rels" file (n = 1, 2, ...) which already has been
  read and stored in the sheetData.
  The rels file contains also the second part of the hyperlink data. }
procedure TsSpreadOOXMLReader.ReadComments(AStream: TStream; AWorksheet: TsBasicWorksheet);
var
  sheetData: TSheetData;
  sheetIndex: Integer;
  XMLStream: TStream;
  fn: String;
  Doc: TXMLDocument = nil;
begin
  sheetIndex := TsWorksheet(AWorksheet).Index;
  sheetData := TSheetData(FSheetlist[sheetIndex]);
  if sheetData.Comments_File <> '' then
  begin
    // Extract texts from the comments file found and apply to worksheet.
    fn := MakeXLPath(sheetData.Comments_File);
    XMLStream := CreateXMLStream;
    try
      if UnzipToStream(AStream, fn, XMLStream) then
      begin
        try
          ReadXMLStream(Doc, XMLStream);
          ReadComments(Doc.DocumentElement.FindNode('commentList'), AWorksheet);
        finally
          FreeAndNil(Doc);
        end;
      end;
    finally
      XMLStream.Free;
    end;
  end;
end;

procedure TsSpreadOOXMLReader.ReadConditionalFormatting(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  childNode: TDOMNode;
  nodeName: string;
  range: TsCellRange;
  s: String;
  dxf: TDifferentialFormatData;
  dxfId: Integer;
  fmtIdx: Integer;
begin
  while ANode <> nil do
  begin
    nodeName := ANode.NodeName;
    if (nodeName = 'conditionalFormatting') then
    begin
      s := GetAttrValue(ANode, 'sqref');
      if ParseCellRangeString(s, range) then
      begin
        childNode := ANode.FirstChild;
        while childNode <> nil do
        begin
          nodeName := childNode.NodeName;
          if nodeName = 'cfRule' then
          begin
            // Get format index
            s := GetAttrValue(childNode, 'dxfId');
            if TryStrToInt(s, dxfId) then
            begin
              dxf := TDifferentialFormatData(FDifferentialFormatList[dxfId]);
              fmtIdx := dxf.CellFormatIndex;
            end else
              fmtIdx := 0;

            s := GetAttrValue(childNode, 'type');
            case s of
              'cellIs':
                ReadCFCellFormat(childNode, AWorksheet, range, fmtIdx);
              'aboveAverage':
                ReadCFAverage(childNode, AWorksheet, range, fmtIdx);
              'top10':
                ReadCFTop10(childNode, AWorksheet, range, fmtIdx);
              'uniqueValues', 'duplicateValues', 'containsErrors', 'notContainsErrors':
                ReadCFMisc(childNode, AWorksheet, range, fmtIdx);
              'containsText', 'notContainsText', 'beginsWith', 'endsWith':
                ReadCFMisc(childNode, AWorksheet, range, fmtIdx);
              'timePeriod':
                ReadCFDateFormat(childNode, AWorksheet, range, fmtIdx);
              'expression':
                ReadCFExpression(childNode, AWorksheet, range, fmtIdx);
              'colorScale':
                ReadCFColorRange(childNode, AWorksheet, range);
              'dataBar':
                ReadCFDataBars(childNode, AWorksheet, range);
              'iconSet':
                ReadCFIconSet(childNode, AWorksheet, range);
            end;
          end;
          childNode := childNode.NextSibling;
        end;
      end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadDateMode(ANode: TDOMNode);
var
  s: String;
begin
  if Assigned(ANode) then begin
    s := Lowercase(GetAttrValue(ANode, 'date1904'));
    if StrIsTrue(s) then // (s = '1') or (s = 'true') then
      FDateMode := dm1904
  end;
end;

procedure TsSpreadOOXMLReader.ReadDefinedNames(ANode: TDOMNode);
var
  node: TDOMNode;
  nodeName: String;
  r1,c1,r2,c2: Cardinal;
  flags: TsRelFlags;
  id, j, p: Integer;
  book: TsWorkbook;
  sheet: TsWorksheet;
  localSheetID: String;
  namestr: String;
  s, sheetname1, sheetName2: String;
  L: TStringList;
begin
  if ANode = nil then
    exit;

  book := TsWorkbook(FWorkbook);

  node := ANode.FirstChild;
  while node <> nil do begin
    nodename := node.NodeName;
    if (nodename = 'definedName') or (nodename = 'x:definedName') then
    begin
      id := -1;
      sheet := nil;
      localSheetID := GetAttrValue(node, 'localSheetId');
      if (localSheetID <> '') then begin
        if not TryStrToInt(localSheetID, id) then
        begin
          FWorkbook.AddErrorMsg('Invalid localSheetID in "definedName" node');
          node := node.NextSibling;
          Continue;
        end;
        sheet := (FWorkbook as TsWorkbook).GetWorksheetByIndex(id);
      end;

      namestr := GetAttrValue(node, 'name');
      if namestr = '_xlnm.Print_Area' then
      begin
        if sheet = nil then begin
          FWorkbook.AddErrorMsg('No localSheetID found for defined name "_xlnm.Print_Area"');
          continue;
        end;
        L := TStringList.Create;
        try
          L.Delimiter := ',';
          L.StrictDelimiter := true;
          L.DelimitedText := GetNodeValue(node);
          for j:=0 to L.Count-1 do
          begin
            s := StringReplace(L[j], '''', '', [rfReplaceAll]);
            p := pos(':', s);
            if p = 0 then
            begin
              FWorkbook.AddErrorMsg('invalid cell range reference in "definedName" node');
              break;
            end;
            ParseSheetCellString(Copy(s, 1, p-1), sheetname1, r1, c1);
            ParseSheetCellString(Copy(s, p+1, MaxInt), sheetname2, r2, c2);
            sheet.PageLayout.AddPrintRange(r1, c1, r2, c2);
          end;
        finally
          L.Free;
        end;
      end else
      if nameStr = '_xlnm.Print_Titles' then
      begin
        if sheet = nil then begin
          FWorkbook.AddErrorMsg('No localSheetID found for defined name "_xlnm.Print_Titles"');
          continue;
        end;
        L := TStringList.Create;
        try
          L.Delimiter := ',';
          L.StrictDelimiter := true;
          L.DelimitedText := GetNodeValue(node);
          for j:=0 to L.Count-1 do
          begin
            //s := ReplaceStr(L[j], '''', '');   // wp: replaced by next line due to Laz 1.0
            s := StringReplace(L[j], '''', '', [rfReplaceAll]);
            p := pos('!', s);
            if p > 0 then s := Copy(s, p+1, MaxInt);
            p := pos(':', s);
            if not ParseCellColString(copy(s, 1, p-1), c1) then
              c1 := UNASSIGNED_ROW_COL_INDEX;
            if not ParseCellColString(copy(s, p+1, MaxInt), c2) then
              c2 := UNASSIGNED_ROW_COL_INDEX;
            if not ParseCellRowString(copy(s, 1, p-1), r1) then
              r1 := UNASSIGNED_ROW_COL_INDEX;
            if not ParseCellRowString(copy(s, p+1, MaxInt), r2) then
              r2 := UNASSIGNED_ROW_COL_INDEX;
            if (r1 <> cardinal(UNASSIGNED_ROW_COL_INDEX)) then
              sheet.PageLayout.SetRepeatedRows(r1, r2);
            if (c1 <> cardinal(UNASSIGNED_ROW_COL_INDEX)) then
              sheet.PageLayout.SetRepeatedCols(c1, c2);
          end;
        finally
          L.Free;
        end;
      end;

      // "Normal" defined names
      s := GetNodeValue(node);
      if ParseCellRangeString(s, sheetName1, sheetName2, r1, c1, r2, c2, flags) then
        book.DefinedNames.Add(nameStr, sheetName1, sheetName2, r1, c1, r2, c2);
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadDifferentialFormat(ANode: TDOMNode);
var
  nodeName: String;
  childNode: TDOMNode;
  pattNode: TDOMNode;
  fontStyles: TsFontStyles;
  fontColor: TsColor;
  bgColor: TsColor;
  fgColor: TsColor;
  fillPatt: String;
  borders: TsCellBorders;
  borderStyles: TsCellBorderStyles;
  uff: TsUsedFormattingFields;
  dxf: TDifferentialFormatData;
  fmt: TsCellFormat;
  numFmtStr: String;
  numFmtParams: TsNumFormatParams;
begin
  uff := [];
  borders := [];
  fontStyles := [];
  fontColor := scNotDefined;
  numFmtStr := '';

  ANode := ANode.FirstChild;
  while ANode <> nil do
  begin
    nodeName := ANode.NodeName;
    if nodeName = 'font' then
    begin
      // read font
      uff := uff + [uffFont];
      childNode := ANode.FirstChild;
      while childNode <> nil do
      begin
        nodeName := childNode.NodeName;
        if (nodename = 'b') then
          fontStyles := fontStyles + [fssBold]
        else
        if (nodename = 'i') then
          fontStyles := fontStyles + [fssItalic]
        else
        if (nodeName = 'strike') then
          fontStyles := fontStyles + [fssStrikeOut]
        else
        if (nodename = 'color') then
          fontColor := ReadColor(childNode);
        childNode := childNode.NextSibling;
      end;
    end else
    if nodeName = 'border' then
    begin
      // read border
      uff := uff + [uffBorder];
      childNode := ANode.FirstChild;
      while childNode <> nil do
      begin
        nodeName := childNode.NodeName;
        case nodeName of
          'left': if ReadBorderStyle(childNode, borderStyles[cbWest]) then Include(borders, cbWest);
          'right': if ReadBorderStyle(childNode, borderStyles[cbEast]) then Include(borders, cbEast);
          'top': if ReadBorderStyle(childNode, borderStyles[cbNorth]) then Include(borders, cbNorth);
          'bottom': if ReadBorderStyle(childNode, borderStyles[cbSouth]) then Include(borders, cbSouth);
        end;
        childNode := childNode.NextSibling;
      end;
    end else
    if nodeName = 'fill'  then
    begin
      // read fill
      uff := uff + [uffBackground];
      pattNode := ANode.FirstChild;
      while pattNode <> nil do begin
        nodeName := pattNode.NodeName;
        if nodeName = 'patternFill' then
        begin
          fillPatt := GetAttrValue(pattNode, 'patternType');
          childNode := pattNode.FirstChild;
          while childNode <> nil do
          begin
            nodeName := childNode.NodeName;
            if nodeName = 'bgColor' then
              bgColor := ReadColor(childNode)
            else if nodeName = 'fgColor' then
              fgColor := ReadColor(childNode);
            childNode := childNode.NextSibling;
          end;
        end;
        pattNode := pattNode.NextSibling;
      end;
    end else
    if nodeName = 'numFmt' then
    begin
      // Nuzmber format
      uff := uff + [uffNumberFormat];
      numFmtStr := GetAttrValue(ANode, 'formatCode');
    end;
    ANode := ANode.NextSibling;
  end;

  InitFormatRecord(fmt);
  if uff <> [] then
  begin
    dxf := TDifferentialFormatData.Create;
    dxf.UsedFormattingFields := uff;
    fmt.UsedFormattingFields := uff;
    if (uffBackground in uff) then
    begin
      dxf.FillPatternType := StrToFillStyle(fillPatt);
      dxf.FillBgColor := bgColor;
      dxf.FillFgColor := fgColor;
      fmt.SetBackground(dxf.FillPatternType, dxf.FillFgColor, dxf.FillBgColor);
    end;
    if (uffBorder in uff) then
    begin
      dxf.Borders := borders;
      dxf.BorderStyles := borderStyles;
      fmt.Border := borders;
      fmt.BorderStyles := borderStyles;
    end;
    if (uffFont in uff) then
    begin
      dxf.FontColor := fontColor;
      dxf.FontStyles := fontStyles;
      fmt.FontIndex := TsWorkbook(FWorkbook).AddFont('', -1, fontStyles, fontColor);
      // Excel does not allow to change font name and font size.
    end else
    if (uffNumberFormat in uff) then
    begin
      fmt.NumberFormatIndex := TsWorkbook(FWorkbook).AddNumberFormat(numFmtStr);
      numFmtParams := TsWorkbook(FWorkbook).GetNumberFormat(fmt.NumberFormatIndex);
      fmt.NumberFormatStr := numFmtParams.NumFormatStr;
      fmt.NumberFormat := numFmtParams.NumFormat;
    end;
    dxf.CellFormatIndex := TsWorkbook(FWorkbook).AddCellFormat(fmt);
    FDifferentialFormatList.Add(dxf);
  end;
end;

procedure TsSpreadOOXMLReader.ReadDifferentialFormats(ANode: TDOMNode);
var
  nodeName: String;
begin
  if ANode = nil then
    exit;
  ANode := ANode.FirstChild;
  while ANode <> nil do
  begin
    nodeName := ANode.NodeName;
    if nodeName = 'dxf' then
      ReadDifferentialFormat(ANode);
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadDimension(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  ref: String;
  r1, c1: Cardinal;
  ok: Boolean;
begin
  Unused(AWorksheet);

  FLastRow := MaxInt;
  FLastCol := MaxInt;

  if ANode = nil then
    exit;

  ref := GetAttrValue(ANode, 'ref');
  if ref <> '' then begin
    // Normally the range of worksheets is specified as, e.g., 'A1:K5'
    ok := ParseCellRangeString(ref, r1, c1, FLastRow, FLastCol);
    // But for empty worksheets it is specified as only 'A1'
    if not ok then
      ParseCellString(ref, FLastRow, FLastCol);
  end;
end;

{ Reads the parameters of the embedded images defined as children of the 
  specified node which is in a drawingX.xml file. }
procedure TsSpreadOOXMLReader.ReadDrawing(ANode: TDOMNode; 
  AWorksheet: TsBasicWorksheet);

  procedure ReadXdrFromTo(ANode: TDOMNode;
    out ARow, ACol: Integer; out ARowOffs, AColOffs: Double);
  var
    nodeName: string;
  begin
    if ANode = nil then
      exit;
    while Assigned(ANode) do begin
      nodeName := ANode.NodeName;
      case nodeName of
        'xdr:col':
          ACol := StrToIntDef(GetNodeValue(ANode), -1);
        'xdr:row':
          ARow := StrToIntDef(GetNodeValue(ANode), -1);
        'xdr:colOff':
          AColOffs := EMUToMM(StrToInt64Def(GetNodeValue(ANode), 0));
        'xdr:rowOff':
          ARowOffs := EMUToMM(StrToInt64Def(GetNodeValue(ANode), 0));
      end;
      ANode := ANode.NextSibling;
    end;
  end;

  procedure ReadXdrPic(ANode: TDOMNode; out rId: String; out AFileName: String);
  var
    nodeName: String;
    child: TDOMNode;
  begin
    if ANode = nil then
      exit;
    while Assigned(ANode) do
    begin
      nodeName := ANode.NodeName;
      case nodeName of
        'xdr:blipFill':
          begin
            child := ANode.FirstChild;
            while Assigned(child) do
            begin
              nodeName := child.NodeName;
              if nodeName = 'a:blip' then
                rID := GetAttrValue(child, 'r:embed');
              child := child.NextSibling;
            end;
          end;
        'xdr:nvPicPr':
          begin
            child := ANode.FirstChild;
            while Assigned(child) do
            begin
              nodeName := child.NodeName;
              if nodeName = 'xdr:cNvPr' then
                AFileName := GetAttrValue(child, 'descr');
              child := child.NextSibling;
            end;
          end;
      end;
      ANode := ANode.NextSibling;
    end;
  end;

  procedure ReadXdrGraphicFrame(ANode: TDOMNode; out AGraphicFrameName: String;
    out rID: String; out IsChart: Boolean);
  var
    nodeName: String;
    child, child2: TDOMNode;
  begin
    if ANode = nil then
      exit;
    while Assigned(ANode) do
    begin
      nodeName := ANode.NodeName;
      case nodeName of
        'xdr:nvGraphicFramePr':
          begin
            child := ANode.FirstChild;
            while Assigned(child) do
            begin
              nodeName := child.NodeName;
              if nodeName = 'xdr:cNvPr' then
              begin
                AGraphicFrameName := GetAttrValue(child, 'name');
              end;
              child := child.NextSibling;
            end;
          end;
        'a:graphic':
          begin
            child := ANode.FirstChild;
            while Assigned(child) do
            begin
              nodename := child.Nodename;
              if nodename = 'a:graphicData' then
              begin
                child2 := child.Firstchild;
                while Assigned(child2) do
                begin
                  nodeName := child2.NodeName;
                  if nodename = 'c:chart' then
                  begin
                    rId := GetAttrValue(child2, 'r:id');
                    if rId <> '' then
                    begin
                      isChart := true;
                    end;
                  end;
                  child2 := child2.NextSibling;
                end;
              end;
              child := child.NextSibling;
            end;
          end;
      end;
      ANode := ANode.NextSibling;
    end;
  end;

var
  node: TDOMNode;
  nodeName: String = '';
  rID, fileName: String;
  xPos, yPos, horExt, vertExt: Double;
  fromCol, fromRow, toCol, toRow: Integer;
  fromColOffs, fromRowOffs, toColOffs, toRowOffs: Double;
  isChart: Boolean;
  data: TEmbeddedObjData;
  sheetData: TSheetData;
  graphicFrameName: String;

begin
  if ANode = nil then
    exit;

  sheetData := TSheetData(FSheetList[TsWorksheet(AWorksheet).Index]);
  
  ANode := ANode.FirstChild;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    horExt := -1.0;    vertExt := -1.0;
    fromCol := -1;     fromColOffs := 0.0;
    fromRow := -1;     fromRowOffs := 0.0;
    toCol := -1;       toColOffs := 0.0;
    toRow := -1;       toRowOffs := 0.0;
    rID := '';         fileName := '';
    isChart := false;
    graphicFrameName := '';
    if nodeName = 'xdr:absoluteAnchor' then
    begin
      node := ANode.FirstChild;
      while Assigned(node) do begin
        nodeName := node.NodeName;
        case nodeName of
          'xdr:pos':
            begin
              xpos := EMUToMM(StrToInt64Def(GetAttrValue(node, 'x'), 0));
              ypos := EMUToMM(StrToInt64Def(GetAttrValue(node, 'y'), 0));
            end;
          'xdr:ext':
            begin
              horExt := EMUToMM(StrToInt64Def(GetAttrValue(node, 'cx'), 0));
              vertExt := EMUToMM(StrToInt64Def(GetAttrValue(node, 'cy'), 0));
            end;
          'xdr:graphicFrame':
            ReadXdrGraphicFrame(node.FirstChild, graphicFrameName, rID, isChart);
        end;
        node := node.NextSibling
      end;
    end
    else
    if nodeName = 'xdr:twoCellAnchor' then
    begin
      node := ANode.FirstChild;
      while Assigned(node) do begin 
        nodeName := node.NodeName;
        case nodeName of
          'xdr:from':
            ReadXdrFromTo(node.FirstChild, fromRow, fromCol, fromRowOffs, fromColOffs);
          'xdr:to':
            ReadXdrFromTo(node.FirstChild, toRow, toCol, toRowOffs, toColOffs);
          'xdr:pic':
            ReadXdrPic(node.FirstChild, rID, filename);
          'xdr:graphicFrame':
            ReadXdrGraphicFrame(node.FirstChild, graphicFrameName, rID, isChart);
        end;
        node := node.NextSibling;
      end;
    end;

    if (rID <> '') and (
       ( (fromCol <> -1) and (toCol <> -1) and (fromRow <> -1) and (toRow <> -1) and (rID <> '')) or
       ( (horExt <> -1) and (vertext <> -1) ) ) then
    begin
      data := TEmbeddedObjData.Create;
      data.XPos := xPos;
      data.YPos := yPos;
      data.Width := horExt;
      data.Height := vertExt;
      data.FromCol := fromCol;
      data.FromColOffs := fromColOffs;
      data.ToCol := toCol;
      data.ToColOffs := toColOffs;
      data.FromRow := fromRow;
      data.FromRowOffs := fromRowOffs;
      data.ToRow := toRow;
      data.ToRowOffs := toRowOffs;
      data.RelId := rId;
      data.FileName := fileName;
      data.MediaName := MakeXLPath(sheetData.DrawingRels.FindTarget(rID));
      data.ImgIndex := -1;
      data.Worksheet := AWorksheet;
      data.IsChart := isChart;
      data.IsHeaderFooter := false;
      FEmbeddedObjList.Add(data);
    end;
      
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadDrawingRels(ANode: TDOMNode; ASheet: TsBasicWorksheet);
var
  {%H-}nodeName: String;
  relID, relTarget, relType: String;
  data: TEmbeddedObjData;
  j: Integer;
begin
  if ANode = nil then
    exit;
  nodeName := ANode.NodeName;
  while Assigned(ANode) do
  begin 
    nodeName := ANode.NodeName;
    relID := GetAttrValue(ANode, 'Id');
    relTarget := GetAttrValue(ANode, 'Target');
    relType := GetAttrValue(ANode, 'Type');
    if (relID <> '') and (relTarget <> '') then
    begin
      case relType of
        SCHEMAS_IMAGE:
          begin
            relTarget := MakeXLPath(relTarget);
            for j := 0 to FEmbeddedObjList.Count-1 do
            begin
              data := TEmbeddedObjData(FEmbeddedObjList[j]);
              if (data.Worksheet = ASheet) and (data.RelID = relID) then
              begin
                data.MediaName := relTarget;
                break;
              end;
            end;
          end;
        SCHEMAS_CHART:
          begin
          end;
      end;
      ANode := ANode.NextSibling;
    end;
  end;
end;

{ Reads embedded images. 
  Information about them is scattered over several places.
  - From the worksheets' rel files we get the drawing*.xml files which contain
    image position information
  - The related drawing*.xml.rels files contain the location of the media
    files
  - The media files themselves. }
procedure TsSpreadOOXMLReader.ReadEmbeddedObjs(AStream: TStream);
var
  i, j: Integer;
  fn, {%H-}relsFn: String;
  XMLStream: TStream;
  doc: TXMLDocument;
  sheet: TsWorksheet;
  data: TEmbeddedObjData;
  sheetData: TSheetData;
  w, h: Double;
  embObj: TsEmbeddedObj;
  scaleX, scaleY: Double;
  pageIdx: Integer;
  {$ifdef FPS_Charts}
  chart: TsChart;
  {$endif}
begin
  doc := nil;
  j := 1;
  try
    for i := 0 to FSheetList.Count-1 do
    begin
      sheetData := TSheetData(FSheetList[i]);
      sheet := (FWorkbook as TsWorkbook).GetWorksheetByIndex(i);
      
      // Read the drawing<n>.xml file
      fn := sheetData.Drawing_File;
      if fn <> '' then 
      begin
        XMLStream := CreateXMLStream;
        try
          // Read drawings parameters and store them in the FEmbeddedObjList.
          if not UnzipToStream(AStream, fn, XMLStream) then
            raise EFPSpreadsheetReader.CreateFmt(rsDefectiveInternalFileStructure, ['xlsx']);
          ReadXMLStream(doc, XMLStream);
          relsFn := RelsFileFor(fn);
          ReadDrawing(doc.DocumentElement, sheet);
          FreeAndNil(doc);
        finally
          XMLStream.Free;
        end;
      end;
        
      // Now repeat the same with the vmlDrawing<n> file
      fn := sheetData.VmlDrawing_File;
      if fn <> '' then
      begin
        XMLStream := CreateXMLStream;
        try
          // Read vmDrawings parameters and store them in the FEmbeddedObjList, too.
          if not UnzipToStream(AStream, fn, XMLStream) then
            raise EFPSpreadsheetReader.CreateFmt(rsDefectiveInternalFileStructure, ['xlsx']);
          ReadXMLStream(doc, XMLStream);
          relsFn := RelsFileFor(fn);
          ReadVmlDrawing(doc.DocumentElement, sheet);
          FreeAndNil(doc);
        finally
          XMLStream.Free;
        end;
      end;
    end;
              
    // Read the embedded streams, add them to the workbook...
    ReadMedia(AStream);

    // ... and insert them in the worksheet
    for i := 0 to FEmbeddedObjList.Count-1 do
    begin
      data := TEmbeddedObjData(FEmbeddedObjList[i]);
      sheet := TsWorksheet(data.Worksheet);

      if sheet = nil then
        Continue;

      // Add chart
      {$ifdef FPS_Charts}
      if data.IsChart then
      begin
        w := -data.FromColOffs + data.ToColOffs;
        h := -data.FromRowOffs + data.ToRowOffs;
        for j := data.FromCol to data.ToCol-1 do
          w := w + sheet.GetColWidth(j, suMillimeters);
        for j := data.FromRow to data.ToRow-1 do
          h := h + sheet.GetRowHeight(j, suMillimeters);
        chart := TsWorkbook(FWorkbook).AddChart(sheet, data.FromRow, data.FromCol, w, h, data.FromRowOffs, data.FromColOffs);
        TsSpreadOOXMLChartReader(FChartReader).ReadChartXML(AStream, chart, data.MediaName);
      end else
      {$endif}
      if (data.ImgIndex > -1) then
      begin
        embObj := TsWorkbook(FWorkbook).GetEmbeddedObj(data.ImgIndex);
        // Add header/footer image
        if data.IsHeaderFooter then
        begin
          if data.HFImgPosition <> '' then
          begin
            scaleX := data.HFImgWidth / embObj.ImageWidth;
            scaleY := data.HFImgHeight / embObj.ImageHeight;
            // Scale factor calculation is very inaccurate. We try to round to integers
            if (scaleX > 0.99) and SameValue(scaleX, round(scaleX), 0.01) then
              scaleX := round(scaleX);
            if (scaleY > 0.99) and SameValue(scaleY, round(scaleY), 0.01) then
              scaleY := round(scaleY);            
            pageIdx := HEADER_FOOTER_INDEX_ALL;  // wp: Don't know which is the correct index...
            case data.HFImgPosition[2] of
              'H':
                case data.HFImgPosition[1] of
                  'L': sheet.PageLayout.AddHeaderImage(pageIdx, hfsLeft, data.ImgIndex, scaleX, scaleY);
                  'C': sheet.PageLayout.AddHeaderImage(pageIdx, hfsCenter, data.ImgIndex, scaleX, scaleY);
                  'R': sheet.PageLayout.AddHeaderImage(pageIdx, hfsRight, data.ImgIndex, scaleX, scaleY);
                end;
              'F':
                case data.HFImgPosition[1] of
                  'L': sheet.PageLayout.AddFooterImage(pageIdx, hfsLeft, data.ImgIndex, scaleX, scaleY);
                  'C': sheet.PageLayout.AddFooterImage(pageIdx, hfsCenter, data.ImgIndex, scaleX, scaleY);
                  'R': sheet.PageLayout.AddFooterImage(pageIdx, hfsRight, data.ImgIndex, scaleX, scaleY);
                end;
            end;
          end;
        end else
        begin
          // Add image
          w := -data.FromColOffs + data.ToColOffs;
          h := -data.FromRowOffs + data.ToRowOffs;
          for j := data.FromCol to data.ToCol-1 do 
            w := w + sheet.GetColWidth(j, suMillimeters);
          for j := data.FromRow to data.ToRow-1 do 
            h := h + sheet.GetRowHeight(j, suMillimeters);
          scaleX := w / embObj.ImageWidth;
          scaleY := h / embObj.ImageHeight;
          // Scale factor calculation is very inaccurate. We try to round to integers
          if (scaleX > 0.99) and SameValue(scaleX, round(scaleX), 0.01) then
            scaleX := round(scaleX);
          if (scaleY > 0.99) and SameValue(scaleY, round(scaleY), 0.01) then
            scaleY := round(scaleY);
          sheet.WriteImage(data.FromRow, data.FromCol,
            data.ImgIndex,
            data.FromRowOffs, data.FromColOffs,
            scaleX, scaleY
          );
        end;
      end;
    end;

  finally
    doc.Free;
  end;
end;

procedure TsSpreadOOXMLReader.ReadFileVersion(ANode: TDOMNode);
begin
  FWrittenByFPS := GetAttrValue(ANode, 'appName') = 'fpspreadsheet';
end;

procedure TsSpreadOOXMLReader.ReadFills(ANode: TDOMNode);
var
  fillNode, patternNode, colorNode, stopNode: TDOMNode;
  nodeName: String;
  filldata: TFillListData;
  patt: String;
  fgclr: TsColor;
  bgclr: TsColor;
begin
  if ANode = nil then
    exit;

  fillNode := ANode.FirstChild;
  while Assigned(fillNode) do begin
    nodename := fillNode.NodeName;
    patternNode := fillNode.FirstChild;
    while Assigned(patternNode) do begin
      nodename := patternNode.NodeName;
      if (nodename = 'patternFill') or (nodename = 'x:patternFill') then begin
        patt := GetAttrValue(patternNode, 'patternType');
        fgclr := scWhite;
        bgclr := scBlack;
        colorNode := patternNode.FirstChild;
        while Assigned(colorNode) do begin
          nodeName := colorNode.NodeName;
          if (nodeName = 'fgColor') or (nodeName = 'x:fgColor')then
            fgclr := ReadColor(colorNode)
          else
          if (nodeName = 'bgColor') or (nodeName = 'x:bgColor') then
            bgclr := ReadColor(colorNode);
          colorNode := colorNode.NextSibling;
        end;

        // Store in FFillList
        fillData := TFillListData.Create;
        fillData.PatternType := patt;
        fillData.FgColor := fgclr;
        fillData.BgColor := bgclr;
        FFillList.Add(fillData);
      end else
      if (nodeName = 'gradientFill') or (nodeName = 'x:gradientFill') then
      begin
        // We do not support gradient fills yet. As a replacement, we read
        // the first color of the gradient and use it for a solid fill
        // This is required in order to keep the fill numbering intact.
        stopNode := patternNode.FirstChild;
        while Assigned(stopNode) do begin
          nodeName := stopNode.NodeName;
          if (nodeName = 'stop') or (nodeName = 'x:stop') then begin
            colorNode := stopNode.FirstChild;
            while Assigned(colorNode) do begin
              nodeName := colorNode.NodeName;
              if (nodeName = 'color') or (nodeName = 'x:color') then
              begin
                bgclr := ReadColor(colorNode);
                fgclr := bgclr;
                break;
              end;
              colorNode := colorNode.NextSibling;
            end;
            break;
          end;
          stopNode := stopNode.NextSibling;
        end;

        // Store in FFillList:
        fillData := TFillListData.Create;
        fillData.PatternType := 'GradientFill '+IntToStr(FFillList.Count);
        fillData.FgColor := fgclr;
        fillData.BgColor := bgclr;
        FFillList.Add(fillData);
      end else
      if nodeName <> '#text' then
      begin
        // This pattern type is unknown to fpspreadsheet. We must store a dummy
        // fill data record to keep the numbering intact.
        fillData := TFillListData.Create;
        fillData.PatternType := 'non';
        FFillList.Add(fillData);
        Workbook.AddErrorMsg('ReadFills: Unsupported pattern node ' + nodeName);
      end;

      patternNode := patternNode.NextSibling;
    end;
    fillNode := fillNode.NextSibling;
  end;
end;

{ Reads the font described by the specified node and stores it in the reader's
  FontList. }
function TsSpreadOOXMLReader.ReadFont(ANode: TDOMNode): Integer;
var
  node: TDOMNode;
  fnt: TsFont;
  fntName: String;
  fntSize: Single;
  fntStyles: TsFontStyles;
  fntColor: TsColor;
  fntPos: TsFontPosition;
  nodename: String;
  s: String;
  acceptDuplicates: Boolean;
  isDefaultFont: Boolean;
begin
  isDefaultFont := FFontList.Count = 0;

  fnt := (Workbook as TsWorkbook).GetDefaultFont;
  if fnt <> nil then
  begin
    fntName := fnt.FontName;
    fntSize := fnt.Size;
    fntStyles := fnt.Style;
    fntColor := fnt.Color;
    fntPos := fnt.Position;
  end else
  begin
    fntName := DEFAULT_FONTNAME;
    fntSize := DEFAULT_FONTSIZE;
    fntStyles := [];
    fntColor := scBlack;
    fntPos := fpNormal;
  end;

  acceptDuplicates := true;
  node := ANode.FirstChild;
  while node <> nil do
  begin
    nodename := node.NodeName;
    if (nodename = 'name') or (nodename = 'x:name') or
       (nodename = 'rFont') or (nodename = 'x:rFont') then
    begin
      s := GetAttrValue(node, 'val');
      if s <> '' then fntName := s;
      if (nodename = 'rFont') or (nodename = 'x:rFont') then acceptDuplicates := false;
    end
    else
    if (nodename = 'sz') or (nodeName = 'x:sz') then
    begin
      s := GetAttrValue(node, 'val');
      if s <> '' then fntSize := StrToFloat(s, FPointSeparatorSettings);
    end
    else
    if (nodename = 'b') or (nodename = 'x:b') then
    begin
      s := GetAttrValue(node, 'val');
      if (s = '') or StrIsTrue(s) then  //   if GetAttrValue(node, 'val') <> 'false' then
        fntStyles := fntStyles + [fssBold];
    end
    else
    if (nodename = 'i') or (nodename = 'x:i') then
    begin
      s := GetAttrValue(node, 'val');
      if (s = '') or StrIsTrue(s) then  // if GetAttrValue(node, 'val') <> 'false' then
        fntStyles := fntStyles + [fssItalic];
    end
    else
    if (nodename = 'u') or (nodename = 'x:u') then
    begin
      s := GetAttrValue(node, 'val');
      if not StrIsFalse(s) then   // can have many values, not just booleans
        fntStyles := fntStyles+ [fssUnderline]
    end
    else
    if (nodename = 'strike') or (nodename = 'x:strike') then
    begin
      s := GetAttrValue(node, 'val');
      if not StrIsFalse(s) then  // can have several values, not just booleans
        fntStyles := fntStyles + [fssStrikeout];
    end
    else
    if (nodename = 'vertAlign') or (nodename = 'x:vertAlign') then
    begin
      s := GetAttrValue(node, 'val');
      if s = 'superscript' then
        fntPos := fpSuperscript
      else
      if s = 'subscript' then
        fntPos := fpSubscript
      else
        fntPos := fpNormal;
    end
    else
    if (nodename = 'color') or (nodename = 'x:color') then
      fntColor := ReadColor(node);
    node := node.NextSibling;
  end;

  // If this method is called when reading the sharedstrings.xml duplicate
  // fonts should not be added to the reader's fontList.
  // As a function result we return the index of the already existing font.
  if not acceptDuplicates then
    for Result := 0 to FFontList.Count-1 do
    begin
      fnt := TsFont(FFontList[Result]);
      if SameText(fnt.FontName, fntName) and
         (fnt.Size = fntSize) and
         (fnt.Style = fntStyles) and
         (fnt.Color = fntColor) and
         (fnt.Position = fntPos)
      then
        exit;
  end;

  // Create a font record and store it in the reader's fontlist.
  // In case of fonts in styles.xml (nodename = "name"), do no look for
  // duplicates because this will screw up the font index
  // used in the xf records
  fnt := TsFont.Create;
  fnt.FontName := fntName;
  fnt.Size := fntSize;
  fnt.Style := fntStyles;
  fnt.Color := fntColor;
  fnt.Position := fntPos;

  Result := FFontList.Add(fnt);

  if isDefaultFont then
    (Workbook as TsWorkbook).SetDefaultFont(fnt.FontName, fnt.Size);
end;

procedure TsSpreadOOXMLReader.ReadFonts(ANode: TDOMNode);
var
  node: TDOMNode;
begin
  if ANode = nil then
    exit;

  node := ANode.FirstChild;
  while node <> nil do begin
    ReadFont(node);
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadHeaderFooter(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  node: TDOMNode;
  nodeName: String;
  s: String;
  sheet: TsWorksheet absolute AWorksheet;

  function FixLineEnding(AString: String): String;
  begin
    Result := StringReplace(AString, #10, LineEnding, [rfReplaceAll]);
  end;

begin
  if ANode = nil then
    exit;

  s := GetAttrValue(ANode, 'differentOddEven');
  if StrIsTrue(s)  then //(s = '1') or (s = 'true') then
    with sheet.PageLayout do Options := Options + [poDifferentOddEven];

  s := Lowercase(GetAttrValue(ANode, 'differentFirst'));
  if StrIsTrue(s) then  //(s = '1') or (s = 'true') then
    with sheet.PageLayout do Options := Options + [poDifferentFirst];

  node := ANode.FirstChild;
  while node <> nil do
  begin
    nodeName := node.NodeName;
    if pos('x:', nodename) = 1 then Delete(nodeName, 1, 2);
    case nodeName of
      'firstHeader': sheet.PageLayout.Headers[0] := FixLineEnding(GetNodeValue(node));
      'oddHeader'  : sheet.PageLayout.Headers[1] := FixLineEnding(GetNodeValue(node));
      'evenHeader' : sheet.PageLayout.Headers[2] := FixLineEnding(GetNodeValue(node));
      'firstFooter': sheet.PageLayout.Footers[0] := FixLineEnding(GetNodeValue(node));
      'oddFooter'  : sheet.PageLayout.Footers[1] := FixLineEnding(GetNodeValue(node));
      'evenFooter' : sheet.PageLayout.Footers[2] := FixLineEnding(GetNodeValue(node));
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadHyperlinks(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  node: TDOMNode;
  nodeName: String;
  sheetData: TSheetData;
  hyperlinkData: THyperlinkListData;

  function FindHyperlinkID(ID: String): THyperlinkListData;
  var
    i: Integer;
  begin
    for i:=0 to FHyperlinkList.Count-1 do
      if THyperlinkListData(FHyperlinkList.Items[i]).ID = ID then
      begin
        Result := THyperlinkListData(FHyperlinkList.Items[i]);
        exit;
      end;
  end;

begin
  sheetData := TSheetData(FSheetList[TsWorksheet(AWorksheet).Index]);
  
  if Assigned(ANode) then begin
    nodename := ANode.NodeName;
    if (nodename = 'hyperlinks') or (nodename = 'x:hyperlinks') then
    begin
      node := ANode.FirstChild;
      while Assigned(node) do
      begin
        nodename := node.NodeName;
        if (nodename = 'hyperlink') or (nodename = 'x:hyperlink') then begin
          hyperlinkData := THyperlinkListData.Create;
          hyperlinkData.CellRef := GetAttrValue(node, 'ref');
          hyperlinkData.ID := GetAttrValue(node, 'r:id');
          if hyperlinkData.ID <> '' then
            hyperlinkData.Target := sheetData.HyperlinkRels.FindTarget(hyperlinkData.ID);
          hyperlinkData.TextMark := GetAttrValue(node, 'location');
          hyperlinkData.Display := GetAttrValue(node, 'display');
          hyperlinkData.Tooltip := GetAttrValue(node, 'tooltip');
          hyperlinkData.Worksheet := AWorksheet;
        end;
        FHyperlinkList.Add(hyperlinkData);
        node := node.NextSibling;
      end;
    end;
  end;
  
  ApplyHyperlinks(AWorksheet);
end;

procedure TsSpreadOOXMLReader.ReadMedia(AStream: TStream);
var
  memstream: TMemoryStream;
  unzip: TStreamUnzipper;
  i: Integer;
  data: TEmbeddedObjData;
  ext: String;
begin 
  unzip := TStreamUnzipper.Create(AStream);
  try
    unzip.Examine;
    for i := 0 to FEmbeddedObjList.Count-1 do
    begin
      data := TEmbeddedObjData(FEmbeddedObjList[i]);
      if (data.MediaName <> '') and not data.IsChart then
      begin
        memStream := TMemoryStream.Create;
        unzip.UnzipFile(data.MediaName, memStream);
        memStream.Position := 0;
        if memStream.Size > 0 then
        begin
          data.ImgIndex := (FWorkbook as TsWorkbook).AddEmbeddedObj(memStream, ExtractFileName(data.Filename));
          ext := ExtractFileExt(data.MediaName);
          if data.FileName = '' then 
            data.FileName := ExtractFileName(data.MediaName)
          else
          if ExtractFileExt(data.FileName) = '' then
            data.FileName := data.FileName + ext;
        end;
        memStream.Free;
      end;
    end;
  finally
    unzip.Free;
  end;         
end;

procedure TsSpreadOOXMLReader.ReadMetaData(ANode: TDOMNode);
var
  childNode: TDOMNode;
  nodeName: string;
  book: TsWorkbook;
  s: String;
  name: String;
begin
  if ANode = nil then
    exit;

  book := TsWorkbook(FWorkbook);

  ANode := ANode.FirstChild;
  while ANode <> nil do
  begin
    nodeName := ANode.NodeName;
    s := GetNodeValue(ANode);
    case nodeName of
      // These fields are from "core.xml"
      'dc:title':
        book.MetaData.Title := s;
      'dc:subject':
        book.MetaData.Subject := s;
      'dc:creator':
        book.MetaData.CreatedBy := s;
      'cp:lastModifiedBy':
        book.MetaData.LastModifiedBy := s;
      'dc:description':
        if s <> '' then
        begin
          s := StringReplace(s, '_x000d_', Lineending, [rfReplaceAll]);
          book.MetaData.Comments.Text := s;
        end else
          book.MetaData.Comments.Clear;
      'cp:keywords':
        if s <> '' then
          book.MetaData.Keywords.CommaText := s
        else
          book.MetaData.Keywords.Clear;
      'dcterms:created':
        if s <> '' then
          book.MetaData.DateCreated := ISO8601StrToDateTime(s);
      'dcterms:modified':
        if s <> '' then
          book.MetaData.DateLastModified :=ISO8601StrToDateTime(s);

      // This field is from "custom.xml"
      'property':
        begin
          name := GetAttrValue(ANode, 'name');
          childNode := ANode.Firstchild;
          while childNode <> nil do
          begin
            nodeName := childNode.NodeName;
            s := GetNodeValue(childNode);
            if (s <> '') then
              book.MetaData.AddCustom(name, s);
            break;
          end;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadMergedCells(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  node: TDOMNode;
  nodename: String;
  s: String;
begin
  if Assigned(ANode) then begin
    node := ANode.FirstChild;
    while Assigned(node) do
    begin
      nodename := node.NodeName;
      if (nodename = 'mergeCell') or (nodename = 'x:mergeCell') then begin
        s := GetAttrValue(node, 'ref');
        if s <> '' then
          (AWorksheet as TsWorksheet).MergeCells(s);
      end;
      node := node.NextSibling;
    end;
  end;
end;

procedure TsSpreadOOXMLReader.ReadNumFormats(ANode: TDOMNode);
var
  node: TDOMNode;
  idStr: String;
  fmtStr: String;
  nodeName: String;
  id: Integer;
begin
  if Assigned(ANode) then
  begin
    node := ANode.FirstChild;
    while Assigned(node) do
    begin
      nodeName := node.NodeName;
      if (nodeName = 'numFmt') or (nodeName = 'x:numFmt') then
      begin
        fmtStr := GetAttrValue(node, 'formatCode');
        idStr := GetAttrValue(node, 'numFmtId');
        id := StrToInt(idStr);
        while id >= NumFormatList.Count do
          NumFormatList.Add('');
        NumFormatList[id] := fmtStr;
      end;
      node := node.NextSibling;
    end;
  end;
end;

procedure TsSpreadOOXMLReader.ReadPageMargins(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  s: String;
  sheet: TsWorksheet absolute AWorksheet;
begin
  if (ANode = nil) or (AWorksheet = nil) then     // just to make sure...
    exit;

  s := GetAttrValue(ANode, 'left');
  if s <> '' then
    sheet.PageLayout.LeftMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));

  s := GetAttrValue(ANode, 'right');
  if s <> '' then
    sheet.PageLayout.RightMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));

  s := GetAttrValue(ANode, 'top');
  if s <> '' then
    sheet.PageLayout.TopMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));

  s := GetAttrValue(ANode, 'bottom');
  if s <> '' then
    sheet.PageLayout.BottomMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));

  s := GetAttrValue(ANode, 'header');
  if s <> '' then
    sheet.PageLayout.HeaderMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));

  s := GetAttrValue(ANode, 'footer');
  if s <> '' then
    sheet.PageLayout.FooterMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));
end;

procedure TsSpreadOOXMLReader.ReadPageSetup(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  s: String;
  n: Integer;
  sheet: TsWorksheet absolute AWorksheet;
begin
  if ANode = nil then
    exit;

  // Paper size
  s := GetAttrValue(ANode, 'paperSize');
  if s <> '' then
  begin
    n := StrToInt(s);
    if (n >= 0) and (n <= High(PAPER_SIZES)) then
    begin
      sheet.PageLayout.PageWidth := PAPER_SIZES[n, 0];
      sheet.PageLayout.PageHeight := PAPER_SIZES[n, 1];
    end;
  end;

  // Orientation
  s := GetAttrValue(ANode, 'orientation');
  if s = 'portrait' then
    sheet.PageLayout.Orientation := spoPortrait
  else if s = 'landscape' then
    sheet.PageLayout.Orientation := spoLandscape;

  // Scaling factor
  s := GetAttrValue(ANode, 'scale');
  if s <> '' then
    sheet.PageLayout.ScalingFactor := StrToInt(s);
    // poFitPages is automatically excluded

  // Fit print job to pages
  s := GetAttrValue(ANode, 'fitToHeight');
  if s <> '' then
    sheet.PageLayout.FitHeightToPages := StrToInt(s);
    // poFitPages is automatically included.

  s := GetAttrValue(ANode, 'fitToWidth');
  if s <> '' then
    sheet.PageLayout.FitWidthToPages := StrToInt(s);
    // poFitPages is automatially included.

  // First page number
  s := GetAttrValue(ANode, 'firstPageNumber');
  if s <> '' then
    sheet.PageLayout.StartPageNumber := StrToInt(s);

  s := Lowercase(GetAttrValue(ANode, 'useFirstPageNumber'));
  with sheet.PageLayout do
    if (s = '1') or (s = 'true') then
      Options := Options + [poUseStartPageNumber]
    else
      Options := Options - [poUseStartPageNumber];

  // Print order
  s := GetAttrValue(ANode, 'pageOrder');
  if s = 'overThenDown' then
    with sheet.PageLayout do Options := Options + [poPrintPagesByRows];

  // Monochrome
  s := LowerCase(GetAttrValue(ANode, 'blackAndWhite'));
  if StrIsTrue(s) then //(s = '1') or (s = 'true') then
    with sheet.PageLayout do Options := Options + [poMonochrome];

  // Quality
  s := Lowercase(GetAttrValue(ANode, 'draft'));
  if StrIsTrue(s) then //(s = '1') or (s = 'true') then
    with sheet.PageLayout do Options := Options + [poDraftQuality];
end;

procedure TsSpreadOOXMLReader.ReadPalette(ANode: TDOMNode);
var
  node, colornode: TDOMNode;
  nodename: String;
  s: string;
  cidx: Integer;    // color index
  rgb: TsColor;
  n: Integer;
begin
  // OOXML sometimes specifies color by index even if a palette ("indexedColors")
  // is not loaeded. Therefore, we use the BIFF8 palette as default because
  // the default indexedColors are identical to it.
  FPalette.Clear;
  FPalette.AddBuiltinColors; // This adds the BIFF2 colors 0..7
  FPalette.AddExcelColors;   // This adds the BIFF8 colors 8..63
  n := FPalette.Count;

  if ANode = nil then
    exit;

  cidx := 0;
  node := ANode.FirstChild;
  while Assigned(node) do
  begin
    nodename := node.NodeName;
    if (nodename = 'indexedColors') or (nodename = 'x:indexedColors') then
    begin
      colornode := node.FirstChild;
      while Assigned(colornode) do
      begin
        nodename := colornode.NodeName;
        if (nodename = 'rgbColor') or (nodename = 'x:rgbColor') then begin
          s := GetAttrValue(colornode, 'rgb');
          if s <> '' then begin
            rgb := HTMLColorStrToColor('#' + s);
            if cidx < n then begin
              FPalette[cidx] := rgb;
              inc(cidx);
            end
            else
              FPalette.AddColor(rgb);
          end;
        end;
        colornode := colorNode.NextSibling;
      end;
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadPrintOptions(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  s: String;
  sheet: TsWorksheet absolute AWorksheet;
begin
  if ANode = nil then
    exit;
  
  s := GetAttrValue(ANode, 'headings');
  if StrIsTrue(s) then 
    with sheet.PageLayout do Options := Options + [poPrintHeaders];

  s := GetAttrValue(ANode, 'gridLines');
  if StrIsTrue(s) then 
    with sheet.PageLayout do Options := Options + [poPrintGridLines];
end;

procedure TsSpreadOOXMLReader.ReadRels(AStream: TStream; ARelsFile: String;
  ARelsList: TFPList);
var
  XMLStream: TStream;
  doc: TXMLDocument;
  node: TDOMNode;
  relID, relSchema, relTarget: String;
  nodeName: String;
begin
  // Avoid adding items again and again to the same list if this function is 
  // called multiple times.
  if ARelsList.Count > 0 then 
    exit;
  
  doc := nil;
  XMLStream := CreateXMLStream;
  try
    // Unzip the requested file or exit if the file does not exist.
    if not UnzipToStream(AStream, ARelsFile, XMLStream) then
      exit;
    ReadXMLStream(doc, XMLStream);
    node := doc.DocumentElement.FindNode('Relationship');
    while Assigned(node) do begin
      nodeName := node.NodeName;
      if nodeName = 'Relationship' then
      begin
        relID := GetAttrValue(node, 'Id');
        relSchema := GetAttrValue(node, 'Type');
        relTarget := GetAttrValue(node, 'Target');
        if (relID <> '') and (relTarget <> '') then
          (ARelsList as TXlsxRelationshipList).Add(relID, relSchema, relTarget);
      end;
      node := node.NextSibling;
    end;
  finally
    XMLStream.Free;
    doc.Free;
  end;
end;

procedure TsSpreadOOXMLReader.ReadRow(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; var ARowIndex: Cardinal);
var
  s: String;
  r: Cardinal;
  lRow: TRow;
  fmt: PsCellFormat;
  idx: Integer;
begin
  if ANode = nil then
    exit;

  { Row height type }
  s := Lowercase(GetAttrValue(ANode, 'customHeight'));
  if StrIsTrue(s) then //(s = '1') or (s = 'true') then
    lRow.RowHeightType := rhtCustom
  else
    lRow.RowHeightType := rhtAuto;

  { Row height value, in points - if there is no "ht" attribute we assume that
    it is the default row height }
  s := GetAttrValue(ANode, 'ht');
  if s = '' then begin
    lRow.Height := (AWorksheet as TsWorksheet).ReadDefaultRowHeight(FWorkbook.Units);
    lRow.RowHeightType := rhtDefault;
  end else
    lRow.Height := (FWorkbook as TsWorkbook).ConvertUnits(
      StrToFloat(s, FPointSeparatorSettings),
      suPoints,
      FWorkbook.Units
    );

  { Row index }
  s := GetAttrValue(ANode, 'r');
  if s = '' then
    r := ARowIndex
  else
    r := StrToInt(s) - 1;

  { Row format }
  lRow.FormatIndex := 0;  // Default format
  s := Lowercase(GetAttrValue(ANode, 'customFormat'));
  if StrIsTrue(s) then  //(s = '1') or (s = 'true') then
  begin
    s := GetAttrValue(ANode, 's');
    if s <> '' then begin
      idx := FCellFormatList.FindIndexOfID(StrToInt(s));
      fmt := FCellFormatList.Items[idx];
      lRow.FormatIndex := (FWorkbook as TsWorkbook).AddCellFormat(fmt^);
    end;
  end;

  { Row visibility }
  lRow.Options := [];
  s := GetAttrvalue(ANode, 'hidden');
  if StrIsTrue(s) then
    Include(lRow.Options, croHidden);

  { Write out }
  if (lRow.RowHeightType <> rhtDefault) or (lRow.FormatIndex <> 0) or (lRow.Options <> []) then
    (AWorksheet as TsWorksheet).WriteRowInfo(r, lRow);
end;

procedure TsSpreadOOXMLReader.ReadSharedStrings(ANode: TDOMNode);
var
  valuenode: TDOMNode;
  childnode: TDOMNode;
  nodename: String;
  totaltxt, sval: String;
  fntIndex: Integer;
  rtParams: TsRichTextParams = nil;
  ms: TMemoryStream;
  fnt: TsFont;
begin
  while Assigned(ANode) do begin
    nodename := ANode.NodeName;
    if (nodeName = 'si') or (nodeName = 'x:si') then begin
      totaltxt := '';
      SetLength(rtParams, 0);
      valuenode := ANode.FirstChild;
      while valuenode <> nil do begin
        nodename := valuenode.NodeName;
        if (nodename = 't') or (nodename = 'x:t') then
          // this is unformatted text
          totaltxt := GetNodeValue(valuenode)
        else
        if (nodename = 'r') or (nodename = 'x:r') then begin
          // all rich-text formatted texts are defined by r nodes
          fntIndex := -1;
          childnode := valuenode.FirstChild;
          while childnode <> nil do begin
            nodename := childnode.NodeName;
            if (nodename = 't') or (nodename = 'x:t') then
            begin
              sval := GetNodeValue(childNode);
              totaltxt := totaltxt + sval;
            end
            else
            if (nodename = 'rPr') or (nodename = 'x:rPr') then begin
              fntIndex := ReadFont(childnode);
              // Here we store the font in the internal font list of the reader.
              // But this fontindex may be different from the one needed for the
              // workbook's font list. We fix this here.
              fnt := TsFont(FFontList[fntIndex]);
              fntIndex := (Workbook as TsWorkbook).FindFont(
                fnt.FontName, fnt.Size, fnt.style, fnt.Color, fnt.Position
              );
              if fntIndex = -1 then
                fntIndex := (Workbook as TsWorkbook).AddFont(
                  fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position
                );
              SetLength(rtParams, Length(rtParams)+1);
              rtParams[High(rtParams)].FirstIndex := UTF8Length(totaltxt) + 1;
              rtParams[High(rtParams)].FontIndex := fntIndex;
              rtParams[High(rtParams)].HyperlinkIndex := -1;
            end;
            childnode := childnode.NextSibling;
          end;
        end;
        valuenode := valuenode.NextSibling;
      end;
      if Length(rtParams) = 0 then
        FSharedStrings.Add(totaltxt)
      else
      begin
        ms := TMemoryStream.Create;
        ms.WriteWord(Length(rtParams));
        ms.WriteBuffer(rtParams[0], SizeOf(TsRichTextParam)*Length(rtParams));
        FSharedStrings.AddObject(totaltxt, ms);
      end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadSheetFormatPr(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  w, h: Double;
  s: String;
begin
  if ANode = nil then
    exit;

  s := GetAttrValue(ANode, 'defaultColWidth');   // is in characters
  if (s <> '') and TryStrToFloat(s, w, FPointSeparatorSettings) then
    (AWorksheet as TsWorksheet).WriteDefaultColWidth(w, suChars);

  s := GetAttrValue(ANode, 'defaultRowHeight');  // is in points
  if (s <> '') and TryStrToFloat(s, h, FPointSeparatorSettings) then
    (AWorksheet as TsWorksheet).WriteDefaultRowHeight(h, suPoints);
end;

{ Reads the sheet parameters from the workbook.xml (node sheets) }
procedure TsSpreadOOXMLReader.ReadSheetList(ANode: TDOMNode);
var
  node: TDOMNode;
  nodename: String;
  sheetData: TSheetData;
begin
  if ANode = nil then
    exit;

  node := ANode.FirstChild;
  while node <> nil do begin
    nodename := node.NodeName;
    if (nodename = 'sheet') or (nodename = 'x:sheet') then
    begin
      sheetData := TSheetData.Create;
      sheetData.Name := GetAttrValue(node, 'name');
      sheetData.ID := GetAttrvalue(node, 'sheetId');
      sheetData.RelID := GetAttrvalue(node, 'r:id');
      sheetData.Hidden := GetAttrValue(node, 'state') = 'hidden';
      // Add the sheetdata to the SheetList.
      FSheetList.Add(sheetData);
      // Create worksheet - needed because of 3d references
      (FWorkbook as TsWorkbook).AddWorksheet(sheetData.Name, true);
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadSheetPR(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  node: TDOMNode;
  nodename: String;
begin
  if ANode = nil then
    exit;

  node := ANode.FirstChild;
  while node <> nil do begin
    nodeName := node.NodeName;
    if nodeName = 'tabColor' then
    begin
      TsWorksheet(AWorksheet).TabColor := ReadColor(node);
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadSheetProtection(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  s: String;
  shc: TsCryptoInfo;
  shp: TsWorksheetProtections;
begin
  if ANode = nil then
    exit;

  InitCryptoInfo(shc);
  s := GetAttrValue(ANode, 'password');
  if s <> '' then begin
    shc.PasswordHash := s;
    shc.Algorithm := caExcel;
  end else
  begin
    s := GetAttrValue(ANode, 'hashValue');
    if s <> '' then begin
      shc.PasswordHash := s;

      s := GetAttrValue(ANode, 'algorithmName');
      shc.Algorithm := StrToAlgorithm(s);
      if shc.Algorithm = caUnknown then
        Workbook.AddErrorMsg('Found unknown encryption algorithm "%s" for worksheet protection', [s]);

      s := GetAttrValue(ANode, 'saltValue');
      shc.SaltValue := s;

      s := GetAttrValue(ANode, 'spinCount');
      shc.SpinCount := StrToIntDef(s, 0);
    end;
  end;
  (AWorksheet as TsWorksheet).CryptoInfo := shc;

  shp := DEFAULT_SHEET_PROTECTION;

  // Attribute not found -> property = false
  s := GetAttrValue(ANode, 'sheet');
  if StrIsTrue(s) then //(s = '1') or (s = 'true') then
    Include(shp, spCells)
  else if (s = '') or StrIsFalse(s) then //(s = '0') or (s = '') or (s = 'false') then
    Exclude(shp, spCells);

  s := GetAttrValue(ANode, 'selectLockedCells');
  if StrIsTrue(s) then //(s = '1') or (s = 'true') then
    Include(shp, spSelectLockedCells)
  else if (s = '') or StrIsFalse(s) then  //(s = '0') or (s = '') or (s = 'false') then
    Exclude(shp, spSelectLockedCells);

  s := GetAttrValue(ANode, 'selectUnlockedCells');
  if StrIsTrue(s) then //(s = '1') or (s = 'true') then
    Include(shp, spSelectUnlockedCells)
  else if (s = '') or StrIsFalse(s) then  //if (s = '') or (s = '0') or (s = 'false') then
    Exclude(shp, spSelectUnlockedCells);

  s := GetAttrValue(ANode, 'objects');
  if StrIsTrue(s) then  // (s = '1') or (s = 'true') then
    Include(shp, spObjects)
  else
  if (s = '') or StrIsFalse(s) then  //(s = '') or (s = '0') or (s = 'false') then
    Exclude(shp, spObjects);

  // these options are currently not supported by fpspreadsheet
  {
  s := GetAttrValue(ANode, 'scenarios');
  if StrIsTrue(s) then
    Include(shp, spScenarios)
  else if (s = '') or StrIsFalse(s) then
    Exclude(shp, spScenarios);
  }

  // Attribute not found -> property = true
  {
  s := GetAttrValue(ANode, 'autoFilter');
  if StrIsFalse(s) then
    Exclude(shp, spAutoFilter)
  else if (s = '') or StrIsTrue(s) then
    Include(shp, spAutoFilter);
  }

  s := GetAttrValue(ANode, 'deleteColumns');
  if StrIsFalse(s) then  // (s = '0') or (s = 'true') then
    Exclude(shp, spDeleteColumns)
  else
  if (s = '') or StrIsTrue(s) then   // (s = '1') or (s = 'false') then
    Include(shp, spDeleteColumns);

  s := GetAttrValue(ANode, 'deleteRows');
  if StrIsFalse(s) then  //(s = '0') or (s = 'false') then
    Exclude(shp, spDeleteRows)
  else
  if (s = '') or StrIsTrue(s) then //(s = '1') or (s = 'true') then
    Include(shp, spDeleteRows);

  s := GetAttrValue(ANode, 'formatCells');
  if StrIsFalse(s) then   // (s = '0') or (s = 'false') then
    Exclude(shp, spFormatCells)
  else
  if (s = '') or StrIsTrue(s) then  //(s = '1') or (s = 'true') then
    Include(shp, spFormatCells);

  s := GetAttrValue(ANode, 'formatColumns');
  if StrIsFalse(s) then //(s = '0') or (s = 'false') then
    Exclude(shp, spFormatColumns)
  else
  if (s = '') or StrIsTrue(s) then   // (s = '1') or (s = 'true') then
    Include(shp, spFormatColumns);

  s := GetAttrValue(ANode, 'formatRows');
  if StrIsFalse(s) then   // (s = '0') or (s = 'false') then
    Exclude(shp, spFormatRows)
  else
  if (s = '') or StrIsTrue(s) then   // (s = '1') or (s = 'true') then
    Include(shp, spFormatRows);

  s := GetAttrValue(ANode, 'insertColumns');
  if StrIsFalse(s) then   // (s = '0') or (s = 'false') then
    Exclude(shp, spInsertColumns)
  else
  if (s = '') or StrIsTrue(s) then    // (s = '1') or (s = 'true') then
    Include(shp, spInsertColumns);

  s := GetAttrValue(ANode, 'insertHyperlinks');
  if StrIsFalse(s) then   //(s = '0') or (s = 'false') then
    Exclude(shp, spInsertHyperlinks)
  else
  if (s = '') or StrIsTrue(s) then   //(s = '1') or (s = 'true') then
    Include(shp, spInsertHyperlinks);

  s := GetAttrValue(ANode, 'insertRows');
  if StrIsFalse(s) then // (s = '0') then
    Exclude(shp, spInsertRows)
  else if (s = '') or StrIsTrue(s) then  // (s = '1') then
    Include(shp, spInsertRows);

  s := GetAttrValue(ANode, 'sort');
  if StrIsFalse(s) then   // (s = '0') or (s = 'false') then
    Exclude(shp, spSort)
  else
  if (s = '') or StrIsTrue(s) then   // (s = '1') or (s = 'true') then
    Include(shp, spSort);

  // Currently no pivottable support in fpspreadsheet
  {
  s := GetAttrValue(ANode, 'pivotTables');
  if StrIsFalse(s) then
    Exclude(shp, spPivotTables)
  else if (s = '') or StrIsTrue(s) then
    Include(shp, spPivotTables);
  }

  with AWorksheet as TsWorksheet do begin
    Protection := shp;
    Protect(true);
  end;
end;

{ Store the contents of the worksheets/_rels/sheet<n>.xml.rels files in the
  SheetData lists. 
  They contain the names of the drawing, vmlDrawing etc files needed 
  at several places.}
procedure TsSpreadOOXMLReader.ReadSheetRels(AStream: TStream);
var
  i: Integer;
  sheetData: TSheetData;
begin
  for i := 0 to FSheetList.Count-1 do
  begin
    sheetData := TSheetData(FSheetList[i]);
    sheetData.Target := TXlsxRelationshipList(FWorkbookRels).FindTarget(sheetData.RelID);
    sheetData.SheetRelsFile := OOXML_PATH_XL_WORKSHEETS + '_rels/' + ExtractFileName(sheetData.Target) + '.rels';
    ReadRels(AStream, sheetData.SheetRelsFile, sheetData.SheetRels);
  end;
end;

{ Extract the rels files related to the specified worksheet 
  (drawing<n>.xml.rel, vmlDrawing<n>.xml.rel) }
procedure TsSpreadOOXMLReader.ReadSheetRels(AStream: TStream; ANode: TDOMNode; 
  AWorksheet: TsBasicWorksheet);
var
  nodeName: string;
  sheetData: TSheetData;
  sheetIndex: Integer;
  relID: String;
  sheetRelsFile: String;
  rels: TXlsxRelationshipList;
  i: Integer;
begin
  sheetIndex := TsWorksheet(AWorksheet).Index;
  sheetData := TSheetData(FSheetList[sheetIndex]);
  sheetRelsFile := sheetData.SheetRelsFile;
  
  ANode := ANode.FirstChild;
  while Assigned(ANode) do 
  begin
    nodeName := ANode.NodeName;
    if nodeName = 'drawing' then
    begin
      // This node points to the drawing.xml file with parameters for embedded 
      // images.
      rels := TXlsxRelationshipList.Create;
      try
        ReadRels(AStream, sheetRelsFile, rels);
        relID := GetAttrValue(ANode, 'r:id');
        sheetData.Drawing_File := MakeXLPath(rels.FindTarget(relID));
        ReadRels(AStream, RelsFileFor(sheetData.Drawing_File), sheetData.DrawingRels);
      finally
        rels.Free;
      end;
    end else 
    if nodeName = 'legacyDrawingHF' then
    begin
      // This is the node pointer to parameters for heater/footer images.
      rels := TXlsxRelationshipList.Create;
      try
        ReadRels(AStream, sheetRelsFile, rels);
        relID := GetAttrValue(ANode, 'r:id');
        sheetData.VMLDrawing_File := MakeXLPath(rels.FindTarget(relID));
        ReadRels(AStream, RelsFileFor(sheetData.VMLDrawing_File), sheetData.VmlDrawingRels);
      finally
        rels.Free;
      end;
    end else
    if nodeName = 'legacyDrawing' then
    begin
      // This node is for comment size & position. We do not support this.
      // But it indicates the presence of comments and extract the name of the
      // comments.xml file from the sheet<n>.xml.rels file
      rels := TXlsxRelationshipList.Create;
      try
        ReadRels(AStream, sheetRelsFile, rels);
        for i := 0 to rels.Count-1 do
          if rels[i].Schema = SCHEMAS_COMMENTS then
          begin
            sheetData.Comments_File := rels[i].Target;
            break;
          end;
      finally
        rels.Free;
      end;
    end else
    if nodeName = 'hyperlinks' then
    begin
      // This node contains the hyperlink data and will be handled separately.
      // It contains also the relationship ids to external files listed in
      // the sheet<n>.xml.rels file. We read these relationships here and store
      // them in the HyperlinkRels of the sheetdata.
      rels := TXlsxRelationshipList.Create;
      try
        ReadRels(AStream, sheetRelsfile, rels);
        for i := 0 to rels.Count-1 do
          if rels[i].Schema = SCHEMAS_HYPERLINK then
            sheetData.HyperlinkRels.Add(rels[i].RelID, rels[i].Schema, rels[i].Target);
      finally
        rels.Free;
      end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadSheetViews(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  sheetViewNode: TDOMNode;
  childNode: TDOMNode;
  nodeName: String;
  s: String;
  actRow, actCol: Cardinal;
  sheet: TsWorksheet absolute AWorksheet;
begin
  if ANode = nil then
    exit;

  sheetViewNode := ANode.FirstChild;
  while Assigned(sheetViewNode) do begin
    nodeName := sheetViewNode.NodeName;
    if (nodeName = 'sheetView') or (nodeName = 'x:sheetView') then begin
      s := GetAttrValue(sheetViewNode, 'showGridLines');
      if StrIsFalse(s) then //(s = '0') or (s = 'false') then
        sheet.Options := AWorksheet.Options - [soShowGridLines];

      s := GetAttrValue(sheetViewNode, 'showRowColHeaders');
      if StrIsFalse(s) then //(s = '0') or (s = 'false') then
         sheet.Options := AWorksheet.Options - [soShowHeaders];

      s := GetAttrValue(sheetViewNode, 'tabSelected');
      if StrIsTrue(s) then  //(s = '1') or (s = 'true') then
        (FWorkbook as TsWorkbook).ActiveWorksheet := sheet;

      s := Lowercase(GetAttrValue(sheetViewNode, 'windowProtection'));
      if StrIsTrue(s) then //(s = '1') or (s = 'true') then
        sheet.Options := sheet.Options + [soPanesProtection];

      s := GetAttrValue(sheetViewNode, 'zoomScale');
      if s <> '' then
        sheet.ZoomFactor := StrToFloat(s, FPointSeparatorSettings) * 0.01;

      s := Lowercase(GetAttrValue(sheetViewNode, 'rightToLeft'));
      if (s = '') or StrIsFalse(s) then  //(s = '0') or (s = 'false') then
        sheet.BiDiMode := bdLTR
      else if StrIsTrue(s) then //(s = '1') or (s = 'true') then
        sheet.BiDiMode := bdRTL;

      childNode := sheetViewNode.FirstChild;
      while Assigned(childNode) do begin
        nodeName := childNode.NodeName;
        if (nodeName = 'pane') or (nodeName = 'x:pane') then begin
          s := GetAttrValue(childNode, 'state');
          if s = 'frozen' then begin
            sheet.Options := sheet.Options + [soHasFrozenPanes];
            s := GetAttrValue(childNode, 'xSplit');
            if s <> '' then
              sheet.LeftPaneWidth := round(StrToFloat(s, FPointSeparatorSettings));
            s := GetAttrValue(childNode, 'ySplit');
            if s <> '' then
              sheet.TopPaneHeight := round(StrToFloat(s, FPointSeparatorSettings));
          end;
        end else
        if (nodeName = 'selection') or (nodeName = 'x:selection') then begin
          s := GetAttrValue(childnode, 'activeCell');
          if s <> '' then
          begin
            ParseCellString(s, actRow, actCol);
            sheet.SelectCell(actRow, actCol);
          end;
        end;
        childNode := childNode.NextSibling;
      end;
    end;
    sheetViewNode := sheetViewNode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadThemeColors(ANode: TDOMNode);
var
  clrNode: TDOMNode;
  nodeName: String;

  procedure AddColor(AColorStr: String);
  begin
    if AColorStr <> '' then begin
      SetLength(FThemeColors, Length(FThemeColors)+1);
      FThemeColors[Length(FThemeColors)-1] := HTMLColorStrToColor('#' + AColorStr);
    end;
  end;

begin
  if not Assigned(ANode) then
    exit;

  SetLength(FThemeColors, 0);
  clrNode := ANode.FirstChild;
  while Assigned(clrNode) do begin
    nodeName := clrNode.NodeName;
    if nodeName = 'a:dk1' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'lastClr'))
    else
    if nodeName = 'a:lt1' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'lastClr'))
    else
    if nodeName = 'a:dk2' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'val'))
    else
    if nodeName = 'a:lt2' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'val'))
    else
    if nodeName = 'a:accent1' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'val'))
    else
    if nodeName = 'a:accent2' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'val'))
    else
    if nodeName = 'a:accent3' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'val'))
    else
    if nodeName = 'a:accent4' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'val'))
    else
    if nodeName = 'a:accent5' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'val'))
    else
    if nodeName = 'a:accent6' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'val'))
    else
    if nodeName = 'a:hlink' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'val'))
    else
    if nodeName = 'a:folHlink' then
      AddColor(GetAttrValue(clrNode.FirstChild, 'aval'));
    clrNode := clrNode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadThemeElements(ANode: TDOMNode);
var
  childNode: TDOMNode;
  nodeName: String;
begin
  if not Assigned(ANode) then
    exit;
  childNode := ANode.FirstChild;
  while Assigned(childNode) do begin
    nodeName := childNode.NodeName;
    if nodeName = 'a:clrScheme' then
      ReadThemeColors(childNode);
    childNode := childNode.NextSibling;
  end;
end;

{ Reads the parameters of the header/footer images defined as children of the 
  specified node which is in a vmlDrawingX.xml file.  }
procedure TsSpreadOOXMLReader.ReadVmlDrawing(ANode: TDOMNode; 
  AWorksheet: TsBasicWorksheet);

  function ExtractFromStyle(AStyle, AKey: String): String;
  var
    sa: TStringArray;
    i: Integer;
  begin
    {$IFDEF FPS_NO_STRING_SPLIT}
    sa := SplitString(AStyle, ',');
    {$ELSE}
    sa := AStyle.Split(';');
    {$ENDIF}
    for i := 0 to High(sa) do
      if pos(AKey, sa[i]) = 1 then
      begin
        Result := Copy(sa[i], Length(AKey)+1, MaxInt);
        exit;
      end;
  end;
  
  function ExtractMMFromStyle(AStyle, AKey: String): Double;
  var
    s: String;
    sval: String = '';
    sunit: String = '';
    i: Integer;
  begin
    s := ExtractFromStyle(AStyle, AKey);
    for i := 1 to Length(s) do
      if s[i] in ['0'..'9', '.', '+', '-'] then
        sval := sval + s[i]
      else 
        sunit := sunit + s[i];
    Result := StrToFloat(sval, FPointSeparatorSettings);
    case sunit of
      'pt': Result := PtsToMM(Result);
      'mm': ;
      else  raise Exception.Create('Unit not supported.');
    end;
  end;
  
var
  nodeName: String;
  node: TDOMNode;
  relID: String;
  title: String;
  style: String;
  data: TEmbeddedObjData;
  id: String;
  sheetData: TSheetData;
begin
  sheetData := TSheetData(FSheetList[TsWorksheet(AWorksheet).Index]);
  
  ANode := ANode.FirstChild;
  while Assigned(ANode) do
  begin
    id := '';
    relID := '';
    title := '';
    style := '';
    nodeName := ANode.NodeName;
    if nodeName = 'v:shape' then
    begin
      id := GetAttrValue(ANode, 'id');
      style := GetAttrValue(ANode, 'style');
      node := ANode.FirstChild;
      while Assigned(node) do
      begin
        nodeName := node.NodeName;
        if nodeName = 'v:imagedata' then
        begin
          relID := GetAttrValue(node, 'o:relid');
          title := GetAttrValue(node, 'o:title');
        end;
        node := node.NextSibling;
      end;
      if relID <> '' then
      begin
        data := TEmbeddedObjData.Create;
        data.FromCol := UNASSIGNED_ROW_COL_INDEX;
        data.FromColOffs := 0.0;
        data.ToCol := UNASSIGNED_ROW_COL_INDEX;
        data.ToColOffs := 0.0;
        data.FromRow := UNASSIGNED_ROW_COL_INDEX;
        data.FromRowOffs := 0.0;
        data.ToRow := UNASSIGNED_ROW_COL_INDEX;
        data.ToRowOffs := 0.0;
        data.RelId := relId;
        data.FileName := title;
        data.MediaName := MakeXLPath(sheetData.VMLDrawingRels.FindTarget(relID));
        data.ImgIndex := -1;
        data.Worksheet := AWorksheet;
        data.IsHeaderFooter := true;
        data.IsChart := false;
        data.HFImgPosition := id;
        data.HFImgWidth := ExtractMMFromStyle(style, 'width:');
        data.HFImgHeight := ExtractMMFromStyle(style, 'height:');
        FEmbeddedObjList.Add(data);
      end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadWorkbookProtection(ANode: TDOMNode);
var
  s : string;
  wbc: TsCryptoInfo;
  wbp: TsWorkbookProtections;
begin
  if ANode = nil then
    Exit;

  s := '';
  wbp := [];

  InitCryptoInfo(wbc);
  s := GetAttrValue(ANode, 'workbookPassword');
  if s <> '' then
    wbc.PasswordHash := s
  else
  begin
    s := GetAttrValue(ANode, 'workbookHashVal');
    if s <> '' then begin
      wbc.PasswordHash := s;

      s := GetAttrValue(ANode, 'workbookAlgorithmName');
      wbc.Algorithm := StrToAlgorithm(s);
      if wbc.Algorithm = caUnknown then
        Workbook.AddErrorMsg('Found unknown encryption algorithm "%s" for workbook protection', [s]);

      wbc.SaltValue := GetAttrValue(ANode, 'workbookSaltValue');

      wbc.SpinCount := StrToIntDef(GetAttrValue(ANode, 'workbookSpinCount'), 0);
    end;
  end;
  (Workbook as TsWorkbook).CryptoInfo := wbc;

  {
  InitCryptoInfo(wbc);
  s := GetAttrValue(ANode, 'revisionsPassword');
  if s <> '' then
    wbc.Password := s
  else
  begin
    s := GetAttrValue(ANode, 'revisionsHashValue');
    if s <> '' then begin
      wbc.HashValue := s;
      wbc.AlgorithmName := GetAttrValue(ANode, 'revisionsAlgorithm');
      wbc.SaltValue := GetAttrValue(ANode, 'revisionsSaltValue');
      wbc.SpinCount := StrToIntDef(GetAttrValue(ANode, 'revisionsSpinCount'), 0);
    end;
  end;
  Workbook.RevisionsCrypto := wbc;
  }
  s := GetAttrValue(ANode, 'lockStructure');
  if StrIsTrue(s) then  // (s = '1') or (s = 'true')then
    Include(wbp, bpLockStructure);

  s := GetAttrValue(ANode, 'lockWindows');
  if StrIsTrue(s) then  //(s = '1') or (s = 'true') then
    Include(wbp, bpLockWindows);

  s := GetAttrValue(ANode, 'lockRevision');
  if StrIsTrue(s) then  //(s = '1') or (s = 'true') then
    Include(wbp, bpLockRevision);

  Workbook.Protection := wbp;
end;

procedure TsSpreadOOXMLReader.ReadWorksheet(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  rownode: TDOMNode;
  cellnode: TDOMNode;
  nodename: String;
  r, c: Cardinal;
begin
  rownode := ANode.FirstChild;
  r := 0;
  while Assigned(rownode) do begin
    nodeName := rownode.NodeName;
    if (nodeName = 'row') or (nodeName = 'x:row') then begin
      ReadRow(rownode, AWorksheet, r);
      cellnode := rownode.FirstChild;
      c := 0;
      while Assigned(cellnode) do begin
        nodename := cellnode.NodeName;
        if (nodeName = 'c') or (nodeName = 'x:c') then
          ReadCell(cellnode, AWorksheet, r, c);
        cellnode := cellnode.NextSibling;
      end;
      inc(r);
    end;
    rownode := rownode.NextSibling;
  end;
  FixCols(AWorksheet);
  FixRows(AWorksheet);
end;

procedure TsSpreadOOXMLReader.ReadFromStream(AStream: TStream;
  APassword: String = ''; AParams: TsStreamParams = []);
var
  Doc : TXMLDocument;
  i, j: Integer;
  fn: String;
  XMLStream: TStream;
  actSheetIndex: Integer;
  sheetData: TSheetData;

  function Doc_FindNode(ANodeName: String): TDOMNode;
  begin
    Result := Doc.DocumentElement.FindNode(ANodeName);
    if Result = nil then
      Result := Doc.DocumentElement.FindNode('x:' + ANodeName);
  end;

begin
  Unused(APassword, AParams);
  Doc := nil;

  CheckPassword(AStream, APassword);

  try
    // Retrieve theme colors
    XMLStream := CreateXMLStream;
    try
      if UnzipToStream(AStream, OOXML_PATH_XL_THEME, XMLStream) then
      begin
        ReadXMLStream(Doc, XMLStream);
        ReadThemeElements(Doc.DocumentElement.FindNode('a:themeElements'));
        FreeAndNil(Doc);
      end;
    finally
      XMLStream.Free;
    end;

    // process the workbook.xml file (1st run)
    XMLStream := CreateXMLStream;
    try
      if not UnzipToStream(AStream, OOXML_PATH_XL_WORKBOOK, XMLStream) then
        raise EFPSpreadsheetReader.CreateFmt(rsDefectiveInternalFileStructure, ['xlsx']);
      ReadXMLStream(Doc, XMLStream);
      ReadFileVersion(Doc_FindNode('fileVersion'));
      ReadDateMode(Doc_FindNode('workbookPr'));
      ReadWorkbookProtection(Doc_FindNode('workbookProtection'));
      ReadSheetList(Doc_FindNode('sheets'));    // This creates the worksheets!
      ReadRels(AStream, OOXML_PATH_XL_RELS_RELS, FWorkbookRels);
      ReadSheetRels(AStream);
      ReadDefinedNames(Doc.DocumentElement.FindNode('definedNames'));
      ReadActiveSheet(Doc_FindNode('bookViews'), actSheetIndex);
      FreeAndNil(Doc);
    finally
      XMLStream.Free;
    end;

    // process the styles.xml file
    XMLStream := CreateXMLStream;
    try
      // Should always exist, just to make sure...
      if UnzipToStream(AStream, OOXML_PATH_XL_STYLES, XMLStream) then
      begin
        ReadXMLStream(Doc, XMLStream);
        ReadPalette(Doc_FindNode('colors'));
        ReadFonts(Doc_FindNode('fonts'));
        ReadFills(Doc_FindNode('fills'));
        ReadBorders(Doc_FindNode('borders'));
        ReadNumFormats(Doc_FindNode('numFmts'));
        ReadCellXfs(Doc_FindNode('cellXfs'));
        ReadDifferentialFormats(DOC_FindNode('dxfs'));
        FreeAndNil(Doc);
      end;
    finally
      XMLStream.Free;
    end;

    // process the sharedstrings.xml file
    // To do: Use buffered stream instead since shared strings may be large
    XMLStream := CreateXMLStream;
    try
      if UnzipToStream(AStream, OOXML_PATH_XL_STRINGS, XMLStream) then
      begin
        ReadXMLStream(Doc, XMLStream);
        ReadSharedStrings(Doc_FindNode('si'));
        FreeAndNil(Doc);
      end;
    finally
      XMLStream.Free;
    end;

    // read worksheet contents
    for i:=0 to FSheetList.Count-1 do begin
      sheetData := TSheetData(FSheetList[i]);
      
      // Worksheets are already created...
      FWorksheet := (FWorkbook as TsWorkbook).GetWorksheetByName(sheetData.Name);
      if sheetData.Hidden then
        FWorksheet.Options := FWorksheet.Options + [soHidden];

      // unzip sheet file
      XMLStream := CreateXMLStream;
      try
        fn := TXlsxRelationshipList(FWorkbookRels).FindTarget(sheetData.RelID);
        if fn = '' then
          fn := OOXML_PATH_XL_WORKSHEETS + Format('sheet%d.xml', [i+1])
        else
          fn := OOXML_PATH_XL_WORKSHEETS + ExtractFileName(fn);
        if not UnzipToStream(AStream, fn, XMLStream) then
          Continue;
        ReadXMLStream(Doc, XMLStream);
      finally
        XMLStream.Free;
      end;

      // clear sharedformulabase list
      for j:=FSharedFormulaBaseList.Count-1 downto 0 do
        TObject(FSharedFormulaBaseList[j]).Free;
      FSharedFormulaBaseList.Clear;

      // Sheet data, formats, etc.
      ReadSheetRels(AStream, Doc.DocumentElement, FWorksheet);
      ReadSheetPr(Doc_FindNode('sheetPr'), FWorksheet);
      ReadDimension(Doc_FindNode('dimension'), FWorksheet);
      ReadSheetViews(Doc_FindNode('sheetViews'), FWorksheet);
      ReadSheetFormatPr(Doc_FindNode('sheetFormatPr'), FWorksheet);
      ReadCols(Doc_FindNode('cols'), FWorksheet);
      ReadWorksheet(Doc_FindNode('sheetData'), FWorksheet);
      ReadConditionalFormatting(Doc_FindNode('conditionalFormatting'), FWorksheet);
      ReadSheetProtection(Doc_FindNode('sheetProtection'), FWorksheet);
      ReadMergedCells(Doc_FindNode('mergeCells'), FWorksheet);
      ReadHyperlinks(Doc_FindNode('hyperlinks'), FWorksheet);
      ReadPrintOptions(Doc_FindNode('printOptions'), FWorksheet);
      ReadPageMargins(Doc_FindNode('pageMargins'), FWorksheet);
      ReadPageSetup(Doc_FindNode('pageSetup'), FWorksheet);
      ReadColRowBreaks(Doc_FindNode('rowBreaks'), FWorksheet);
      ReadColRowBreaks(Doc_FindNode('colBreaks'), FWorksheet);
      ReadHeaderFooter(Doc_FindNode('headerFooter'), FWorksheet);

      FreeAndNil(Doc);

      ReadComments(AStream, FWorksheet);

      // Active worksheet
      if i = actSheetIndex then
        (FWorkbook as TsWorkbook).SelectWorksheet(FWorksheet as TsWorksheet);
    end;  // for

    // Read embedded images
    ReadEmbeddedObjs(AStream);
    
    // MetaData
    XMLStream := CreateXMLStream;
    try
      if UnzipToStream(AStream, OOXML_PATH_DOCPROPS_CORE, XMLStream) then
      begin
        ReadXMLStream(Doc, XMLStream);
        ReadMetaData(Doc.DocumentElement);
        FreeAndNil(Doc);
      end;
    finally
      XMLStream.Free;
    end;
    // custom meta data
    XMLStream := CreateXMLStream;
    try
      if UnzipToStream(AStream, OOXML_PATH_DOCPROPS_CUSTOM, XMLStream) then
      begin
        ReadXMLStream(Doc, XMLStream);
        ReadMetaData(Doc.DocumentElement);
        FreeAndNil(Doc);
      end;
    finally
      XMLStream.Free;
    end;

  finally
    FreeAndNil(Doc);
  end;
end;


{------------------------------------------------------------------------------}
{                             TsSpreadOOXMLWriter                              }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the OOXML writer

  Defines the date mode and the limitations of the file format.
  Initializes the format settings to be used when writing to xml.
-------------------------------------------------------------------------------}
constructor TsSpreadOOXMLWriter.Create(AWorkbook: TsBasicWorkbook);
begin
  inherited Create(AWorkbook);

 {$ifdef FPS_CHARTS}
  FChartWriter := TsSpreadOOXMLChartWriter.Create(self);
 {$endif}

  // Initial base date in case it won't be set otherwise.
  // Use 1900 to get a bit more range between 1900..1904.
  FDateMode := XlsxSettings.DateMode;

  // Special version of FormatSettings using a point decimal separator for sure.
  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';

  InitOOXMLLimitations(FLimitations);
end;


destructor TsSpreadOOXMLWriter.Destroy;
begin
 {$ifdef FPS_CHARTS}
  FChartWriter.Free;
 {$endif}
  inherited;
end;


{@@ ----------------------------------------------------------------------------
  Looks for the combination of border attributes of the given format record in
  the FBorderList and returns its index.
-------------------------------------------------------------------------------}
function TsSpreadOOXMLWriter.FindBorderInList(AFormat: PsCellFormat): Integer;
var
  i: Integer;
  fmt: PsCellFormat;
begin
  // No cell, or border-less --> index 0
  if (AFormat = nil) or not (uffBorder in AFormat^.UsedFormattingFields) then begin
    Result := 0;
    exit;
  end;

  for i:=0 to High(FBorderList) do begin
    fmt := FBorderList[i];
    if SameCellBorders(fmt, AFormat) then begin
      Result := i;
      exit;
    end;
  end;

  // Not found --> return -1
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
  Looks for the combination of fill attributes of the given format record in the
  FFillList and returns its index.
-------------------------------------------------------------------------------}
function TsSpreadOOXMLWriter.FindFillInList(AFormat: PsCellFormat): Integer;
var
  i: Integer;
  fmt: PsCellFormat;
begin
  if (AFormat = nil) or not (uffBackground in AFormat^.UsedFormattingFields)
  then begin
    Result := 0;
    exit;
  end;

  // Index 0 is "no fill" which already has been handled.
  for i:=1 to High(FFillList) do begin
    fmt := FFillList[i];
    if (fmt <> nil) and (uffBackground in fmt^.UsedFormattingFields) then
    begin
      if (AFormat^.Background.Style = fmt^.Background.Style) and
         (AFormat^.Background.BgColor = fmt^.Background.BgColor) and
         (AFormat^.Background.FgColor = fmt^.Background.FgColor)
      then begin
        Result := i;
        exit;
      end;
    end;
  end;

   // Not found --> return -1
  Result := -1;
end;

{ Calculates the rIds for comments, hyperlinks, image, charts and
  header/footer images of the specified worksheet }
procedure TsSpreadOOXMLWriter.Get_rId(AWorksheet: TsBasicWorksheet;
  out AComment_rId, AFirstHyperlink_rId, ADrawing_rId, ADrawingHF_rId: Integer);
var
  next_rId: Integer;
  sheet: TsWorksheet;
  book: TsWorkbook;
begin
  sheet := AWorksheet as TsWorksheet;
  book := sheet.Workbook;

  AComment_rId := -1;
  AFirstHyperlink_rId := -1;
  ADrawing_rId := -1;
  ADrawingHF_rId := -1;
  next_rId := 1;

  // Comments first
  if sheet.Comments.Count > 0 then
  begin
    AComment_rId := next_rId;
    inc(next_rId, 2);  // there are two .rels entries in case of comments
  end;

  // Charts or embedded images next
  if (sheet.GetImageCount > 0) {$ifdef FPS_CHARTS}or (book.GetChartCount > 0){$endif} then
  begin
    ADrawing_rId := next_rId;
    inc(next_rId);
  end;

  // HeaderFooter images next
  if sheet.PageLayout.HasHeaderFooterImages then
  begin
    ADrawingHF_rId := next_rId;
    inc(next_rId);
  end;

  // Hyperlinks at the end because it is not clear how many rIds will be
  // used without analyzing the hyperlink.
  if sheet.Hyperlinks.Count > 0 then
    AFirstHyperlink_rId := next_rId;
end;

function TsSpreadOOXMLWriter.GetActiveTab: String;
var
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;
  Result := IfThen(book.ActiveWorksheet = nil, '',
    ' activeTab="' + IntToStr(book.GetWorksheetIndex(book.ActiveWorksheet)) + '"');
end;

{ Determines the formatting index which a given cell has in list of
  "FormattingStyles" which correspond to the section cellXfs of the styles.xml
  file. }
function TsSpreadOOXMLWriter.GetStyleIndex(ACell: PCell): Cardinal;
begin
  Result := ACell^.FormatIndex;
end;

{ Creates a list of all border styles found in the workbook.
  The list contains indexes into the array FFormattingStyles for each unique
  combination of border attributes.
  To be used for the styles.xml. }
procedure TsSpreadOOXMLWriter.ListAllBorders;
var
  //styleCell: PCell;
  i, n : Integer;
  fmt: PsCellFormat;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;

  // first list entry is a no-border cell
  n := 1;
  SetLength(FBorderList, n);
  FBorderList[0] := nil;

  for i := 0 to book.GetNumCellFormats - 1 do
  begin
    fmt := book.GetPointerToCellFormat(i);
    if FindBorderInList(fmt) = -1 then
    begin
      SetLength(FBorderList, n+1);
      FBorderList[n] := fmt;
      inc(n);
    end;
  end;
end;

{ FDifferentialFormatIndexList stores the indexes of the cells formats used
  in conditional formatting. }
procedure TsSpreadOOXMLWriter.ListAllDifferentialFormats;
var
  book: TsWorkbook;
  n: Integer;
  idx: Integer;
  j, k, d: Integer;
  CF: TsConditionalFormat;
  rule: TsCFCellRule;
begin
  n := 0;
  SetLength(FDifferentialFormatIndexList, n);

  book := TsWorkbook(FWorkbook);
  for j := 0 to book.GetNumConditionalFormats-1 do
  begin
    CF := book.GetConditionalFormat(j);
    for k := 0 to CF.RulesCount-1 do
      if CF.Rules[k] is TsCFCellRule then
      begin
        rule := TsCFCellRule(CF.Rules[k]);
        idx := -1;
        for d := 0 to High(FDifferentialFormatIndexList) do
          if FDifferentialFormatIndexList[d] = rule.FormatIndex then
          begin
            idx := d;
            break;
           end;
        if idx = -1 then
        begin
          SetLength(FDifferentialFormatIndexList, n+1);
          FDifferentialFormatIndexList[n] := rule.FormatIndex;
          inc(n);
        end;
      end;
  end;
end;

{ Creates a list of all fill styles found in the workbook.
  The list contains indexes into the array FFormattingStyles for each unique
  combination of fill attributes.
  Currently considers only backgroundcolor, fill style is always "solid".
  To be used for styles.xml. }
procedure TsSpreadOOXMLWriter.ListAllFills;
var
  i, n: Integer;
  fmt: PsCellFormat;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;

  // Add built-in fills first.
  n := 2;
  SetLength(FFillList, n);
  FFillList[0] := nil;  // built-in "no fill"
  FFillList[1] := nil;  // built-in "gray125"

  for i := 0 to book.GetNumCellFormats - 1 do
  begin
    fmt := book.GetPointerToCellFormat(i);
    if FindFillInList(fmt) = -1 then
    begin
      SetLength(FFillList, n+1);
      FFillList[n] := fmt;
      inc(n);
    end;
  end;
end;

procedure TsSpreadOOXMLWriter.WriteBorderList(AStream: TStream);

  procedure WriteBorderStyle(AStream: TStream; AFormatRecord: PsCellFormat;
    ABorder: TsCellBorder; ABorderName: String);
  { border names found in xlsx files for Excel selections:
    "thin", "hair", "dotted", "dashed", "dashDotDot", "dashDot", "mediumDashDotDot",
    "slantDashDot", "mediumDashDot", "mediumDashed", "medium", "thick", "double" }
  var
    styleName: String;
    colorStr: String;
    rgb: TsColor;
  begin
    if (ABorder in AFormatRecord^.Border) then begin
      // Line style
      styleName := LINESTYLE_TYPES[AFormatRecord^.BorderStyles[ABorder].LineStyle];

      // Border color
      rgb := AFormatRecord^.BorderStyles[ABorder].Color;
      colorStr := ColorToExcelStr(rgb);
      AppendToStream(AStream, Format(
        '<%s style="%s"><color rgb="%s" /></%s>',
          [ABorderName, styleName, colorStr, ABorderName]
        ));
    end else
      AppendToStream(AStream, Format(
        '<%s />', [ABorderName]));
  end;

var
  i: Integer;
  diag: String;
begin
  AppendToStream(AStream, Format(
    '<borders count="%d">', [Length(FBorderList)]));

  // index 0 -- built-in "no borders"
  AppendToStream(AStream,
      '<border>',
        '<left /><right /><top /><bottom /><diagonal />',
      '</border>');

  for i:=1 to High(FBorderList) do begin
    diag := '';
    if (cbDiagUp in FBorderList[i]^.Border) then
      diag := diag + ' diagonalUp="1"';
    if (cbDiagDown in FBorderList[i]^.Border) then diag := diag + ' diagonalDown="1"';
    AppendToStream(AStream,
      '<border' + diag + '>');
        WriteBorderStyle(AStream, FBorderList[i], cbWest, 'left');
        WriteBorderStyle(AStream, FBorderList[i], cbEast, 'right');
        WriteBorderStyle(AStream, FBorderList[i], cbNorth, 'top');
        WriteBorderStyle(AStream, FBorderList[i], cbSouth, 'bottom');
        // OOXML uses the same border style for both diagonals. In agreement with
        // the biff implementation we select the style from the diagonal-up line.
        WriteBorderStyle(AStream, FBorderList[i], cbDiagUp, 'diagonal');
    AppendToStream(AStream,
      '</border>');
  end;

  AppendToStream(AStream,
    '</borders>');
end;


procedure TsSpreadOOXMLWriter.WriteCFCellRule(AStream: TStream;
  ARule: TsCFCellRule; ARange: TsCellRange; APriority: Integer);
const
  FORMULA: array[cfcBeginsWith..cfcNextMonth] of String = (
    'LEFT(%0:s,LEN("%1:s"))="%1:s"',      // cfcBeginsWith
    'RIGHT(%0:s,Len("%1:s"))="%1:s"',     // cfcEndsWidth
    'NOT(ISERROR(SEARCH("%1:s",%0:s)))',  // cfcContainsText
    'ISERROR(SEARCH("%1:s",%0:s))',       // cfcNotContainsText
    'ISERROR(%0:s)',                      // cfcContainsErrors
    'NOT(ISERROR(%0:s))',                 // cfcNotContainsErrors
    'FLOOR(%s,1)=TODAY()-1',              // cfcYesterday
    'FLOOR(%s,1)=TODAY()',                // cfcToday
    'FLOOR(%s,1)=TODAY()+1',              // cfcTomorrow
    'AND(TODAY()-FLOOR(%0:s,1)&lt;=6,FLOOR(%0:s,1)&lt;=TODAY())',                                                // cfcLast7Days
    'AND(TODAY()-ROUNDDOWN(%0:s,0)&gt;=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(%0:s,0)&lt;(WEEKDAY(TODAY())+7))',   // cfcLastWeek
    'AND(TODAY()-ROUNDDOWN(%0:s,0)&lt;=WEEKDAY(TODAY())-1,ROUNDDOWN(%0:s,0)-TODAY()&lt;=7-WEEKDAY(TODAY()))',    // cfcThisWeek
    'AND(ROUNDDOWN(%0:s,0)-TODAY()&gt;(7-WEEKDAY(TODAY())),ROUNDDOWN(C15,0)-TODAY()&lt;(15-WEEKDAY(TODAY())))',  // cfcNextWeek
    'AND(MONTH(%0:s)=MONTH(EDATE(TODAY(),0-1)),YEAR(%0:s)=YEAR(EDATE(TODAY(),0-1)))',   // cfcLastMonth
    'AND(MONTH(%0:s)=MONTH(TODAY()),YEAR(%0:s)=YEAR(TODAY()))',                         // cfcThisMonth
    'AND(MONTH(%0:s)=MONTH(EDATE(TODAY(),0+1)),YEAR(%0:s)=YEAR(EDATE(TODAY(),0+1)))'    // cfcNextMonth
  );
var
  i: Integer;
  dxfID: Integer;
  typeStr, opStr, formula1Str, formula2Str, param1Str, param2Str, param3Str: String;
  firstCellOfRange: String;
  s: String;
begin
  dxfID := -1;
  for i := 0 to High(FDifferentialFormatIndexList) do
    if FDifferentialFormatIndexList[i] = ARule.FormatIndex then
    begin
      dxfID := i;
      break;
    end;

  typeStr := CF_TYPE_NAMES[ARule.Condition];
  if CF_OPERATOR_NAMES[ARule.Condition] = '' then
    opStr := ''
  else
    opStr := ' operator="' + CF_OPERATOR_NAMES[ARule.Condition] + '"';
  formula1Str := '';
  formula2Str := '';
  param1Str := '';
  param2Str := '';
  param3Str := '';
  case ARule.Condition of
    cfcEqual..cfcNotBetween:
      begin
        s := CFOperandToStr(ARule.Operand1, FPointSeparatorSettings);
        formula1Str := Format('<formula>%s</formula>', [s]);
        if (ARule.Condition in [cfcBetween, cfcNotBetween]) then
        begin
          s := CFOperandToStr(ARule.Operand2, FPointSeparatorSettings);
          formula2Str := Format('<formula>%s</formula>', [s]);
        end;
      end;
    cfcAboveAverage..cfcBelowEqualAverage:
      begin
        if (ARule.Condition in [cfcBelowAverage, cfcBelowEqualAverage]) then
          param1Str := ' aboveAverage="0"';
        if (ARule.Condition in [cfcAboveEqualAverage, cfcBelowEqualAverage]) then
          param2Str := ' equalAverage="1"';
        if VarIsNumeric(ARule.Operand1) or (ARule.Operand1 = 0) then
          param3Str := Format(' stdDev="%g"', [double(ARule.Operand1)], FPointSeparatorSettings);
      end;
    cfcTop, cfcBottom, cfcTopPercent, cfcBottomPercent:
      begin
        // <cfRule type="top10" dxfId="0" priority="1" percent="1" bottom="1" rank="30" />    // = bottom 30 percent
        if ARule.Condition in [cfcBottom, cfcBottomPercent] then
          param1Str := ' bottom="1"';
        if ARule.Condition in [cfcTopPercent, cfcBottomPercent] then
          param2Str := ' percent="1"';
        param3Str := ' rank="' + VarToStr(ARule.Operand1) + '"';
      end;
    cfcDuplicate, cfcUnique:
      ;
    cfcBeginsWith..cfcNotContainsErrors:
      begin
        firstCellOfRange := GetCellString(ARange.Row1, ARange.Col1);
        formula1Str :=
          '<formula>' +
            Format(FORMULA[ARule.Condition], [firstcellOfRange, ARule.Operand1]) +
          '</formula>';
        param1Str := ' text="' + VarToStr(ARule.Operand1) + '"';
      end;
    cfcYesterday..cfcNextMonth:
      begin
        firstCellOfrange := GetCellString(ARange.Row1, ARange.Col1);
        s := Format(FORMULA[ARule.Condition], [firstcellOfRange]);
        formula1Str := '<formula>' + s + '</formula>';
        param1Str := Format(' timePeriod="%s"', [CF_OPERATOR_NAMES[ARule.Condition]]);
        opStr := '';
      end;
    cfcLastYear..cfcNextYear:
      begin
        firstCellOfRange := GetCellString(ARange.Row1, ARange.Col1);
        s := Format(CF_YEAR_FORMULAS[ARule.Condition], [firstCellOfRange]);
        formula1Str := '<formula>' + s + '</formula>';
      end;
    cfcExpression:
      begin
        s := ARule.Operand1;
        if (s <> '') and (s[1] = '=') then Delete(s, 1, 1);
        formula1Str := '<formula>' + s + '</formula>';
      end;
    else
      FWorkbook.AddErrorMsg('ConditionalFormat operator not supported.');
  end;

  if formula1Str = '' then
    s := Format(
      '<cfRule type="%s" dxfId="%d" priority="%d"%s%s%s%s />', [
      typeStr, dxfId, APriority, opStr, param1Str, param2Str, param3Str
    ])
  else
    s := Format(
      '<cfRule type="%s" dxfId="%d" priority="%d"%s%s%s%s>' +
        '%s%s' +
      '</cfRule>', [
      typeStr, dxfId, APriority, opStr, param1Str, param2Str, param3Str,
      formula1Str, formula2Str
    ]);
  AppendToStream(AStream, s);
end;


procedure TsSpreadOOXMLWriter.WriteCFColorRangeRule(AStream: TStream;
  ARule: TsCFColorRangeRule; APriority: Integer);
{ example:
    <cfRule type="colorScale" priority="3">
      <colorScale>
        <cfvo type="min" />
        <cfvo type="percentile" val="50" />
        <cfvo type="max" />
        <color rgb="FFF8696B" />
        <color rgb="FFFFEB84" />
        <color rgb="FF63BE7B" />
      </colorScale>
    </cfRule> }
begin
  AppendToStream(AStream,
    '<cfRule type="colorScale" priority="' + IntToStr(APriority) + '">' +
      '<colorScale>');
  AppendToStream(AStream,
        CF_ValueNode(ARule.StartValueKind, ARule.StartValue),
        IfThen(ARule.ThreeColors, CF_ValueNode(ARule.CenterValueKind, ARule.CenterValue), ''),
        CF_ValueNode(ARule.EndValueKind, ARule.EndValue)
  );
  AppendToStream(AStream,
        CF_ColorNode(ARule.StartColor),
        IfThen(ARule.ThreeColors, CF_ColorNode(ARule.CenterColor), ''),
        CF_ColorNode(ARule.EndColor)
  );
  AppendToStream(AStream,
      '</colorScale>' +
    '</cfRule>');
end;


procedure TsSpreadOOXMLWriter.WriteCFDatabarRule(AStream: TStream;
  ARule: TsCFDataBarRule; APriority: Integer);
(* example from test file:
      <cfRule type="dataBar" priority="1">
        <dataBar>
          <cfvo type="min" />
          <cfvo type="max" />
          <color rgb="FF638EC6" />
        </dataBar>
        <extLst>
          <ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <x14:id>{A620EE03-2FEC-4D54-872C-66BDB99CB07E}</x14:id>
          </ext>
        </extLst>
      </cfRule>   *)
begin
  AppendToStream(AStream,
    '<cfRule type="dataBar" priority="' + IntToStr(APriority) + '">' +
       '<dataBar>');

  AppendToStream(AStream,
         CF_ValueNode(ARule.StartValueKind, ARule.StartValue),
         CF_ValueNode(ARule.EndValueKind, ARule.EndValue),
         CF_ColorNode(ARule.Color) );

  AppendToStream(AStream,
      '</dataBar>' +
    '</cfRule>');
end;


procedure TsSpreadOOXMLWriter.WriteCFIconSetRule(AStream: TStream;
  ARule: TsCFIconSetRule; APriority: Integer);
{ <cfRule type="iconSet" priority="13">
    <iconSet iconSet="3Symbols2">
      <cfvo type="percent" val="0" />
      <cfvo type="percent" val="33" />
      <cfvo type="percent" val="67" />
    </iconSet>
  </cfRule> }
var
  i: Integer;
begin
  AppendToStream(AStream, Format(
    '<cfRule type="iconSet" priority="%d">' +
      '<iconSet iconSet="%s">' +
        '<cfvo type="percent" val="0" />', [
      APriority, CF_ICON_SET[Arule.IconSet]
    ]));

  for i := 0 to ARule.IconCount-2 do
    AppendToStream(AStream,
        CF_ValueNode(ARule.ValueKinds[i], ARule.Values[i])
    );
  AppendToStream(AStream,
      '</iconSet>' +
    '</cfRule>');
end;

procedure TsSpreadOOXMLWriter.WriteColBreaks(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  sheet: TsWorksheet absolute AWorksheet;
  n: Integer;
  i: Integer;
begin
  n := 0;
  for i := 0 to sheet.Cols.Count - 1 do
    if (croPageBreak in PCol(sheet.Cols[i])^.Options) then inc(n);
  if n = 0 then
    exit;

  AppendToStream(AStream, Format(
    '<colBreaks count="%d" manualBreakCount="%d">', [n, n]));
    for i := 0 to sheet.Cols.Count - 1 do
      if (croPageBreak in PCol(sheet.Cols[i])^.Options) then
        AppendToStream(AStream, Format(
          '<brk id="%d" max="1048575" man="1" />', [PCol(sheet.Cols[i])^.Col]));
  AppendToStream(AStream,
    '</colBreaks>');
end;

procedure TsSpreadOOXMLWriter.WriteCols(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  lCol: PCol;
  c: Integer;
  w: Single;
  customWidth: String;
  customStyle: String;
  sheet: TsWorksheet absolute AWorksheet;
  hiddenStr: String;
begin
  AppendToStream(AStream,
    '<cols>');

  for c:=0 to sheet.GetLastColIndex do begin
    customWidth := '';
    customStyle := '';
    lCol := sheet.FindCol(c);

    // The column width is needed in suChars here.
    w := sheet.ReadDefaultColWidth(suChars);
    hiddenStr := '';
    if lCol <> nil then begin
      if lCol^.ColWidthType = cwtCustom then begin
        w := (FWorkbook as TsWorkbook).ConvertUnits(lCol^.Width, FWorkbook.Units, suChars);
        customWidth := 'customWidth="1" ';
      end;
      if lCol^.FormatIndex > 0 then
        customStyle := Format('style="%d" ', [lCol^.FormatIndex]);
      if (croHidden in lCol^.Options) then hiddenStr := ' hidden="1"';
    end;
    AppendToStream(AStream, Format(
      '<col min="%d" max="%d" width="%.2f" %s%s%s />',
      [c+1, c+1, w, customWidth, customStyle, hiddenStr], FPointSeparatorSettings)
    );
  end;

  AppendToStream(AStream,
    '</cols>');
end;

procedure TsSpreadOOXMLWriter.WriteComments(AWorksheet: TsBasicWorksheet);
var
  comment: PsComment;
  txt: String;
begin
  if (AWorksheet as TsWorksheet).Comments.Count = 0 then
    exit;

  // Create the comments stream
  SetLength(FSComments, FCurSheetNum + 1);
  FSComments[FCurSheetNum] := CreateTempStream(FWorkbook, Format('fpsCMNT%d', [FCurSheetNum]));

  // Header
  AppendToStream(FSComments[FCurSheetNum],
    XML_HEADER);
  AppendToStream(FSComments[FCurSheetNum], Format(
    '<comments xmlns="%s">', [SCHEMAS_SPREADML]));
  AppendToStream(FSComments[FCurSheetNum],
      '<authors>'+
        '<author />'+   // Not necessary to specify an author here. But the node must exist!
      '</authors>');
  AppendToStream(FSComments[FCurSheetNum],
      '<commentList>');

  // Comments
  for comment in (AWorksheet as TsWorksheet).Comments do
  begin
    txt := comment^.Text;
    ValidXMLText(txt);

    // Write comment text to Comments stream
    AppendToStream(FSComments[FCurSheetNum], Format(
        '<comment ref="%s" authorId="0">', [GetCellString(comment^.Row, comment^.Col)]) +
          '<text>'+
            '<r>'+
              '<rPr>'+  // thie entire node could be omitted, but then Excel uses some ugly default font
                '<sz val="9"/>'+
                '<color rgb="000000" />'+  // Excel files have color index 81 here, but it could be that this does not exist in fps files --> use rgb instead
                '<rFont val="Arial" />'+   // It is not harmful to Excel if the font does not exist.
                '<charset val="1" />'+
              '</rPr>'+
              '<t xml:space="preserve">' + txt + '</t>' +
            '</r>' +
          '</text>' +
        '</comment>');
  end;

  // Footer
  AppendToStream(FSComments[FCurSheetNum],
      '</commentList>');
  AppendToStream(FSComments[FCurSheetNum],
    '</comments>');
end;

procedure TsSpreadOOXMLWriter.WriteConditionalFormat(AStream: TStream;
  AFormat: TsConditionalFormat; var APriority: Integer);
var
  rangeStr: String;
  i: Integer;
  rule: TsCFRule;
begin
  with AFormat.CellRange do
    rangeStr := GetCellRangeString(Row1, Col1, Row2, Col2,rfAllRel, true);
  AppendToStream(AStream, Format(
    '<conditionalFormatting sqref="%s">', [rangeStr]));
  for i := 0 to AFormat.RulesCount-1 do
  begin
    rule := AFormat.Rules[i];
    if rule is TsCFCellRule then
      WriteCFCellRule(AStream, TsCFCellRule(rule), AFormat.CellRange, APriority)
    else
    if rule is TsCFColorRangeRule then
      WriteCFColorRangeRule(AStream, TsCFColorRangeRule(rule), APriority)
    else
    if rule is TsCFDataBarRule then
      WriteCFDataBarRule(AStream, TsCFDataBarRule(rule), APriority)
    else
    if rule is TsCFIconSetRule then
      WriteCFIconSetRule(AStream, TsCFIconSetRule(rule), APriority)
    else
      exit;
    dec(APriority);
  end;
  AppendToStream(AStream,
    '</conditionalFormatting>');
end;

procedure TsSpreadOOXMLWriter.WriteConditionalFormats(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  book: TsWorkbook;
  worksheet: TsWorksheet absolute AWorksheet;
  i: Integer;
  CF: TsConditionalFormat;
  priority: Integer = 0;
  ncf: Integer;
begin
  book := TsWorkbook(FWorkbook);
  ncf := book.GetNumConditionalFormats;
  if ncf = 0 then
    exit;

  for i := 0 to ncf-1 do
  begin
    CF := book.GetConditionalFormat(i);
    if CF.Worksheet = AWorksheet then
      inc(priority, CF.RulesCount);
  end;

  for i := 0 to ncf-1 do begin
    CF := book.GetConditionalFormat(i);
    if CF.Worksheet = AWorksheet then
      WriteConditionalFormat(AStream, CF, priority);
  end;
end;

procedure TsSpreadOOXMLWriter.WriteCustomMetaData(AStream: TStream);
var
  book: TsWorkbook;
  i: Integer;
  id: Integer;
begin
  book := TsWorkbook(FWorkbook);
  if book.MetaData.Custom.Count = 0 then
    exit;

  AppendToStream(AStream,
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" ' +
      'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">');

  id := 2;
  for i := 0 to book.MetaData.Custom.Count-1 do
  begin
    AppendToStream(AStream, Format(
      '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="%d" name="%s">' +
        '<vt:lpwstr>%s</vt:lpwstr>' +
      '</property>', [
      id, book.MetaData.Custom.Names[i],
      book.MetaData.Custom.ValueFromIndex[i]
    ]));
    inc(id);
  end;

  AppendToStream(AStream,
    '</Properties>');
end;

procedure TsSpreadOOXMLWriter.WriteDimension(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  r1,c1,r2,c2: Cardinal;
  dim: String;
begin
  GetSheetDimensions(AWorksheet, r1, r2, c1, c2);
  if (r1=r2) and (c1=c2) then
    dim := GetCellString(r1, c1)
  else
    dim := GetCellRangeString(r1, c1, r2, c2);
  AppendToStream(AStream, Format(
    '<dimension ref="%s" />', [dim]));
end;

procedure TsSpreadOOXMLWriter.WriteFillList(AStream: TStream);
var
  i: Integer;
  pt, bc, fc: string;
begin
  AppendToStream(AStream, Format(
    '<fills count="%d">', [Length(FFillList)]));

  // index 0 -- built-in empty fill
  AppendToStream(AStream,
      '<fill>',
        '<patternFill patternType="none" />',
      '</fill>');

  // index 1 -- built-in gray125 pattern
  AppendToStream(AStream,
      '<fill>',
        '<patternFill patternType="gray125" />',
      '</fill>');

  // user-defined fills
  for i:=2 to High(FFillList) do begin
    pt := PATTERN_TYPES[FFillList[i]^.Background.Style];
    if FFillList[i]^.Background.FgColor = scTransparent then
      fc := 'auto="1"'
    else
      fc := Format('rgb="%s"', [ColorToExcelStr(FFillList[i]^.Background.FgColor)]);
    if FFillList[i]^.Background.BgColor = scTransparent then
      bc := 'auto="1"'
    else
      bc := Format('rgb="%s"', [ColorToExcelStr(FFillList[i]^.Background.BgColor)]);
    AppendToStream(AStream,
      '<fill>');
    AppendToStream(AStream, Format(
        '<patternFill patternType="%s">', [pt]) + Format(
          '<fgColor %s />', [fc]) + Format(
          '<bgColor %s />', [bc]) +
//          '<bgColor indexed="64" />' +
        '</patternFill>' +
      '</fill>');
  end;

  AppendToStream(FSStyles,
    '</fills>');
end;

{ Writes font parameters to the stream.
  ATag is "font" for the entry in "styles.xml", or "rPr" for the entry for
  richtext parameters in the shared string list.
  ANameTag is "name" for the entry in "styles.xml", or "rFont" for the entry}
procedure TsSpreadOOXMLWriter.WriteFont(AStream: TStream; AFont: TsFont;
  UseInStyleNode: Boolean);
const
  TAG: Array[boolean] of string = ('rPr', 'font');
  NAME_TAG: Array[boolean] of String = ('rFont', 'name');
var
  s: String;
begin
  s := '';
  if AFont.Size > 0 then
    s := s + Format('<sz val="%g" />', [AFont.Size], FPointSeparatorSettings);
  if AFont.FontName <> '' then
    s := s + Format('<%s val="%s" />', [NAME_TAG[UseInStyleNode], AFont.FontName]);
  if (fssBold in AFont.Style) then
    s := s + '<b />';
  if (fssItalic in AFont.Style) then
    s := s + '<i />';
  if (fssUnderline in AFont.Style) then
    s := s + '<u />';
  if (fssStrikeout in AFont.Style) then
    s := s + '<strike />';
  if AFont.Color <> scBlack then
    s := s + Format('<color rgb="%s" />', [ColorToExcelStr(AFont.Color)]);
  case AFont.Position of
    fpSubscript  : s := s + '<vertAlign val="subscript" />';
    fpSuperscript: s := s + '<vertAlign val="superscript" />';
  end;
  AppendToStream(AStream, Format(
    '<%s>%s</%s>', [TAG[UseInStyleNode], s, TAG[UseInStyleNode]]));
end;

{ Writes the fontlist of the workbook to the stream. }
procedure TsSpreadOOXMLWriter.WriteFontList(AStream: TStream);
var
  i: Integer;
  font: TsFont;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;
  AppendToStream(AStream, Format(
    '<fonts count="%d">', [book.GetFontCount]));
  for i:=0 to book.GetFontCount-1 do begin
    font := book.GetFont(i);
    WriteFont(AStream, font, true);
  end;
  AppendToStream(AStream,
    '</fonts>');
end;

procedure TsSpreadOOXMLWriter.WriteHeaderFooter(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  s: String;
begin
  with (AWorksheet as TsWorksheet).PageLayout do
  begin
    if not (HasHeader or HasFooter) then
      exit;

    s := '';
    if poDifferentFirst in Options then
      s := s + ' differentFirst="1"';
    if poDifferentOddEven in Options then
      s := s + ' differentOddEven="1"';

    AppendToStream(AStream,
        '<headerFooter' + s + '>');

    if Headers[HEADER_FOOTER_INDEX_ODD] <> '' then
      AppendToStream(AStream,
          '<oddHeader>' + UTF8TextToXMLText(Headers[HEADER_FOOTER_INDEX_ODD]) + '</oddHeader>');
    if Footers[HEADER_FOOTER_INDEX_ODD] <> '' then
      AppendToStream(AStream,
          '<oddFooter>' + UTF8TextToXMLText(Footers[HEADER_FOOTER_INDEX_ODD]) + '</oddFooter>');

    if poDifferentFirst in (AWorksheet as TsWorksheet).PageLayout.Options then
    begin
      if Headers[HEADER_FOOTER_INDEX_FIRST] <> '' then
        AppendToStream(AStream,
          '<firstHeader>' + UTF8TextToXMLText(Headers[HEADER_FOOTER_INDEX_FIRST]) + '</firstHeader>');
      if Footers[HEADER_FOOTER_INDEX_FIRST] <> '' then
        AppendToStream(AStream,
          '<firstFooter>' + UTF8TextToXMLText(Footers[HEADER_FOOTER_INDEX_FIRST]) + '</firstFooter>');
    end;

    if poDifferentOddEven in Options then
    begin
      AppendToStream(AStream,
          '<evenHeader>' + UTF8TextToXMLText(Headers[HEADER_FOOTER_INDEX_EVEN]) + '</evenHeader>');
      AppendToStream(AStream,
          '<evenFooter>' + UTF8TextToXMLText(Footers[HEADER_FOOTER_INDEX_EVEN]) + '</evenFooter>');
    end;

    AppendToStream(AStream,
        '</headerFooter>');
  end;
end;

procedure TsSpreadOOXMLWriter.WriteHyperlinks(AStream: TStream;
  AWorksheet: TsBasicWorksheet; rId: Integer);
var
  hyperlink: PsHyperlink;
  target, bookmark: String;
  s: String;
  txt: String;
  AVLNode: TAvgLvlTreeNode;
  sheet: TsWorksheet absolute AWorksheet;
begin
  if sheet.Hyperlinks.Count = 0 then
    exit;

  AppendToStream(AStream,
    '<hyperlinks>');

  AVLNode := sheet.Hyperlinks.FindLowest;
  while AVLNode <> nil do begin
    hyperlink := PsHyperlink(AVLNode.Data);
    SplitHyperlink(hyperlink^.Target, target, bookmark);
    s := Format('ref="%s"', [GetCellString(hyperlink^.Row, hyperlink^.Col)]);
    if target <> '' then
    begin
      s := Format('%s r:id="rId%d"', [s, rId]);
      inc(rId);
    end;
    if bookmark <> '' then //target = '' then
      s := Format('%s location="%s"', [s, bookmark]);
    txt := UTF8TextToXMLText(sheet.ReadAsText(hyperlink^.Row, hyperlink^.Col));
    if (txt <> '') and (txt <> hyperlink^.Target) then
      s := Format('%s display="%s"', [s, txt]);
    if hyperlink^.ToolTip <> '' then begin
      txt := hyperlink^.Tooltip;
      ValidXMLText(txt);
      s := Format('%s tooltip="%s"', [s, txt]);
    end;
    AppendToStream(AStream,
        '<hyperlink ' + s + ' />');
    AVLNode := sheet.Hyperlinks.FindSuccessor(AVLNode);
  end;

  AppendToStream(AStream,
    '</hyperlinks>');
end;

procedure TsSpreadOOXMLWriter.WriteMergedCells(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  rng: PsCellRange;
  n: Integer;
begin
  n := (AWorksheet as TsWorksheet).MergedCells.Count;
  if n = 0 then
    exit;
  AppendToStream(AStream, Format(
    '<mergeCells count="%d">', [n]) );
  for rng in (AWorksheet as TsWorksheet).MergedCells do
    AppendToStream(AStream, Format(
      '<mergeCell ref="%s" />', [GetCellRangeString(rng^.Row1, rng^.Col1, rng^.Row2, rng^.Col2)]));
  AppendToStream(AStream,
    '</mergeCells>');
end;

{ Writes all number formats to the stream. Saving starts at the item with the
  FirstFormatIndexInFile. }
procedure TsSpreadOOXMLWriter.WriteNumFormatList(AStream: TStream);
var
  i, n: Integer;
  numFmtStr: String;
  xmlStr: String;
  parser: TsNumFormatParser;
begin
  xmlStr := '';
  n := 0;
  for i:= FFirstNumFormatIndexInFile to NumFormatList.Count-1 do
  begin
    numFmtStr := NumFormatList[i];
    parser := TsExcelNumFormatParser.Create(numFmtStr, Workbook.FormatSettings);
    try
      numFmtStr := UTF8TextToXMLText(parser.FormatString);
      xmlStr := xmlStr + Format('<numFmt numFmtId="%d" formatCode="%s" />',
        [i, numFmtStr]);
      inc(n);
    finally
      parser.Free;
    end;
  end;

  if n > 0 then
    AppendToStream(AStream, Format(
      '<numFmts count="%d">', [n]),
        xmlStr,
      '</numFmts>'
    );
end;

{ In older versions, the workbook had a color palette which was written here.
  Now there is no palette any more. }
procedure TsSpreadOOXMLWriter.WritePalette(AStream: TStream);
begin
  // just keep it here in case we'd need it later...
  Unused(AStream);
end;

procedure TsSpreadOOXMLWriter.WritePageMargins(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
begin
  with (AWorksheet as TsWorksheet).PageLayout do
    AppendToStream(AStream, Format(
      '<pageMargins left="%g" right="%g" top="%g" bottom="%g" header="%g" footer="%g" />', [
      mmToIn(LeftMargin), mmToIn(RightMargin), mmToIn(TopMargin), mmToIn(BottomMargin),
      mmToIn(HeaderMargin), mmToIn(FooterMargin) ],
      FPointSeparatorSettings
    ));
end;

procedure TsSpreadOOXMLWriter.WritePageSetup(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  s: String;
  i: Integer;
  sheet: TsWorksheet absolute AWorksheet;
begin
  s := '';

  // Paper size
  for i:=0 to High(PAPER_SIZES) do
    if (SameValue(PAPER_SIZES[i,0], sheet.PageLayout.PageHeight) and
        SameValue(PAPER_SIZES[i,1], sheet.PageLayout.PageWidth))
    or (SameValue(PAPER_SIZES[i,1], sheet.PageLayout.PageHeight) and
        SameValue(PAPER_SIZES[i,0], sheet.PageLayout.PageWidth))
    then begin
      s := Format('%s paperSize="%d"', [s, i]);
      break;
    end;

  if poFitPages in sheet.PageLayout.Options then
  begin
    // Fit width to pages
    s := Format('%s fitToWidth="%d"', [s, sheet.PageLayout.FitWidthToPages]);
    // Fit height to pages
    s := Format('%s fitToHeight="%d"', [s, sheet.PageLayout.FitHeightToPages]);
  end else
    // Scaling factor
    s := Format('%s scale="%d"', [s, sheet.PageLayout.ScalingFactor]);

  // Orientation
  s := Format('%s orientation="%s"', [
    s, IfThen(sheet.PageLayout.Orientation = spoPortrait, 'portrait', 'landscape')
  ]);

  // First page number
  if poUseStartPageNumber in sheet.PageLayout.Options then
    s := Format('%s useFirstPageNumber="1"', [s]);

  s := Format('%s firstPageNumber="%d"', [s, sheet.PageLayout.StartPageNumber]);

  // Print order
  if poPrintPagesByRows in sheet.PageLayout.Options then
    s := s + ' pageOrder="overThenDown"';

  // Monochrome
  if poMonochrome in sheet.PageLayout.Options then
    s := s + ' blackAndWhite="1"';

  // Quality
  if poDraftQuality in sheet.PageLayout.Options then
    s := s + ' draft="1"';

  if s <> '' then
    AppendToStream(AStream,
      '<pageSetup' + s + ' />');
end;

procedure TsSpreadOOXMLWriter.WritePrintOptions(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  s: String;
begin
  s := '';
  if poPrintGridLines in (AWorksheet as TsWorksheet).PageLayout.Options then
    s := s + ' gridLines="1"';
  if poPrintHeaders in (AWorksheet as TsWorksheet).PageLayout.Options then
    s := s + ' headings="1"';

  if s <> '' then
    AppendToStream(AStream,
      '<printOptions' + s + ' />');
end;

procedure TsSpreadOOXMLWriter.WriteRowBreaks(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  i, n: Integer;
  sheet: TsWorksheet absolute AWorksheet;
begin
  n := 0;
  for i := 0 to sheet.Rows.Count-1 do
    if (croPageBreak in PRow(sheet.Rows[i])^.Options) then inc(n);
  if n = 0 then
    exit;

  AppendToStream(AStream, Format(
    '<rowBreaks count="%d" manualBreakCount="%d">', [n, n]));
    for i := 0 to sheet.Rows.Count - 1 do
      if (croPageBreak in PRow(sheet.Rows[i])^.Options) then
        AppendToStream(AStream, Format(
          '<brk id="%d" max="16383" man="1" />', [PRow(sheet.Rows[i])^.Row]));
  AppendToStream(AStream,
    '</rowBreaks>');
end;

procedure TsSpreadOOXMLWriter.WriteSheetData(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  r, r1, r2: Cardinal;
  c, c1, c2: Cardinal;
  row: PRow;
  value: Variant;
  lCell: TCell;
  styleCell: PCell;
  cell: PCell;
  s: String;
  sheet: TsWorksheet;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;
  sheet := AWorksheet as TsWorksheet;
  {$IF FPC_FullVersion >= 30000}
  lCell := Default(TCell);
  {$IFEND}

  AppendToStream(AStream,
      '<sheetData>');

  GetSheetDimensions(AWorksheet, r1, r2, c1, c2);

  if (boVirtualMode in Workbook.Options) then begin
    if Assigned(sheet.OnWriteCellData) and
      (sheet.VirtualColCount > 0) and (sheet.VirtualRowCount > 0)
    then begin
      for r := 0 to r2 do begin
        row := sheet.FindRow(r);
        s := '';
        if row <> nil then begin
          s := s + Format(' ht="%.2f"',
            [book.ConvertUnits(row^.Height, book.Units, suPoints)],
            FPointSeparatorSettings
          );
          if row^.RowHeightType = rhtCustom then
            s := s + ' customHeight="1"';
          if row^.FormatIndex > 0 then
            s := s + Format(' s="%d" customFormat="1"', [row^.FormatIndex]);
          if (croHidden in row^.Options) then
            s := s + ' hidden="1"';
        end;
        AppendToStream(AStream, Format(
          '<row r="%d" spans="1:%d"%s>', [r+1, sheet.VirtualColCount, s]));
        for c := 0 to c2 do begin
          InitCell(sheet, r, c, lCell);
          value := varNull;
          styleCell := nil;
          sheet.OnWriteCellData(sheet, r, c, value, styleCell);
          if styleCell <> nil then
            lCell.FormatIndex := styleCell^.FormatIndex;
          if VarIsNull(value) then
          begin
            if styleCell <> nil then
              lCell.ContentType := cctEmpty
            else
              Continue;
          end else
          if VarIsNumeric(value) then
          begin
            lCell.ContentType := cctNumber;
            lCell.NumberValue := value;
          end else
          if VarType(value) = varDate then
          begin
            lCell.ContentType := cctDateTime;
            lCell.DateTimeValue := VarToDateTime(value);
          end else
          if VarIsStr(value) then
          begin
            lCell.ContentType := cctUTF8String;
            lCell.UTF8StringValue := VarToStrDef(value, '');
          end else
          if VarIsBool(value) then
          begin
            lCell.ContentType := cctBool;
            lCell.BoolValue := value <> 0;
          end;
          WriteCellToStream(AStream, @lCell);
          varClear(value);
        end;
        AppendToStream(AStream,
          '</row>');
      end;
    end;
  end    // end of virtual mode writing
  else
  begin
    // The cells need to be written in order, row by row, cell by cell
    for r := r1 to r2 do begin
      // If the row has a custom or auto height and/or custom format
      // then add them to the <row> specification
      row := sheet.FindRow(r);
      s := '';
      if row <> nil then begin
        s := s + Format(' ht="%.2f"',
          [book.ConvertUnits(row^.Height, book.Units, suPoints)],
          FPointSeparatorSettings
        );
        if row^.RowHeightType = rhtCustom then
          s := s + ' customHeight="1"';
        if row^.FormatIndex > 0 then
          s := s + Format(' s="%d" customFormat="1"', [row^.FormatIndex]);
        if (croHidden in row^.Options) then
          s := s + ' hidden="1"';
      end;
      AppendToStream(AStream, Format(
        '<row r="%d" spans="%d:%d"%s>', [r+1, c1+1, c2+1, s]));

      // Write cells belonging to this row.
      {
      // Strange: the RowEnumerator is very slow here... ?!
      for cell in AWorksheet.Cells.GetRowEnumerator(r) do
        WriteCellToStream(AStream, cell);
      }
      for c := c1 to c2 do begin
        cell := sheet.FindCell(r, c);
        if Assigned(cell) then
          WriteCellToStream(AStream, cell);
      end;

      AppendToStream(AStream,
        '</row>');
    end;
  end;
  AppendToStream(AStream,
      '</sheetData>');
end;

procedure TsSpreadOOXMLWriter.WriteSheetFormatPr(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  w, h: Single;
begin
  // Excel has column width in characters, and row heights in pts.
  w := (AWorksheet as TsWorksheet).ReadDefaultColWidth(suChars);
  h := (AWorksheet as TsWorksheet).ReadDefaultRowHeight(suPoints);
  AppendToStream(AStream, Format(
    '<sheetFormatPr baseColWidth="10" defaultColWidth="%.2f" defaultRowHeight="%.2f" customHeight="true" />',
    [w, h],
    FPointSeparatorSettings));
end;

procedure TsSpreadOOXMLWriter.WriteSheetPr(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  s: String;
  sheet: TsWorksheet absolute AWorksheet;
begin
  s := '';

  if (sheet.PageLayout.FitWidthToPages > 0) or
     (sheet.PageLayout.FitHeightToPages > 0) then
  s := s + ' fitToPage="1"';
  if s <> '' then s := '<pageSetUpPr' + s + ' />';

  if sheet.TabColor <> scNotDefined then
    s := s + Format('<tabColor rgb="%s" />', [ColorToExcelStr(sheet.TabColor)]);

  if s <> '' then
    AppendToStream(AStream,
      '<sheetPr>' + s + '</sheetPr>');
end;

procedure TsSpreadOOXMLWriter.WriteSheetProtection(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  s: String;
  sheet: TsWorksheet absolute AWorksheet;
begin
  s := '';

  // No attribute -> attr="0"
  if sheet.IsProtected then
    s := ' sheet="1" scenarios="1"'
  else
    Exit; //exit if sheet not protected

  if sheet.CryptoInfo.PasswordHash <> '' then begin
    if sheet.CryptoInfo.Algorithm = caExcel then
      s := s + ' password="' + sheet.CryptoInfo.PasswordHash + '"'
    else
    begin
      s := s + ' hashValue="' + sheet.CryptoInfo.PasswordHash + '"';

      if sheet.CryptoInfo.Algorithm <> caUnknown then
        s := s + ' algorithmName="' + AlgorithmToStr(sheet.CryptoInfo.Algorithm, auExcel) + '"';

      if sheet.CryptoInfo.SaltValue <> '' then
        s := s + ' saltValue="' + sheet.CryptoInfo.SaltValue + '"';

      if sheet.CryptoInfo.SpinCount <> 0 then
        s := s + ' spinCount="' + IntToStr(sheet.CryptoInfo.SpinCount) + '"';
    end;
  end;

  if spObjects in sheet.Protection then
    s := s + ' objects="1"';

  {
  if spScenarios in sheet.Protection then     // to do: Remove from default above
    s := s + ' scenarios="1"';
  }

  if spSelectLockedCells in sheet.Protection then
    s := s + ' selectLockedCells="1"';

  if spSelectUnlockedCells in sheet.Protection then
    s := s + ' selectUnlockedCells="1"';

  // No attribute -> attr="1"
  {
  if not (spAutoFilter in sheet.Protection) then
    s := s + ' autoFilter="0"';
  }
  if not (spDeleteColumns in sheet.Protection) then
    s := s + ' deleteColumns="0"';

  if not (spDeleteRows in sheet.Protection) then
    s := s + ' deleteRows="0"';

  if not (spFormatCells in sheet.Protection) then
    s := s + ' formatCells="0"';

  if not (spFormatColumns in sheet.Protection) then
    s := s + ' formatColumns="0"';

  if not (spFormatRows in sheet.Protection) then
    s := s + ' formatRows="0"';

  if not (spInsertColumns in sheet.Protection) then
    s := s + ' insertColumns="0"';

  if not (spInsertHyperlinks in sheet.Protection) then
    s := s + ' insertHyperlinks="0"';

  if not (spInsertRows in sheet.Protection) then
    s := s + ' insertRows="0"';

  {
  if not (spPivotTables in sheet.Protection) then
    s := s + ' pivotTables="0"';
  }
  if not (spSort in sheet.Protection) then
    s := s + ' sort="0"';

  if s <> '' then
    AppendToStream(AStream,
      '<sheetProtection' + s + ' />');
end;

procedure TsSpreadOOXMLWriter.WriteSheets(AStream: TStream);
var
  counter: Integer;
  sheet: TsWorksheet;
  sheetName: String;
  sheetState: String;
begin
  AppendToStream(AStream,
    '<sheets>');
  for counter := 1 to (Workbook as TsWorkbook).GetWorksheetCount do
  begin
    sheet := (Workbook as TsWorkbook).GetWorksheetByIndex(counter-1);
    sheetname := UTF8TextToXMLText(sheet.Name);
    sheetState := IfThen(soHidden in sheet.Options, ' state="hidden"', '');
    AppendToStream(AStream, Format(
      '<sheet name="%s" sheetId="%d" r:id="rId%d"%s />',
      [sheetname, counter, counter, sheetstate]));
  end;
  AppendToStream(AStream,
    '</sheets>');
end;

procedure TsSpreadOOXMLWriter.WriteSheetViews(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
const
  ZOOM_EPS = 1E-3;
var
  showGridLines: String;
  showHeaders: String;
  topRightCell: String;
  bottomLeftCell: String;
  bottomRightCell: String;
  actCell: String;
  tabSel: String;
  bidi: String;
  zoomscale: String;
  attr: String;
  windowProtection: String;
  book: TsWorkbook;
  sheet: TsWorksheet absolute AWorksheet;
begin
  book := FWorkbook as TsWorkbook;

  // Show gridlines ?
  showGridLines := StrUtils.IfThen(soShowGridLines in sheet.Options, '', ' showGridLines="0"');

  // Show headers?
  showHeaders := StrUtils.IfThen(soShowHeaders in sheet.Options, '', ' showRowColHeaders="0"');

  // Zoom factor
  if boWriteZoomFactor in FWorkbook.Options then
    zoomscale := StrUtils.IfThen(SameValue(sheet.ZoomFactor, 1.0, ZOOM_EPS), '',
      Format(' zoomScale="%.0f"', [sheet.ZoomFactor*100]))
  else
    zoomscale := '';

  // BiDiMode
  case sheet.BiDiMode of
    bdDefault: bidi := '';
    bdLTR    : bidi := ' rightToLeft="0"';
    bdRTL    : bidi := ' rightToLeft="1"';
  end;

  // Active cell
  if (sheet.ActiveCellRow <> cardinal(-1)) and (sheet.ActiveCellCol <> cardinal(-1)) then
    actCell := GetCellString(sheet.ActiveCellRow, sheet.ActiveCellCol) else
    actCell := '';

  // Selected tab?
  tabSel := StrUtils.IfThen(sheet = book.ActiveWorksheet, ' tabSelected="1"', '');

  // Window protection
  if (soPanesProtection in sheet.Options) and book.IsProtected then
    windowProtection := ' windowProtection="1"'
  else
    windowProtection := '';

  // All SheetView attributes
  attr := windowProtection + showGridLines + showHeaders + tabSel + zoomScale + bidi;

  // No frozen panes
  if not (soHasFrozenPanes in sheet.Options) or
     ((sheet.LeftPaneWidth = 0) and (sheet.TopPaneHeight = 0))
  then
  begin
    if actCell = '' then actCell := 'A1';
    AppendToStream(AStream, Format(
      '<sheetViews>' +
        '<sheetView workbookViewId="0"%s>' +
          '<selection activeCell="%s" sqref="%s" />' +
        '</sheetView>' +
      '</sheetViews>', [
      attr,
      actCell, actCell
    ]))
  end else
  begin  // Frozen panes
    topRightCell := GetCellString(0, sheet.LeftPaneWidth, [rfRelRow, rfRelCol]);
    bottomLeftCell := GetCellString(sheet.TopPaneHeight, 0, [rfRelRow, rfRelCol]);
    bottomRightCell := GetCellString(sheet.TopPaneHeight, sheet.LeftPaneWidth, [rfRelRow, rfRelCol]);
    if (sheet.LeftPaneWidth > 0) and (sheet.TopPaneHeight > 0) then
    begin
      if actCell = '' then
        actCell := bottomRightcell;
      AppendToStream(AStream, Format(
        '<sheetViews>' +
          '<sheetView workbookViewId="0"%s>'+
            '<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen" />' +
            '<selection pane="topRight" activeCell="%s" sqref="%s" />' +
            '<selection pane="bottomLeft" activeCell="%s" sqref="%s" />' +
            '<selection pane="bottomRight" activeCell="%s" sqref="%s" />' +
          '</sheetView>' +
        '</sheetViews>', [
        attr,
        sheet.LeftPaneWidth, sheet.TopPaneHeight, bottomRightCell,
        topRightCell, topRightCell,
        bottomLeftCell, bottomLeftCell,
        actCell, actCell
      ]))
    end else
    if (sheet.LeftPaneWidth > 0) then
    begin
      if actCell = '' then
        actCell := topRightCell;
      AppendToStream(AStream, Format(
        '<sheetViews>' +
          '<sheetView workbookViewId="0"%s>'+
            '<pane xSplit="%d" topLeftCell="%s" activePane="topRight" state="frozen" />' +
            '<selection pane="topRight" activeCell="%s" sqref="%s" />' +
          '</sheetView>' +
        '</sheetViews>', [
        attr,
        sheet.LeftPaneWidth, topRightCell,
        actCell, actCell
      ]))
    end else
    if (sheet.TopPaneHeight > 0) then
    begin
      if actCell = '' then
        actCell := bottomLeftCell;
      AppendToStream(AStream, Format(
        '<sheetViews>'+
          '<sheetView workbookViewId="0"%s>'+
             '<pane ySplit="%d" topLeftCell="%s" activePane="bottomLeft" state="frozen" />'+
             '<selection pane="bottomLeft" activeCell="%s" sqref="%s" />' +
          '</sheetView>'+
        '</sheetViews>', [
        attr,
        sheet.TopPaneHeight, bottomLeftCell,
        actCell, actCell
      ]));
    end;
  end;
end;

procedure TsSpreadOOXMLWriter.WriteStyle(AStream: TStream; ANodeName: String;
  AFormat: PsCellFormat);
var
  s: String;
  sAlign: String;
  sProtected: String;
  book: TsWorkbook;
  numFmtParams: TsNumFormatParams;
  numFmtStr: String;
  fontID: Integer;
  fillID: Integer;
  borderID: Integer;
  idx: Integer;
begin
  book := TsWorkbook(FWorkbook);

  s := '';
  sAlign := '';
  sProtected := '';

  { Number format }
  if (uffNumberFormat in AFormat^.UsedFormattingFields) then
  begin
    numFmtParams := book.GetNumberFormat(AFormat^.NumberFormatIndex);
    if numFmtParams <> nil then
    begin
      numFmtStr := numFmtParams.NumFormatStr;
      idx := NumFormatList.IndexOf(numFmtStr);
    end else
      idx := 0;  // "General" format is at index 0
    s := s + Format('numFmtId="%d" applyNumberFormat="1" ', [idx]);
  end else
    s := s + 'numFmtId="0" ';

  { Font }
  fontId := 0;
  if (uffFont in AFormat^.UsedFormattingFields) then
    fontID := AFormat^.FontIndex;
  s := s + Format('fontId="%d" ', [fontId]);
  if fontID > 0 then s := s + 'applyFont="1" ';

  if ANodeName = 'xf' then s := s + 'xfId="0" ';
//  if ANodeName = 'cellXfs' then s := s + 'xfId="0" ';

  { Text rotation }
  if (uffTextRotation in AFormat^.UsedFormattingFields) then
    case AFormat^.TextRotation of
      trHorizontal:
        ;
      rt90DegreeClockwiseRotation:
        sAlign := sAlign + Format('textRotation="%d" ', [180]);
      rt90DegreeCounterClockwiseRotation:
        sAlign := sAlign + Format('textRotation="%d" ',  [90]);
      rtStacked:
        sAlign := sAlign + Format('textRotation="%d" ', [255]);
    end;

  { Text alignment }
  if (uffHorAlign in AFormat^.UsedFormattingFields) and (AFormat^.HorAlignment <> haDefault)
  then
    case AFormat^.HorAlignment of
      haLeft       : sAlign := sAlign + 'horizontal="left" ';
      haCenter     : sAlign := sAlign + 'horizontal="center" ';
      haRight      : sAlign := sAlign + 'horizontal="right" ';
      haJustified  : sAlign := sAlign + 'horizontal="justify" ';
      hadistributed: sAlign := sAlign + 'horizontal="distributed" ';
      haFilled     : sAlign := sAlign + 'horizontal="fill" ';
    end;

  if (uffVertAlign in AFormat^.UsedFormattingFields) and (AFormat^.VertAlignment <> vaDefault)
  then
    case AFormat^.VertAlignment of
      vaTop   : sAlign := sAlign + 'vertical="top" ';
      vaCenter: sAlign := sAlign + 'vertical="center" ';
      vaBottom: sAlign := sAlign + 'vertical="bottom" ';
    end;

  { Word wrap }
  if (uffWordWrap in AFormat^.UsedFormattingFields) then
    sAlign := sAlign + 'wrapText="1" ';

  { BiDi mode }
  if (uffBiDi in Aformat^.UsedFormattingFields) and (AFormat^.BiDiMode <> bdDefault) then
    sAlign := sAlign + Format('readingOrder="%d" ', [Ord(AFormat^.BiDiMode)]);

  if sAlign <> '' then
  begin
    s := s + 'applyAlignment="1" ';
    sAlign := '<alignment ' + sAlign + '/>';
  end;

  { Fill }
  if (uffBackground in AFormat^.UsedFormattingFields) then
  begin
    fillID := FindFillInList(AFormat);
    if fillID = -1 then fillID := 0;
    s := s + Format('fillId="%d" applyFill="1" ', [fillID]);
  end;

  { Border }
  if (uffBorder in AFormat^.UsedFormattingFields) then
  begin
    borderID := FindBorderInList(AFormat);
    if borderID = -1 then borderID := 0;
    s := s + Format('borderId="%d" applyBorder="1" ', [borderID]);
  end;

  { Protection }
  if not (cpLockCell in AFormat^.Protection) then
    sProtected := 'locked="0" ';

  if (cpHideFormulas in AFormat^.Protection) then
    sProtected := sProtected + 'hidden="1" ';

  if sProtected <> '' then
  begin
    s := s + 'applyProtection="1" ';
    sProtected := '<protection ' + sProtected + '/>';
  end;

  { Write everything to stream }
  if (sAlign = '') and (sProtected = '') then
    AppendToStream(AStream,
      Format('<%s %s />', [ANodeName, s]))
  else
    AppendToStream(AStream,
      Format('<%s %s>', [ANodeName, s]),
        sAlign + sProtected,
      Format('</%s>', [ANodeName]));
end;

{ Writes the style list which the workbook has collected in its FormatList }
procedure TsSpreadOOXMLWriter.WriteStyleList(AStream: TStream; ANodeName: String);
var
  fmt: PsCellFormat;
  i: Integer;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;

  AppendToStream(AStream, Format(
    '<%s count="%d">', [ANodeName, book.GetNumCellFormats]));

  for i:=0 to book.GetNumCellFormats-1 do
  begin
    fmt := book.GetPointerToCellFormat(i);
    WriteStyle(AStream, 'xf', fmt);
  end;

  AppendToStream(FSStyles, Format(
    '</%s>', [ANodeName]));
end;

procedure TsSpreadOOXMLWriter.WriteDrawings(AWorksheet: TsBasicWorksheet);

  function AnchorAsXML(Indent: Integer; AName: String; ACol, ARow: Integer;
    AColOffset, ARowOffset: Double): String;
  var
    ind: String;
  begin
    ind := DupeString(' ', Indent);
    Result := Format(
      ind + '<%s>' + LE +
      ind + '  <xdr:col>%d</xdr:col>' + LE +
      ind + '  <xdr:colOff>%d</xdr:colOff>' + LE +
      ind + '  <xdr:row>%d</xdr:row>' + LE +
      ind + '  <xdr:rowOff>%d</xdr:rowOff>' + LE +
      ind + '</%s>', [
      AName,
      ACol, mmToEMU(AColOffset),
      ARow, mmToEMU(ARowOffset),
      AName
    ]);
  end;

  procedure DoWriteImage(AStream: TStream; AIndex: Integer; var RelID: Integer);
  var
    img: TsImage;
    book: TsWorkbook;
    sheet: TsWorksheet;
    r1, c1, r2, c2: Cardinal;
    roffs1, coffs1, roffs2, coffs2: Double;
    x, y, w, h: Double;
    descr: String;
    hlink: String;
    xdr_cNvPr: String;
  begin
    book := FWorkbook as TsWorkbook;
    sheet := TsWorksheet(AWorksheet);

    img := sheet.GetImage(AIndex);
    if book.GetEmbeddedObj(img.Index).ImageType = itUnknown then
      exit;

    sheet.CalcImageExtent(AIndex, true,
      r1, c1, r2, c2,
      roffs1, coffs1, roffs2, coffs2,  // mm
      x, y, w, h);                     // mm;

    descr := ExtractFileName(book.GetEmbeddedObj(img.index).Filename);
    if descr = '' then descr := 'image';

    // This part defines the relationship to the graphic and, if available, to
    // a hyperlink.
    xdr_cNvPr := Format('id="%d" name="Graphic %d" descr="%s"', [AIndex+3, AIndex+2, descr]);
    if img.HyperlinkTarget <> '' then
    begin
      hlink := Format('<a:hlinkClick xmlns:r="%s" r:id="rId%d" ', [
        SCHEMAS_DOC_RELS, RelId
      ]);
      inc(RelId);
      if img.HyperlinkToolTip <> '' then
        hlink := hlink + Format('tooltip="%s" ', [img.HyperlinkToolTip]);
      hlink := hlink + '/>';

      xdr_cNvPr := '        <xdr:cNvPr ' + xdr_cNvPr + '>' + LE +
                   '          ' + hlink + LE +
                   '        </xdr:cNvPr>' + LE;
    end else
      xdr_cNvPr := '        <xdr:cNvPr ' + xdr_cNvPr + ' />';

    AppendToStream(AStream,
      '  <xdr:twoCellAnchor editAs="oneCell">' + LE);
    AppendToStream(AStream,
      AnchorAsXML(4, 'xdr:from', c1, r1, coffs1, roffs1) + LE
    );
    AppendToStream(AStream,
      AnchorAsXML(4, 'xdr:to', c2, r2, coffs2, roffs2) + LE
    );
    AppendToStream(AStream, Format(
        '    <xdr:pic>' + LE +
        '      <xdr:nvPicPr>' + LE +
                 xdr_cNvPr + LE +
        '        <xdr:cNvPicPr>' + LE +
        '          <a:picLocks noChangeAspect="1"/>' + LE +
        '        </xdr:cNvPicPr>' + LE +
        '      </xdr:nvPicPr>' + LE +
        '      <xdr:blipFill>' + LE +
        '        <a:blip xmlns:r="%s" r:embed="rId%d" cstate="print"/>' + LE +
        '        <a:stretch>' + LE +
        '          <a:fillRect/>' + LE +
        '        </a:stretch>' + LE +
        '      </xdr:blipFill>' + LE +
        '      <xdr:spPr>' + LE +
        '        <a:xfrm>' + LE +
        '          <a:off x="%d" y="%d"/>' + LE +
        '          <a:ext cx="%d" cy="%d"/>' + LE +  // size in EMU
        '        </a:xfrm>' + LE +
        '        <a:prstGeom prst="rect">' + LE +
        '          <a:avLst/>' + LE +
        '        </a:prstGeom>' + LE +
        '      </xdr:spPr>' + LE +
        '    </xdr:pic>' + LE +
        '    <xdr:clientData/>' + LE, [
       // i + 3, i+2, descr,
        SCHEMAS_DOC_RELS, RelID,
        mmToEMU(x), mmToEMU(y),
        mmToEMU(w), mmToEMU(h)
    ]));
    AppendToStream(AStream,
      '  </xdr:twoCellAnchor>' + LE);
  end;

  procedure DoWriteChart(AStream: TStream; AChart: TsChart;
    AChartNoInSheet: Integer; var RelID: Integer);
  var
    r1, c1, r2, c2: Cardinal;
    roffs1, coffs1, roffs2, coffs2: Double;
    x, y, w, h: Double;
    sheet: TsWorksheet;
    guid: TGUID;
  begin
    r1 := AChart.Row;
    c1 := AChart.Col;
    rOffs1 := AChart.OffsetX;
    cOffs1 := AChart.OffsetY;
    w := AChart.Width;
    h := AChart.Height;
    sheet := TsWorksheet(AChart.Worksheet);
    sheet.CalcDrawingExtent(true, w, h, r1, c1, r2, c2, rOffs1, cOffs1, rOffs2, cOffs2, x, y);

    AppendToStream(AStream,
      '  <xdr:twoCellAnchor>' + LE);

    AppendToStream(AStream,
      AnchorAsXML(4, 'xdr:from', c1, r1, coffs1, roffs1) + LE
    );
    AppendToStream(AStream,
      AnchorAsXML(4, 'xdr:to', c2, r2, coffs2, roffs2) + LE
    );

    CreateGUID(guid);
    AppendToStream(AStream, Format(
      '    <xdr:graphicFrame macro="">' + LE +
      '      <xdr:nvGraphicFramePr>' + LE +
      '        <xdr:cNvPr id="%d" name="Diagram %d">' + LE +         // line 1
      //{
      '          <a:extLst>' + LE +
      '            <a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">' + LE +
      '              <a16:creationId  xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" ' +
                       'id="' + GUIDToString(guid) + '"/>' + LE +
       //                'id="{00000000-0008-0000-0000-000002000000}"/>' + LE +
      '            </a:ext>' + LE +
      '          </a:extLst>' + LE +
      //}
      '        </xdr:cNvPr>' + LE +
      '        <xdr:cNvGraphicFramePr />' + LE +
      '      </xdr:nvGraphicFramePr>' + LE +
      '      <xdr:xfrm>' + LE +
      '        <a:off x="0" y="0" />' + LE +
      '        <a:ext cx="0" cy="0" />' + LE +
      '      </xdr:xfrm>' + LE +
      '      <a:graphic>' + LE +
      '        <a:graphicData uri="%s">' + LE +                        // line 2
      '          <c:chart xmlns:c="%s" xmlns:r="%s" r:id="rId%d" />'+ LE +  // line 3
      '        </a:graphicData>' + LE +
      '      </a:graphic>' + LE +
      '    </xdr:graphicFrame>' + LE +
      '    <xdr:clientData />' + LE,
      [
        AChartNoInSheet + 1, AChartNoInSheet,               // --> line 1
        SCHEMAS_DRAWINGML_CHART,                            // --> line 2
        SCHEMAS_DRAWINGML_CHART, SCHEMAS_DOC_RELS, RelID    // --> line 3
    ]));

    AppendToStream(AStream,
      '  </xdr:twoCellAnchor>' + LE);
  end;

var
  i, j: Integer;
  rId: Integer;
  sheet: TsWorksheet absolute AWorksheet;
  {$ifdef FPS_CHARTS}
  chart: TsChart;
  {$endif}
begin
  if (sheet.GetImageCount = 0) {$ifdef FPS_CHARTS}and (sheet.GetChartCount = 0){$endif}  then
    exit;

  SetLength(FSDrawings, FCurSheetNum + 1);
  FSDrawings[FCurSheetNum] := CreateTempStream(FWorkbook, Format('fpsD%d', [FCurSheetNum]));

  // Header
  AppendToStream(FSDrawings[FCurSheetNum],
    XML_HEADER + LE +
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '+
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + LE
  );

  // Repeat for each image
  rId := 1;
  for i:=0 to sheet.GetImageCount - 1 do
  begin
    DoWriteImage(FSDrawings[FCurSheetNum], i, rId);
    inc(rId, 1);
  end;

  {$ifdef FPS_CHARTS}
  // Repeat for each chart
  j := 1;  // Counts the charts in the current sheet
  for i := 0 to TsWorkbook(FWorkbook).GetChartCount - 1 do
  begin
    chart := TsWorkbook(FWorkbook).GetChartByIndex(i);
    if chart.Worksheet = AWorksheet then
    begin
      DoWriteChart(FSDrawings[FCurSheetNum], chart, j, rId);
      inc(j);
      inc(rId);
    end;
  end;
  {$endif}

  // Close node
  AppendToStream(FSDrawings[FCurSheetNum],
    '</xdr:wsDr>');
end;

// For each sheet, writes a "drawingN.xml.rels" file to
// folder "../drawings/_rels".
// N is a sequential number (starting 1), not neccessarily identical with the
// sheet index.
// See also: WriteVmlDrawingRels
procedure TsSpreadOOXMLWriter.WriteDrawingRels(AWorksheet: TsBasicWorksheet);
var
  i: Integer;
  ext: String;
  img: TsImage;
  chart: TsChart;
  rId: Integer;
  target, bookmark: String;
  u: TURI;
  sheet: TsWorksheet absolute AWorksheet;
begin
  if (sheet.GetImageCount = 0) {$ifdef FPS_CHARTS}and (sheet.GetChartCount = 0){$endif} then
    exit;

  SetLength(FSDrawingsRels, FCurSheetNum + 1);
  FSDrawingsRels[FCurSheetNum] := CreateTempStream(FWorkbook, Format('fpsDR%d', [FCurSheetNum]));

  // Header
  AppendToStream(FSDrawingsRels[FCurSheetNum],
    XML_HEADER + LineEnding,
    '<Relationships xmlns="' + SCHEMAS_RELS + '">' + LE);

  // Repeat for each image
  rId := 1;
  for i:=0 to sheet.GetImageCount - 1 do
  begin
    img := sheet.GetImage(i);

    if img.HyperlinkTarget <> '' then begin
      SplitHyperlink(img.HyperlinkTarget, target, bookmark);
      if (target <> '') and (pos('file:', target) = 0) then
      begin
        u := ParseURI(target);
        if u.Protocol = '' then
          target := '../' + target;
      end;
      if (bookmark <> '') then
        target := target + '#' + bookmark;

      AppendToStream(FSDrawingsRels[FCurSheetNum], Format(
      '  <Relationship Id="rId%d" Type="%s" Target="%s" TargetMode="External"/>' + LE, [
         rId, SCHEMAS_HYPERLINK, target
      ]));
      inc(rId);
    end;

    ext := GetImageTypeExt((FWorkbook as TsWorkbook).GetEmbeddedObj(img.Index).Imagetype);
    AppendToStream(FSDrawingsRels[FCurSheetNum], Format(
      '  <Relationship Id="rId%d" Type="%s" Target="../media/image%d.%s"/>' + LE, [
         rId, SCHEMAS_IMAGE, img.Index+1, ext
    ]));
    inc(rId);
  end;

  {$ifdef FPS_CHARTS}
  // Repeat for each chart
  for i := 0 to TsWorkbook(FWorkbook).GetChartCount - 1 do
  begin
    chart := TsWorkbook(FWorkbook).GetChartByIndex(i);
    if chart.Worksheet = AWorksheet then
    begin
      AppendToStream(FSDrawingsRels[FCurSheetNum], Format(
       '  <Relationship Id="rId%d" Type="%s" Target="../charts/chart%d.xml"/>' + LE, [
           rId, SCHEMAS_CHART, i+1
      ]));
      inc(rId);
    end;
  end;
  {$endif}

  AppendToStream(FSDrawingsRels[FCurSheetNum],
    '</Relationships>');
end;

{@ -----------------------------------------------------------------------------
  Writes a VmlDrawings file for the specified worksheet.

  This file contains information on drawing of shapes etc.
  Currently fpspreadsheet supports only comments and embedded header/footer
  images.

  Each worksheet writes a vmlDrawing file if it contains comments or
  header/footer images. All comments are packed into the same file, all
  images as well. The comments file is written first, the Images file next.
  All files are numbered consecutively for ALL sheets.

  Example
    vmlDrawing1.vml  --> Sheet 1 comments
    vmlDrawing2.vml  --> Sheet 1 header/footer images
    vmlDrawing3.vml  --> Sheet 2 header/footer images
    vmlDrawing4.vml  --> Sheet 3 comments
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.WriteVmlDrawings(AWorksheet: TsBasicWorksheet);
begin
  // At first write the VmlDrawings related to comments
  WriteVmlDrawings_Comments(AWorksheet);

  // Now write the vmlDrawings related to headers/footers
  WriteVmlDrawings_HeaderFooterImages(AWorksheet);
end;

procedure TsSpreadOOXMLWriter.WriteVMLDrawings_Comments(
  AWorksheet: TsBasicWorksheet);
var
  comment: PsComment;
  fileindex: Integer;
  index: Integer;
  id: Integer;
begin
  if (AWorksheet as TsWorksheet).Comments.Count = 0 then
    exit;

  fileIndex := Length(FSVmlDrawings);

  SetLength(FSVmlDrawings, fileIndex+1);
  FSVmlDrawings[fileIndex] := CreateTempStream(FWorkbook, Format('fpsVMLD%', [fileIndex+1]));

  // Header of file
  AppendToStream(FSVmlDrawings[fileIndex],
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" '+
         'xmlns:o="urn:schemas-microsoft-com:office:office" '+
         'xmlns:x="urn:schemas-microsoft-com:office:excel">' + LineEnding);
  // My xml viewer does not format vml files property --> format in code.
  AppendToStream(FSVmlDrawings[fileIndex],
    '  <o:shapelayout v:ext="edit">' + LineEnding +
    '    <o:idmap v:ext="edit" data="1" />' + LineEnding +
         // "data" is a comma-separated list with the ids of groups of 1024 comments -- really?
    '  </o:shapelayout>' + LineEnding);
  AppendToStream(FSVmlDrawings[fileIndex],
    '  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">'+LineEnding+
    '    <v:stroke joinstyle="miter"/>' + LineEnding +
    '    <v:path gradientshapeok="t" o:connecttype="rect"/>' + LineEnding +
    '  </v:shapetype>' + LineEnding);

  // Write vmlDrawings for each comment (formatting and position of comment box)
  index := 1;
  for comment in (AWorksheet as TsWorksheet).Comments do
  begin
    id := 1024*(FCurSheetNum+1) + index;     // if more than 1024 comments then use data="1,2,etc" above! -- not implemented yet
    AppendToStream(FSVmlDrawings[fileIndex], LineEnding + Format(
    '  <v:shape id="_x0000_s%d" type="#_x0000_t202" ', [id]) + LineEnding + Format(
    '       style="position:absolute; width:108pt; height:52.5pt; z-index:%d; visibility:hidden" ', [index]) + LineEnding +
            // it is not necessary to specify margin-left and margin-top here!

  //            'style=''position:absolute; margin-left:71.25pt; margin-top:1.5pt; ' + Format(
  //                   'width:108pt; height:52.5pt; z-index:%d; visibility:hidden'' ', [FDrawingCounter+1]) +
                  //          'width:108pt; height:52.5pt; z-index:1; visibility:hidden'' ' +

    '       fillcolor="#ffffe1" o:insetmode="auto"> '+ LineEnding +
    '    <v:fill color2="#ffffe1" />'+LineEnding+
    '    <v:shadow on="t" color="black" obscured="t" />'+LineEnding+
    '    <v:path o:connecttype="none" />'+LineEnding+
    '    <v:textbox style="mso-direction-alt:auto">'+LineEnding+
    '      <div style="text-align:left"></div>'+LineEnding+
    '    </v:textbox>' + LineEnding +
    '    <x:ClientData ObjectType="Note">'+LineEnding+
    '      <x:MoveWithCells />'+LineEnding+
    '      <x:SizeWithCells />'+LineEnding+
    '      <x:Anchor> 1, 15, 0, 2, 2, 79, 4, 4</x:Anchor>'+LineEnding+
    '      <x:AutoFill>False</x:AutoFill>'+LineEnding + Format(
    '      <x:Row>%d</x:Row>', [comment^.Row]) + LineEnding + Format(
    '      <x:Column>%d</x:Column>', [comment^.Col]) + LineEnding +
    '    </x:ClientData>'+ LineEnding+
    '  </v:shape>' + LineEnding);
  end;

  // Footer of file
  AppendToStream(FSVmlDrawings[fileIndex],
    '</xml>');
end;

(*
<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
  <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1"/>
  </o:shapelayout>
  <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
    <v:stroke joinstyle="miter"/>
    <v:formulas>
      <v:f eqn="if lineDrawn pixelLineWidth 0"/>
      <v:f eqn="sum @0 1 0"/>
      <v:f eqn="sum 0 0 @1"/>
      <v:f eqn="prod @2 1 2"/>
      <v:f eqn="prod @3 21600 pixelWidth"/>
      <v:f eqn="prod @3 21600 pixelHeight"/>
      <v:f eqn="sum @0 0 1"/>
      <v:f eqn="prod @6 1 2"/>
      <v:f eqn="prod @7 21600 pixelWidth"/>
      <v:f eqn="sum @8 21600 0"/>
      <v:f eqn="prod @7 21600 pixelHeight"/>
      <v:f eqn="sum @10 21600 0"/>
    </v:formulas>
    <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
    <o:lock v:ext="edit" aspectratio="t"/>
 </v:shapetype>
 <v:shape id="CF" o:spid="_x0000_s1025" type="#_x0000_t75" style='position:absolute;margin-left:0;margin-top:0;width:12pt;height:12pt; z-index:1'>
   <v:imagedata o:relid="rId1" o:title="arrow_down"/>
   <o:lock v:ext="edit" rotation="t"/>
 </v:shape>
 <v:shape id="RH" o:spid="_x0000_s1026" type="#_x0000_t75" style='position:absolute; margin-left:0;margin-top:0;width:12pt;height:12pt;z-index:2'>
   <v:imagedata o:relid="rId2" o:title="arrow_right"/>
   <o:lock v:ext="edit" rotation="t"/>
 </v:shape>
 <v:shape id="LH" o:spid="_x0000_s1027" type="#_x0000_t75" style='position:absolute; margin-left:0;margin-top:0;width:12pt;height:12pt;z-index:3'>
   <v:imagedata o:relid="rId3" o:title="arrow_left"/>
   <o:lock v:ext="edit" rotation="t"/>
 </v:shape>
</xml>
 *)

procedure TsSpreadOOXMLWriter.WriteVMLDrawings_HeaderFooterImages(
  AWorksheet: TsBasicWorksheet);
var
  book: TsWorkbook;
  sheet: TsWorksheet absolute AWorksheet;

  { AName = 'header' or 'footer'
    ATag  = 'L', 'C', 'R', 'x', or ' '
    AChar = 'H' or 'F' }
  procedure Process(AStream: TStream; AName: String; ATag, AChar: Char;
    AImage: TsHeaderFooterImage; var id, index: Integer);
  var
    fn: String;
    w, h: Double;
    img: TsEmbeddedObj;
  begin
    if AImage.Index = -1 then
      exit;
    if ATag = 'x' then
    begin
      book.AddErrorMsg(rsOnlyOneHeaderFooterImageAllowed, [AName]);
      exit;
    end;
    if ATag = ' ' then
    begin
      book.AddErrorMsg(rsIncorrectPositionOfImageInHeaderFooter, [AName]);
      exit;
    end;
    
    img := book.GetEmbeddedObj(AImage.Index);
    if img = nil then 
      exit;
    
    fn := ChangeFileExt(img.FileName, '');
    if fn = '' then fn := 'image';
    w := mmToPts(img.ImageWidth * AImage.ScaleX);
    h := mmToPts(img.ImageHeight * AImage.ScaleY);
    AppendToStream(AStream, Format(
      ' <v:shape id="%s" o:spid="_x0000_s%d" type="#_x0000_t75"' + LineEnding +
      //    e.g.    "CH"         _x0000_s1025
      '   style=''position:absolute;margin-left:0;margin-top:0;width:%.1fpt;height:%.1fpt;z-index:%d''>' + LineEnding +
      //    e.g.                                                     width         height         z-index:1
      '   <v:imagedata o:relid="rId%d" o:title="%s"/>' + LineEnding +
      //    e.g.               "rId1"          "arrow_down"
      '   <o:lock v:ext="edit" rotation="t" />' + LineEnding +
      ' </v:shape>' + LineEnding, [
      ATag + AChar, id, 
      w, h, index, 
      index, fn
    ], FPointSeparatorSettings));
    inc(id);
    inc(index);
  end;

var
  fileindex: Integer;
  id, index: Integer;
  tagIndex: Integer;
  img: TsHeaderFooterImage;
  sec: TsHeaderFooterSectionIndex;
  headerTags, footerTags: String;
begin
  book := FWorkbook as TsWorkbook;
  if not sheet.PageLayout.HasHeaderFooterImages then
    exit;

  fileIndex := Length(FSVmlDrawings);
  SetLength(FSVmlDrawings, fileIndex+1);
  FSVmlDrawings[fileIndex] := CreateTempStream(FWorkbook, Format('fpsVMLD%d', [fileIndex+1]));

  // Header of file
  AppendToStream(FSVmlDrawings[fileIndex],
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" ' + LineEnding +
    '     xmlns:o="urn:schemas-microsoft-com:office:office" ' + LineEnding +
    '     xmlns:x="urn:schemas-microsoft-com:office:excel"> ' + LineEnding +
    '  <o:shapelayout v:ext="edit">' + LineEnding +
    '    <o:idmap v:ext="edit" data="1"/>' + LineEnding +
    '  </o:shapelayout>' + LineEnding +
    '  <v:shapetype id="_x0000_t75" coordsize="21600,21600" ' + LineEnding +
    '               o:spt="75" o:preferrelative="t" ' + LineEnding +
    '               path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">' + LineEnding +
    '    <v:stroke joinstyle="miter"/>' + LineEnding +
    '    <v:formulas>' + LineEnding +
    '      <v:f eqn="if lineDrawn pixelLineWidth 0"/>' + LineEnding +
    '      <v:f eqn="sum @0 1 0"/>' + LineEnding +
    '      <v:f eqn="sum 0 0 @1"/>' + LineEnding +
    '      <v:f eqn="prod @2 1 2"/>' + LineEnding +
    '      <v:f eqn="prod @3 21600 pixelWidth"/>' + LineEnding +
    '      <v:f eqn="prod @3 21600 pixelHeight"/>' + LineEnding +
    '      <v:f eqn="sum @0 0 1"/>' + LineEnding +
    '      <v:f eqn="prod @6 1 2"/>' + LineEnding +
    '      <v:f eqn="prod @7 21600 pixelWidth"/>' + LineEnding +
    '      <v:f eqn="sum @8 21600 0"/>' + LineEnding +
    '      <v:f eqn="prod @7 21600 pixelHeight"/>' + LineEnding +
    '      <v:f eqn="sum @10 21600 0"/>' + LineEnding +
    '    </v:formulas>' + LineEnding +
    '    <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>' + LineEnding +
    '    <o:lock v:ext="edit" aspectratio="t"/>' + LineEnding +
    ' </v:shapetype>' + LineEnding);

  index := 1;
  id := 1024 * (FCurSheetNum+1) + index;

  sheet.PageLayout.GetImageSections(headerTags, footerTags);

  // Write the data for the image in each section of the header
  for sec in TsHeaderFooterSectionIndex do
  begin
    tagIndex := ord(sec) + 1;
    img := sheet.PageLayout.HeaderImages[sec];
    Process(FSVmlDrawings[fileIndex], rsHeader, headerTags[tagIndex], 'H', img, id, index);
  end;
  // Repeat with footer
  for sec in TsHeaderFooterSectionIndex do
  begin
    img := sheet.PageLayout.FooterImages[sec];
    tagIndex := ord(sec) + 1;
    Process(FSVmlDrawings[fileIndex], rsFooter, footerTags[tagIndex], 'F', img, id, index);
  end;

  // Footer of file
  AppendToStream(FSVmlDrawings[fileIndex],
    '</xml>');
end;

{<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
      Target="../media/image3.png"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
      Target="../media/image2.png"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
      Target="../media/image1.png"/>
  </Relationships> }

{@@ ----------------------------------------------------------------------------
  Writes a relationship file (*.rels) for a vmlDrawing.xml file to a media file.
  Destination folder will be "../drawings/_rels".
  Needed for header/footer images.
  Note: vmlDrawing files of comments do not have a correspondig rels file.
  The index of the rels file must match that of the vmlDrawingX.vml file.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.WriteVmlDrawingRels(AWorksheet: TsBasicWorksheet);
var
  fileindex: Integer;
  sec: TsHeaderFooterSectionIndex;
  rId: Integer;
  img: TsHeaderFooterImage;
  ext: String;
  sheet: TsWorksheet absolute AWorksheet;
begin
  if not sheet.PageLayout.HasHeaderFooterImages then
    exit;

  fileIndex := Length(FSVmlDrawingsRels);
  if sheet.Comments.Count > 0 then
    inc(fileIndex);  // skip comments for numbering

  SetLength(FSVmlDrawingsRels, fileIndex+1);
  FsVmlDrawingsRels[fileIndex] := CreateTempStream(FWorkbook, Format('fpsVMSDR%d', [fileIndex]));

  // Write file header
  AppendToStream(FSVmlDrawingsRels[fileIndex],
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + LineEnding +
    '<Relationships xmlns="' + SCHEMAS_RELS + '">' + LineEnding
  );

  // Write entry for each header/footer image
  // Note: use same order as for writing of VmlDrawing files.

  rId := 1;

  // Write the data for the image in each section of the header
  for sec in TsHeaderFooterSectionIndex do begin
    img := sheet.PageLayout.HeaderImages[sec];
    if img.Index = -1 then
      continue;
    ext := GetImageTypeExt((FWorkbook as TsWorkbook).GetEmbeddedObj(img.Index).ImageType);
//    imgIdx := FWorkbook.FindEmbeddedObj(imgName);
    AppendToStream(FSVmlDrawingsRels[fileIndex], Format(
      '  <Relationship Id="rId%d" Target="../media/image%d.%s" '+
         'Type="' + SCHEMAS_IMAGE + '" />' + LineEnding, [
      rId,                   // Id="rID1"
      img.Index + 1, ext     // Target="../media/image1.png"
    ]));
   inc(rId);
  end;

  // Repeat with footer
  for sec in TsHeaderFooterSectionIndex do begin
    img := sheet.PageLayout.FooterImages[sec];
    if img.Index = -1 then
      continue;
    ext := GetImageTypeExt((FWorkbook as TsWorkbook).GetEmbeddedObj(img.Index).Imagetype);
    AppendToStream(FSVmlDrawingsRels[fileIndex], Format(
      '  <Relationship Id="rId%d" Target="../media/image%d.%s" '+  //
         //  e.g.         "rId1"         "..(media/image1.png"
         'Type="' + SCHEMAS_IMAGE + '" />', [
      rId,
      img.Index + 1, ext
    ]));
   inc(rId);
  end;

  // Write file footer
  AppendToStream(FSVmlDrawingsRels[fileIndex],
    '</Relationships>');
end;

procedure TsSpreadOOXMLWriter.WriteWorksheetRels(AWorksheet: TsBasicWorksheet);
var
  AVLNode: TAvgLvlTreeNode;
  hyperlink: PsHyperlink;
  s: String;
  target, bookmark: String;
  rId_Comments, rId_Hyperlink, rId_Drawing, rId_DrawingHF: Integer;
  sheet: TsWorksheet absolute AWorksheet;
begin
  // Extend stream array
  // NOTE: If no .rels file is written for this sheet at least an empty stream
  // must be provided to keep the numbering intact.
  SetLength(FSSheetRels, FCurSheetNum + 1);

  // Anything to write?
  if (sheet.Comments.Count = 0) and (sheet.Hyperlinks.Count = 0) and
     (sheet.GetImageCount = 0) {$ifdef FPS_CHARTS}and (sheet.GetChartCount = 0){$endif} and
     (not (sheet.PageLayout.HasHeaderFooterImages))
  then
    exit;

  Get_rId(AWorksheet, rID_Comments, rId_Hyperlink, rId_Drawing, rId_DrawingHF);

  // Create stream
  FSSheetRels[FCurSheetNum] := CreateTempStream(FWorkbook, Format('fpsWSR%d', [FCurSheetNum]));

  // Header
  AppendToStream(FSSheetRels[FCurSheetNum],
    XML_HEADER + LineEnding);
  AppendToStream(FSSheetRels[FCurSheetNum], Format(
    '<Relationships xmlns="%s">' + LineEnding, [SCHEMAS_RELS]));

  // Relationships for comments
  if sheet.Comments.Count > 0 then
  begin
    AppendToStream(FSSheetRels[FCurSheetNum], Format(
      '  <Relationship Id="rId%d" Target="../comments%d.xml" Type="%s" />' + LineEnding,
          [rId_Comments+1, FCurSheetNum+1, SCHEMAS_COMMENTS]));
    AppendToStream(FSSheetRels[FCurSheetNum], Format(
      '  <Relationship Id="rId%d" Target="../drawings/vmlDrawing%d.vml" Type="%s" />' + LineEnding,
          [rId_Comments, vmlDrawingCounter, SCHEMAS_VMLDRAWING]));
    inc(vmlDrawingCounter);
  end;

  // Relationships for hyperlinks
  if sheet.Hyperlinks.Count > 0 then
  begin
    AVLNode := sheet.Hyperlinks.FindLowest;
    while Assigned(AVLNode) do
    begin
      hyperlink := PsHyperlink(AVLNode.Data);
      SplitHyperlink(hyperlink^.Target, target, bookmark);
      if target <> '' then
      begin
        if (pos('file:', target) = 0) and FileNameIsAbsolute(target) then
          FileNameToURI(target);
        s := Format('Id="rId%d" Target="%s" TargetMode="External" Type="%s"',
          [rId_Hyperlink, target, SCHEMAS_HYPERLINK]);
        AppendToStream(FSSheetRels[FCurSheetNum],
          '  <Relationship ' + s + ' />' + LineEnding);
        inc(rId_Hyperlink);
      end;
      AVLNode := sheet.Hyperlinks.FindSuccessor(AVLNode);
    end;
  end;

  // Relationships for charts or embedded images
  // relationship with to the ../drawings/drawingN.xml file containing all
  // chart/image infos. N is the 1-based sheet index
  if (sheet.GetImageCount > 0) {$ifdef FPS_CHARTS}or (sheet.GetChartCount > 0){$endif} then
    AppendToStream(FSSheetRels[FCurSheetNum], Format(
      '  <Relationship Id="rId%d" Target="../drawings/drawing%d.xml" Type="%s" />' + LineEnding,
      [rId_Drawing, FCurSheetNum + 1, SCHEMAS_DRAWING]
    ));

  // Relationships for embedded header/footer images
  if sheet.PageLayout.HasHeaderFooterImages then
  begin
    AppendToStream(FSSheetRels[FCurSheetnum], Format(
      '  <Relationship Id="rId%d" Target="../drawings/vmlDrawing%d.vml" Type="%s" />' + LineEnding,
        [rId_DrawingHF, vmlDrawingCounter, SCHEMAS_VMLDRAWING]));
    inc(vmlDrawingCounter);
  end;

  // Footer
  AppendToStream(FSSheetRels[FCurSheetNum],
    '</Relationships>');
end;

procedure TsSpreadOOXMLWriter.WriteGlobalFiles;
begin
  { --- Content Types --- }
  // Will be written at the end of WriteToStream when all Sheet.rels files are
  // known

  { --- meta data ---- }
  WriteMetaData(FSMetaData);
  WriteCustomMetaData(FSCustomMetaData);

  { --- _rels/.rels --- }
  AppendToStream(FSRelsRels,
    XML_HEADER + LineEnding);

  AppendToStream(FSRelsRels, Format(
    '<Relationships xmlns="%s">' + LineEnding,
      [SCHEMAS_RELS]));

  AppendToStream(FSRelsRels, Format(
      '<Relationship Id="rId1" Target="xl/workbook.xml" Type="%s" />' + LineEnding,
      [SCHEMAS_DOCUMENT]));

  AppendToStream(FSRelsRels, Format(
      '<Relationship Id="rId2" Target="docProps/core.xml" Type="%s" />' + LineEnding,
      [SCHEMAS_META_CORE]));

  if TsWorkbook(FWorkbook).MetaData.Custom.Count > 0 then
    AppendToStream(FSRelsRels, Format(
      '<Relationship Id="rId3" Target="docProps/custom.xml" Type="%s" />' + LineEnding,
      [SCHEMAS_META_CUSTOM]));

  AppendToStream(FSRelsRels,
    '</Relationships>');

  { --- xl/styles --- }
  AppendToStream(FSStyles,
    XML_Header);
  AppendToStream(FSStyles, Format(
    '<styleSheet xmlns="%s">', [SCHEMAS_SPREADML]));

  // Number formats
  WriteNumFormatList(FSStyles);

  // Fonts
  WriteFontList(FSStyles);

  // Fill patterns
  WriteFillList(FSStyles);

  // Borders
  WriteBorderList(FSStyles);

  // Style records
  AppendToStream(FSStyles,
      '<cellStyleXfs count="1">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" />' +
      '</cellStyleXfs>'
  );
  WriteStyleList(FSStyles, 'cellXfs');

  // Cell style records
  AppendToStream(FSStyles,
      '<cellStyles count="1">' +
        '<cellStyle name="Normal" xfId="0" builtinId="0" />' +
      '</cellStyles>');

  // Conditional format styles
  WriteDifferentialFormats(FSStyles);

  // Misc
  AppendToStream(FSStyles,
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16" />');

  // Palette
  WritePalette(FSStyles);

  AppendToStream(FSStyles,
    '</styleSheet>');
end;

{ Write folder "media" with embedded streams }
procedure TsSpreadOOXMLWriter.WriteMedia(AZip: TZipper);
var
  i: Integer;
  stream: TMemoryStream;
  embObj: TsEmbeddedObj;
  embName: String;
  ext: String;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;
  for i:=0 to book.GetEmbeddedObjCount-1 do
  begin
    embObj := book.GetEmbeddedObj(i);
    stream := embObj.Stream;
    stream.Position := 0;
    ext := GetImageTypeExt(embObj.ImageType);
    embName := Format('image%d.%s', [i+1, ext]);
    AZip.Entries.AddFileEntry(stream, OOXML_PATH_XL_MEDIA + embname);
  end;
end;

procedure TsSpreadOOXMLWriter.WriteMetaData(AStream: TStream);
{<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
  <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:title>test file meta data äöü</dc:title>
    <dc:creator>Donald Duck</dc:creator>
    <dc:description>this is a comment_x000d_ in two lines äöü.</dc:description>
    <cp:lastModifiedBy>Donald Duck</cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">2015-06-05T18:19:34Z</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">2020-07-27T21:23:27Z</dcterms:modified>
  </cp:coreProperties>    }
var
  book: TsWorkbook;
  s: String;
  dt: TDateTime;
begin
  book := TsWorkbook(FWorkbook);

  AppendToStream(AStream,
    XML_HEADER);

  AppendToStream(AStream,
   '<cp:coreProperties '+
     'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '+
     'xmlns:dc="http://purl.org/dc/elements/1.1/" '+
     'xmlns:dcterms="http://purl.org/dc/terms/" '+
     'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '+
     'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">');

  if book.MetaData.Title <> '' then
    AppendToStream(AStream, Format(
      '<dc:title>%s</dc:title>', [UTF8TextToXMLText(book.MetaData.Title)]));

  if book.MetaData.Subject <> '' then
    AppendToStream(AStream, Format(
      '<dc:subject>%s</dc:subject>', [UTF8TextToXMLText(book.Metadata.Subject)]));

  if book.MetaData.CreatedBy <> '' then
    AppendToStream(AStream, Format(
      '<dc:creator>%s</dc:creator>', [UTF8TextToXMLText(book.MetaData.CreatedBy)]));

  if book.MetaData.Keywords.Count > 0 then
  begin
    s := book.MetaData.KeyWords.CommaText;
    AppendToStream(AStream, Format(
      '<cp:keywords>%s</cp:keywords>', [UTF8TextToXMLText(s)]));
  end;

  if book.MetaData.Comments.Count > 0 then
  begin
    s := book.MetaData.Comments.Text;
    while (s <> '') and (s[Length(s)] in [#10, #13]) do
      Delete(s, Length(s), 1);
    s := StringReplace(s, LineEnding, '_x000d_', [rfReplaceAll]);
    AppendToStream(AStream, Format(
      '<dc:description>%s</dc:description>', [UTF8TextToXMLText(s)]));
  end;

  if book.MetaData.LastModifiedBy <> '' then
  begin
    s := book.MetaData.LastModifiedBy;
    AppendToStream(AStream, Format(
      '<cp:lastModifiedBy>%s</cp:lastModifiedBy>', [UTF8TextToXMLText(s)]));
  end;

  if book.MetaData.DateCreated > 0 then
  begin
    dt := book.MetaData.DateCreated + GetLocalTimeOffset / MinsPerDay;
    s := FormatDateTime(ISO8601FormatExtendedUTC, dt);
    AppendToStream(AStream, Format(
      '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>', [UTF8TextToXMLText(s)]));
  end;

  if book.MetaData.DateLastModified >0 then
  begin
    dt := book.MetaData.DateLastModified + GetLocalTimeOffset / MinsPerDay;
    s := FormatDateTime(ISO8601FormatExtendedUTC, dt);
    AppendToStream(AStream, Format(
      '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified>', [s]));
  end;

  AppendToStream(AStream,
    '</cp:coreProperties>');
end;


{
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    Target="theme/theme1.xml"/>
</Relationships>
}
procedure TsSpreadOOXMLWriter.WriteContent;
var
  i: Integer;
begin
  { Global workbook data - Mark all sheets }
  WriteWorkbook(FSWorkbook);

  { Preparation for shared strings }
  FSharedStringsCount := 0;

  { Write all worksheets which fills also the shared strings.
    Also: write comments, Drawings, vmlDrawings and relationship files }
  for i := 0 to (Workbook as TsWorkbook).GetWorksheetCount - 1 do
  begin
    FWorksheet := (Workbook as TsWorkbook).GetWorksheetByIndex(i);
    WriteWorksheet(FWorksheet);
    WriteComments(FWorksheet);
    WriteVmlDrawings(FWorksheet);
    WriteVmlDrawingRels(FWorksheet);
    WriteDrawings(FWorksheet);
    WriteDrawingRels(FWorksheet);
    WriteWorksheetRels(FWorksheet);
  end;

  { Finalization of the shared strings document }
  if FSharedStringsCount > 0 then
  begin
    AppendToStream(FSSharedStrings_complete,
      XML_HEADER, Format(
        '<sst xmlns="%s" count="%d" uniqueCount="%d">', [
        SCHEMAS_SPREADML, FSharedStringsCount, FSharedStringsCount
    ]));
    ResetStream(FSSharedStrings);
    FSSharedStrings_complete.CopyFrom(FSSharedStrings, FSSharedStrings.Size);
    AppendToStream(FSSharedStrings_complete,
        '</sst>');
  end;

 {$ifdef FPS_CHARTS}
  { Write all charts }
  FChartWriter.WriteCharts;
 {$endif}

  { Workbook relations - Mark relation to all sheets }
  WriteWorkbookRels(FSWorkbookRels);
end;

procedure TsSpreadOOXMLWriter.WriteContentTypes;
var
  i, j, n: Integer;
  imgext: TStringList;
  ext: String;
  sheet: TsWorksheet;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;

  AppendToStream(FSContentTypes,
    XML_HEADER + LineEnding);
  AppendToStream(FSContentTypes,
    '<Types xmlns="' + SCHEMAS_TYPES + '">' + LineEnding);

  AppendToStream(FSContentTypes, Format(
      '<Default Extension="rels" ContentType="%s" />' + LineEnding, [MIME_RELS]));
  AppendToStream(FSContentTypes, Format(
      '<Default Extension="xml" ContentType="%s" />' + LineEnding, [MIME_XML]));
   AppendToStream(FSContentTypes, Format(
      '<Default Extension="vml" ContentType="%s" />' + LineEnding, [MIME_VMLDRAWING]));

  if book.GetEmbeddedObjCount > 0 then
  begin
    imgExt := TStringList.Create;
    try
      for i:=0 to book.GetEmbeddedObjCount-1 do
      begin
        ext := GetImageTypeExt(book.GetEmbeddedObj(i).ImageType);
        j := imgExt.IndexOf(ext);
        if j = -1 then
          imgExt.Add(ext);
      end;
      for i := 0 to imgExt.Count-1 do
        AppendToStream(FSContentTypes, Format(
          '<Default Extension="%s" ContentType="image/%s" />' + LineEnding, [imgExt[i], imgExt[i]]));
    finally
      imgExt.Free;
    end;
  end;

  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/workbook.xml" ContentType="' + MIME_SHEET + '" />' + LineEnding);

  for i:=1 to book.GetWorksheetCount do
  begin
    AppendToStream(FSContentTypes, Format(
      '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="%s" />' + LineEnding,
        [i, MIME_WORKSHEET]));
    sheet := book.GetWorksheetByIndex(i-1);
    if (sheet.GetImageCount > 0) {$ifdef FPS_CHARTS} or (sheet.GetChartCount > 0){$endif} then
      AppendToStream(FSContentTypes, Format(
        '<Override PartName="/xl/drawings/drawing%d.xml" ContentType="%s"/>' + LineEnding,
        [i, MIME_DRAWING]));
  end;

  for i:=1 to Length(FSComments) do
    AppendToStream(FSContentTypes, Format(
      '<Override PartName="/xl/comments%d.xml" ContentType="%s" />' + LineEnding,
        [i, MIME_COMMENTS]));

  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/styles.xml" ContentType="' + MIME_STYLES + '" />' + LineEnding);
  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/sharedStrings.xml" ContentType="' + MIME_STRINGS + '" />' + LineEnding);

  {$ifdef FPS_CHARTS}
  TsSpreadOOXMLChartWriter(FChartWriter).WriteChartContentTypes(FSContentTypes);
  {$endif}

  AppendToStream(FSContentTypes,
      '<Override PartName="/docProps/core.xml" ContentType="' + MIME_CORE + '" />');

  if book.MetaData.Custom.Count > 0 then
    AppendToStream(FSContentTypes,
      '<Override PartName="/docProps/custom.xml" ContentType="' + MIME_CUSTOM + '" />');

  AppendToStream(FSContentTypes,
    '</Types>');
end;

procedure TsSpreadOOXMLWriter.WriteDefinedNames(AStream: TStream);
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  stotal, srng, sheetname: String;
  i, j: Integer;
  prng: TsCellRange;
  firstIndex, lastIndex: Integer;
  defName: TsDefinedName;
begin
  stotal := '';

  // Write global defined names
  book := TsWorkbook(FWorkbook);
  for i := 0 to book.DefinedNames.Count-1 do
  begin
    defName := book.DefinedNames[i];
    sTotal := sTotal + Format('<definedName name="%s">%s</definedName>',
      [ defName.Name, defName.RangeAsString ]
    );
  end;


  // Write print ranges and repeatedly printed rows and columns
  for i := 0 to (Workbook as TsWorkbook).GetWorksheetCount-1 do
  begin
    sheet := (Workbook as TsWorkbook).GetWorksheetByIndex(i);
    sheetname := '''' + UTF8TextToXMLText(sheet.Name) + '''';

    // Cell block of print range
    srng := '';
    for j := 0 to sheet.PageLayout.NumPrintRanges - 1 do
    begin
      prng := sheet.PageLayout.GetPrintRange(j);
      srng := srng + ',' + Format('%s!%s', [
        sheetname, GetCellRangeString(prng.Row1, prng.Col1, prng.Row2, prng.Col2, [])
      ]);
    end;
    if srng <> '' then
    begin
      Delete(srng, 1, 1);
      stotal := stotal + Format(
        '<definedName name="_xlnm.Print_Area" localSheetId="%d">%s</definedName>',
        [i, srng]
      );
    end;

    // repeated columns ...
    srng := '';
    if sheet.PageLayout.RepeatedCols.FirstIndex <> UNASSIGNED_ROW_COL_INDEX then
    begin
      firstindex := sheet.PageLayout.RepeatedCols.FirstIndex;
      lastindex := IfThen(sheet.PageLayout.RepeatedCols.LastIndex = UNASSIGNED_ROW_COL_INDEX,
        firstindex, sheet.PageLayout.RepeatedCols.LastIndex);
      srng := srng + ',' + Format('%s!$%s:$%s', [sheetname, GetColString(firstindex), GetColString(lastindex)]);
    end;
    // ... and repeated rows
    if sheet.PageLayout.RepeatedRows.FirstIndex <> UNASSIGNED_ROW_COL_INDEX then
    begin
      firstindex := sheet.PageLayout.RepeatedRows.FirstIndex;
      lastindex := IfThen(sheet.PageLayout.RepeatedRows.LastIndex = UNASSIGNED_ROW_COL_INDEX,
        firstindex, sheet.PageLayout.RepeatedRows.LastIndex);
      srng := srng + ',' + Format('%s!$%d:$%d', [sheetname, firstindex+1, lastindex+1]);
    end;
    if srng <> '' then begin
      Delete(srng, 1,1);
      stotal := stotal + Format(
        '<definedName name="_xlnm.Print_Titles" localSheetId="%d">%s</definedName>',
        [i, srng]
      );
    end;
  end;

  // Write to stream if any defined names exist
  if stotal <> '' then
    AppendToStream(AStream,
      '<definedNames>' + stotal + '</definedNames>');
end;

procedure TsSpreadOOXMLWriter.WriteDifferentialFormat(AStream: TStream;
  AFormat: PsCellFormat);

  procedure WriteBorderStyle(AStream: TStream; AFormatRecord: PsCellFormat;
    ABorder: TsCellBorder; ABorderName: String);
  { border names found in xlsx files for Excel selections:
    "thin", "hair", "dotted", "dashed", "dashDotDot", "dashDot", "mediumDashDotDot",
    "slantDashDot", "mediumDashDot", "mediumDashed", "medium", "thick", "double" }
  var
    styleName: String;
    colorStr: String;
    rgb: TsColor;
  begin
    if (ABorder in AFormatRecord^.Border) then begin
      // Line style
      styleName := LINESTYLE_TYPES[AFormatRecord^.BorderStyles[ABorder].LineStyle];

      // Border color
      rgb := AFormatRecord^.BorderStyles[ABorder].Color;
      colorStr := ColorToExcelStr(rgb);
      AppendToStream(AStream, Format(
        '<%s style="%s"><color rgb="%s" /></%s>',
          [ABorderName, styleName, colorStr, ABorderName]
        ));
    end else
      AppendToStream(AStream, Format(
        '<%s />', [ABorderName]));
  end;

var
  pt, bc, fc, diag: string;
  font: TsFont;
  nfp: TsNumFormatParams;
  nfs: String;
  nfi: Integer;
  nfId: String;
begin
  AppendToStream(AStream,
    '<dxf>');

  { font }
  if (uffFont in AFormat^.UsedFormattingFields) then
  begin
    font := TsWorkbook(FWorkbook).GetFont(AFormat^.FontIndex);
    if font <> nil then
    begin
      AppendToStream(AStream, '<font>');
      if font.Color <> scNotDefined then
      begin
        fc := ColorToExcelStr(font.Color);
        AppendToStream(AStream, Format('<color rgb="%s" />', [fc] ));
      end;
      if fssBold in font.Style then
        AppendToStream(AStream, '<b />');
      if fssItalic in font.Style then
        AppendToStream(AStream, '<i />');
      if fssStrikeout in font.Style then
        AppendToStream(AStream, '<strike />');
      // Font name, font size, and style underline not supported
      AppendToStream(AStream, '</font>');
    end;
  end;

  { number format }
  if (uffNumberFormat in AFormat^.UsedFormattingFields) then
  begin
    nfp := TsWorkbook(FWorkbook).GetNumberFormat(AFormat^.NumberFormatIndex);
//    nfs := UTF8TextToXMLText(nfp.NumFormatStr);
    nfs := nfp.NumFormatStr;
    nfi := FNumFormatList.IndexOf(nfs);
    if nfi > -1 then nfId := Format('numFmtId="%d" ', [nfi]) else nfId := '';
    AppendToStream(AStream, Format('<numFmt %sformatCode="%s" />', [nfId, nfs]));
  end;

  { background fill }
  if (uffBackground in AFormat^.UsedFormattingFields) then
  begin
    pt := PATTERN_TYPES[AFormat^.Background.Style];
    if AFormat^.Background.FgColor <> scTransparent then
      fc := Format('rgb="%s"', [ColorToExcelStr(AFormat^.Background.FgColor)]);
    if AFormat^.Background.BgColor = scTransparent then
      bc := 'auto="1"'
    else
      bc := Format('rgb="%s"', [ColorToExcelStr(AFormat^.Background.BgColor)]);
    AppendToStream(AStream,
      '<fill>' + Format(
        '<patternFill patternType="%s">', [pt]) + Format(
          '<fgColor %s />', [fc]) + Format(
          '<bgColor %s />', [bc]) +
        '</patternFill>' +
      '</fill>');
  end;

  { cell borders }
  if (uffBorder in AFormat^.UsedFormattingFields) then
  begin
    diag := '';
    if (cbDiagUp in AFormat^.Border) then
      diag := diag + ' diagonalUp="1"';
    if (cbDiagDown in AFormat^.Border) then
      diag := diag + ' diagonalDown="1"';
    AppendToStream(AStream,
      '<border' + diag + '>');
        WriteBorderStyle(AStream, AFormat, cbWest, 'left');
        WriteBorderStyle(AStream, AFormat, cbEast, 'right');
        WriteBorderStyle(AStream, AFormat, cbNorth, 'top');
        WriteBorderStyle(AStream, AFormat, cbSouth, 'bottom');
        // OOXML uses the same border style for both diagonals. In agreement with
        // the biff implementation we select the style from the diagonal-up line.
        WriteBorderStyle(AStream, AFormat, cbDiagUp, 'diagonal');
    AppendToStream(AStream,
      '</border>');
  end;

  AppendToStream(AStream,
    '</dxf>');
end;

procedure TsSpreadOOXMLWriter.WriteDifferentialFormats(AStream: TStream);
var
  book: TsWorkbook;
  i: Integer;
  fmtIndex: Integer;
  fmt: PsCellFormat;
begin
  if Length(FDifferentialFormatIndexList) = 0 then
  begin
    AppendToStream(AStream, '<dxfs count="0" />');
    exit;
  end;

  AppendToStream(AStream, Format(
    '<dxfs count="%d">', [Length(FDifferentialFormatIndexList)]));

  book := TsWorkbook(FWorkbook);
  for i := 0 to High(FDifferentialFormatIndexList) do
  begin
    fmtIndex := FDifferentialFormatIndexList[i];
    fmt := book.GetPointerToCellFormat(fmtIndex);
    WriteDifferentialFormat(AStream, fmt);
  end;

  AppendToStream(AStream,
    '</dxfs>');
  {
  AppendToStream(AStream,
    '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16" />');
    }
end;

procedure TsSpreadOOXMLWriter.WriteWorkbook(AStream: TStream);
begin
  AppendToStream(AStream,
    XML_HEADER);
  AppendToStream(AStream, Format(
    '<workbook xmlns="%s" xmlns:r="%s">', [SCHEMAS_SPREADML, SCHEMAS_DOC_RELS]));
  AppendToStream(AStream,
      '<fileVersion appName="fpspreadsheet" />');
  AppendToStream(AStream,
      '<workbookPr defaultThemeVersion="124226" />');
  WriteWorkbookProtection(AStream);
  AppendToStream(AStream,
      '<bookViews>' +
        '<workbookView xWindow="480" yWindow="90" ' +
          'windowWidth="15195" windowHeight="12525"' + GetActiveTab + '/>' +
      '</bookViews>');
  WriteSheets(AStream);
  WriteDefinedNames(AStream);

  AppendToStream(AStream,
      '<calcPr calcId="114210" />');
  AppendToStream(AStream,
    '</workbook>');
end;

procedure TsSpreadOOXMLWriter.WriteWorkbookProtection(AStream: TStream);
var
  s: String;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;
  s := '';

  if book.CryptoInfo.PasswordHash <> '' then
  begin
    if book.CryptoInfo.Algorithm = caExcel then
      s := s + ' workbookPassword="' + book.CryptoInfo.PasswordHash + '"'
    else
    begin
      s:= s + ' workbookHashVal="' + book.CryptoInfo.PasswordHash + '"';
      if book.CryptoInfo.Algorithm <> caUnknown then
        s:= s + ' workbookAlgorithmName="' + AlgorithmToStr(book.CryptoInfo.Algorithm, auExcel) + '"';

      if book.CryptoInfo.SaltValue <> '' then
        s:= s + ' workbookSaltValue="' + book.CryptoInfo.SaltValue + '"';

      if book.CryptoInfo.SpinCount <> 0 then
        s:= s + ' workbookSpinCount="' + IntToStr(book.CryptoInfo.SpinCount) + '"';
    end;
  end;

  {
  if book.RevisionsCrypto.Password <> '' then
    s:= s + ' revisionsPassword="' + book.RevisionsCrypto.Password +'"'
  else
  if book.RevisionsCrypto.HashValue <> '' then
  begin
    s:= s + ' revisionsHashValue="' + book.RevisionsCrypto.HashValue +'"';

    if book.RevisionsCrypto.AlgorithmName <> '' then
      s:= s + ' revisionsAlgorithm="' + book.RevisionsCrypto.AlgorithmName +'"';

    if book.RevisionsCrypto.SaltValue <> '' then
      s:= s + ' revisionsSaltValue="' + book.RevisionsCrypto.SaltValue +'"';

    if book.RevisionsCrypto.SpinCount <> 0 then
      s:= s + ' revisionsSpinCount="' + IntToStr( book.RevisionsCrypto.SpinCount ) +'"';
  end;
  }

  if bpLockStructure in book.Protection then
    s := s + ' lockStructure="1"';

  if bpLockWindows in book.Protection then
    s := s + ' lockWindows="1"';

  if bpLockRevision in book.Protection then
    s := s + ' lockRevision="1"';

  if s <> '' then
    AppendToStream(AStream, '<workbookProtection' + s +' />');
end;

procedure TsSpreadOOXMLWriter.WriteWorkbookRels(AStream: TStream);
var
  counter: Integer;
begin
  AppendToStream(AStream,
    XML_HEADER + LineEnding,
    '<Relationships xmlns="' + SCHEMAS_RELS + '">' + LineEnding);

  counter := 1;
  while counter <= (Workbook as TsWorkbook).GetWorksheetCount do begin
    AppendToStream(AStream, Format(
      '  <Relationship Id="rId%d" Target="worksheets/sheet%d.xml" Type="%s" />' + LineEnding,
        [counter, counter, SCHEMAS_WORKSHEET]));
    inc(counter);
  end;

  AppendToStream(AStream, Format(
      '  <Relationship Id="rId%d" Target="styles.xml" Type="%s" />' + LineEnding,
        [counter, SCHEMAS_STYLES]));
  inc(counter);

  if FSharedStringsCount > 0 then begin
    AppendToStream(AStream, Format(
      '  <Relationship Id="rId%d" Target="sharedStrings.xml" Type="%s" />' + LineEnding,
        [counter, SCHEMAS_STRINGS]));
    inc(counter);
  end;

  AppendToStream(AStream,
    '</Relationships>');
end;

procedure TsSpreadOOXMLWriter.WriteWorksheet(AWorksheet: TsBasicWorksheet);
var
  rId_Comments: Integer;
  rId_FirstHyperlink: Integer;
  rId_Drawing, rId_DrawingHF: Integer;
  worksheet: TsWorksheet;
begin
  worksheet := TsWorksheet(AWorksheet);

  FCurSheetNum := Length(FSSheets);
  SetLength(FSSheets, FCurSheetNum + 1);

  Get_rId(AWorksheet, rID_Comments, rId_FirstHyperlink, rId_Drawing,
    rId_DrawingHF);

  // Create the stream
  FSSheets[FCurSheetNum] := CreateTempStream(FWorkbook, Format('fpsSH%d', [FCurSheetNum]));

  // Header
  AppendToStream(FSSheets[FCurSheetNum],
    XML_HEADER);
  AppendToStream(FSSheets[FCurSheetNum], Format(
    '<worksheet xmlns="%s" xmlns:r="%s">', [SCHEMAS_SPREADML, SCHEMAS_DOC_RELS]));

  WriteSheetPr(FSSheets[FCurSheetNum], AWorksheet);
  WriteDimension(FSSheets[FCurSheetNum], AWorksheet);
  WriteSheetViews(FSSheets[FCurSheetNum], AWorksheet);
  WriteSheetFormatPr(FSSheets[FCurSheetNum], AWorksheet);
  WriteCols(FSSheets[FCurSheetNum], AWorksheet);
  WriteSheetData(FSSheets[FCurSheetNum], AWorksheet);
  WriteSheetProtection(FSSheets[FCurSheetNum], AWorksheet);
  WriteMergedCells(FSSheets[FCurSheetNum], AWorksheet);
  WriteConditionalFormats(FSSheets[FCurSheetNum], AWorksheet);   // Must be before hyperlinks!!!
  WriteHyperlinks(FSSheets[FCurSheetNum], AWorksheet, rId_FirstHyperlink);

  WritePrintOptions(FSSheets[FCurSheetNum], AWorksheet);
  WritePageMargins(FSSheets[FCurSheetNum], AWorksheet);
  WritePageSetup(FSSheets[FCurSheetNum], AWorksheet);
  WriteRowBreaks(FSSheets[FCurSheetNum], AWorksheet);
  WriteColBreaks(FSSheets[FCurSheetNum], AWorksheet);
  WriteHeaderFooter(FSSheets[FCurSheetNum], AWorksheet);

  { This item is required for all embedded images and charts.
    There must be a matching file in "drawingN.xml" file in "../drawings"
    which contains the image/chart-related data of all images/charts in this sheet.
    The file in turn requires an entry "drawingN.xml.rels" in the drawings rels
    folder }
  if (worksheet.GetImageCount > 0) {$ifdef FPS_CHARTS}or (worksheet.GetChartCount > 0){$endif} then
    AppendToStream(FSSheets[FCurSheetNum], Format(
      '<drawing r:id="rId%d" />', [rId_Drawing]));

  { This item is required for all comments of a worksheet.
    Comments have two entries in the sheet's .rels file, one for the
    "../comments.xml" file, and one for the "../drawings/vmlDrawingX.vml" file.
    The vmlDrawing file must have an entry "vmlDrawingX.vml.rels" in the drawings
    rels folder. }
  if (AWorksheet as TsWorksheet).Comments.Count > 0 then
    AppendToStream(FSSheets[FCurSheetNum], Format(
      '<legacyDrawing r:id="rId%d" />', [rId_Comments]));

  { This item is required for all images embedded to a header/footer.
    There must be a corresponding "vmlDrawingX.vml" file in "../drawings". }
  if (AWorksheet as TsWorksheet).PageLayout.HasHeaderFooterImages then
    AppendToStream(FSSheets[FCurSheetNum], Format(
      '<legacyDrawingHF r:id="rId%d" />', [rId_DrawingHF]));

  // Footer
  AppendToStream(FSSheets[FCurSheetNum],
    '</worksheet>');
end;

{@@ ----------------------------------------------------------------------------
  Adds the built-in number formats to the NumFormatList.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.AddBuiltinNumFormats;
begin
  FFirstNumFormatIndexInFile := 164;
  AddBuiltInBiffFormats(
    FNumFormatList, Workbook.FormatSettings, FFirstNumFormatIndexInFile-1
  );
end;

{@@ ----------------------------------------------------------------------------
  Creates the basic streams for the individual data files.
  Will be zipped into a single xlsx file.
  Other stream depending on the count of sheets will be created when needed.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.CreateStreams;
begin
  FSContentTypes := CreateTempStream(FWorkbook, 'fpsCT');
  FSRelsRels := CreateTempStream(FWorkbook, 'fpsRR');
  FSWorkbookRels := CreateTempStream(FWorkbook, 'fpsWBR');
  FSWorkbook := CreateTempStream(FWorkbook, 'fpsWB');
  FSStyles := CreateTempStream(FWorkbook, 'fpsSTY');
  FSSharedStrings := CreateTempStream(FWorkbook, 'fpsSS');
  FSSharedStrings_complete := CreateTempStream(FWorkbook, 'fpsSSC');
  FSMetaData := CreateTempStream(FWorkbook, 'fpsMETA');
  FSCustomMetaData := CreateTempStream(FWorkbook, 'fpsCM');
 {$IFDEF FPS_CHARTS}
  FChartWriter.CreateStreams;
 {$ENDIF}
  {
  if boFileStream in FWorkbook.Options then
  begin
    FSContentTypes := TFileStream.Create(GetTempFileName('', 'fpsCT'), fmCreate);
    FSRelsRels := TFileStream.Create(GetTempFileName('', 'fpsRR'), fmCreate);
    FSWorkbookRels := TFileStream.Create(GetTempFileName('', 'fpsWBR'), fmCreate);
    FSWorkbook := TFileStream.Create(GetTempFileName('', 'fpsWB'), fmCreate);
    FSStyles := TFileStream.Create(GetTempFileName('', 'fpsSTY'), fmCreate);
    FSSharedStrings := TFileStream.Create(GetTempFileName('', 'fpsSS'), fmCreate);
    FSSharedStrings_complete := TFileStream.Create(GetTempFileName('', 'fpsSSC'), fmCreate);
  end else
  if (boBufStream in Workbook.Options) then
  begin
    FSContentTypes := TBufStream.Create(GetTempFileName('', 'fpsCT'));
    FSRelsRels := TBufStream.Create(GetTempFileName('', 'fpsRR'));
    FSWorkbookRels := TBufStream.Create(GetTempFileName('', 'fpsWBR'));
    FSWorkbook := TBufStream.Create(GetTempFileName('', 'fpsWB'));
    FSStyles := TBufStream.Create(GetTempFileName('', 'fpsSTY'));
    FSSharedStrings := TBufStream.Create(GetTempFileName('', 'fpsSS'));
    FSSharedStrings_complete := TBufStream.Create(GetTempFileName('', 'fpsSSC'));
  end else begin;
    FSContentTypes := TMemoryStream.Create;
    FSRelsRels := TMemoryStream.Create;
    FSWorkbookRels := TMemoryStream.Create;
    FSWorkbook := TMemoryStream.Create;
    FSStyles := TMemoryStream.Create;
    FSSharedStrings := TMemoryStream.Create;
    FSSharedStrings_complete := TMemoryStream.Create;
  end;
  }
  // FSSheets will be created when needed.
end;

{@@ ----------------------------------------------------------------------------
  Destroys the streams that were created by the writer
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.DestroyStreams;
var
  stream: TStream;
begin
  DestroyTempStream(FSCustomMetaData);
  DestroyTempStream(FSMetaData);
  DestroyTempStream(FSContentTypes);
  DestroyTempStream(FSRelsRels);
  DestroyTempStream(FSWorkbookRels);
  DestroyTempStream(FSWorkbook);
  DestroyTempStream(FSStyles);
  DestroyTempStream(FSSharedStrings);
  DestroyTempStream(FSSharedStrings_complete);
  for stream in FSSheets do DestroyTempStream(stream);
  SetLength(FSSheets, 0);
 {$ifdef FPS_CHARTS}
  FChartWriter.DestroyStreams;
 {$endif}
  for stream in FSComments do DestroyTempStream(stream);
  SetLength(FSComments, 0);
  for stream in FSSheetRels do DestroyTempStream(stream);
  SetLength(FSSheetRels, 0);
  for stream in FSVmlDrawings do DestroyTempStream(stream);
  SetLength(FSVmlDrawings, 0);
  for stream in FSVmlDrawingsRels do DestroyTempStream(stream);
  SetLength(FSVmlDrawingsRels, 0);
  for stream in FSDrawings do DestroyTempStream(stream);
  SetLength(FSDrawings, 0);
  for stream in FSDrawingsRels do DestroyTempStream(stream);
  Setlength(FSDrawings, 0);
end;

{@@ ----------------------------------------------------------------------------
  Prepares a string formula for writing: Deletes the leading = sign and makes
  sure that it is a valid xml string.
-------------------------------------------------------------------------------}
function TsSpreadOOXMLWriter.PrepareFormula(const AFormula: String): String;
begin
  Result := AFormula;
  if (Result <> '') and (Result[1] = '=') then Delete(Result, 1, 1);
  Result := UTF8TextToXMLText(Result)
end;

{@@ ----------------------------------------------------------------------------
  Is called before zipping the individual file parts. Rewinds the streams.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.ResetStreams;
var
  i: Integer;
begin
  ResetStream(FSContentTypes);
  ResetStream(FSRelsRels);
  ResetStream(FSWorkbookRels);
  ResetStream(FSWorkbook);
  ResetStream(FSStyles);
  ResetStream(FSSharedStrings_complete);
  ResetStream(FSMetaData);
  ResetStream(FSCustomMetaData);
  for i:=0 to High(FSSheets) do ResetStream(FSSheets[i]);
  for i:=0 to High(FSSheetRels) do ResetStream(FSSheetRels[i]);
  for i:=0 to High(FSComments) do ResetStream(FSComments[i]);
  for i:=0 to High(FSVmlDrawings) do ResetStream(FSVmlDrawings[i]);
  for i:=0 to High(FSVmlDrawingsRels) do ResetStream(FSVmlDrawingsRels[i]);
  for i:=0 to High(FSDrawings) do ResetStream(FSDrawings[i]);
  for i:=0 to High(FSDrawingsRels) do ResetStream(FSDrawingsRels[i]);
end;

{@@ ----------------------------------------------------------------------------
  Writes a string to a file. Helper convenience method.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.WriteStringToFile(AFileName, AString: string);
var
  stream : TFileStream;
  S : String;
begin
  stream := TFileStream.Create(AFileName, fmCreate);
  try
    S := AString;
    stream.WriteBuffer(Pointer(S)^, Length(S));
  finally
    stream.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an OOXML document to a stream
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.WriteToStream(AStream: TStream;
  AParams: TsStreamParams = []);
var
  FZip: TZipper;
  i: Integer;
begin
  Unused(AParams);
  vmlDrawingCounter := 1;

  { Analyze the workbook and collect all information needed }
  ListAllNumFormats;
  ListAllFills;
  ListAllBorders;
  ListAllDifferentialFormats;

  { Create the streams that will hold the file contents }
  CreateStreams;

  { Fill the streams with the contents of the files }
  WriteGlobalFiles;
  WriteContent;
  WriteContentTypes;

  // Stream positions must be at beginning, they were moved to end during adding of xml strings.
  ResetStreams;

  { Now compress the files }
  FZip := TZipper.Create;
  try
    FZip.FileName := GetTempFilename;   // needed if the zipped file is too big for in-memory processing
    FZip.Entries.AddFileEntry(FSContentTypes, OOXML_PATH_TYPES);
    FZip.Entries.AddFileEntry(FSRelsRels, OOXML_PATH_RELS_RELS);
    FZip.Entries.AddFileEntry(FSWorkbookRels, OOXML_PATH_XL_RELS_RELS);
    FZip.Entries.AddFileEntry(FSWorkbook, OOXML_PATH_XL_WORKBOOK);
    FZip.Entries.AddFileEntry(FSStyles, OOXML_PATH_XL_STYLES);
    if FSSharedStrings_complete.Size > 0 then
      FZip.Entries.AddFileEntry(FSSharedStrings_complete, OOXML_PATH_XL_STRINGS);
    FZip.Entries.AddFileEntry(FSMetaData, OOXML_PATH_DOCPROPS_CORE);
    if TsWorkbook(FWorkbook).MetaData.Custom.Count > 0 then
      FZip.Entries.AddFileEntry(FSCustomMetaData, OOXML_PATH_DOCPROPS_CUSTOM);

    // Write embedded images
    WriteMedia(FZip);

    // Write worksheets
    for i:=0 to High(FSSheets) do begin
      FSSheets[i].Position:= 0;
      FZip.Entries.AddFileEntry(FSSheets[i], OOXML_PATH_XL_WORKSHEETS + Format('sheet%d.xml', [i+1]));
    end;

    // Write comments
    for i:=0 to High(FSComments) do begin
      if (FSComments[i] = nil) or (FSComments[i].Size = 0) then continue;
      FSComments[i].Position := 0;
      FZip.Entries.AddFileEntry(FSComments[i], OOXML_PATH_XL + Format('comments%d.xml', [i+1]));
    end;

    // Write worksheet relationships
    for i:=0 to High(FSSheetRels) do begin
      if (FSSheetRels[i] = nil) or (FSSheetRels[i].Size = 0) then continue;
      FSSheetRels[i].Position := 0;
      FZip.Entries.AddFileEntry(FSSheetRels[i], OOXML_PATH_XL_WORKSHEETS_RELS + Format('sheet%d.xml.rels', [i+1]));
    end;

   {$ifdef FPS_CHARTS}
   // Write charts
   TsSpreadOOXMLChartWriter(FChartWriter).AddChartsToZip(FZip);
   {$endif}

    // Write drawings
    for i:=0 to High(FSDrawings) do begin
      if (FSDrawings[i] = nil) or (FSDrawings[i].Size = 0) then continue;
      FSDrawings[i].Position := 0;
      FZip.Entries.AddFileEntry(FSDrawings[i], OOXML_PATH_XL_DRAWINGS + Format('drawing%d.xml', [i+1]));
    end;
    for i:=0 to High(FSVmlDrawings) do begin
      if (FSVmlDrawings[i] = nil) or (FSVmlDrawings[i].Size = 0) then continue;
      FSVmlDrawings[i].Position := 0;
      FZip.Entries.AddFileEntry(FSVmlDrawings[i], OOXML_PATH_XL_DRAWINGS + Format('vmlDrawing%d.vml', [i+1]));
    end;
    for i:=0 to High(FSDrawingsRels) do begin
      if (FSDrawingsRels[i] = nil) or (FSDrawingsRels[i].Size = 0) then continue;
      FSDrawingsRels[i].Position := 0;
      FZip.Entries.AddFileEntry(FSDrawingsRels[i], OOXML_PATH_XL_DRAWINGS_RELS + Format('drawing%d.xml.rels', [i+1]));
    end;
    for i:=0 to High(FSVmlDrawingsRels) do begin
      if (FSVmlDrawingsRels[i] = nil) or (FSVmlDrawingsRels[i].Size = 0) then continue;
      FSVmlDrawingsRels[i].Position := 0;
      FZip.Entries.AddFileEntry(FSVmlDrawingsRels[i], OOXML_PATH_XL_DRAWINGS_RELS + Format('vmlDrawing%d.vml.rels', [i+1]));
    end;

    FZip.SaveToStream(AStream);

  finally
    DestroyStreams;
    FZip.Free;
  end;
end;

procedure TsSpreadOOXMLWriter.WriteBlank(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
var
  cellPosText: String;
  lStyleIndex: Integer;
begin
  cellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);
  AppendToStream(AStream, Format(
    '<c r="%s" s="%d">', [CellPosText, lStyleIndex]),
      '<v></v>',
    '</c>');
end;

{@@ ----------------------------------------------------------------------------
  Writes a boolean value to the stream
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.WriteBool(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: Boolean; ACell: PCell);
var
  CellPosText: String;
  CellValueText: String;
  lStyleIndex: Integer;
begin
  CellPosText := GetCellString(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);
  CellValueText := StrUtils.IfThen(AValue, '1', '0');
  AppendToStream(AStream, Format(
    '<c r="%s" s="%d" t="b"><v>%s</v></c>', [CellPosText, lStyleIndex, CellValueText]));
end;

{@@ ----------------------------------------------------------------------------
  Writes an error value to the specified cell.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
var
  CellPosText: String;
  CellValueText: String;
  lStyleIndex: Integer;
begin
  Unused(AValue);
  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);
  CellValueText := GetErrorValueStr(ACell^.ErrorValue);
  AppendToStream(AStream, Format(
    '<c r="%s" s="%d" t="e"><v>%s</v></c>', [CellPosText, lStyleIndex, CellValueText]));
end;

{@@ ----------------------------------------------------------------------------
  Writes a string formula to the given cell.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.WriteFormula(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
var
  cellPosText: String;
  lStyleIndex: Integer;
  t, v: String;
  formula: PsFormula;
  formulaStr: String;
begin
  cellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);

  formula := TsWorksheet(FWorksheet).Formulas.FindFormula(ARow, ACol);
  formulaStr := PrepareFormula(formula^.Text);

  case ACell^.ContentType of
    cctFormula:
      begin
        t := '';
        v := '';
      end;
    cctUTF8String:
      begin
        t := ' t="str"';
        v := Format('<v>%s</v>', [UTF8TextToXMLText(ACell^.UTF8StringValue)]);
      end;
    cctNumber:
      begin
        t := '';
        v := Format('<v>%g</v>', [ACell^.NumberValue], FPointSeparatorSettings);
      end;
    cctDateTime:
      begin
        t := '';
        v := Format('<v>%g</v>', [ACell^.DateTimeValue], FPointSeparatorSettings);
      end;
    cctBool:
      begin
        t := ' t="b"';
        if ACell^.BoolValue then
          v := '<v>1</v>'
        else
          v := '<v>0</v>';
      end;
    cctError:
      begin
        t := ' t="e"';
        v := Format('<v>%s</v>', [GetErrorValueStr(ACell^.ErrorValue)]);
      end;
  end;

  AppendToStream(AStream, Format(
      '<c r="%s" s="%d"%s>' +
        '<f>%s</f>' +
        '%s' +
      '</c>', [
      CellPosText, lStyleIndex, t,
      formulaStr,
      v
  ]));
end;

{@@ ----------------------------------------------------------------------------
  Writes a string to the stream

  If the string length exceeds 32767 bytes, the string will be truncated and a
  warning will be written to the workbook's log.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
const
  MAXBYTES = 32767;  // limit for this format
var
  CellPosText: string;
  lStyleIndex: Cardinal;
  ResultingValue: string;
  fnt: TsFont;
  i, n, L: Integer;
  rtParam: TsRichTextParam;
  txt: String;
begin
  // Office 2007-2010 (at least) supports no more characters in a cell;
  if Length(AValue) > MAXBYTES then
  begin
    ResultingValue := Copy(AValue, 1, MAXBYTES); //may chop off multicodepoint UTF8 characters but well...
    Workbook.AddErrorMsg(rsTruncateTooLongCellText, [
      MAXBYTES, GetCellString(ARow, ACol)
    ]);
  end
  else
    ResultingValue := AValue;

  { Check for invalid characters }
  if not ValidXMLText(ResultingValue) then
    Workbook.AddErrorMsg(
      rsInvalidCharacterInCell, [
      GetCellString(ARow, ACol)
    ]);

  { Write string to SharedString table }
  if Length(ACell^.RichTextParams) = 0 then
    // unformatted string
    AppendToStream(FSSharedStrings,
      '<si>' +
        '<t xml:space="preserve">' + ResultingValue + '</t>' +
      '</si>')
  else
  begin
    // rich-text formatted string
    FixLineEndings(ResultingValue, ACell^.RichTextParams);
    L := UTF8Length(Resultingvalue);
    AppendToStream(FSSharedStrings,
      '<si>');
    rtParam := ACell^.RichTextParams[0];
    if rtParam.FirstIndex > 1 then
    begin
      // Unformatted part before first format
      txt := UTF8Copy(ResultingValue, 1, rtParam.FirstIndex - 1);
      ValidXMLText(txt);
      AppendToStream(FSSharedStrings,
        '<r>' +
          '<t xml:space="preserve">' + txt + '</t>' +
        '</r>'
      );
    end;
    for i := 0 to High(ACell^.RichTextParams) do
    begin
      rtParam := ACell^.RichTextParams[i];
      fnt := (FWorkbook as TsWorkbook).GetFont(rtParam.FontIndex);
      // Calculate count of characters in this format section
      if i = High(ACell^.RichTextParams) then
        n := L - rtParam.FirstIndex + 1 else
        n := ACell^.RichTextParams[i+1].FirstIndex - rtParam.FirstIndex;
      // Partial string having this format
      txt := UTF8Copy(Resultingvalue, rtParam.FirstIndex, n);
      ValidXMLText(txt);
      AppendToStream(FSSharedStrings,
        '<r>');
      WriteFont(FSSharedStrings, fnt, false);  // <rPr> ... font data ... </rPr>
      AppendToStream(FSSharedStrings,
          '<t xml:space="preserve">' + txt + '</t>' +
        '</r>'
      );
    end;
    AppendToStream(FSSharedStrings,
      '</si>');
  end;

  { Write shared string index to cell record }
  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);
  AppendToStream(AStream, Format(
    '<c r="%s" s="%d" t="s">'+
      '<v>%d</v>'+
    '</c>',
    [CellPosText, lStyleIndex, FSharedStringsCount]
  ));
  inc(FSharedStringsCount);
end;

{@@ ----------------------------------------------------------------------------
  Writes a number (64-bit IEE 754 floating point) to the stream
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double; ACell: PCell);
var
  CellPosText: String;
  CellValueText: String;
  lStyleIndex: Integer;
begin
  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);
  CellValueText := FloatToStr(AValue, FPointSeparatorSettings);
  AppendToStream(AStream, Format(
    '<c r="%s" s="%d" t="n"><v>%s</v></c>', [CellPosText, lStyleIndex, CellValueText]));
end;

{@@ ----------------------------------------------------------------------------
  Writes a date/time value as a number

  Respects DateMode of the file
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLWriter.WriteDateTime(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell);
var
  ExcelDateSerial: double;
  cf: TsCellFormat;
  nfp: TsNumFormatParams;
begin
  // We must correct the bug of Lotus 1-2-3 which had ignored that year 1900 was
  // a leap year, but only for "normal" date format, not for time-interval formats
  cf := TsWorksheet(FWorksheet).ReadCellFormat(ACell);
  nfp := TsWorkbook(FWorkbook).GetNumberFormat(cf.NumberFormatIndex);
  if IsTimeIntervalFormat(nfp) then
    ExcelDateSerial := AValue
  else
    ExcelDateSerial := ConvertDateTimeToExcelDateTime(AValue, FDateMode);
  WriteNumber(AStream, ARow, ACol, ExcelDateSerial, ACell);
end;


initialization

  // Registers this reader / writer on fpSpreadsheet
  sfidOOXML := RegisterSpreadFormat(sfOOXML,
    TsSpreadOOXMLReader, TsSpreadOOXMLWriter,
    STR_FILEFORMAT_EXCEL_XLSX, 'OOXML', [STR_OOXML_EXCEL_EXTENSION, '.xlsm']
  );

end.

