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

{$ifdef fpc}
  {$mode delphi}
{$endif}

interface

uses
  Classes, SysUtils,
  laz2_xmlread, laz2_DOM,
  AVL_Tree,
 {$IF FPC_FULLVERSION >= 20701}
  zipper,
 {$ELSE}
  fpszipper,
 {$ENDIF}
  fpsTypes, fpSpreadsheet, fpsUtils, fpsReaderWriter, fpsNumFormat,
  fpsxmlcommon, xlsCommon;
  
type

  { TsOOXMLFormatList }
  TsOOXMLNumFormatList = class(TsCustomNumFormatList)
  protected
    procedure AddBuiltinFormats; override;
  public
    procedure ConvertBeforeWriting(var AFormatString: String;
      var ANumFormat: TsNumberFormat); override;
  end;

  { TsSpreadOOXMLReader }

  TsSpreadOOXMLReader = class(TsSpreadXMLReader)
  private
    FDateMode: TDateMode;
    FPointSeparatorSettings: TFormatSettings;
    FSharedStrings: TStringList;
    FFillList: TFPList;
    FBorderList: TFPList;
    FHyperlinkList: TFPList;
    FThemeColors: array of TsColorValue;
    FWrittenByFPS: Boolean;
    procedure ApplyCellFormatting(ACell: PCell; XfIndex: Integer);
    procedure ApplyHyperlinks(AWorksheet: TsWorksheet);
    function FindCommentsFileName(ANode: TDOMNode): String;
    procedure ReadBorders(ANode: TDOMNode);
    procedure ReadCell(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadCellXfs(ANode: TDOMNode);
    function  ReadColor(ANode: TDOMNode): TsColor;
    procedure ReadCols(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadComments(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadDateMode(ANode: TDOMNode);
    procedure ReadFileVersion(ANode: TDOMNode);
    procedure ReadFills(ANode: TDOMNode);
    procedure ReadFont(ANode: TDOMNode);
    procedure ReadFonts(ANode: TDOMNode);
    procedure ReadHyperlinks(ANode: TDOMNode);
    procedure ReadMergedCells(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadNumFormats(ANode: TDOMNode);
    procedure ReadPalette(ANode: TDOMNode);
    procedure ReadRowHeight(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadSharedStrings(ANode: TDOMNode);
    procedure ReadSheetFormatPr(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadSheetList(ANode: TDOMNode; AList: TStrings);
    procedure ReadSheetViews(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadThemeElements(ANode: TDOMNode);
    procedure ReadThemeColors(ANode: TDOMNode);
    procedure ReadWorksheet(ANode: TDOMNode; AWorksheet: TsWorksheet);
  protected
    procedure CreateNumFormatList; override;
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;
    procedure ReadFromFile(AFileName: string); override;
    procedure ReadFromStream(AStream: TStream); override;
  end;

  { TsSpreadOOXMLWriter }

  TsSpreadOOXMLWriter = class(TsCustomSpreadWriter)
  private
    FNext_rId: Integer;
    procedure WriteVmlDrawingsCallback(AComment: PsComment;
      ACommentIndex: Integer; AStream: TStream);

  protected
    FDateMode: TDateMode;
    FPointSeparatorSettings: TFormatSettings;
    FSharedStringsCount: Integer;
    FFillList: array of PsCellFormat;
    FBorderList: array of PsCellFormat;
  protected
    { Helper routines }
    procedure CreateNumFormatList; override;
    procedure CreateStreams;
    procedure DestroyStreams;
    function  FindBorderInList(AFormat: PsCellFormat): Integer;
    function  FindFillInList(AFormat: PsCellFormat): Integer;
    function GetStyleIndex(ACell: PCell): Cardinal;
    procedure ListAllBorders;
    procedure ListAllFills;
    function  PrepareFormula(const AFormula: String): String;
    procedure ResetStreams;
    procedure WriteBorderList(AStream: TStream);
    procedure WriteCols(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteComments(AWorksheet: TsWorksheet);
    procedure WriteDimension(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteFillList(AStream: TStream);
    procedure WriteFontList(AStream: TStream);
    procedure WriteHyperlinks(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteMergedCells(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteNumFormatList(AStream: TStream);
    procedure WritePalette(AStream: TStream);
    procedure WriteSheetData(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteSheetViews(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteStyleList(AStream: TStream; ANodeName: String);
    procedure WriteVmlDrawings(AWorksheet: TsWorksheet);
    procedure WriteWorksheet(AWorksheet: TsWorksheet);
    procedure WriteWorksheetRels(AWorksheet: TsWorksheet);
  protected
    { Streams with the contents of files }
    FSContentTypes: TStream;
    FSRelsRels: TStream;
    FSWorkbook: TStream;
    FSWorkbookRels: TStream;
    FSStyles: TStream;
    FSSharedStrings: TStream;
    FSSharedStrings_complete: TStream;
    FSSheets: array of TStream;
    FSSheetRels: array of TStream;
    FSComments: array of TStream;
    FSVmlDrawings: array of TStream;
    FCurSheetNum: Integer;
  protected
    { Routines to write the files }
    procedure WriteContent;
    procedure WriteContentTypes;
    procedure WriteGlobalFiles;
  protected
    { Record writing methods }
    //todo: add WriteDate
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
    constructor Create(AWorkbook: TsWorkbook); override;
    { General writing methods }
    procedure WriteStringToFile(AFileName, AString: string);
    procedure WriteToFile(const AFileName: string; const AOverwriteExisting: Boolean = False); override;
    procedure WriteToStream(AStream: TStream); override;
  end;


  TXlsxSettings = record
    DateMode: TDateMode;
  end;

var
  XlsxSettings: TXlsxSettings = (
    DateMode: dm1900;
  );


implementation

uses
  variants, fileutil, strutils, math, lazutf8, uriparser,
  {%H-}fpsPatches, fpsStrings, fpsStreams, fpsNumFormatParser;

const
  { OOXML general XML constants }
     XML_HEADER               = '<?xml version="1.0" encoding="utf-8" ?>';

  { OOXML Directory structure constants }
  // Note: directory separators are always / because the .xlsx is a zip file which
  // requires / instead of \, even on Windows; see 
  // http://www.pkware.com/documents/casestudies/APPNOTE.TXT
  // 4.4.17.1 All slashes MUST be forward slashes '/' as opposed to backwards slashes '\'
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
     OOXML_PATH_XL_DRAWINGS        = 'xl/drawings/';
     OOXML_PATH_XL_THEME           = 'xl/theme/theme1.xml';

     { OOXML schemas constants }
     SCHEMAS_TYPES        = 'http://schemas.openxmlformats.org/package/2006/content-types';
     SCHEMAS_RELS         = 'http://schemas.openxmlformats.org/package/2006/relationships';
     SCHEMAS_DOC_RELS     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
     SCHEMAS_DOCUMENT     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
     SCHEMAS_WORKSHEET    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
     SCHEMAS_STYLES       = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
     SCHEMAS_STRINGS      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
     SCHEMAS_COMMENTS     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
     SCHEMAS_DRAWINGS     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
     SCHEMAS_HYPERLINKS   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
     SCHEMAS_SPREADML     = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

     { OOXML mime types constants }
{%H-}MIME_XML             = 'application/xml';
     MIME_RELS            = 'application/vnd.openxmlformats-package.relationships+xml';
     MIME_OFFICEDOCUMENT  = 'application/vnd.openxmlformats-officedocument';
     MIME_SPREADML        = MIME_OFFICEDOCUMENT + '.spreadsheetml';
     MIME_SHEET           = MIME_SPREADML + '.sheet.main+xml';
     MIME_WORKSHEET       = MIME_SPREADML + '.worksheet+xml';
     MIME_STYLES          = MIME_SPREADML + '.styles+xml';
     MIME_STRINGS         = MIME_SPREADML + '.sharedStrings+xml';
     MIME_COMMENTS        = MIME_SPREADML + '.comments+xml';
     MIME_VMLDRAWING      = MIME_OFFICEDOCUMENT + '.vmlDrawing';

     LAST_PALETTE_COLOR   = $3F;  // 63

var
  // the palette of the 64 default colors as "big-endian color" values
  // (identical to BIFF8)
  PALETTE_OOXML: array[$00..LAST_PALETTE_COLOR] of TsColorValue = (
    $000000,  // $00: black            // 8 built-in default colors
    $FFFFFF,  // $01: white
    $FF0000,  // $02: red
    $00FF00,  // $03: green
    $0000FF,  // $04: blue
    $FFFF00,  // $05: yellow
    $FF00FF,  // $06: magenta
    $00FFFF,  // $07: cyan

    $000000,  // $08: EGA black
    $FFFFFF,  // $09: EGA white
    $FF0000,  // $0A: EGA red
    $00FF00,  // $0B: EGA green
    $0000FF,  // $0C: EGA blue
    $FFFF00,  // $0D: EGA yellow
    $FF00FF,  // $0E: EGA magenta
    $00FFFF,  // $0F: EGA cyan

    $800000,  // $10: EGA dark red
    $008000,  // $11: EGA dark green
    $000080,  // $12: EGA dark blue
    $808000,  // $13: EGA olive
    $800080,  // $14: EGA purple
    $008080,  // $15: EGA teal
    $C0C0C0,  // $16: EGA silver
    $808080,  // $17: EGA gray
    $9999FF,  // $18:
    $993366,  // $19:
    $FFFFCC,  // $1A:
    $CCFFFF,  // $1B:
    $660066,  // $1C:
    $FF8080,  // $1D:
    $0066CC,  // $1E:
    $CCCCFF,  // $1F:

    $000080,  // $20:
    $FF00FF,  // $21:
    $FFFF00,  // $22:
    $00FFFF,  // $23:
    $800080,  // $24:
    $800000,  // $25:
    $008080,  // $26:
    $0000FF,  // $27:
    $00CCFF,  // $28:
    $CCFFFF,  // $29:
    $CCFFCC,  // $2A:
    $FFFF99,  // $2B:
    $99CCFF,  // $2C:
    $FF99CC,  // $2D:
    $CC99FF,  // $2E:
    $FFCC99,  // $2F:

    $3366FF,  // $30:
    $33CCCC,  // $31:
    $99CC00,  // $32:
    $FFCC00,  // $33:
    $FF9900,  // $34:
    $FF6600,  // $35:
    $666699,  // $36:
    $969696,  // $37:
    $003366,  // $38:
    $339966,  // $39:
    $003300,  // $3A:
    $333300,  // $3B:
    $993300,  // $3C:
    $993366,  // $3D:
    $333399,  // $3E:
    $333333   // $3F:
  );

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

  THyperlinkListData = class
    ID: String;
    CellRef: String;
    Target: String;
    TextMark: String;
    Display: String;
    Tooltip: String;
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
    'darkTellis',      // fsHatchTickDiag
    'lightGrid'        // fsHatchThinHor
    );




{ TsOOXMLNumFormatList }

{ These are the built-in number formats as expected in the biff spreadsheet file.
  Identical to BIFF8. These formats are not written to file but they are used
  for lookup of the number format that Excel used. They are specified here in
  fpc dialect. }
procedure TsOOXMLNumFormatList.AddBuiltinFormats;
var
  fs: TFormatSettings;
  cs: String;
begin
  fs := Workbook.FormatSettings;
  cs := AnsiToUTF8(Workbook.FormatSettings.CurrencyString);

  AddFormat( 0, nfGeneral, '');
  AddFormat( 1, nfFixed, '0');
  AddFormat( 2, nfFixed, '0.00');
  AddFormat( 3, nfFixedTh, '#,##0');
  AddFormat( 4, nfFixedTh, '#,##0.00');
  AddFormat( 5, nfCurrency, '"'+cs+'"#,##0_);("'+cs+'"#,##0)');
  AddFormat( 6, nfCurrencyRed, '"'+cs+'"#,##0_);[Red]("'+cs+'"#,##0)');
  AddFormat( 7, nfCurrency, '"'+cs+'"#,##0.00_);("'+cs+'"#,##0.00)');
  AddFormat( 8, nfCurrencyRed, '"'+cs+'"#,##0.00_);[Red]("'+cs+'"#,##0.00)');
  AddFormat( 9, nfPercentage, '0%');
  AddFormat(10, nfPercentage, '0.00%');
  AddFormat(11, nfExp, '0.00E+00');
  // fraction formats 12 ('# ?/?') and 13 ('# ??/??') not supported
  AddFormat(14, nfShortDate, fs.ShortDateFormat);                       // 'M/D/YY'
  AddFormat(15, nfLongDate, fs.LongDateFormat);                         // 'D-MMM-YY'
  AddFormat(16, nfCustom, 'd/mmm');                                     // 'D-MMM'
  AddFormat(17, nfCustom, 'mmm/yy');                                    // 'MMM-YY'
  AddFormat(18, nfShortTimeAM, AddAMPM(fs.ShortTimeFormat, fs));        // 'h:mm AM/PM'
  AddFormat(19, nfLongTimeAM, AddAMPM(fs.LongTimeFormat, fs));          // 'h:mm:ss AM/PM'
  AddFormat(20, nfShortTime, fs.ShortTimeFormat);                       // 'h:mm'
  AddFormat(21, nfLongTime, fs.LongTimeFormat);                         // 'h:mm:ss'
  AddFormat(22, nfShortDateTime, fs.ShortDateFormat + ' ' + fs.ShortTimeFormat);  // 'M/D/YY h:mm' (localized)
  // 23..36 not supported
  AddFormat(37, nfCurrency, '_(#,##0_);(#,##0)');
  AddFormat(38, nfCurrencyRed, '_(#,##0_);[Red](#,##0)');
  AddFormat(39, nfCurrency, '_(#,##0.00_);(#,##0.00)');
  AddFormat(40, nfCurrencyRed, '_(#,##0.00_);[Red](#,##0.00)');
  AddFormat(41, nfCustom, '_("'+cs+'"* #,##0_);_("'+cs+'"* (#,##0);_("'+cs+'"* "-"_);_(@_)');
  AddFormat(42, nfCustom, '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)');
  AddFormat(43, nfCustom, '_("'+cs+'"* #,##0.00_);_("'+cs+'"* (#,##0.00);_("'+cs+'"* "-"??_);_(@_)');
  AddFormat(44, nfCustom, '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)');
  AddFormat(45, nfCustom, 'nn:ss');
  AddFormat(46, nfTimeInterval, '[h]:nn:ss');
  AddFormat(47, nfCustom, 'nn:ss.z');
  AddFormat(48, nfCustom, '##0.0E+00');
  // 49 ("Text") not supported

  // All indexes from 0 to 163 are reserved for built-in formats.
  // The first user-defined format starts at 164.
  FFirstNumFormatIndexInFile := 164;
  FNextNumFormatIndex := 164;
end;

procedure TsOOXMLNumFormatList.ConvertBeforeWriting(var AFormatString: String;
  var ANumFormat: TsNumberFormat);
var
  parser: TsNumFormatParser;
begin
  parser := TsNumFormatParser.Create(Workbook, AFormatString, ANumFormat);
  try
    if parser.Status = psOK then begin
      // For writing, we have to convert the fpc format string to Excel dialect
      AFormatString := parser.FormatString[nfdExcel];
      ANumFormat := parser.NumFormat;
    end;
  finally
    parser.Free;
  end;
end;


{ TsSpreadOOXMLReader }

constructor TsSpreadOOXMLReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FDateMode := XlsxSettings.DateMode;
  // Set up the default palette in order to have the default color names correct.
  Workbook.UseDefaultPalette;

  FSharedStrings := TStringList.Create;
  FFillList := TFPList.Create;
  FBorderList := TFPList.Create;
  FHyperlinkList := TFPList.Create;
  FCellFormatList := TsCellFormatList.Create(true);
  // Allow duplicates because xf indexes used in cell records cannot be found any more.

  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';
end;

destructor TsSpreadOOXMLReader.Destroy;
var
  j: Integer;
begin
  for j := FFillList.Count-1 downto 0 do TObject(FFillList[j]).Free;
  FFillList.Free;

  for j := FBorderList.Count-1 downto 0 do TObject(FBorderList[j]).Free;
  FBorderList.Free;

  for j := FHyperlinkList.Count-1 downto 0 do TObject(FHyperlinkList[j]).Free;
  FHyperlinkList.Free;

  FSharedStrings.Free;

  // FCellFormatList and FFontList are destroyed by ancestor

  inherited Destroy;
end;

procedure TsSpreadOOXMLReader.ApplyCellFormatting(ACell: PCell; XFIndex: Integer);
var
  i: Integer;
  fmt: PsCellFormat;
begin
  if Assigned(ACell) then begin
    i := FCellFormatList.FindIndexOfID(XFIndex);
    fmt := FCellFormatList.Items[i];
    ACell^.FormatIndex := FWorkbook.AddCellFormat(fmt^);
  end;
end;

procedure TsSpreadOOXMLReader.ApplyHyperlinks(AWorksheet: TsWorksheet);
var
  i: Integer;
  hyperlinkData: THyperlinkListData;
  r1, c1, r2, c2, r, c: Cardinal;
begin
  for i:=0 to FHyperlinkList.Count-1 do
  begin
    hyperlinkData := THyperlinkListData(FHyperlinkList.Items[i]);
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
            AWorksheet.WriteHyperlink(r, c, '#'+TextMark, ToolTip)
          else
          if TextMark = '' then
            AWorksheet.WriteHyperlink(r, c, Target, ToolTip)
          else
            AWorksheet.WriteHyperlink(r, c, Target+'#'+TextMark, ToolTip);
  end;
end;

function TsSpreadOOXMLReader.FindCommentsFileName(ANode: TDOMNode): String;
var
  s: String;
begin
  while ANode <> nil do
  begin
    s := GetAttrValue(ANode, 'Type');
    if s = SCHEMAS_COMMENTS then
    begin
      Result := ExtractFileName(GetAttrValue(ANode, 'Target'));
      exit;
    end;
    ANode := ANode.NextSibling;
  end;
  Result := '';
end;

procedure TsSpreadOOXMLReader.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsOOXMLNumFormatList.Create(Workbook);
end;

procedure TsSpreadOOXMLReader.ReadBorders(ANode: TDOMNode);

  function ReadBorderStyle(ANode: TDOMNode; out ABorderStyle: TsCellBorderStyle): Boolean;
  var
    s: String;
    colorNode: TDOMNode;
    nodeName: String;
  begin
    Result := false;
    ABorderStyle.LineStyle := lsThin;
    ABorderStyle.Color := scBlack;

    s := GetAttrValue(ANode, 'style');
    if s = '' then
      exit;

    if s = 'thin' then
      ABorderStyle.LineStyle := lsThin
    else if s = 'medium' then
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
      ABorderStyle.LineStyle := lsHair;

    colorNode := ANode.FirstChild;
    while Assigned(colorNode) do begin
      nodeName := colorNode.NodeName;
      if nodeName = 'color' then
        ABorderStyle.Color := ReadColor(colorNode);
      colorNode := colorNode.NextSibling;
    end;
    Result := true;
  end;

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
    if nodeName = 'border' then begin
      borders := [];
      s := GetAttrValue(borderNode, 'diagonalUp');
      if s = '1' then
        Include(borders, cbDiagUp);
      s := GetAttrValue(borderNode, 'diagonalDown');
      if s = '1' then
        Include(borders, cbDiagDown);
      edgeNode := borderNode.FirstChild;
      while Assigned(edgeNode) do begin
        nodeName := edgeNode.NodeName;
        if nodeName = 'left' then begin
          if ReadBorderStyle(edgeNode, borderStyles[cbWest]) then
            Include(borders, cbWest);
        end
        else if nodeName = 'right' then begin
          if ReadBorderStyle(edgeNode, borderStyles[cbEast]) then
            Include(borders, cbEast);
        end
        else if nodeName = 'top' then begin
          if ReadBorderStyle(edgeNode, borderStyles[cbNorth]) then
            Include(borders, cbNorth);
        end
        else if nodeName = 'bottom' then begin
          if ReadBorderStyle(edgeNode, borderStyles[cbSouth]) then
            Include(borders, cbSouth);
        end
        else if nodeName = 'diagonal' then begin
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

procedure TsSpreadOOXMLReader.ReadCell(ANode: TDOMNode; AWorksheet: TsWorksheet);
var
  addr, s: String;
  rowIndex, colIndex: Cardinal;
  cell: PCell;
  datanode: TDOMNode;
  dataStr: String;
  formulaStr: String;
  sstIndex: Integer;
  number: Double;
  fmt: TsCellFormat;
  rng: TsCellRange;
  r,c: Cardinal;
begin
  if ANode = nil then
    exit;

  // get row and column address
  addr := GetAttrValue(ANode, 'r');       // cell address, like 'A1'
  ParseCellString(addr, rowIndex, colIndex);

  // create cell
  if FIsVirtualMode then
  begin
    InitCell(rowIndex, colIndex, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := AWorksheet.GetCell(rowIndex, colIndex);

  // get style index
  s := GetAttrValue(ANode, 's');
  if s <> '' then begin
    ApplyCellFormatting(cell, StrToInt(s));
    fmt := Workbook.GetCellFormat(cell^.FormatIndex);
  end else
    InitFormatRecord(fmt);

  // get data
  datanode := ANode.FirstChild;
  dataStr := '';
  formulaStr := '';
  while Assigned(datanode) do
  begin
    if datanode.NodeName = 'v' then
      dataStr := GetNodeValue(datanode)
    else
    if (boReadFormulas in FWorkbook.Options) and (datanode.NodeName = 'f') then
    begin
      // Formula to cell
      formulaStr := GetNodeValue(datanode);

      s := GetAttrValue(datanode, 't');
      if s = 'shared' then
      begin
        // Shared formula
        s := GetAttrValue(datanode, 'ref');
        if (s <> '') then      // This is the shared formula range
        begin
          // Split shared formula into single-cell formulas
          ParseCellRangeString(s, rng);
          for r := rng.Row1 to rng.Row2 do
            for c := rng.Col1 to rng.Col2 do
              FWorksheet.CopyFormula(cell, r, c);
        end;
      end
      else
        // "Normal" formula
        AWorksheet.WriteFormula(cell, formulaStr);
    end;
    datanode := datanode.NextSibling;
  end;

  // get data type
  s := GetAttrValue(ANode, 't');   // "t" = data type
  if (s = '') and (dataStr = '') then
    AWorksheet.WriteBlank(cell)
  else
  if (s = '') or (s = 'n') then begin
    // Number or date/time, depending on format
    number := StrToFloat(dataStr, FPointSeparatorSettings);
    if IsDateTimeFormat(fmt.NumberFormatStr) then begin
      if fmt.NumberFormat <> nfTimeInterval then   // no correction of time origin for "time interval" format
        number := ConvertExcelDateTimeToDateTime(number, FDateMode);
      AWorksheet.WriteDateTime(cell, number, fmt.NumberFormatStr)
    end
    else
      AWorksheet.WriteNumber(cell, number);
  end
  else
  if s = 's' then begin
    // String from shared strings table
    sstIndex := StrToInt(dataStr);
    AWorksheet.WriteUTF8Text(cell, FSharedStrings[sstIndex]);
  end else
  if s = 'str' then
    // literal string
    AWorksheet.WriteUTF8Text(cell, datastr)
  else
  if s = 'b' then
    // boolean
    AWorksheet.WriteBoolValue(cell, dataStr='1')
  else
  if s = 'e' then begin
    // error value
    if dataStr = '#NULL!' then
      AWorksheet.WriteErrorValue(cell, errEmptyIntersection)
    else if dataStr = '#DIV/0!' then
      AWorksheet.WriteErrorValue(cell, errDivideByZero)
    else if dataStr = '#VALUE!' then
      AWorksheet.WriteErrorValue(cell, errWrongType)
    else if dataStr = '#REF!' then
      AWorksheet.WriteErrorValue(cell, errIllegalRef)
    else if dataStr = '#NAME?' then
      AWorksheet.WriteErrorValue(cell, errWrongName)
    else if dataStr = '#NUM!' then
      AWorksheet.WriteErrorValue(cell, errOverflow)
    else if dataStr = '#N/A' then
      AWorksheet.WriteErrorValue(cell, errArgError)
    else
      raise Exception.Create(rsUnknownErrorType);
  end else
    raise Exception.Create(rsUnknownDataType);

  if FIsVirtualMode then
    Workbook.OnReadCellData(Workbook, rowIndex, colIndex, cell);
end;

procedure TsSpreadOOXMLReader.ReadCellXfs(ANode: TDOMNode);
var
  node: TDOMNode;
  childNode: TDOMNode;
  nodeName: String;
  fmt: TsCellFormat;
  fs: TsFillStyle;
  s1, s2: String;
  i, numFmtIndex, fillIndex, borderIndex: Integer;
  numFmtData: TsNumFormatData;
  fillData: TFillListData;
  borderData: TBorderListData;
  fnt: TsFont;
begin
  node := ANode.FirstChild;
  while Assigned(node) do
  begin
    nodeName := node.NodeName;
    if nodeName = 'xf' then
    begin
      InitFormatRecord(fmt);
      fmt.ID := FCellFormatList.Count;
      fmt.Name := '';

      // strange: sometimes the "apply*" are missing. Therefore, it may be better
      // to check against "<>0" instead of "=1"
      s1 := GetAttrValue(node, 'numFmtId');
      s2 := GetAttrValue(node, 'applyNumberFormat');
      if (s1 <> '') and (s2 <> '0') then
      begin
        numFmtIndex := StrToInt(s1);
        i := NumFormatList.FindByIndex(numFmtIndex);
        if i > -1 then
        begin
          numFmtData := NumFormatList.Items[i];
          fmt.NumberFormat := numFmtData.NumFormat;
          fmt.NumberFormatStr := numFmtData.FormatString;
          if numFmtData.NumFormat <> nfGeneral then
            Include(fmt.UsedFormattingFields, uffNumberFormat);
        end;
      end;

      s1 := GetAttrValue(node, 'fontId');
      s2 := GetAttrValue(node, 'applyFont');
      if (s1 <> '') and (s2 <> '0') then
      begin
        fnt := TsFont(FFontList.Items[StrToInt(s1)]);
        fmt.FontIndex := Workbook.FindFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color);
        if fmt.FontIndex = -1 then
          fmt.FontIndex := Workbook.AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color);
        {
        if fmt.FontIndex = BOLD_FONTINDEX then
          Include(fmt.UsedFormattingFields, uffBold)
        else }
        if fmt.FontIndex > 0 then
          Include(fmt.UsedFormattingFields, uffFont);
      end;

      s1 := GetAttrValue(node, 'fillId');
      s2 := GetAttrValue(node, 'applyFill');
      if (s1 <> '') and (s2 <> '0') then
      begin
        fillIndex := StrToInt(s1);
        fillData := FFillList[fillIndex];
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
      s2 := GetAttrValue(node, 'applyBorder');
      if (s1 <> '') and (s2 <> '0') then
      begin
        borderIndex := StrToInt(s1);
        borderData := FBorderList[borderIndex];
        if (borderData <> nil) then
        begin
          fmt.BorderStyles := borderData.BorderStyles;
          fmt.Border := borderData.Borders;
        end;
      end;

      s2 := GetAttrValue(node, 'applyAlignment');
      if (s2 <> '0') and (s2 <> '') then begin
        childNode := node.FirstChild;
        while Assigned(childNode) do begin
          nodeName := childNode.NodeName;
          if nodeName = 'alignment' then begin
            s1 := GetAttrValue(childNode, 'horizontal');
            if s1 = 'left' then
              fmt.HorAlignment := haLeft
            else
            if s1 = 'center' then
              fmt.HorAlignment := haCenter
            else
            if s1 = 'right' then
              fmt.HorAlignment := haRight;

            s1 := GetAttrValue(childNode, 'vertical');
            if s1 = 'top' then
              fmt.VertAlignment := vaTop
            else
            if s1 = 'center' then
              fmt.VertAlignment := vaCenter
            else
            if s1 = 'bottom' then
              fmt.VertAlignment := vaBottom;

            s1 := GetAttrValue(childNode, 'wrapText');
            if (s1 <> '0') then
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
      FCellFormatList.Add(fmt);
    end;
    node := node.NextSibling;
  end;
end;

function TsSpreadOOXMLReader.ReadColor(ANode: TDOMNode): TsColor;
var
  s: String;
  rgb: TsColorValue;
  idx: Integer;
  tint: Double;
  n: Integer;
begin
  Assert(ANode <> nil);

  s := GetAttrValue(ANode, 'auto');
  if s = '1' then begin
    if ANode.NodeName = 'fgColor' then
      Result := scBlack
    else
      Result := scTransparent;
    exit;
  end;

  s := GetAttrValue(ANode, 'rgb');
  if s <> '' then begin
    Result := FWorkbook.AddColorToPalette(HTMLColorStrToColor('#' + s));
    exit;
  end;

  s := GetAttrValue(ANode, 'indexed');
  if s <> '' then begin
    Result := StrToInt(s);
    n := FWorkbook.GetPaletteSize;
    if (Result <= LAST_PALETTE_COLOR) and (Result < n) then
      exit;
    // System colors
    // taken from OpenOffice docs
    case Result of
      $0040: Result := scBlack;  // Default border color
      $0041: Result := scWhite;  // Default background color
      $0043: Result := scGray;   // Dialog background color
      $004D: Result := scBlack;  // Text color, chart border lines
      $004E: Result := scGray;   // Background color for chart areas
      $004F: Result := scBlack;  // Automatic color for chart border lines
      $0050: Result := scBlack;  // ???
      $0051: Result := scBlack;  // ??
      $7FFF: Result := scBlack;  // ??
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
      Result := FWorkBook.AddColorToPalette(rgb);
      exit;
    end;
  end;

  Result := scBlack;
end;

procedure TsSpreadOOXMLReader.ReadCols(ANode: TDOMNode; AWorksheet: TsWorksheet);
const
  EPS = 1e-2;
var
  colNode: TDOMNode;
  col, col1, col2: Cardinal;
  w: Double;
  s: String;
begin
  if ANode = nil then
    exit;

  colNode := ANode.FirstChild;
  while Assigned(colNode) do begin
    s := GetAttrValue(colNode, 'customWidth');
    if s = '1' then begin
      s := GetAttrValue(colNode, 'min');
      if s <> '' then col1 := StrToInt(s)-1 else col1 := 0;
      s := GetAttrValue(colNode, 'max');
      if s <> '' then col2 := StrToInt(s)-1 else col2 := col1;
      s := GetAttrValue(colNode, 'width');
      if (s <> '') and TryStrToFloat(s, w, FPointSeparatorSettings) then
        if not SameValue(w, AWorksheet.DefaultColWidth, EPS) then
          for col := col1 to col2 do
            AWorksheet.WriteColWidth(col, w);
    end;
    colNode := colNode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadComments(ANode: TDOMNode;
  AWorksheet: TsWorksheet);
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
            if nodename = 't' then begin
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
          AWorksheet.WriteComment(r, c, comment);
        end;
        txtNode := txtNode.NextSibling;
      end;
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadDateMode(ANode: TDOMNode);
var
  s: String;
begin
  if Assigned(ANode) then begin
    s := GetAttrValue(ANode, 'date1904');
    if s = '1' then FDateMode := dm1904
  end;
end;

procedure TsSpreadOOXMLReader.ReadFileVersion(ANode: TDOMNode);
begin
  FWrittenByFPS := GetAttrValue(ANode, 'appName') = 'fpspreadsheet';
end;

procedure TsSpreadOOXMLReader.ReadFills(ANode: TDOMNode);
var
  fillNode, patternNode, colorNode: TDOMNode;
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
      if nodename = 'patternFill' then begin
        patt := GetAttrValue(patternNode, 'patternType');
        fgclr := scWhite;
        bgclr := scBlack;
        colorNode := patternNode.FirstChild;
        while Assigned(colorNode) do begin
          nodeName := colorNode.NodeName;
          if nodeName = 'fgColor' then
            fgclr := ReadColor(colorNode)
          else
          if nodeName = 'bgColor' then
            bgclr := ReadColor(colorNode);
          colorNode := colorNode.NextSibling;
        end;

        // Store in FFillList
        fillData := TFillListData.Create;
        fillData.PatternType := patt;
        fillData.FgColor := fgclr;
        fillData.BgColor := bgclr;
        FFillList.Add(fillData);
      end;
      patternNode := patternNode.NextSibling;
    end;
    fillNode := fillNode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadFont(ANode: TDOMNode);
var
  node: TDOMNode;
  fnt: TsFont;
  fntName: String;
  fntSize: Single;
  fntStyles: TsFontStyles;
  fntColor: TsColor;
  nodename: String;
  s: String;
begin
  fnt := Workbook.GetDefaultFont;
  if fnt <> nil then begin
    fntName := fnt.FontName;
    fntSize := fnt.Size;
    fntStyles := fnt.Style;
    fntColor := fnt.Color;
  end else begin
    fntName := DEFAULT_FONTNAME;
    fntSize := DEFAULT_FONTSIZE;
    fntStyles := [];
    fntColor := scBlack;
  end;

  node := ANode.FirstChild;
  while node <> nil do begin
    nodename := node.NodeName;
    if nodename = 'name' then begin
      s := GetAttrValue(node, 'val');
      if s <> '' then fntName := s;
    end
    else
    if nodename = 'sz' then begin
      s := GetAttrValue(node, 'val');
      if s <> '' then fntSize := StrToFloat(s);
    end
    else
    if nodename = 'b' then begin
      if GetAttrValue(node, 'val') <> 'false'
        then fntStyles := fntStyles + [fssBold];
    end
    else
    if nodename = 'i' then begin
      if GetAttrValue(node, 'val') <> 'false'
        then fntStyles := fntStyles + [fssItalic];
    end
    else
    if nodename = 'u' then begin
      if GetAttrValue(node, 'val') <> 'false'
        then fntStyles := fntStyles+ [fssUnderline]
    end
    else
    if nodename = 'strike' then begin
      if GetAttrValue(node, 'val') <> 'false'
        then fntStyles := fntStyles + [fssStrikeout];
    end
    else
    if nodename = 'color' then
      fntColor := ReadColor(node);
    node := node.NextSibling;
  end;

  fnt := TsFont.Create;
  fnt.FontName := fntName;
  fnt.Size := fntSize;
  fnt.Style := fntStyles;
  fnt.Color := fntColor;

  FFontList.Add(fnt);
end;

procedure TsSpreadOOXMLReader.ReadFonts(ANode: TDOMNode);
var
  node: TDOMNode;
begin
  node := ANode.FirstChild;
  while node <> nil do begin
    ReadFont(node);
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadHyperlinks(ANode: TDOMNode);
var
  node: TDOMNode;
  nodeName: String;
  hyperlinkData: THyperlinkListData;
  s: String;

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
  if Assigned(ANode) then begin
    nodename := ANode.NodeName;
    if nodename = 'hyperlinks' then
    begin
      node := ANode.FirstChild;
      while Assigned(node) do
      begin
        nodename := node.NodeName;
        if nodename = 'hyperlink' then begin
          hyperlinkData := THyperlinkListData.Create;
          hyperlinkData.CellRef := GetAttrValue(node, 'ref');
          hyperlinkData.ID := GetAttrValue(node, 'r:id');
          hyperlinkData.Target := '';
          hyperlinkData.TextMark := GetAttrValue(node, 'location');
          hyperlinkData.Display := GetAttrValue(node, 'display');
          hyperlinkData.Tooltip := GetAttrValue(node, 'tooltip');
        end;
        FHyperlinkList.Add(hyperlinkData);
        node := node.NextSibling;
      end;
    end else
    if nodename = 'Relationship' then
    begin
      node := ANode;
      while Assigned(node) do
      begin
        nodename := node.NodeName;
        if nodename = 'Relationship' then
        begin
          s := GetAttrValue(node, 'Type');
          if s = SCHEMAS_HYPERLINKS then
          begin
            s := GetAttrValue(node, 'Id');
            if s <> '' then
            begin
              hyperlinkData := FindHyperlinkID(s);
              if hyperlinkData <> nil then begin
                s := GetAttrValue(node, 'Target');
                if s <> '' then hyperlinkData.Target := s;
                s := GetAttrValue(node, 'TargetMode');
                if s <> 'External' then   // Only "External" accepted!
                begin
                  hyperlinkData.Target := '';
                  hyperlinkData.TextMark := '';
                end;
              end;
            end;
          end;
        end;
        node := node.NextSibling;
      end;
    end;
  end;
end;

procedure TsSpreadOOXMLReader.ReadMergedCells(ANode: TDOMNode;
  AWorksheet: TsWorksheet);
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
      if nodename = 'mergeCell' then begin
        s := GetAttrValue(node, 'ref');
        if s <> '' then
          AWorksheet.MergeCells(s);
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
begin
  if Assigned(ANode) then begin
    node := ANode.FirstChild;
    while Assigned(node) do begin
      nodeName := node.NodeName;
      if nodeName = 'numFmt' then begin
        idStr := GetAttrValue(node, 'numFmtId');
        fmtStr := GetAttrValue(node, 'formatCode');
        NumFormatList.AnalyzeAndAdd(StrToInt(idStr), fmtStr);
      end;
      node := node.NextSibling;
    end;
  end;
end;

procedure TsSpreadOOXMLReader.ReadPalette(ANode: TDOMNode);
var
  node, colornode: TDOMNode;
  nodename: String;
  s: string;
  clr: TsColor;
  rgb: TsColorValue;
  n: Integer;
begin
  // OOXML sometimes specifies color by index even if a palette ("indexedColors")
  // is not loaeded. Therefore, we use the BIFF8 palette as default because
  // the default indexedColors are identical to it.
  n := Length(PALETTE_OOXML);
  FWorkbook.UsePalette(@PALETTE_OOXML, n);
  if ANode = nil then
    exit;

  clr := 0;
  node := ANode.FirstChild;
  while Assigned(node) do begin
    nodename := node.NodeName;
    if nodename = 'indexedColors' then begin
      colornode := node.FirstChild;
      while Assigned(colornode) do begin
        nodename := colornode.NodeName;
        if nodename = 'rgbColor' then begin
          s := GetAttrValue(colornode, 'rgb');
          if s <> '' then begin
            rgb := HTMLColorStrToColor('#' + s);
            if clr < n then begin
              FWorkbook.SetPaletteColor(clr, rgb);
              inc(clr);
            end
            else
              FWorkbook.AddColorToPalette(rgb);
          end;
        end;
        colornode := colorNode.NextSibling;
      end;
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadRowHeight(ANode: TDOMNode; AWorksheet: TsWorksheet);
var
  s: String;
  ht: Single;
  r: Cardinal;
  row: PRow;
begin
  if ANode = nil then
    exit;
  s := GetAttrValue(ANode, 'customHeight');
  if s = '1' then begin
    s := GetAttrValue(ANode, 'r');
    r := StrToInt(s) - 1;
    s := GetAttrValue(ANode, 'ht');
    ht := StrToFloat(s, FPointSeparatorSettings);    // seems to be in "Points"
    row := AWorksheet.GetRow(r);
    row^.Height := ht / FWorkbook.GetDefaultFontSize;
    if row^.Height > ROW_HEIGHT_CORRECTION then
      row^.Height := row^.Height - ROW_HEIGHT_CORRECTION
    else
      row^.Height := 0;
  end;
end;

procedure TsSpreadOOXMLReader.ReadSharedStrings(ANode: TDOMNode);
var
  valuenode: TDOMNode;
  childnode: TDOMNode;
  nodename: String;
  s: String;
begin
  while Assigned(ANode) do begin
    if ANode.NodeName = 'si' then begin
      s := '';
      valuenode := ANode.FirstChild;
      while valuenode <> nil do begin
        nodename := valuenode.NodeName;
        if nodename = 't' then
          s := GetNodeValue(valuenode)
        else
        if nodename = 'r' then begin
          childnode := valuenode.FirstChild;
          while childnode <> nil do begin
            s := s + GetNodeValue(childnode);
            childnode := childnode.NextSibling;
          end;
        end;
        valuenode := valuenode.NextSibling;
      end;
      FSharedStrings.Add(s);
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadSheetFormatPr(ANode: TDOMNode;
  AWorksheet: TsWorksheet);
var
  w, h: Single;
  s: String;
begin
  if ANode = nil then
    exit;

  s := GetAttrValue(ANode, 'defaultColWidth');   // is in characters
  if (s <> '') and TryStrToFloat(s, w, FPointSeparatorSettings) then
    AWorksheet.DefaultColWidth := w;

  s := GetAttrValue(ANode, 'defaultRowHeight');  // in in points
  if (s <> '') and TryStrToFloat(s, h, FPointSeparatorSettings) then begin
    h := h / Workbook.GetDefaultFontSize;
    if h > ROW_HEIGHT_CORRECTION then begin
      h := h - ROW_HEIGHT_CORRECTION;
      AWorksheet.DefaultRowHeight := h;
    end;
  end;
end;

procedure TsSpreadOOXMLReader.ReadSheetList(ANode: TDOMNode; AList: TStrings);
var
  node: TDOMNode;
  nodename: String;
  sheetName: String;
  sheetId: String;
begin
  node := ANode.FirstChild;
  while node <> nil do begin
    nodename := node.NodeName;
    if nodename = 'sheet' then
    begin
      sheetName := GetAttrValue(node, 'name');
      sheetId := GetAttrValue(node, 'sheetId');
      AList.AddObject(sheetName, TObject(ptrInt(StrToInt(sheetID))));
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadSheetViews(ANode: TDOMNode; AWorksheet: TsWorksheet);
var
  sheetViewNode: TDOMNode;
  childNode: TDOMNode;
  nodeName: String;
  s: String;
begin
  if ANode = nil then
    exit;

  sheetViewNode := ANode.FirstChild;
  while Assigned(sheetViewNode) do begin
    nodeName := sheetViewNode.NodeName;
    if nodeName = 'sheetView' then begin
      s := GetAttrValue(sheetViewNode, 'showGridLines');
      if s = '0' then
        AWorksheet.Options := AWorksheet.Options - [soShowGridLines];
      s := GetAttrValue(sheetViewNode, 'showRowColHeaders');
      if s = '0' then
         AWorksheet.Options := AWorksheet.Options - [soShowHeaders];

      childNode := sheetViewNode.FirstChild;
      while Assigned(childNode) do begin
        nodeName := childNode.NodeName;
        if nodeName = 'pane' then begin
          s := GetAttrValue(childNode, 'state');
          if s = 'frozen' then begin
            AWorksheet.Options := AWorksheet.Options + [soHasFrozenPanes];
            s := GetAttrValue(childNode, 'xSplit');
            if s <> '' then AWorksheet.LeftPaneWidth := StrToInt(s);
            s := GetAttrValue(childNode, 'ySplit');
            if s <> '' then AWorksheet.TopPaneHeight := StrToInt(s);
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

procedure TsSpreadOOXMLReader.ReadWorksheet(ANode: TDOMNode; AWorksheet: TsWorksheet);
var
  rownode: TDOMNode;
  cellnode: TDOMNode;
begin
  rownode := ANode.FirstChild;
  while Assigned(rownode) do begin
    if rownode.NodeName = 'row' then begin
      ReadRowHeight(rownode, AWorksheet);
      cellnode := rownode.FirstChild;
      while Assigned(cellnode) do begin
        if cellnode.NodeName = 'c' then
          ReadCell(cellnode, AWorksheet);
        cellnode := cellnode.NextSibling;
      end;
    end;
    rownode := rownode.NextSibling;
  end;
  FixCols(AWorksheet);
  FixRows(AWorksheet);
end;

procedure TsSpreadOOXMLReader.ReadFromFile(AFileName: string);
var
  Doc : TXMLDocument;
  FilePath : string;
  UnZip : TUnZipper;
  FileList : TStringList;
  SheetList: TStringList;
  i: Integer;
  fn: String;
  fn_comments: String;
begin
  //unzip "content.xml" of "AFileName" to folder "FilePath"
  FilePath := GetTempDir(false);
  UnZip := TUnZipper.Create;
  FileList := TStringList.Create;
  try
    FileList.Add(OOXML_PATH_XL_STYLES);   // styles
    FileList.Add(OOXML_PATH_XL_STRINGS);  // sharedstrings
    FileList.Add(OOXML_PATH_XL_WORKBOOK); // workbook
    FileList.Add(OOXML_PATH_XL_THEME);    // theme
    UnZip.OutputPath := FilePath;
    Unzip.UnZipFiles(AFileName,FileList);
  finally
    FreeAndNil(FileList);
    FreeAndNil(UnZip);
  end; //try

  Doc := nil;
  SheetList := TStringList.Create;
  try
    // Retrieve theme colors
    if FileExists(FilePath + OOXML_PATH_XL_THEME) then begin
      ReadXMLFile(Doc, FilePath + OOXML_PATH_XL_THEME);
      DeleteFile(FilePath + OOXML_PATH_XL_THEME);
      ReadThemeElements(Doc.DocumentElement.FindNode('a:themeElements'));
      FreeAndNil(Doc);
    end;

    // process the sharedstrings.xml file
    if FileExists(FilePath + OOXML_PATH_XL_STRINGS) then begin
      ReadXMLFile(Doc, FilePath + OOXML_PATH_XL_STRINGS);
      DeleteFile(FilePath + OOXML_PATH_XL_STRINGS);
      ReadSharedStrings(Doc.DocumentElement.FindNode('si'));
      FreeAndNil(Doc);
    end;

    // process the workbook.xml file
    if not FileExists(FilePath + OOXML_PATH_XL_WORKBOOK) then
      raise Exception.CreateFmt(rsDefectiveInternalStructure, ['xlsx']);
    ReadXMLFile(Doc, FilePath + OOXML_PATH_XL_WORKBOOK);
    DeleteFile(FilePath + OOXML_PATH_XL_WORKBOOK);
    ReadFileVersion(Doc.DocumentElement.FindNode('fileVersion'));
    ReadDateMode(Doc.DocumentElement.FindNode('workbookPr'));
    ReadSheetList(Doc.DocumentElement.FindNode('sheets'), SheetList);
    FreeAndNil(Doc);

    // process the styles.xml file
    if FileExists(FilePath + OOXML_PATH_XL_STYLES) then begin // should always exist, just to make sure...
      ReadXMLFile(Doc, FilePath + OOXML_PATH_XL_STYLES);
      DeleteFile(FilePath + OOXML_PATH_XL_STYLES);
      ReadPalette(Doc.DocumentElement.FindNode('colors'));
      ReadFonts(Doc.DocumentElement.FindNode('fonts'));
      ReadFills(Doc.DocumentElement.FindNode('fills'));
      ReadBorders(Doc.DocumentElement.FindNode('borders'));
      ReadNumFormats(Doc.DocumentElement.FindNode('numFmts'));
      ReadCellXfs(Doc.DocumentElement.FindNode('cellXfs'));
      FreeAndNil(Doc);
    end;

    // read worksheets
    for i:=0 to SheetList.Count-1 do begin
      // Create worksheet
      FWorksheet := FWorkbook.AddWorksheet(SheetList[i], true);

      // unzip sheet file
      fn := OOXML_PATH_XL_WORKSHEETS + Format('sheet%d.xml', [i+1]);
      UnzipFile(AFileName, fn, FilePath);
      ReadXMLFile(Doc, FilePath + fn);
      DeleteFile(FilePath + fn);

      // Sheet data, formats, etc.
      ReadSheetViews(Doc.DocumentElement.FindNode('sheetViews'), FWorksheet);
      ReadSheetFormatPr(Doc.DocumentElement.FindNode('sheetFormatPr'), FWorksheet);
      ReadCols(Doc.DocumentElement.FindNode('cols'), FWorksheet);
      ReadWorksheet(Doc.DocumentElement.FindNode('sheetData'), FWorksheet);
      ReadMergedCells(Doc.DocumentElement.FindNode('mergeCells'), FWorksheet);
      ReadHyperlinks(Doc.DocumentElement.FindNode('hyperlinks'));

      FreeAndNil(Doc);

      // Comments:
      // The comments are stored in separate "comments<n>.xml" files (n = 1, 2, ...)
      // The relationship which comment belongs to which sheet file must be
      // retrieved from the "sheet<n>.xml.rels" file (n = 1, 2, ...).
      // The rels file contains also the second part of the hyperlink data.
      fn := OOXML_PATH_XL_WORKSHEETS_RELS + Format('sheet%d.xml.rels', [i+1]);
      UnzipFile(AFilename, fn, FilePath);
      if FileExists(FilePath + fn) then begin
        // find exact name of comments<n>.xml file
        ReadXMLFile(Doc, FilePath + fn);
        DeleteFile(FilePath + fn);
        fn_comments := FindCommentsFileName(Doc.DocumentElement.FindNode('Relationship'));
        ReadHyperlinks(Doc.DocumentElement.FindNode('Relationship'));
        FreeAndNil(Doc);
      end else
      if (SheetList.Count = 1) then
        // if the wookbook has only 1 sheet then the sheet.xml.rels file is missing
        fn_comments := 'comments1.xml'
      else
        // this sheet does not have any cell comments
        continue;
      // Extract texts from the comments file found and apply to worksheet.
      if fn_comments <> '' then
      begin
        fn := OOXML_PATH_XL + fn_comments;
        UnzipFile(AFileName, fn, FilePath);
        if FileExists(FilePath + fn) then begin
          ReadXMLFile(Doc, FilePath + fn);
          DeleteFile(FilePath + fn);
          ReadComments(Doc.DocumentElement.FindNode('commentList'), FWorksheet);
          FreeAndNil(Doc);
        end;
      end;
      ApplyHyperlinks(FWorksheet);
    end;  // for

  finally
    SheetList.Free;
    FreeAndNil(Doc);
  end;
end;

procedure TsSpreadOOXMLReader.ReadFromStream(AStream: TStream);
begin
  Unused(AStream);
  raise Exception.Create('[TsSpreadOOXMLReader.ReadFromStream] '+
                         'Method not implemented. Use "ReadFromFile" instead.');
end;


{ TsSpreadOOXMLWriter }

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
  if (AFormat = nil) or not (uffBorder in AFormat.UsedFormattingFields) then begin
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

  {
  // Index 1 is also pre-defined (gray 25%)
  for i:=2 to High(FFillList) do begin
    fmt := FFillList[i];
    if (fmt <> nil) and (uffBackgroundColor in fmt^.UsedFormattingFields) then
      if (AFormat^.BackgroundColor = fmt^.BackgroundColor) then
      begin
        Result := i;
        exit;
      end;
  end;
   }

   // Not found --> return -1
  Result := -1;
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
begin
  // first list entry is a no-border cell
  n := 1;
  SetLength(FBorderList, n);
  FBorderList[0] := nil;

  for i := 0 to FWorkbook.GetNumCellFormats - 1 do
  begin
    fmt := FWorkbook.GetPointerToCellFormat(i);
    if FindBorderInList(fmt) = -1 then
    begin
      SetLength(FBorderList, n+1);
      FBorderList[n] := fmt;
      inc(n);
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
begin
  // Add built-in fills first.
  n := 2;
  SetLength(FFillList, n);
  FFillList[0] := nil;  // built-in "no fill"
  FFillList[1] := nil;  // built-in "gray125"

  for i := 0 to FWorkbook.GetNumCellFormats - 1 do
  begin
    fmt := FWorkbook.GetPointerToCellFormat(i);
    if FindFillInList(fmt) = -1 then
    begin
      SetLength(FFillList, n+1);
      FFillList[n] := fmt;
      inc(n);
    end;
  end;
end;

procedure TsSpreadOOXMLWriter.WriteBorderList(AStream: TStream);
const
  LINESTYLE_NAME: Array[TsLineStyle] of String = (
     'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair');

  procedure WriteBorderStyle(AStream: TStream; AFormatRecord: PsCellFormat;
    ABorder: TsCellBorder; ABorderName: String);
  { border names found in xlsx files for Excel selections:
    "thin", "hair", "dotted", "dashed", "dashDotDot", "dashDot", "mediumDashDotDot",
    "slantDashDot", "mediumDashDot", "mediumDashed", "medium", "thick", "double" }
  var
    styleName: String;
    colorName: String;
    rgb: TsColorValue;
  begin
    if (ABorder in AFormatRecord^.Border) then begin
      // Line style
      styleName := LINESTYLE_NAME[AFormatRecord^.BorderStyles[ABorder].LineStyle];

      // Border color
      rgb := Workbook.GetPaletteColor(AFormatRecord^.BorderStyles[ABorder].Color);
      //rgb := Workbook.GetPaletteColor(ACell^.BorderStyles[ABorder].Color);
      colorName := ColorToHTMLColorStr(rgb, true);
      AppendToStream(AStream, Format(
        '<%s style="%s"><color rgb="%s" /></%s>',
          [ABorderName, styleName, colorName, ABorderName]
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
    if (cbDiagUp in FBorderList[i].Border) then diag := diag + ' diagonalUp="1"';
    if (cbDiagDown in FBorderList[i].Border) then diag := diag + ' diagonalDown="1"';
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

procedure TsSpreadOOXMLWriter.WriteCols(AStream: TStream; AWorksheet: TsWorksheet);
var
  col: PCol;
  c: Integer;
begin
  if AWorksheet.Cols.Count = 0 then
    exit;

  AppendToStream(AStream,
    '<cols>');

  for c:=0 to AWorksheet.GetLastColIndex do begin
    col := AWorksheet.FindCol(c);
    if col <> nil then
      AppendToStream(AStream, Format(
        '<col min="%d" max="%d" width="%g" customWidth="1" />',
        [c+1, c+1, col.Width], FPointSeparatorSettings)
      );
  end;

  AppendToStream(AStream,
    '</cols>');
end;

procedure TsSpreadOOXMLWriter.WriteComments(AWorksheet: TsWorksheet);
var
  comment: PsComment;
  txt: String;
begin
  if AWorksheet.Comments.Count = 0 then
    exit;

  // Create the comments stream
  SetLength(FSComments, FCurSheetNum + 1);
  if (boBufStream in Workbook.Options) then
    FSComments[FCurSheetNum] := TBufStream.Create(GetTempFileName('', Format('fpsCMNT%d', [FCurSheetNum])))
  else
    FSComments[FCurSheetNum] := TMemoryStream.Create;

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
  //IterateThroughComments(FSComments[FCurSheetNum], AWorksheet.Comments, WriteCommentsCallback);

  for comment in AWorksheet.Comments do
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
                '<fFont vel="Arial" />'+   // It is not harmful to Excel if the font does not exist.
                '<charset val="1" />'+
              '</rPr>'+
              '<t xml:space="preserve">' + txt + '</t>' +
            '</r>' +
          '</text>' +
        '</comment>');
  end;

  (*
  procedure TsSpreadOOXMLWriter.WriteCommentsCallback(AComment: PsComment;
    ACommentIndex: Integer; AStream: TStream);
  var
    comment: String;
  begin
    Unused(ACommentIndex);

    comment := AComment^.Text;
    ValidXMLText(comment);

    // Write comment to Comments stream
    AppendToStream(AStream, Format(
      '<comment ref="%s" authorId="0">', [GetCellString(AComment^.Row, AComment^.Col)]));
    AppendToStream(AStream,
        '<text>'+
          '<r>'+
            '<rPr>'+     // this entire node could be omitted, but then Excel uses some default font out of control
              '<sz val="9"/>'+
              '<color rgb="000000" />'+   // It could be that color index 81 does not exist in fps files --> use rgb instead
              '<rFont val="Arial"/>'+     // It is not harmful to Excel if the font does not exist.
              '<charset val="1"/>'+
            '</rPr>'+
            '<t xml:space="preserve">' + comment + '</t>' +
          '</r>'+
        '</text>');
    AppendToStream(AStream,
      '</comment>');
  end;
  *)

  // Footer
  AppendToStream(FSComments[FCurSheetNum],
      '</commentList>');
  AppendToStream(FSComments[FCurSheetNum],
    '</comments>');
end;
                              (*
procedure TsSpreadOOXMLWriter.WriteCommentsCallback(AComment: PsComment;
  ACommentIndex: Integer; AStream: TStream);
var
  comment: String;
begin
  Unused(ACommentIndex);

  comment := AComment^.Text;
  ValidXMLText(comment);

  // Write comment to Comments stream
  AppendToStream(AStream, Format(
    '<comment ref="%s" authorId="0">', [GetCellString(AComment^.Row, AComment^.Col)]));
  AppendToStream(AStream,
      '<text>'+
        '<r>'+
          '<rPr>'+     // this entire node could be omitted, but then Excel uses some default font out of control
            '<sz val="9"/>'+
            '<color rgb="000000" />'+   // It could be that color index 81 does not exist in fps files --> use rgb instead
            '<rFont val="Arial"/>'+     // It is not harmful to Excel if the font does not exist.
            '<charset val="1"/>'+
          '</rPr>'+
          '<t xml:space="preserve">' + comment + '</t>' +
        '</r>'+
      '</text>');
  AppendToStream(AStream,
    '</comment>');
end;                            *)

procedure TsSpreadOOXMLWriter.WriteDimension(AStream: TStream;
  AWorksheet: TsWorksheet);
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
      fc := Format('rgb="%s"', [Copy(Workbook.GetPaletteColorAsHTMLStr(FFillList[i]^.Background.FgColor), 2, 255)]);
    if FFillList[i].Background.BgColor = scTransparent then
      bc := 'auto="1"'
    else
      bc := Format('rgb="%s"', [Copy(Workbook.GetPaletteColorAsHTMLStr(FFillList[i]^.Background.BgColor), 2, 255)]);
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

{ Writes the fontlist of the workbook to the stream. The font id used in xf
  records is given by the index of a font in the list. Therefore, we have
  to write an empty record for font #4 which is nil due to compatibility with BIFF }
procedure TsSpreadOOXMLWriter.WriteFontList(AStream: TStream);
var
  i: Integer;
  font: TsFont;
  s: String;
  rgb: TsColorValue;
begin
  AppendToStream(FSStyles, Format(
      '<fonts count="%d">', [Workbook.GetFontCount]));
  for i:=0 to Workbook.GetFontCount-1 do begin
    font := Workbook.GetFont(i);
    {
    if font = 4 then
//    if font = nil then
      AppendToStream(AStream, '<font />')
      // Font #4 is missing in fpspreadsheet due to BIFF compatibility. We write
      // an empty node to keep the numbers in sync with the stored font index.
    else begin}
      s := Format('<sz val="%g" /><name val="%s" />', [font.Size, font.FontName], FPointSeparatorSettings);
      if (fssBold in font.Style) then
        s := s + '<b />';
      if (fssItalic in font.Style) then
        s := s + '<i />';
      if (fssUnderline in font.Style) then
        s := s + '<u />';
      if (fssStrikeout in font.Style) then
        s := s + '<strike />';
      if font.Color <> scBlack then begin
        if font.Color < 64 then
          s := s + Format('<color indexed="%d" />', [font.Color])
        else begin
          rgb := Workbook.GetPaletteColor(font.Color);
          s := s + Format('<color rgb="%s" />', [Copy(ColorToHTMLColorStr(rgb), 2, 255)]);
        end;
      end;
      AppendToStream(AStream,
        '<font>', s, '</font>');
//    end;
  end;
  AppendToStream(AStream,
      '</fonts>');
end;

procedure TsSpreadOOXMLWriter.WriteHyperlinks(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  hyperlink: PsHyperlink;
  target, bookmark: String;
  s: String;
  txt: String;
  AVLNode: TAVLTreeNode;
begin
  if AWorksheet.Hyperlinks.Count = 0 then
    exit;

  AppendToStream(AStream,
    '<hyperlinks>');

  // Keep in sync with WriteWorksheetRels !
  FNext_rID := IfThen(AWorksheet.Comments.Count = 0, 1, 3);

  AVLNode := AWorksheet.Hyperlinks.FindLowest;
  while AVLNode <> nil do begin
    hyperlink := PsHyperlink(AVLNode.Data);
    SplitHyperlink(hyperlink^.Target, target, bookmark);
    s := Format('ref="%s"', [GetCellString(hyperlink^.Row, hyperlink^.Col)]);
    if target <> '' then
    begin
      s := Format('%s r:id="rId%d"', [s, FNext_rId]);
      inc(FNext_rId);
    end;
    if bookmark <> '' then //target = '' then
      s := Format('%s location="%s"', [s, bookmark]);
    txt := AWorksheet.ReadAsUTF8Text(hyperlink^.Row, hyperlink^.Col);
    if (txt <> '') and (txt <> hyperlink^.Target) then
      s := Format('%s display="%s"', [s, txt]);
    if hyperlink^.ToolTip <> '' then begin
      txt := hyperlink^.Tooltip;
      ValidXMLText(txt);
      s := Format('%s tooltip="%s"', [s, txt]);
    end;
    AppendToStream(AStream,
        '<hyperlink ' + s + ' />');
    AVLNode := AWorksheet.Hyperlinks.FindSuccessor(AVLNode);
  end;

  AppendToStream(AStream,
    '</hyperlinks>');
end;

procedure TsSpreadOOXMLWriter.WriteMergedCells(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  rng: PsCellRange;
  n: Integer;
begin
  n := AWorksheet.MergedCells.Count;
  if n = 0 then
    exit;
  AppendToStream(AStream, Format(
    '<mergeCells count="%d">', [n]) );
  for rng in AWorksheet.MergedCells do
    AppendToStream(AStream, Format(
      '<mergeCell ref="%s" />', [GetCellRangeString(rng.Row1, rng.Col1, rng.Row2, rng.Col2)]));
  AppendToStream(AStream,
    '</mergeCells>');
end;

{ Writes all number formats to the stream. Saving starts at the item with the
  FirstFormatIndexInFile. }
procedure TsSpreadOOXMLWriter.WriteNumFormatList(AStream: TStream);
var
  i: Integer;
  item: TsNumFormatData;
  s: String;
  n: Integer;
begin
  s := '';
  n := 0;
  i := NumFormatList.FindByIndex(NumFormatList.FirstNumFormatIndexInFile);
  if i > -1 then begin
    while i < NumFormatList.Count do begin
      item := NumFormatList[i];
      if item <> nil then begin
        s := s + Format('<numFmt numFmtId="%d" formatCode="%s" />',
          [item.Index, UTF8TextToXMLText(NumFormatList.FormatStringForWriting(i))]);
        inc(n);
      end;
      inc(i);
    end;
    if n > 0 then
      AppendToStream(AStream, Format(
        '<numFmts count="%d">', [n]),
          s,
        '</numFmts>'
      );
  end;
end;

{ Writes the workbook's color palette to the file }
procedure TsSpreadOOXMLWriter.WritePalette(AStream: TStream);
var
  rgb: TsColorValue;
  i: Integer;
begin
  AppendToStream(AStream,
    '<colors>' +
      '<indexedColors>');

  // There must not be more than 64 palette entries because the next colors
  // are system colors.
  for i:=0 to Min(LAST_PALETTE_COLOR, Workbook.GetPaletteSize-1) do begin
    rgb := Workbook.GetPaletteColor(i);
    AppendToStream(AStream,
        '<rgbColor rgb="'+ColorToHTMLColorStr(rgb, true) + '" />');
  end;

  AppendToStream(AStream,
      '</indexedColors>' +
    '</colors>');
end;

procedure TsSpreadOOXMLWriter.WriteSheetData(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  r, r1, r2: Cardinal;
  c, c1, c2: Cardinal;
  row: PRow;
  value: Variant;
  lCell: TCell;
  styleCell: PCell;
  cell: PCell;
  rh: String;
  h0: Single;
begin
  h0 := Workbook.GetDefaultFontSize;  // Point size of default font

  AppendToStream(AStream,
      '<sheetData>');

  GetSheetDimensions(AWorksheet, r1, r2, c1, c2);

  if (boVirtualMode in Workbook.Options) and Assigned(Workbook.OnWriteCellData)
  then begin
    for r := 0 to r2 do begin
      row := AWorksheet.FindRow(r);
      if row <> nil then
        rh := Format(' ht="%g" customHeight="1"', [
          (row^.Height + ROW_HEIGHT_CORRECTION)*h0],
          FPointSeparatorSettings)
      else
        rh := '';
      AppendToStream(AStream, Format(
        '<row r="%d" spans="1:%d"%s>', [r+1, Workbook.VirtualColCount, rh]));
      for c := 0 to c2 do begin
        lCell.Row := r; // to silence a compiler hint
        InitCell(lCell);
        value := varNull;
        styleCell := nil;
        Workbook.OnWriteCellData(Workbook, r, c, value, styleCell);
        if styleCell <> nil then
          lCell := styleCell^;
        lCell.Row := r;
        lCell.Col := c;
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
          lCell.DateTimeValue := StrToDateTime(VarToStr(value), Workbook.FormatSettings);  // was: StrToDate
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
//        WriteCellCallback(@lCell, AStream);
        varClear(value);
      end;
      AppendToStream(AStream,
        '</row>');
    end;
  end else
  begin
    // The cells need to be written in order, row by row, cell by cell
    for r := r1 to r2 do begin
      // If the row has a custom height add this value to the <row> specification
      row := AWorksheet.FindRow(r);
      if row <> nil then
        rh := Format(' ht="%g" customHeight="1"', [
          (row^.Height + ROW_HEIGHT_CORRECTION)*h0], FPointSeparatorSettings)
      else
        rh := '';
      AppendToStream(AStream, Format(
        '<row r="%d" spans="%d:%d"%s>', [r+1, c1+1, c2+1, rh]));
      // Write cells belonging to this row.
      for cell in AWorksheet.Cells.GetRowEnumerator(r) do
        WriteCellToStream(AStream, cell);
                          {
      for c := c1 to c2 do begin
        cell := AWorksheet.FindCell(r, c);
        if Assigned(cell) then begin
          WriteCellCallback(cell, AStream);
        end;
      end;
      }
      AppendToStream(AStream,
        '</row>');
    end;
  end;
  AppendToStream(AStream,
      '</sheetData>');
end;

procedure TsSpreadOOXMLWriter.WriteSheetViews(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  showGridLines: String;
  showHeaders: String;
  topRightCell: String;
  bottomLeftCell: String;
  bottomRightCell: String;
begin
  // Show gridlines ?
  showGridLines := StrUtils.IfThen(soShowGridLines in AWorksheet.Options, ' ', 'showGridLines="0" ');

  // Show headers?
  showHeaders := StrUtils.IfThen(soShowHeaders in AWorksheet.Options, ' ', 'showRowColHeaders="0" ');

  // No frozen panes
  if not (soHasFrozenPanes in AWorksheet.Options) or
     ((AWorksheet.LeftPaneWidth = 0) and (AWorksheet.TopPaneHeight = 0))
  then
    AppendToStream(AStream, Format(
      '<sheetViews>' +
        '<sheetView workbookViewId="0" %s%s/>' +
      '</sheetViews>', [
      showGridLines, showHeaders
    ]))
  else
  begin  // Frozen panes
    topRightCell := GetCellString(0, AWorksheet.LeftPaneWidth, [rfRelRow, rfRelCol]);
    bottomLeftCell := GetCellString(AWorksheet.TopPaneHeight, 0, [rfRelRow, rfRelCol]);
    bottomRightCell := GetCellString(AWorksheet.TopPaneHeight, AWorksheet.LeftPaneWidth, [rfRelRow, rfRelCol]);
    if (AWorksheet.LeftPaneWidth > 0) and (AWorksheet.TopPaneHeight > 0) then
      AppendToStream(AStream, Format(
        '<sheetViews>' +
          '<sheetView workbookViewId="0" %s%s>'+
            '<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen" />' +
            '<selection pane="topRight" activeCell="%s" sqref="%s" />' +
            '<selection pane="bottomLeft" activeCell="%s" sqref="%s" />' +
            '<selection pane="bottomRight" activeCell="%s" sqref="%s" />' +
          '</sheetView>' +
        '</sheetViews>', [
        showGridLines, showHeaders,
        AWorksheet.LeftPaneWidth, AWorksheet.TopPaneHeight, bottomRightCell,
        topRightCell, topRightCell,
        bottomLeftCell, bottomLeftCell,
        bottomRightCell, bottomrightCell
      ]))
    else
    if (AWorksheet.LeftPaneWidth > 0) then
      AppendToStream(AStream, Format(
        '<sheetViews>' +
          '<sheetView workbookViewId="0" %s%s>'+
            '<pane xSplit="%d" topLeftCell="%s" activePane="topRight" state="frozen" />' +
            '<selection pane="topRight" activeCell="%s" sqref="%s" />' +
          '</sheetView>' +
        '</sheetViews>', [
        showGridLines, showHeaders,
        AWorksheet.LeftPaneWidth, topRightCell,
        topRightCell, topRightCell
      ]))
    else
    if (AWorksheet.TopPaneHeight > 0) then
      AppendToStream(AStream, Format(
        '<sheetViews>'+
          '<sheetView workbookViewId="0" %s%s>'+
             '<pane ySplit="%d" topLeftCell="%s" activePane="bottomLeft" state="frozen" />'+
             '<selection pane="bottomLeft" activeCell="%s" sqref="%s" />' +
          '</sheetView>'+
        '</sheetViews>', [
        showGridLines, showHeaders,
        AWorksheet.TopPaneHeight, bottomLeftCell,
        bottomLeftCell, bottomLeftCell
      ]));
  end;
end;

{ Writes the style list which the workbook has collected in its FormatList }
procedure TsSpreadOOXMLWriter.WriteStyleList(AStream: TStream; ANodeName: String);
var
//  styleCell: TCell;
  s, sAlign: String;
  fontID: Integer;
  numFmtId: Integer;
  fillId: Integer;
  borderId: Integer;
  idx: Integer;
  fmt: PsCellFormat;
  i: Integer;
begin
  AppendToStream(AStream, Format(
    '<%s count="%d">', [ANodeName, FWorkbook.GetNumCellFormats]));

  for i:=0 to FWorkbook.GetNumCellFormats-1 do
  begin
    fmt := FWorkbook.GetPointerToCellFormat(i);
    s := '';
    sAlign := '';

    { Number format }
    if (uffNumberFormat in fmt^.UsedFormattingFields) then
    begin
      idx := NumFormatList.Find(fmt^.NumberFormat, fmt^.NumberFormatStr);
      if idx > -1 then begin
        numFmtID := NumFormatList[idx].Index;
        s := s + Format('numFmtId="%d" applyNumberFormat="1" ', [numFmtId]);
      end;
    end;

    { Font }
    fontId := 0;
    {
    if (uffBold in fmt^.UsedFormattingFields) then
      fontID := BOLD_FONTINDEX;
    }
    if (uffFont in fmt^.UsedFormattingFields) then
      fontID := fmt^.FontIndex;
    s := s + Format('fontId="%d" ', [fontId]);
    if fontID > 0 then s := s + 'applyFont="1" ';

    if ANodeName = 'cellXfs' then s := s + 'xfId="0" ';

    { Text rotation }
    if (uffTextRotation in fmt^.UsedFormattingFields) then
      case fmt^.TextRotation of
        trHorizontal                      : ;
        rt90DegreeClockwiseRotation       : sAlign := sAlign + Format('textRotation="%d" ', [180]);
        rt90DegreeCounterClockwiseRotation: sAlign := sAlign + Format('textRotation="%d" ',  [90]);
        rtStacked                         : sAlign := sAlign + Format('textRotation="%d" ', [255]);
      end;

    { Text alignment }
    if (uffHorAlign in fmt^.UsedFormattingFields) and (fmt^.HorAlignment <> haDefault)
    then
      case fmt.HorAlignment of
        haLeft  : sAlign := sAlign + 'horizontal="left" ';
        haCenter: sAlign := sAlign + 'horizontal="center" ';
        haRight : sAlign := sAlign + 'horizontal="right" ';
      end;

    if (uffVertAlign in fmt^.UsedFormattingFields) and (fmt^.VertAlignment <> vaDefault)
    then
      case fmt.VertAlignment of
        vaTop   : sAlign := sAlign + 'vertical="top" ';
        vaCenter: sAlign := sAlign + 'vertical="center" ';
        vaBottom: sAlign := sAlign + 'vertical="bottom" ';
      end;

    if (uffWordWrap in fmt^.UsedFormattingFields) then
      sAlign := sAlign + 'wrapText="1" ';

    { Fill }
    if (uffBackground in fmt.UsedFormattingFields) then
    begin
      fillID := FindFillInList(fmt);
      if fillID = -1 then fillID := 0;
      s := s + Format('fillId="%d" applyFill="1" ', [fillID]);
    end;

    { Border }
    if (uffBorder in fmt^.UsedFormattingFields) then
    begin
      borderID := FindBorderInList(fmt);
      if borderID = -1 then borderID := 0;
      s := s + Format('borderId="%d" applyBorder="1" ', [borderID]);
    end;

    { Write everything to stream }
    if sAlign = '' then
      AppendToStream(AStream,
        '<xf ' + s + '/>')
    else
      AppendToStream(AStream,
       '<xf ' + s + 'applyAlignment="1">',
         '<alignment ' + sAlign + ' />',
       '</xf>');
  end;

  AppendToStream(FSStyles, Format(
    '</%s>', [ANodeName]));
end;

procedure TsSpreadOOXMLWriter.WriteVmlDrawings(AWorksheet: TsWorksheet);
// My xml viewer does not format vml files property --> format in code.
var
  comment: PsComment;
  index: Integer;
  id: Integer;
begin
  if AWorksheet.Comments.Count = 0 then
    exit;

  SetLength(FSVmlDrawings, FCurSheetNum + 1);
  if (boBufStream in Workbook.Options) then
    FSVmlDrawings[FCurSheetNum] := TBufStream.Create(GetTempFileName('', Format('fpsVMLD%d', [FCurSheetNum])))
  else
    FSVmlDrawings[FCurSheetNum] := TMemoryStream.Create;

  // Header
  AppendToStream(FSVmlDrawings[FCurSheetNum],
    '<xml xmlns:v="urn:schemas-microsoft-com:vml" '+
         'xmlns:o="urn:schemas-microsoft-com:office:office" '+
         'xmlns:x="urn:schemas-microsoft-com:office:excel">' + LineEnding);
  // My xml viewer does not format vml files property --> format in code.
  AppendToStream(FSVmlDrawings[FCurSheetNum],
    '  <o:shapelayout v:ext="edit">' + LineEnding +
    '    <o:idmap v:ext="edit" data="1" />' + LineEnding +
         // "data" is a comma-separated list with the ids of groups of 1024 comments -- really?
    '  </o:shapelayout>' + LineEnding);
  AppendToStream(FSVmlDrawings[FCurSheetNum],
    '  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">'+LineEnding+
    '    <v:stroke joinstyle="miter"/>' + LineEnding +
    '    <v:path gradientshapeok="t" o:connecttype="rect"/>' + LineEnding +
    '  </v:shapetype>' + LineEnding);

  // Write vmlDrawings for each comment (formatting and position of comment box)
  index := 1;
  for comment in AWorksheet.Comments do
  begin
    id := 1024 + index;     // if more than 1024 comments then use data="1,2,etc" above! -- not implemented yet
    AppendToStream(FSVmlDrawings[FCurSheetNum], LineEnding + Format(
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

  //IterateThroughComments(FSVmlDrawings[FCurSheetNum], AWorksheet.Comments, WriteVmlDrawingsCallback);

  // Footer
  AppendToStream(FSVmlDrawings[FCurSheetNum],
    '</xml>');
end;

procedure TsSpreadOOXMLWriter.WriteVmlDrawingsCallback(AComment: PsComment;
  ACommentIndex: integer; AStream: TStream);
var
  id: Integer;
begin
  id := 1025 + ACommentIndex;     // if more than 1024 comments then use data="1,2,etc" above! -- not implemented yet

  // My xml viewer does not format vml files property --> format in code.
  AppendToStream(AStream, LineEnding + Format(
    '  <v:shape id="_x0000_s%d" type="#_x0000_t202" ', [id]) + LineEnding + Format(
    '       style="position:absolute; width:108pt; height:52.5pt; z-index:%d; visibility:hidden" ', [ACommentIndex+1]) + LineEnding +
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
    '      <x:Row>%d</x:Row>', [AComment^.Row]) + LineEnding + Format(
    '      <x:Column>%d</x:Column>', [AComment^.Col]) + LineEnding +
    '    </x:ClientData>'+ LineEnding+
    '  </v:shape>' + LineEnding);
end;

procedure TsSpreadOOXMLWriter.WriteWorksheetRels(AWorksheet: TsWorksheet);
var
  AVLNode: TAVLTreeNode;
  hyperlink: PsHyperlink;
  s: String;
  target, bookmark: String;
begin
  // Extend stream array
  SetLength(FSSheetRels, FCurSheetNum + 1);

  // Anything to write?
  if (AWorksheet.Comments.Count = 0) and (AWorksheet.Hyperlinks.Count = 0) then
    exit;

  // Create stream
  if (boBufStream in Workbook.Options) then
    FSSheetRels[FCurSheetNum] := TBufStream.Create(GetTempFileName('', Format('fpsWSR%d', [FCurSheetNum])))
  else
    FSSheetRels[FCurSheetNum] := TMemoryStream.Create;

  // Header
  AppendToStream(FSSheetRels[FCurSheetNum],
    XML_HEADER);
  AppendToStream(FSSheetRels[FCurSheetNum], Format(
    '<Relationships xmlns="%s">', [SCHEMAS_RELS]));

  FNext_rId := 1;

  // Relationships for comments
  if AWorksheet.Comments.Count > 0 then
  begin
    AppendToStream(FSSheetRels[FCurSheetNum], Format(
      '<Relationship Id="rId1" Type="%s" Target="../drawings/vmlDrawing%d.vml" />',
        [SCHEMAS_DRAWINGS, FCurSheetNum+1]));
    AppendToStream(FSSheetRels[FCurSheetNum], Format(
      '<Relationship Id="rId2" Type="%s" Target="../comments%d.xml" />',
        [SCHEMAS_COMMENTS, FCurSheetNum+1]));
    FNext_rId := 3;
  end;

  // Relationships for hyperlinks
  if AWorksheet.Hyperlinks.Count > 0 then
  begin
    AVLNode := AWorksheet.Hyperlinks.FindLowest;
    while Assigned(AVLNode) do
    begin
      hyperlink := PsHyperlink(AVLNode.Data);
      SplitHyperlink(hyperlink^.Target, target, bookmark);
      if target <> '' then
      begin
        if (pos('file:', target) = 0) and FileNameIsAbsolute(target) then
          FileNameToURI(target);
//          target := 'file:///' + target;
        s := Format('Id="rId%d" Type="%s" Target="%s" TargetMode="External"',
          [FNext_rId, SCHEMAS_HYPERLINKS, target]);
        AppendToStream(FSSheetRels[FCurSheetNum],
          '<Relationship ' + s + ' />');
        inc(FNext_rId);
      end;
      AVLNode := AWorksheet.Hyperlinks.FindSuccessor(AVLNode);
    end;
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

  { --- RelsRels --- }
  AppendToStream(FSRelsRels,
    XML_HEADER);
  AppendToStream(FSRelsRels, Format(
    '<Relationships xmlns="%s">', [SCHEMAS_RELS]));
  AppendToStream(FSRelsRels, Format(
      '<Relationship Type="%s" Target="xl/workbook.xml" Id="rId1" />', [SCHEMAS_DOCUMENT]));
  AppendToStream(FSRelsRels,
    '</Relationships>');

  { --- Styles --- }
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

  // Misc
  AppendToStream(FSStyles,
      '<dxfs count="0" />');
  AppendToStream(FSStyles,
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16" />');

  // Palette
  WritePalette(FSStyles);

  AppendToStream(FSStyles,
    '</styleSheet>');
end;

procedure TsSpreadOOXMLWriter.WriteContent;
var
  i, counter: Integer;
begin
  { --- WorkbookRels ---
  { Workbook relations - Mark relation to all sheets }
  counter := 0;
  AppendToStream(FSWorkbookRels,
    XML_HEADER);
  AppendToStream(FSWorkbookRels,
    '<Relationships xmlns="' + SCHEMAS_RELS + '">');
  while counter <= Workbook.GetWorksheetCount do begin
    inc(counter);
    AppendToStream(FSWorkbookRels, Format(
      '<Relationship Type="%s" Target="worksheets/sheet%d.xml" Id="rId%d" />',
        [SCHEMAS_WORKSHEET, counter, counter]));
  end;
  AppendToStream(FSWorkbookRels, Format(
      '<Relationship Id="rId%d" Type="%s" Target="styles.xml" />',
        [counter+1, SCHEMAS_STYLES]));
  AppendToStream(FSWorkbookRels, Format(
      '<Relationship Id="rId%d" Type="%s" Target="sharedStrings.xml" />',
        [counter+2, SCHEMAS_STRINGS]));
  AppendToStream(FSWorkbookRels,
    '</Relationships>');

  { --- Workbook --- }
  { Global workbook data - Mark all sheets }
  AppendToStream(FSWorkbook,
    XML_HEADER);
  AppendToStream(FSWorkbook, Format(
    '<workbook xmlns="%s" xmlns:r="%s">', [SCHEMAS_SPREADML, SCHEMAS_DOC_RELS]));
  AppendToStream(FSWorkbook,
      '<fileVersion appName="fpspreadsheet" />');
  AppendToStream(FSWorkbook,
      '<workbookPr defaultThemeVersion="124226" />');
  AppendToStream(FSWorkbook,
      '<bookViews>' +
        '<workbookView xWindow="480" yWindow="90" windowWidth="15195" windowHeight="12525" />' +
      '</bookViews>');
  AppendToStream(FSWorkbook,
      '<sheets>');
  for counter:=1 to Workbook.GetWorksheetCount do
    AppendToStream(FSWorkbook, Format(
        '<sheet name="%s" sheetId="%d" r:id="rId%d" />',
          [Workbook.GetWorksheetByIndex(counter-1).Name, counter, counter]));
  AppendToStream(FSWorkbook,
      '</sheets>');
  AppendToStream(FSWorkbook,
      '<calcPr calcId="114210" />');
  AppendToStream(FSWorkbook,
    '</workbook>');

  // Preparation for shared strings
  FSharedStringsCount := 0;

  // Write all worksheets which fills also the shared strings.
  // Also: write comments and related files
  FNext_rId := 1;
  for i := 0 to Workbook.GetWorksheetCount - 1 do
  begin
    FWorksheet := Workbook.GetWorksheetByIndex(i);
    WriteWorksheet(FWorksheet);
    WriteComments(FWorksheet);
    WriteVmlDrawings(FWorksheet);
    WriteWorksheetRels(FWorksheet);
  end;

  // Finalization of the shared strings document
  AppendToStream(FSSharedStrings_complete,
    XML_HEADER, Format(
    '<sst xmlns="%s" count="%d" uniqueCount="%d">', [SCHEMAS_SPREADML, FSharedStringsCount, FSharedStringsCount]
  ));
  ResetStream(FSSharedStrings);
  FSSharedStrings_complete.CopyFrom(FSSharedStrings, FSSharedStrings.Size);
  AppendToStream(FSSharedStrings_complete,
    '</sst>');
end;

procedure TsSpreadOOXMLWriter.WriteContentTypes;
var
  i: Integer;
begin
  AppendToStream(FSContentTypes,
    XML_HEADER);
  AppendToStream(FSContentTypes,
    '<Types xmlns="' + SCHEMAS_TYPES + '">');
    (*
  AppendToStream(FSContentTypes,
      '<Override PartName="/_rels/.rels" ContentType="' + MIME_RELS + '" />');
  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />');
      *)
  AppendToStream(FSContentTypes, Format(
      '<Default Extension="rels" ContentType="%s" />', [MIME_RELS]));
  AppendToStream(FSContentTypes, Format(
      '<Default Extension="xml" ContentType="%s" />', [MIME_XML]));
  AppendToStream(FSContentTypes, Format(
      '<Default Extension="vml" ContentType="%s" />', [MIME_VMLDRAWING]));

  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/workbook.xml" ContentType="' + MIME_SHEET + '" />');

  for i:=1 to Workbook.GetWorksheetCount do
    AppendToStream(FSContentTypes, Format(
      '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="%s" />',
        [i, MIME_WORKSHEET]));

  for i:=1 to Length(FSComments) do
    AppendToStream(FSContentTypes, Format(
      '<Override PartName="/xl/comments%d.xml" ContentType="%s" />',
        [i, MIME_COMMENTS]));

  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/styles.xml" ContentType="' + MIME_STYLES + '" />');
  AppendToStream(FSContentTypes,
      '<Override PartName="/xl/sharedStrings.xml" ContentType="' + MIME_STRINGS + '" />');
  AppendToStream(FSContentTypes,
    '</Types>');
end;

procedure TsSpreadOOXMLWriter.WriteWorksheet(AWorksheet: TsWorksheet);
begin
  FCurSheetNum := Length(FSSheets);
  SetLength(FSSheets, FCurSheetNum + 1);

  // Create the stream
  if (boBufStream in Workbook.Options) then
    FSSheets[FCurSheetNum] := TBufStream.Create(GetTempFileName('', Format('fpsSH%d', [FCurSheetNum])))
  else
    FSSheets[FCurSheetNum] := TMemoryStream.Create;

  // Header
  AppendToStream(FSSheets[FCurSheetNum],
    XML_HEADER);
  AppendToStream(FSSheets[FCurSheetNum], Format(
    '<worksheet xmlns="%s" xmlns:r="%s">', [SCHEMAS_SPREADML, SCHEMAS_DOC_RELS]));

  WriteDimension(FSSheets[FCurSheetNum], AWorksheet);
  WriteSheetViews(FSSheets[FCurSheetNum], AWorksheet);
  WriteCols(FSSheets[FCurSheetNum], AWorksheet);
  WriteSheetData(FSSheets[FCurSheetNum], AWorksheet);
  WriteHyperlinks(FSSheets[FCurSheetNum], AWorksheet);
  WriteMergedCells(FSSheets[FCurSheetNum], AWorksheet);

  // Footer
  if AWorksheet.Comments.Count > 0 then
    AppendToStream(FSSheets[FCurSheetNum],
      '<legacyDrawing r:id="rId1" />');
  AppendToStream(FSSheets[FCurSheetNum],
    '</worksheet>');
end;

constructor TsSpreadOOXMLWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  // Initial base date in case it won't be set otherwise.
  // Use 1900 to get a bit more range between 1900..1904.
  FDateMode := XlsxSettings.DateMode;

  // Special version of FormatSettings using a point decimal separator for sure.
  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';

  // http://en.wikipedia.org/wiki/List_of_spreadsheet_software#Specifications
  FLimitations.MaxColCount := 16384;
  FLimitations.MaxRowCount := 1048576;
end;

procedure TsSpreadOOXMLWriter.CreateNumFormatList;
begin
  FreeAndNil(FNumFormatList);
  FNumFormatList := TsOOXMLNumFormatList.Create(Workbook);
end;

{ Creates the streams for the individual data files. Will be zipped into a
  single xlsx file. }
procedure TsSpreadOOXMLWriter.CreateStreams;
begin
  if (boBufStream in Workbook.Options) then begin
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
  // FSSheets will be created when needed.
end;

{ Destroys the streams that were created by the writer }
procedure TsSpreadOOXMLWriter.DestroyStreams;

  procedure DestroyStream(AStream: TStream);
  var
    fn: String;
  begin
    if AStream is TFileStream then begin
      fn := TFileStream(AStream).Filename;
      DeleteFile(fn);
    end;
    AStream.Free;
  end;

var
  stream: TStream;
begin
  DestroyStream(FSContentTypes);
  DestroyStream(FSRelsRels);
  DestroyStream(FSWorkbookRels);
  DestroyStream(FSWorkbook);
  DestroyStream(FSStyles);
  DestroyStream(FSSharedStrings);
  DestroyStream(FSSharedStrings_complete);
  for stream in FSSheets do DestroyStream(stream);
  SetLength(FSSheets, 0);
  for stream in FSComments do DestroyStream(stream);
  SetLength(FSComments, 0);
  for stream in FSSheetRels do DestroyStream(stream);
  SetLength(FSSheetRels, 0);
  for stream in FSVmlDrawings do DestroyStream(stream);
  SetLength(FSVmlDrawings, 0);
end;

{ Prepares a string formula for writing }
function TsSpreadOOXMLWriter.PrepareFormula(const AFormula: String): String;
begin
  Result := AFormula;
  if (Result <> '') and (Result[1] = '=') then Delete(Result, 1, 1);
  Result := UTF8TextToXMLText(Result)
end;

{ Is called before zipping the individual file parts. Rewinds the streams. }
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
  for i:=0 to High(FSSheets) do ResetStream(FSSheets[i]);
  for i:=0 to High(FSSheetRels) do ResetStream(FSSheetRels[i]);
  for i:=0 to High(FSComments) do ResetStream(FSComments[i]);
  for i:=0 to High(FSVmlDrawings) do ResetStream(FSVmlDrawings[i]);
end;

{
  Writes a string to a file. Helper convenience method.
}
procedure TsSpreadOOXMLWriter.WriteStringToFile(AFileName, AString: string);
var
  TheStream : TFileStream;
  S : String;
begin
  TheStream := TFileStream.Create(AFileName, fmCreate);
  S:=AString;
  TheStream.WriteBuffer(Pointer(S)^,Length(S));
  TheStream.Free;
end;

{
  Writes an OOXML document to the disc
}
procedure TsSpreadOOXMLWriter.WriteToFile(const AFileName: string;
  const AOverwriteExisting: Boolean);
var
  lStream: TStream;
  lMode: word;
begin
  if AOverwriteExisting
    then lMode := fmCreate or fmOpenWrite
    else lMode := fmCreate;

  if (boBufStream in Workbook.Options) then
    lStream := TBufStream.Create(AFileName, lMode)
  else
    lStream := TFileStream.Create(AFileName, lMode);
  try
    WriteToStream(lStream);
  finally
    FreeAndNil(lStream);
  end;
end;

procedure TsSpreadOOXMLWriter.WriteToStream(AStream: TStream);
var
  FZip: TZipper;
  i: Integer;
begin
  { Analyze the workbook and collect all information needed }
  ListAllNumFormats;
  ListAllFills;
  ListAllBorders;

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
    FZip.FileName := '__temp__.tmp';
    FZip.Entries.AddFileEntry(FSContentTypes, OOXML_PATH_TYPES);
    FZip.Entries.AddFileEntry(FSRelsRels, OOXML_PATH_RELS_RELS);
    FZip.Entries.AddFileEntry(FSWorkbookRels, OOXML_PATH_XL_RELS_RELS);
    FZip.Entries.AddFileEntry(FSWorkbook, OOXML_PATH_XL_WORKBOOK);
    FZip.Entries.AddFileEntry(FSStyles, OOXML_PATH_XL_STYLES);
    FZip.Entries.AddFileEntry(FSSharedStrings_complete, OOXML_PATH_XL_STRINGS);

    for i:=0 to High(FSSheets) do begin
      FSSheets[i].Position:= 0;
      FZip.Entries.AddFileEntry(FSSheets[i], OOXML_PATH_XL_WORKSHEETS + Format('sheet%d.xml', [i+1]));
    end;

    for i:=0 to High(FSComments) do begin
      if (FSComments[i] = nil) or (FSComments[i].Size = 0) then continue;
      FSComments[i].Position := 0;
      FZip.Entries.AddFileEntry(FSComments[i], OOXML_PATH_XL + Format('comments%d.xml', [i+1]));
    end;

    for i:=0 to High(FSSheetRels) do begin
      if (FSSheetRels[i] = nil) or (FSSheetRels[i].Size = 0) then continue;
      FSSheetRels[i].Position := 0;
      FZip.Entries.AddFileEntry(FSSheetRels[i], OOXML_PATH_XL_WORKSHEETS_RELS + Format('sheet%d.xml.rels', [i+1]));
    end;

    for i:=0 to High(FSVmlDrawings) do begin
      if (FSVmlDrawings[i] = nil) or (FSVmlDrawings[i].Size = 0) then continue;
      FSVmlDrawings[i].Position := 0;
      FZip.Entries.AddFileEntry(FSVmlDrawings[i], OOXML_PATH_XL_DRAWINGS + Format('vmlDrawing%d.vml', [i+1]));
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

{ Writes a boolean value to the stream }
procedure TsSpreadOOXMLWriter.WriteBool(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: Boolean; ACell: PCell);
var
  CellPosText: String;
  CellValueText: String;
  lStyleIndex: Integer;
begin
  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);
  if AValue then CellValueText := '1' else CellValueText := '0';
  AppendToStream(AStream, Format(
    '<c r="%s" s="%d" t="b"><v>%s</v></c>', [CellPosText, lStyleIndex, CellValueText]));
end;

{ Writes an error value to the specified cell. }
procedure TsSpreadOOXMLWriter.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol);
  Unused(AValue, ACell);
end;

{ Writes a string formula to the given cell. }
procedure TsSpreadOOXMLWriter.WriteFormula(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
var
  cellPosText: String;
  lStyleIndex: Integer;
  t, v: String;
begin
  cellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);

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
      PrepareFormula(ACell^.FormulaValue),
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
  MAXBYTES = 32767; //limit for this format
var
  CellPosText: string;
  lStyleIndex: Cardinal;
  ResultingValue: string;
begin
  // Office 2007-2010 (at least) support no more characters in a cell;
  if Length(AValue) > MAXBYTES then
  begin
    ResultingValue := Copy(AValue, 1, MAXBYTES); //may chop off multicodepoint UTF8 characters but well...
    Workbook.AddErrorMsg(rsTruncateTooLongCellText, [
      MAXBYTES, GetCellString(ARow, ACol)
    ]);
  end
  else
    ResultingValue := AValue;

  if not ValidXMLText(ResultingValue) then
    Workbook.AddErrorMsg(
      rsInvalidCharacterInCell, [
      GetCellString(ARow, ACol)
    ]);

  AppendToStream(FSSharedStrings,
    '<si>' +
      '<t>' + ResultingValue + '</t>' +
    '</si>');

  CellPosText := TsWorksheet.CellPosToText(ARow, ACol);
  lStyleIndex := GetStyleIndex(ACell);
  AppendToStream(AStream, Format(
    '<c r="%s" s="%d" t="s"><v>%d</v></c>', [CellPosText, lStyleIndex, FSharedStringsCount]));

  inc(FSharedStringsCount);
end;

{
  Writes a number (64-bit IEE 754 floating point) to the sheet
}
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

{*******************************************************************
*  TsSpreadOOXMLWriter.WriteDateTime ()
*
*  DESCRIPTION:    Writes a date/time value as a number
*                  Respects DateMode of the file
*******************************************************************}
procedure TsSpreadOOXMLWriter.WriteDateTime(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell);
var
  ExcelDateSerial: double;
begin
  ExcelDateSerial := ConvertDateTimeToExcelDateTime(AValue, FDateMode);
  WriteNumber(AStream, ARow, ACol, ExcelDateSerial, ACell);
end;

{
  Registers this reader / writer on fpSpreadsheet
}
initialization

  RegisterSpreadFormat(TsSpreadOOXMLReader, TsSpreadOOXMLWriter, sfOOXML);
  MakeLEPalette(@PALETTE_OOXML, Length(PALETTE_OOXML));

end.

