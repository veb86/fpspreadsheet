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
  {$mode objfpc}{$H+}
//  {$mode delphi}
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
  fpsTypes, fpSpreadsheet, fpsUtils, fpsReaderWriter, fpsNumFormat, fpsPalette,
  fpsxmlcommon, xlsCommon;
  
type

  { TsSpreadOOXMLReader }

  TsSpreadOOXMLReader = class(TsSpreadXMLReader)
  private
    FDateMode: TDateMode;
    FPointSeparatorSettings: TFormatSettings;
    FSharedStrings: TStringList;
    FFillList: TFPList;
    FBorderList: TFPList;
    FHyperlinkList: TFPList;
    FSharedFormulaBaseList: TFPList;
    FPalette: TsPalette;
    FThemeColors: array of TsColor;
    FLastRow, FLastCol: Cardinal;
    FWrittenByFPS: Boolean;
    procedure ApplyCellFormatting(ACell: PCell; XfIndex: Integer);
    procedure ApplyHyperlinks(AWorksheet: TsWorksheet);
    function FindCommentsFileName(ANode: TDOMNode): String;
    procedure ReadActiveSheet(ANode: TDOMNode; out ActiveSheetIndex: Integer);
    procedure ReadBorders(ANode: TDOMNode);
    procedure ReadCell(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadCellXfs(ANode: TDOMNode);
    function  ReadColor(ANode: TDOMNode): TsColor;
    procedure ReadCols(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadComments(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadDateMode(ANode: TDOMNode);
    procedure ReadDefinedNames(ANode: TDOMNode);
    procedure ReadDimension(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadFileVersion(ANode: TDOMNode);
    procedure ReadFills(ANode: TDOMNode);
    function ReadFont(ANode: TDOMNode): Integer;
    procedure ReadFonts(ANode: TDOMNode);
    procedure ReadHeaderFooter(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadHyperlinks(ANode: TDOMNode);
    procedure ReadMergedCells(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadNumFormats(ANode: TDOMNode);
    procedure ReadPageMargins(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadPageSetup(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadPalette(ANode: TDOMNode);
    procedure ReadPrintOptions(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadRow(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadSharedStrings(ANode: TDOMNode);
    procedure ReadSheetFormatPr(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadSheetList(ANode: TDOMNode; AList: TStrings);
    procedure ReadSheetViews(ANode: TDOMNode; AWorksheet: TsWorksheet);
    procedure ReadThemeElements(ANode: TDOMNode);
    procedure ReadThemeColors(ANode: TDOMNode);
    procedure ReadWorksheet(ANode: TDOMNode; AWorksheet: TsWorksheet);
  protected
    FFirstNumFormatIndexInFile: Integer;
    procedure AddBuiltinNumFormats; override;
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    destructor Destroy; override;
    procedure ReadFromStream(AStream: TStream; AParams: TsStreamParams = []); override;
  end;

  { TsSpreadOOXMLWriter }

  TsSpreadOOXMLWriter = class(TsCustomSpreadWriter)
  private
    FFirstNumFormatIndexInFile: Integer;
    vmlDrawingCounter: Integer;
  protected
    FDateMode: TDateMode;
    FPointSeparatorSettings: TFormatSettings;
    FSharedStringsCount: Integer;
    FFillList: array of PsCellFormat;
    FBorderList: array of PsCellFormat;
    procedure Get_rId(AWorksheet: TsWorksheet;
      out AComment_rId, AFirstHyperlink_rId, ADrawing_rId, ADrawingHF_rId: Integer);
  protected
    procedure AddBuiltinNumFormats; override;
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
    procedure WriteDefinedNames(AStream: TStream);
    procedure WriteDimension(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteDrawings(AWorksheet: TsWorksheet);
    procedure WriteDrawingRels(AWorksheet: TsWorksheet);
//    procedure WriteDrawingsOfSheet(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteFillList(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFont: TsFont; UseInStyleNode: Boolean);
    procedure WriteFontList(AStream: TStream);
    procedure WriteHeaderFooter(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteHyperlinks(AStream: TStream; AWorksheet: TsWorksheet; rId: Integer);
    procedure WriteMergedCells(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteNumFormatList(AStream: TStream);
    procedure WritePalette(AStream: TStream);
    procedure WritePageMargins(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WritePageSetup(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WritePrintOptions(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteSheetData(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteSheetFormatPr(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteSheetPr(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteSheetViews(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteStyleList(AStream: TStream; ANodeName: String);
    procedure WriteVmlDrawings(AWorksheet: TsWorksheet);
    procedure WriteVMLDrawings_Comments(AWorksheet: TsWorksheet);
    procedure WriteVMLDrawings_HeaderFooterImages(AWorksheet: TsWorksheet);
    procedure WriteVMLDrawingRels(AWorksheet: TsWorksheet);
    procedure WriteWorkbook(AStream: TStream);
    procedure WriteWorkbookRels(AStream: TStream);
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


implementation

uses
  variants, strutils, math, lazutf8, LazFileUtils, uriparser,
  {%H-}fpsPatches,
  fpsStrings, fpsStreams, fpsClasses, fpsImages;

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
     OOXML_PATH_XL_DRAWINGS        = 'xl/drawings/';
     OOXML_PATH_XL_DRAWINGS_RELS   = 'xl/drawings/_rels/';
     OOXML_PATH_XL_THEME           = 'xl/theme/theme1.xml';
     OOXML_PATH_XL_MEDIA           = 'xl/media/';

     { OOXML schemas constants }
     SCHEMAS_TYPES        = 'http://schemas.openxmlformats.org/package/2006/content-types';
     SCHEMAS_RELS         = 'http://schemas.openxmlformats.org/package/2006/relationships';
     SCHEMAS_DOC_RELS     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
     SCHEMAS_DOCUMENT     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
     SCHEMAS_WORKSHEET    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
     SCHEMAS_STYLES       = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
     SCHEMAS_STRINGS      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';
     SCHEMAS_COMMENTS     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
     SCHEMAS_DRAWING      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
     SCHEMAS_VMLDRAWING   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing';
     SCHEMAS_HYPERLINK    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
     SCHEMAS_IMAGE        = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
     SCHEMAS_SPREADML     = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

     { OOXML mime types constants }
{%H-}MIME_XML             = 'application/xml';
     MIME_RELS            = 'application/vnd.openxmlformats-package.relationships+xml';
     MIME_OFFICEDOCUMENT  = 'application/vnd.openxmlformats-officedocument';
{%H-}MIME_CORE            = 'application/vnd.openxmlformats-package.core-properties+xml';
     MIME_SPREADML        = MIME_OFFICEDOCUMENT + '.spreadsheetml';
     MIME_SHEET           = MIME_SPREADML + '.sheet.main+xml';
     MIME_WORKSHEET       = MIME_SPREADML + '.worksheet+xml';
     MIME_STYLES          = MIME_SPREADML + '.styles+xml';
     MIME_STRINGS         = MIME_SPREADML + '.sharedStrings+xml';
     MIME_COMMENTS        = MIME_SPREADML + '.comments+xml';
     MIME_DRAWING         = MIME_OFFICEDOCUMENT + '.drawing+xml'; // 'application/vnd.openxmlformats-officedocument.drawing+xml
     MIME_VMLDRAWING      = MIME_OFFICEDOCUMENT + '.vmlDrawing';

     LAST_PALETTE_INDEX   = 63;

     ROWHEIGHT_EPS        = 1E-2;

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

  TSharedFormulaData = class
    Worksheet: TsWorksheet;
    Row: Integer;
    Col: Integer;
    Formula: String;
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


{------------------------------------------------------------------------------}
{                           TsSpreadOOXMLReader                                }
{------------------------------------------------------------------------------}

constructor TsSpreadOOXMLReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FDateMode := XlsxSettings.DateMode;

  FSharedStrings := TStringList.Create;
  FFillList := TFPList.Create;
  FBorderList := TFPList.Create;
  FHyperlinkList := TFPList.Create;
  FCellFormatList := TsCellFormatList.Create(true);
  // Allow duplicates because xf indexes used in cell records cannot be found any more.
  FSharedFormulaBaseList := TFPList.Create;

  FPalette := TsPalette.Create;

  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';
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

  for j := FHyperlinkList.Count-1 downto 0 do
    TObject(FHyperlinkList[j]).Free;
  FHyperlinkList.Free;

  for j := FSharedStrings.Count-1 downto 0 do
    FSharedStrings.Objects[j].Free;
  FSharedStrings.Free;

  for j := FSharedFormulaBaseList.Count-1 downto 0 do
    TObject(FSharedFormulaBaseList[j]).Free;
  FSharedFormulaBaseList.Free;

  // FCellFormatList, FNumFormatList and FFontList are destroyed by ancestor

  FPalette.Free;
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

procedure TsSpreadOOXMLReader.ReadActiveSheet(ANode: TDOMNode;
  out ActiveSheetIndex: Integer);
var
  S: string;
  i: Integer;
begin
  ActiveSheetIndex := -1;
  if ANode = nil then
    Exit;
  S := GetAttrValue(ANode.FindNode('workbookView'), 'activeTab');
  if TryStrToInt(S, i) then
    ActiveSheetIndex := i;
end;

procedure TsSpreadOOXMLReader.ReadBorders(ANode: TDOMNode);

  function ReadBorderStyle(ANode: TDOMNode;
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
    if s = '' then
      exit;

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
  lCell: TCell;
  sharedFormulabase: TSharedFormulaData;
  datanode, tnode: TDOMNode;
  dataStr: String;
  formulaStr: String;
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

  // get row and column address
  addr := GetAttrValue(ANode, 'r');       // cell address, like 'A1'
  ParseCellString(addr, rowIndex, colIndex);

  // create cell
  if FIsVirtualMode then
  begin
    InitCell(rowIndex, colIndex, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := AWorksheet.AddCell(rowIndex, colIndex);

  // get style index
  s := GetAttrValue(ANode, 's');
  if s <> '' then begin
    ApplyCellFormatting(cell, StrToInt(s));
    fmt := Workbook.GetCellFormat(cell^.FormatIndex);
  end else
  begin
    InitFormatRecord(fmt);
    cell^.FormatIndex := 0;
  end;

  // get number format parameters
  numFmt := Workbook.GetNumberFormat(fmt.NumberFormatIndex);

  // get data
  datanode := ANode.FirstChild;
  dataStr := '';
  formulaStr := '';
  while Assigned(datanode) do
  begin
    nodeName := datanode.NodeName;
    if nodeName = 'v' then
      dataStr := GetNodeValue(datanode)
    else
    if nodeName = 'is' then
    begin
      tnode := datanode.FirstChild;
      while Assigned(tnode) do begin
        nodename := tnode.NodeName;
        if nodename = 't' then
          dataStr := dataStr + GetNodeValue(tnode);
        tnode := tnode.NextSibling;
      end;
    end else
    if (boReadFormulas in FWorkbook.Options) and (nodeName = 'f') then
    begin
      // Formula to cell
      formulaStr := GetNodeValue(datanode);

      s := GetAttrValue(datanode, 't');
      if s = 'shared' then
      begin
        // Shared formula
        s := GetAttrValue(datanode, 'ref');
        if (s <> '') then      // This defines the shared formula range
        begin
          AWorksheet.WriteFormula(cell, formulaStr);
          // We store the shared formula base in the SharedFormulaBaseList.
          // The list index is identical with the 'si' attribute of the node.
          sharedformulabase := TSharedFormulaData.Create;
          sharedformulabase.Worksheet := FWorksheet;
          sharedformulabase.Row := rowindex;
          sharedformulabase.Col := colindex;
          sharedformulabase.Formula := formulaStr;
          FSharedFormulaBaseList.Add(sharedformulabase);
        end else
        begin
          // Get index into the SharedFormulaBaseList...
          s := GetAttrValue(datanode, 'si');
          if s <> '' then
          begin
            sharedformulabase := TSharedFormulaData(FSharedFormulaBaseList[StrToInt(s)]);
            // ... and copy shared formula to destination cell
            InitCell(sharedformulabase.Row, sharedformulabase.Col, lCell);
            lCell.Formulavalue := sharedformulabase.Formula;
            lCell.Worksheet := sharedformulabase.Worksheet;
            FWorksheet.CopyFormula(@lCell, cell);
            cell^.ContentType := cctFormula;
          end;
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
  begin
    formulaStr := cell^.FormulaValue;
    AWorksheet.WriteBlank(cell);     // this erases the formula!!!
    cell^.FormulaValue := formulaStr;
  end else
  if (s = '') or (s = 'n') then begin
    // Number or date/time, depending on format
    number := StrToFloat(dataStr, FPointSeparatorSettings);
    if IsDateTimeFormat(numFmt) then
    begin
      if not IsTimeIntervalFormat(numFmt) then   // no correction of time origin for "time interval" format
        number := ConvertExcelDateTimeToDateTime(number, FDateMode);
      AWorksheet.WriteDateTime(cell, number);
    end
    else if IsTextFormat(numFmt) then
      AWorksheet.WriteText(cell, dataStr)
    else
      AWorksheet.WriteNumber(cell, number);
  end
  else
  if s = 's' then begin
    // String from shared strings table
    sstIndex := StrToInt(dataStr);
    AWorksheet.WriteText(cell, FSharedStrings[sstIndex]);
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
  if (s = 'str') or (s = 'inlineStr') then begin
    // literal string
    formulaStr := cell^.FormulaValue;
    AWorksheet.WriteText(cell, datastr);
    cell^.FormulaValue := formulaStr;
  end else
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
  numFmtIndex, fillIndex, borderIndex: Integer;
  numFmtStr: String;
  numFmtParams: TsNumFormatParams;
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
        numFmtStr := NumFormatList[numFmtIndex];
        if SameText(numFmtStr, 'General') then
          numFmtParams := nil
        else
        begin
          fmt.NumberFormatIndex := Workbook.AddNumberFormat(numFmtStr);
          numFmtParams := Workbook.GetNumberFormat(fmt.NumberFormatIndex);
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
      s2 := GetAttrValue(node, 'applyFont');
      if (s1 <> '') and (s2 <> '0') then
      begin
        fnt := TsFont(FFontList.Items[StrToInt(s1)]);
        fmt.FontIndex := Workbook.FindFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
        if fmt.FontIndex = -1 then
          fmt.FontIndex := Workbook.AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
        if fmt.FontIndex > 0 then
          Include(fmt.UsedFormattingFields, uffFont);
      end;

      s1 := GetAttrValue(node, 'fillId');
      s2 := GetAttrValue(node, 'applyFill');
      if (s1 <> '') and (s2 <> '0') then
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
      s2 := GetAttrValue(node, 'applyBorder');
      if (s1 <> '') and (s2 <> '0') then
      begin
        borderIndex := StrToInt(s1);
        borderData := TBorderListData(FBorderList[borderIndex]);
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

            s1 := GetAttrValue(childNode, 'readingOrder');
            if (s1 = '1') or (s1 = '2') then
              fmt.BiDiMode := TsBiDiMode(StrToInt(s1));

            s1 := GetAttrValue(childNode, 'wrapText');
            if (s1 <> '') and (s1 <> '0') then
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
      if fmt.BiDiMode <> bdDefault then
        Include(fmt.UsedFormattingFields, uffBiDi);
      FCellFormatList.Add(fmt);
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

procedure TsSpreadOOXMLReader.ReadCols(ANode: TDOMNode; AWorksheet: TsWorksheet);
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
begin
  if ANode = nil then
    exit;

  colNode := ANode.FirstChild;
  while Assigned(colNode) do begin
    nodeName := colNode.NodeName;
    if nodename = 'col' then
    begin
      s := GetAttrValue(colNode, 'min');
      if s <> '' then col1 := StrToInt(s)-1 else col1 := 0;

      s := GetAttrValue(colNode, 'max');
      if s <> '' then col2 := StrToInt(s)-1 else col2 := col1;

      s := GetAttrValue(colNode, 'customWidth');
      if s = '1' then begin
        s := GetAttrValue(colNode, 'width');
        if (s <> '') and TryStrToFloat(s, w, FPointSeparatorSettings) then
        begin
          if SameValue(w, AWorksheet.ReadDefaultColWidth(suChars), EPS) then  // is this needed?
            lCol.ColWidthType := cwtDefault
          else
            lCol.ColWidthType := cwtCustom;
          lCol.Width := FWorkbook.ConvertUnits(w, suChars, FWorkbook.Units);
        end;
      end else begin
        lCol.ColWidthType := cwtDefault;
        lCol.Width := AWorksheet.ReadDefaultColWidth(FWorkbook.Units);
      end;

      s := GetAttrValue(colNode, 'style');
      if s <> '' then begin
        idx := FCellFormatList.FindIndexOfID(StrToInt(s));
        fmt := FCellFormatList.Items[idx];
        lCol.FormatIndex := FWorkbook.AddCellFormat(fmt^);
      end else
        lCol.FormatIndex := 0;

      if (lCol.ColWidthType = cwtCustom) or (lCol.FormatIndex > 0) then
        for col := col1 to Min(col2, FLastCol) do
          AWorksheet.WriteColInfo(col, lCol);
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

procedure TsSpreadOOXMLReader.ReadDefinedNames(ANode: TDOMNode);
var
  node: TDOMNode;
  nodeName: String;
  r1,c1,r2,c2: Cardinal;
  id, j, p: Integer;
  sheet: TsWorksheet;
  localSheetID: String;
  namestr: String;
  s, sheetname: String;
  L: TStringList;
begin
  if ANode = nil then
    exit;
  node := ANode.FirstChild;
  while node <> nil do begin
    nodename := node.NodeName;
    if nodename = 'definedName' then
    begin
      localSheetID := GetAttrValue(node, 'localSheetId');
      if (localSheetID = '') or not TryStrToInt(localSheetID, id) then
      begin
        FWorkbook.AddErrorMsg('no/invalid localID in "definedName" node');
        node := node.NextSibling;
        Continue;
      end;
      namestr := GetAttrValue(node, 'name');
      sheet := FWorkbook.GetWorksheetByIndex(id);
      if namestr = '_xlnm.Print_Area' then
      begin
        L := TStringList.Create;
        try
          L.Delimiter := ',';
          L.StrictDelimiter := true;
          L.DelimitedText := GetNodeValue(node);
          for j:=0 to L.Count-1 do
          begin
            //s := ReplaceStr(L[j], '''', '');   // wp: replaced by next line because of Laz 1.0
            s := StringReplace(L[j], '''', '', [rfReplaceAll]);
            p := pos(':', s);
            if p = 0 then
            begin
              FWorkbook.AddErrorMsg('invalid cell range reference in "definedName" node');
              break;
            end;
            ParseSheetCellString(Copy(s, 1, p-1), sheetname, r1, c1);
            ParseSheetCellString(Copy(s, p+1, MaxInt), sheetname, r2, c2);
            sheet.PageLayout.AddPrintRange(r1, c1, r2, c2);
          end;
        finally
          L.Free;
        end;
      end else
      if nameStr = '_xlnm.Print_Titles' then
      begin
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
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLReader.ReadDimension(ANode: TDOMNode;
  AWorksheet: TsWorksheet);
var
  ref: String;
  r1, c1: Cardinal;
begin
  FLastRow := MaxInt;
  FLastCol := MaxInt;

  if ANode = nil then
    exit;

  ref := GetAttrValue(ANode, 'ref');
  if ref <> '' then
    ParseCellRangeString(ref, r1, c1, FLastRow, FLastCol);
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
      end else
      if nodeName = 'gradientFill' then
      begin
        // We do not support gradient fills yet. As a replacement, we read
        // the first color of the gradient and use it for a solid fill
        // This is required in order to keep the fill numbering intact.
        stopNode := patternNode.FirstChild;
        while Assigned(stopNode) do begin
          nodeName := stopNode.NodeName;
          if nodeName = 'stop' then begin
            colorNode := stopNode.FirstChild;
            while Assigned(colorNode) do begin
              nodeName := colorNode.NodeName;
              if nodeName = 'color' then
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
begin
  fnt := Workbook.GetDefaultFont;
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
    if (nodename = 'name') or (nodename = 'rFont') then
    begin
      s := GetAttrValue(node, 'val');
      if s <> '' then fntName := s;
      if nodename = 'rFont' then acceptDuplicates := false;
    end
    else
    if nodename = 'sz' then
    begin
      s := GetAttrValue(node, 'val');
      if s <> '' then fntSize := StrToFloat(s, FPointSeparatorSettings);
    end
    else
    if nodename = 'b' then
    begin
      if GetAttrValue(node, 'val') <> 'false' then
        fntStyles := fntStyles + [fssBold];
    end
    else
    if nodename = 'i' then
    begin
      if GetAttrValue(node, 'val') <> 'false' then
        fntStyles := fntStyles + [fssItalic];
    end
    else
    if nodename = 'u' then
    begin
      if GetAttrValue(node, 'val') <> 'false' then
        fntStyles := fntStyles+ [fssUnderline]
    end
    else
    if nodename = 'strike' then
    begin
      if GetAttrValue(node, 'val') <> 'false' then
        fntStyles := fntStyles + [fssStrikeout];
    end
    else
    if nodename = 'vertAlign' then
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
    if nodename = 'color' then
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

procedure TsSpreadOOXMLReader.ReadHeaderFooter(ANode: TDOMNode;
  AWorksheet: TsWorksheet);
var
  node: TDOMNode;
  nodeName: String;
  s: String;

  function FixLineEnding(AString: String): String;
  begin
    Result := StringReplace(AString, #10, LineEnding, [rfReplaceAll]);
  end;

begin
  if ANode = nil then
    exit;

  s := GetAttrValue(ANode, 'differentOddEven');
  if s = '1' then
    with AWorksheet.PageLayout do Options := Options + [poDifferentOddEven];

  s := GetAttrValue(ANode, 'differentFirst');
  if s = '1' then
    with AWorksheet.PageLayout do Options := Options + [poDifferentFirst];

  node := ANode.FirstChild;
  while node <> nil do
  begin
    nodeName := node.NodeName;
    case nodeName of
      'firstHeader': AWorksheet.PageLayout.Headers[0] := FixLineEnding(GetNodeValue(node));
      'oddHeader'  : AWorksheet.PageLayout.Headers[1] := FixLineEnding(GetNodeValue(node));
      'evenHeader' : AWorksheet.PageLayout.Headers[2] := FixLineEnding(GetNodeValue(node));
      'firstFooter': AWorksheet.PageLayout.Footers[0] := FixLineEnding(GetNodeValue(node));
      'oddFooter'  : AWorksheet.PageLayout.Footers[1] := FixLineEnding(GetNodeValue(node));
      'evenFooter' : AWorksheet.PageLayout.Footers[2] := FixLineEnding(GetNodeValue(node));
    end;
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
          if s = SCHEMAS_HYPERLINK then
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
  id: Integer;
begin
  if Assigned(ANode) then
  begin
    node := ANode.FirstChild;
    while Assigned(node) do
    begin
      nodeName := node.NodeName;
      if nodeName = 'numFmt' then
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
  AWorksheet: TsWorksheet);
var
  s: String;
begin
  if (ANode = nil) or (AWorksheet = nil) then     // just to make sure...
    exit;

  s := GetAttrValue(ANode, 'left');
  if s <> '' then
    AWorksheet.PageLayout.LeftMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));

  s := GetAttrValue(ANode, 'right');
  if s <> '' then
    AWorksheet.PageLayout.RightMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));

  s := GetAttrValue(ANode, 'top');
  if s <> '' then
    AWorksheet.PageLayout.TopMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));

  s := GetAttrValue(ANode, 'bottom');
  if s <> '' then
    AWorksheet.PageLayout.BottomMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));

  s := GetAttrValue(ANode, 'header');
  if s <> '' then
    AWorksheet.PageLayout.HeaderMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));

  s := GetAttrValue(ANode, 'footer');
  if s <> '' then
    AWorksheet.PageLayout.FooterMargin := PtsToMM(HtmlLengthStrToPts(s, 'in'));
end;

procedure TsSpreadOOXMLReader.ReadPageSetup(ANode: TDOMNode;
  AWorksheet: TsWorksheet);
var
  s: String;
  n: Integer;
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
      AWorksheet.PageLayout.PageWidth := PAPER_SIZES[n, 0];
      AWorksheet.PageLayout.PageHeight := PAPER_SIZES[n, 1];
    end;
  end;

  // Orientation
  s := GetAttrValue(ANode, 'orientation');
  if s = 'portrait' then
    AWorksheet.PageLayout.Orientation := spoPortrait
  else if s = 'landscape' then
    AWorksheet.PageLayout.Orientation := spoLandscape;

  // Scaling factor
  s := GetAttrValue(ANode, 'scale');
  if s <> '' then
    AWorksheet.PageLayout.ScalingFactor := StrToInt(s);
    // poFitPages is automatically excluded

  // Fit print job to pages
  s := GetAttrValue(ANode, 'fitToHeight');
  if s <> '' then
    AWorksheet.PageLayout.FitHeightToPages := StrToInt(s);
    // poFitPages is automatically included.

  s := GetAttrValue(ANode, 'fitToWidth');
  if s <> '' then
    AWorksheet.PageLayout.FitWidthToPages := StrToInt(s);
    // poFitPages is automatially included.

  // First page number
  s := GetAttrValue(ANode, 'firstPageNumber');
  if s <> '' then
    AWorksheet.PageLayout.StartPageNumber := StrToInt(s);

  s := GetAttrValue(ANode, 'useFirstPageNumber');
  with AWorksheet.PageLayout do
    if s = '1' then
      Options := Options + [poUseStartPageNumber]
    else
      Options := Options - [poUseStartPageNumber];

  // Print order
  s := GetAttrValue(ANode, 'pageOrder');
  if s = 'overThenDown' then
    with AWorksheet.PageLayout do Options := Options + [poPrintPagesByRows];

  // Monochrome
  s := GetAttrValue(ANode, 'blackAndWhite');
  if s = '1' then
    with AWorksheet.PageLayout do Options := Options + [poMonochrome];

  // Quality
  s := GetAttrValue(ANode, 'draft');
  if s = '1' then
    with AWorksheet.PageLayout do Options := Options + [poDraftQuality];
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
    if nodename = 'indexedColors' then
    begin
      colornode := node.FirstChild;
      while Assigned(colornode) do
      begin
        nodename := colornode.NodeName;
        if nodename = 'rgbColor' then begin
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
  AWorksheet: TsWorksheet);
var
  s: String;
begin
  if ANode = nil then
    exit;
  s := GetAttrValue(ANode, 'headings');
  if (s = '1') then
    with AWorksheet.PageLayout do Options := Options + [poPrintHeaders];

  s := GetAttrValue(ANode, 'gridLines');
  if (s = '1') then
    with AWorksheet.PageLayout do Options := Options + [poPrintGridLines];
end;

procedure TsSpreadOOXMLReader.ReadRow(ANode: TDOMNode; AWorksheet: TsWorksheet);
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
  s := GetAttrValue(ANode, 'customHeight');
  if s = '1' then
    lRow.RowHeightType := rhtCustom
  else
    lRow.RowHeightType := rhtAuto;

  { Row height value, in points - if there is no "ht" attribute we assume that
    it is the default row height }
  s := GetAttrValue(ANode, 'ht');
  if s = '' then begin
    lRow.Height := AWorksheet.ReadDefaultRowHeight(FWorkbook.Units);
    lRow.RowHeightType := rhtDefault;
  end else
    lRow.Height := FWorkbook.ConvertUnits(
      StrToFloat(s, FPointSeparatorSettings),
      suPoints,
      FWorkbook.Units
    );

  { Row index }
  s := GetAttrValue(ANode, 'r');
  r := StrToInt(s) - 1;

  { Row format }
  lRow.FormatIndex := 0;  // Default format
  s := GetAttrValue(ANode, 'customFormat');
  if s = '1' then begin
    s := GetAttrValue(ANode, 's');
    if s <> '' then begin
      idx := FCellFormatList.FindIndexOfID(StrToInt(s));
      fmt := FCellFormatList.Items[idx];
      lRow.FormatIndex := FWorkbook.AddCellFormat(fmt^);
    end;
  end;

  { Write out }
  if (lRow.RowHeightType <> rhtDefault) or (lRow.FormatIndex <> 0) then
    AWorksheet.WriteRowInfo(r, lRow);
end;

procedure TsSpreadOOXMLReader.ReadSharedStrings(ANode: TDOMNode);
var
  valuenode: TDOMNode;
  childnode: TDOMNode;
  nodename: String;
  totaltxt, sval: String;
  fntIndex: Integer;
  rtParams: TsRichTextParams;
  ms: TMemoryStream;
  fnt: TsFont;
begin
  while Assigned(ANode) do begin
    if ANode.NodeName = 'si' then begin
      totaltxt := '';
      SetLength(rtParams, 0);
      valuenode := ANode.FirstChild;
      while valuenode <> nil do begin
        nodename := valuenode.NodeName;
        if nodename = 't' then
          // this is unformatted text
          totaltxt := GetNodeValue(valuenode)
        else
        if nodename = 'r' then begin
          // all rich-text formatted texts are defined by r nodes
          fntIndex := -1;
          childnode := valuenode.FirstChild;
          while childnode <> nil do begin
            nodename := childnode.NodeName;
            if nodename = 't' then
            begin
              sval := GetNodeValue(childNode);
              totaltxt := totaltxt + sval;
            end
            else if nodename = 'rPr' then begin
              fntIndex := ReadFont(childnode);
              // Here we store the font in the internal font list of the reader.
              // But this fontindex may be different from the one needed for the
              // workbook's font list. We fix this here.
              fnt := TsFont(FFontList[fntIndex]);
              fntIndex := Workbook.FindFont(fnt.FontName, fnt.Size, fnt.style, fnt.Color, fnt.Position);
              if fntIndex = -1 then
                fntIndex := Workbook.AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color, fnt.Position);
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
  AWorksheet: TsWorksheet);
var
  w, h: Double;
  s: String;
begin
  if ANode = nil then
    exit;

  s := GetAttrValue(ANode, 'defaultColWidth');   // is in characters
  if (s <> '') and TryStrToFloat(s, w, FPointSeparatorSettings) then
    AWorksheet.WriteDefaultColWidth(w, suChars);

  s := GetAttrValue(ANode, 'defaultRowHeight');  // is in points
  if (s <> '') and TryStrToFloat(s, h, FPointSeparatorSettings) then
    AWorksheet.WriteDefaultRowHeight(h, suPoints);
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
  actRow, actCol: Cardinal;
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

      s := GetAttrValue(sheetViewNode, 'zoomScale');
      if s <> '' then
        AWorksheet.ZoomFactor := StrToFloat(s, FPointSeparatorSettings) * 0.01;

      s := GetAttrValue(sheetViewNode, 'rightToLeft');
      if s = '0' then
        AWorksheet.BiDiMode := bdLTR
      else if s = '1' then
        AWorksheet.BiDiMode := bdRTL;

      childNode := sheetViewNode.FirstChild;
      while Assigned(childNode) do begin
        nodeName := childNode.NodeName;
        if nodeName = 'pane' then begin
          s := GetAttrValue(childNode, 'state');
          if s = 'frozen' then begin
            AWorksheet.Options := AWorksheet.Options + [soHasFrozenPanes];
            s := GetAttrValue(childNode, 'xSplit');
            if s <> '' then AWorksheet.LeftPaneWidth := round(StrToFloat(s, FPointSeparatorSettings));
            s := GetAttrValue(childNode, 'ySplit');
            if s <> '' then AWorksheet.TopPaneHeight := round(StrToFloat(s, FPointSeparatorSettings));
          end;
        end else
        if nodeName = 'selection' then begin
          s := GetAttrValue(childnode, 'activeCell');
          if s <> '' then
          begin
            ParseCellString(s, actRow, actCol);
            AWorksheet.SelectCell(actRow, actCol);
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
      ReadRow(rownode, AWorksheet);
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

procedure TsSpreadOOXMLReader.ReadFromStream(AStream: TStream;
  AParams: TsStreamParams = []);
var
  Doc : TXMLDocument;
  RelsNode: TDOMNode;
  SheetList: TStringList;
  i, j: Integer;
  fn: String;
  fn_comments: String;
  XMLStream: TStream;
  actSheetIndex: Integer;

  function CreateXMLStream: TStream;
  begin
    if boFileStream in FWorkbook.Options then
      Result := TFileStream.Create(GetTempFileName, fmCreate)
    else
    if boBufStream in FWorkbook.Options then
      Result := TBufStream.Create(GetTempFileName, fmCreate)
    else
      Result := TMemoryStream.Create;
  end;

begin
  Unused(AParams);
  Doc := nil;
  SheetList := TStringList.Create;
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
        raise Exception.CreateFmt(rsDefectiveInternalFileStructure, ['xlsx']);
      ReadXMLStream(Doc, XMLStream);
      ReadFileVersion(Doc.DocumentElement.FindNode('fileVersion'));
      ReadDateMode(Doc.DocumentElement.FindNode('workbookPr'));
      ReadSheetList(Doc.DocumentElement.FindNode('sheets'), SheetList);
      //ReadDefinedNames(Doc.DocumentElement.FindNode('definedNames'));  -- don't read here because sheets do not yet exist
      ReadActiveSheet(Doc.DocumentElement.FindNode('bookViews'), actSheetIndex);
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
        ReadPalette(Doc.DocumentElement.FindNode('colors'));
        ReadFonts(Doc.DocumentElement.FindNode('fonts'));
        ReadFills(Doc.DocumentElement.FindNode('fills'));
        ReadBorders(Doc.DocumentElement.FindNode('borders'));
        ReadNumFormats(Doc.DocumentElement.FindNode('numFmts'));
        ReadCellXfs(Doc.DocumentElement.FindNode('cellXfs'));
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
        ReadSharedStrings(Doc.DocumentElement.FindNode('si'));
        FreeAndNil(Doc);
      end;
    finally
      XMLStream.Free;
    end;

    // read worksheets
    for i:=0 to SheetList.Count-1 do begin
      // Create worksheet
      FWorksheet := FWorkbook.AddWorksheet(SheetList[i], true);

      // unzip sheet file
      XMLStream := CreateXMLStream;
      try
        fn := OOXML_PATH_XL_WORKSHEETS + Format('sheet%d.xml', [i+1]);
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
      ReadDimension(Doc.DocumentElement.FindNode('dimension'), FWorksheet);
      ReadSheetViews(Doc.DocumentElement.FindNode('sheetViews'), FWorksheet);
      ReadSheetFormatPr(Doc.DocumentElement.FindNode('sheetFormatPr'), FWorksheet);
      ReadCols(Doc.DocumentElement.FindNode('cols'), FWorksheet);
      ReadWorksheet(Doc.DocumentElement.FindNode('sheetData'), FWorksheet);
      ReadMergedCells(Doc.DocumentElement.FindNode('mergeCells'), FWorksheet);
      ReadHyperlinks(Doc.DocumentElement.FindNode('hyperlinks'));
      ReadPrintOptions(Doc.DocumentElement.FindNode('printOptions'), FWorksheet);
      ReadPageMargins(Doc.DocumentElement.FindNode('pageMargins'), FWorksheet);
      ReadPageSetup(Doc.DocumentElement.FindNode('pageSetup'), FWorksheet);
      ReadHeaderFooter(Doc.DocumentElement.FindNode('headerFooter'), FWorksheet);

      FreeAndNil(Doc);

      { Comments:
        The comments are stored in separate "comments<n>.xml" files (n = 1, 2, ...)
        The relationship which comment belongs to which sheet file must be
        retrieved from the "sheet<n>.xml.rels" file (n = 1, 2, ...).
        The rels file contains also the second part of the hyperlink data. }
      fn := OOXML_PATH_XL_WORKSHEETS_RELS + Format('sheet%d.xml.rels', [i+1]);
      XMLStream := CreateXMLStream;
      try
        if UnzipToStream(AStream, fn, XMLStream) then
        begin
          // Find exact name of comments<n>.xml file
          ReadXMLStream(Doc, XMLStream);
          RelsNode := Doc.DocumentElement.FindNode('Relationship');
          fn_comments := FindCommentsFileName(RelsNode);
          // Get hyperlink data
          ReadHyperlinks(RelsNode);
          FreeAndNil(Doc);
        end else
        if (SheetList.Count = 1) then
          // If the workbook has only one sheet then the sheet.xml.rels file
          // is missing
          fn_comments := 'comments1.xml'
        else
          // This sheet does not have any cell comments at all
          fn_comments := '';
          //continue;
      finally
        XMLStream.Free;
      end;

      // Extract texts from the comments file found and apply to worksheet.
      if fn_comments <> '' then
      begin
        fn := OOXML_PATH_XL + fn_comments;
        XMLStream := CreateXMLStream;
        try
          if UnzipToStream(AStream, fn, XMLStream) then
          begin
            ReadXMLStream(Doc, XMLStream);
            ReadComments(Doc.DocumentElement.FindNode('commentList'), FWorksheet);
            FreeAndNil(Doc);
          end;
        finally
          XMLStream.Free;
        end;
      end;

      // Add hyperlinks to cells
      ApplyHyperlinks(FWorksheet);

      // Active worksheet
      if i = actSheetIndex then
        FWorkbook.SelectWorksheet(FWorksheet);
    end;  // for

    // 2nd run for the workbook.xml file
    // Read defined names
    XMLStream := CreateXMLStream;
    try
      if not UnzipToStream(AStream, OOXML_PATH_XL_WORKBOOK, XMLStream) then
        raise Exception.CreateFmt(rsDefectiveInternalFileStructure, ['xlsx']);
      ReadXMLStream(Doc, XMLStream);
      ReadDefinedNames(Doc.DocumentElement.FindNode('definedNames'));
      FreeAndNil(Doc);
    finally
      XMLStream.Free;
    end;

  finally
    SheetList.Free;
    FreeAndNil(Doc);
  end;
end;

                                             (*
procedure TsSpreadOOXMLReader.ReadFromStream(AStream: TStream);
begin
  Unused(AStream);
  raise Exception.Create('[TsSpreadOOXMLReader.ReadFromStream] '+
                         'Method not implemented. Use "ReadFromFile" instead.');
end;
                                               *)

{------------------------------------------------------------------------------}
{                             TsSpreadOOXMLWriter                              }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Constructor of the OOXML writer

  Defines the date mode and the limitations of the file format.
  Initializes the format settings to be used when writing to xml.
-------------------------------------------------------------------------------}
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
  FLimitations.MaxSheetNameLength := 31;
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

{ Calculates the rIds for comments, hyperlinks, image, and
  header/footer images of the specified worksheet }
procedure TsSpreadOOXMLWriter.Get_rId(AWorksheet: TsWorksheet;
  out AComment_rId, AFirstHyperlink_rId, ADrawing_rId, ADrawingHF_rId: Integer);
var
  next_rId: Integer;
begin
  AComment_rId := -1;
  AFirstHyperlink_rId := -1;
  ADrawing_rId := -1;
  ADrawingHF_rId := -1;
  next_rId := 1;

  // Comments first
  if AWorksheet.Comments.Count > 0 then
  begin
    AComment_rId := next_rId;
    inc(next_rId, 2);  // there are two .rels entries in case of comments
  end;

  // Embedded images next
  if AWorksheet.GetImageCount > 0 then
  begin
    ADrawing_rId := next_rId;
    inc(next_rId);
  end;

  // HeaderFooter images next
  if AWorksheet.PageLayout.HasHeaderFooterImages then
  begin
    ADrawingHF_rId := next_rId;
    inc(next_rId);
  end;

  // Hyperlinks at the end because it is not clear how many rIds will be
  // used without analyzing the hyperlink.
  if AWorksheet.Hyperlinks.Count > 0 then
    AFirstHyperlink_rId := next_rId;
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
  // lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair,
  // lsMediumDash, lsDashDot, lsMediumDashDot, lsDashDotDot, lsMediumDashDotDot,
  // lsSlantDashDot
  LINESTYLE_NAME: Array[TsLineStyle] of String = (
     'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair',
     'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot');

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
      styleName := LINESTYLE_NAME[AFormatRecord^.BorderStyles[ABorder].LineStyle];

      // Border color
      rgb := AFormatRecord^.BorderStyles[ABorder].Color;
      colorStr := ColorToHTMLColorStr(rgb, true);
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

procedure TsSpreadOOXMLWriter.WriteCols(AStream: TStream; AWorksheet: TsWorksheet);
var
  lCol: PCol;
  c: Integer;
  w: Single;
  customWidth: String;
  customStyle: String;
begin
  AppendToStream(AStream,
    '<cols>');

  for c:=0 to AWorksheet.GetLastColIndex do begin
    customWidth := '';
    customStyle := '';
    lCol := AWorksheet.FindCol(c);

    // The column width is needed in suChars here.
    w := AWorksheet.ReadDefaultColWidth(suChars);
    if lCol <> nil then begin
      if lCol^.ColWidthType = cwtCustom then begin
        w := FWorkbook.ConvertUnits(lCol^.Width, FWorkbook.Units, suChars);
        customWidth := 'customWidth="1" ';
      end;
      if lCol^.FormatIndex > 0 then customStyle := Format('style="%d" ', [lCol^.FormatIndex]);
    end;
    AppendToStream(AStream, Format(
      '<col min="%d" max="%d" width="%.2f" %s%s />',
      [c+1, c+1, w, customWidth, customStyle], FPointSeparatorSettings)
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

  // Footer
  AppendToStream(FSComments[FCurSheetNum],
      '</commentList>');
  AppendToStream(FSComments[FCurSheetNum],
    '</comments>');
end;

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
      fc := Format('rgb="%s"', [Copy(ColorToHTMLColorStr(FFillList[i]^.Background.FgColor), 2, MaxInt)]);
    if FFillList[i]^.Background.BgColor = scTransparent then
      bc := 'auto="1"'
    else
      bc := Format('rgb="%s"', [Copy(ColorToHTMLColorStr(FFillList[i]^.Background.BgColor), 2, MaxInt)]);
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
  s := s + Format('<sz val="%g" />', [AFont.Size], FPointSeparatorSettings);
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
    s := s + Format('<color rgb="%s" />', [Copy(ColorToHTMLColorStr(AFont.Color), 2, MaxInt)]);
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
begin
  AppendToStream(AStream, Format(
    '<fonts count="%d">', [Workbook.GetFontCount]));
  for i:=0 to Workbook.GetFontCount-1 do begin
    font := Workbook.GetFont(i);
    WriteFont(AStream, font, true);
  end;
  AppendToStream(AStream,
    '</fonts>');
end;

procedure TsSpreadOOXMLWriter.WriteHeaderFooter(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  s: String;
begin
  with AWorksheet.PageLayout do
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

    if poDifferentFirst in AWorksheet.PageLayout.Options then
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
  AWorksheet: TsWorksheet; rId: Integer);
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

  AVLNode := AWorksheet.Hyperlinks.FindLowest;
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
    txt := UTF8TextToXMLText(AWorksheet.ReadAsText(hyperlink^.Row, hyperlink^.Col));
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
    parser := TsNumFormatParser.Create(numFmtStr, Workbook.FormatSettings);
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
  AWorksheet: TsWorksheet);
begin
  with AWorksheet.PageLayout do
    AppendToStream(AStream, Format(
      '<pageMargins left="%g" right="%g" top="%g" bottom="%g" header="%g" footer="%g" />', [
      mmToIn(LeftMargin), mmToIn(RightMargin), mmToIn(TopMargin), mmToIn(BottomMargin),
      mmToIn(HeaderMargin), mmToIn(FooterMargin) ],
      FPointSeparatorSettings
    ));
end;

procedure TsSpreadOOXMLWriter.WritePageSetup(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  s: String;
  i: Integer;
begin
  s := '';

  // Paper size
  for i:=0 to High(PAPER_SIZES) do
    if (SameValue(PAPER_SIZES[i,0], AWorksheet.PageLayout.PageHeight) and
        SameValue(PAPER_SIZES[i,1], AWorksheet.PageLayout.PageWidth))
    or (SameValue(PAPER_SIZES[i,1], AWorksheet.PageLayout.PageHeight) and
        SameValue(PAPER_SIZES[i,0], AWorksheet.PageLayout.PageWidth))
    then begin
      s := Format('%s paperSize="%d"', [s, i]);
      break;
    end;

  if poFitPages in AWorksheet.PageLayout.Options then
  begin
    // Fit width to pages
    s := Format('%s fitToWidth="%d"', [s, AWorksheet.PageLayout.FitWidthToPages]);
    // Fit height to pages
    s := Format('%s fitToHeight="%d"', [s, AWorksheet.PageLayout.FitHeightToPages]);
  end else
    // Scaling factor
    s := Format('%s scale="%d"', [s, AWorksheet.PageLayout.ScalingFactor]);

  // Orientation
  s := Format('%s orientation="%s"', [
    s, IfThen(AWorksheet.PageLayout.Orientation = spoPortrait, 'portrait', 'landscape')
  ]);

  // First page number
  if poUseStartPageNumber in FWorksheet.PageLayout.Options then
    s := Format('%s useFirstPageNumber="1"', [s]);

  s := Format('%s firstPageNumber="%d"', [s, AWorksheet.PageLayout.StartPageNumber]);

  // Print order
  if poPrintPagesByRows in AWorksheet.PageLayout.Options then
    s := s + ' pageOrder="overThenDown"';

  // Monochrome
  if poMonochrome in AWorksheet.PageLayout.Options then
    s := s + ' blackAndWhite="1"';

  // Quality
  if poDraftQuality in AWOrksheet.PageLayout.Options then
    s := s + ' draft="1"';

  if s <> '' then
    AppendToStream(AStream,
      '<pageSetup' + s + ' />');
end;

procedure TsSpreadOOXMLWriter.WritePrintOptions(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  s: String;
begin
  s := '';
  if poPrintGridLines in AWorksheet.PageLayout.Options then
    s := s + ' gridLines="1"';
  if poPrintHeaders in AWorksheet.PageLayout.Options then
    s := s + ' headings="1"';

  if s <> '' then
    AppendToStream(AStream,
      '<printOptions' + s + ' />');
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
  s: String;
begin
  AppendToStream(AStream,
      '<sheetData>');

  GetSheetDimensions(AWorksheet, r1, r2, c1, c2);

  if (boVirtualMode in Workbook.Options) then begin
    if Assigned(AWorksheet.OnWriteCellData) and
      (AWorksheet.VirtualColCount > 0) and (AWorksheet.VirtualRowCount > 0)
    then begin
      for r := 0 to r2 do begin
        row := AWorksheet.FindRow(r);
        s := '';
        if row <> nil then begin
          s := s + Format(' ht="%.2f"',
            [FWorkbook.ConvertUnits(row^.Height, FWorkbook.Units, suPoints)],
            FPointSeparatorSettings);
          if row^.RowHeightType = rhtCustom then
            s := s + ' customHeight="1"';
          if row^.FormatIndex > 0 then
            s := s + Format(' s="%d" customFormat="1"', [row^.FormatIndex]);
        end;
        AppendToStream(AStream, Format(
          '<row r="%d" spans="1:%d"%s>', [r+1, AWorksheet.VirtualColCount, s]));
        for c := 0 to c2 do begin
          lCell.Row := r; // to silence a compiler hint
          InitCell(lCell);
          value := varNull;
          styleCell := nil;
          AWorksheet.OnWriteCellData(AWorksheet, r, c, value, styleCell);
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
      row := AWorksheet.FindRow(r);
      s := '';
      if row <> nil then begin
        s := s + Format(' ht="%.2f"',
          [FWorkbook.ConvertUnits(row^.Height, FWorkbook.Units, suPoints)],
          FPointSeparatorSettings);
        if row^.RowHeightType = rhtCustom then
          s := s + ' customHeight="1"';
        if row^.FormatIndex > 0 then
          s := s + Format(' s="%d" customFormat="1"', [row^.FormatIndex]);
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
        cell := AWorksheet.FindCell(r, c);
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
  AWorksheet: TsWorksheet);
var
  w, h: Single;
begin
  // Excel has column width in characters, and row heights in pts.
  w := AWorksheet.ReadDefaultColWidth(suChars);
  h := AWorksheet.ReadDefaultRowHeight(suPoints);
  AppendToStream(AStream, Format(
    '<sheetFormatPr baseColWidth="10" defaultColWidth="%.2f" defaultRowHeight="%.2f" customHeight="true" />',
    [w, h],
    FPointSeparatorSettings));
end;

procedure TsSpreadOOXMLWriter.WriteSheetPr(AStream: TStream; AWorksheet: TsWorksheet);
var
  s: String;
begin
  s := '';
  if (AWorksheet.PageLayout.FitWidthToPages > 0) or
     (AWorksheet.PageLayout.FitHeightToPages > 0) then
  s := s + ' fitToPage="1"';
  if s <> '' then s := '<pageSetUpPr' + s + ' />';

  if s <> '' then
    AppendToStream(AStream,
      '<sheetPr>' + s + '</sheetPr>');
end;

procedure TsSpreadOOXMLWriter.WriteSheetViews(AStream: TStream;
  AWorksheet: TsWorksheet);
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
begin
  // Show gridlines ?
  showGridLines := StrUtils.IfThen(soShowGridLines in AWorksheet.Options, '', ' showGridLines="0"');

  // Show headers?
  showHeaders := StrUtils.IfThen(soShowHeaders in AWorksheet.Options, '', ' showRowColHeaders="0"');

  // Zoom factor
  if boWriteZoomFactor in FWorkbook.Options then
    zoomscale := StrUtils.IfThen(SameValue(AWorksheet.ZoomFactor, 1.0, ZOOM_EPS), '',
      Format(' zoomScale="%.0f"', [AWorksheet.ZoomFactor*100]))
  else
    zoomscale := '';

  // BiDiMode
  case AWorksheet.BiDiMode of
    bdDefault: bidi := '';
    bdLTR    : bidi := ' rightToLeft="0"';
    bdRTL    : bidi := ' rightToLeft="1"';
  end;

  // Active cell
  if (AWorksheet.ActiveCellRow <> cardinal(-1)) and (AWorksheet.ActiveCellCol <> cardinal(-1)) then
    actCell := GetCellString(AWorksheet.ActiveCellRow, AWorksheet.ActiveCellCol) else
    actCell := '';

  // Selected tab?
  tabSel := StrUtils.IfThen(AWorksheet = FWorkbook.ActiveWorksheet, ' tabSelected="1"', '');

  // SheetView attributes
  attr := showGridLines + showHeaders + tabSel + zoomScale + bidi;

  // No frozen panes
  if not (soHasFrozenPanes in AWorksheet.Options) or
     ((AWorksheet.LeftPaneWidth = 0) and (AWorksheet.TopPaneHeight = 0))
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
    topRightCell := GetCellString(0, AWorksheet.LeftPaneWidth, [rfRelRow, rfRelCol]);
    bottomLeftCell := GetCellString(AWorksheet.TopPaneHeight, 0, [rfRelRow, rfRelCol]);
    bottomRightCell := GetCellString(AWorksheet.TopPaneHeight, AWorksheet.LeftPaneWidth, [rfRelRow, rfRelCol]);
    if (AWorksheet.LeftPaneWidth > 0) and (AWorksheet.TopPaneHeight > 0) then
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
        AWorksheet.LeftPaneWidth, AWorksheet.TopPaneHeight, bottomRightCell,
        topRightCell, topRightCell,
        bottomLeftCell, bottomLeftCell,
        actCell, actCell
      ]))
    end else
    if (AWorksheet.LeftPaneWidth > 0) then
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
        AWorksheet.LeftPaneWidth, topRightCell,
        actCell, actCell
      ]))
    end else
    if (AWorksheet.TopPaneHeight > 0) then
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
        AWorksheet.TopPaneHeight, bottomLeftCell,
        actCell, actCell
      ]));
    end;
  end;
end;

{ Writes the style list which the workbook has collected in its FormatList }
procedure TsSpreadOOXMLWriter.WriteStyleList(AStream: TStream; ANodeName: String);
var
//  styleCell: TCell;
  s, sAlign: String;
  fontID: Integer;
  numFmtParams: TsNumFormatParams;
  numFmtStr: String;
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
      numFmtParams := Workbook.GetNumberFormat(fmt^.NumberFormatIndex);
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
      case fmt^.HorAlignment of
        haLeft  : sAlign := sAlign + 'horizontal="left" ';
        haCenter: sAlign := sAlign + 'horizontal="center" ';
        haRight : sAlign := sAlign + 'horizontal="right" ';
      end;

    if (uffVertAlign in fmt^.UsedFormattingFields) and (fmt^.VertAlignment <> vaDefault)
    then
      case fmt^.VertAlignment of
        vaTop   : sAlign := sAlign + 'vertical="top" ';
        vaCenter: sAlign := sAlign + 'vertical="center" ';
        vaBottom: sAlign := sAlign + 'vertical="bottom" ';
      end;

    { Word wrap }
    if (uffWordWrap in fmt^.UsedFormattingFields) then
      sAlign := sAlign + 'wrapText="1" ';

    { BiDi mode }
    if (uffBiDi in fmt^.UsedFormattingFields) and (fmt^.BiDiMode <> bdDefault) then
      sAlign := sAlign + Format('readingOrder="%d" ', [Ord(fmt^.BiDiMode)]);

    { Fill }
    if (uffBackground in fmt^.UsedFormattingFields) then
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

procedure TsSpreadOOXMLWriter.WriteDrawings(AWorksheet: TsWorksheet);
var
  i: Integer;
  img: TsImage;
  r1, c1, r2, c2: Cardinal;
  roffs1, coffs1, roffs2, coffs2: Double;
  x, y, w, h: Double;
  descr: String;
begin
  if AWorksheet.GetImageCount= 0 then
    exit;

  SetLength(FSDrawings, FCurSheetNum + 1);
  FSDrawings[FCurSheetNum] := CreateTempStream(FWorkbook, Format('fpsD%d', [FCurSheetNum]));

  // Header
  AppendToStream(FSDrawings[FCurSheetNum],
    XML_HEADER,
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '+
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">');

  // Repeat for each image
  for i:=0 to AWorksheet.GetImageCount - 1 do
  begin
    img := AWorksheet.GetImage(i);
    if FWorkbook.GetEmbeddedObj(img.Index).ImageType = itUnknown then
      Continue;
    AWorksheet.CalcImageExtent(i, true,
      r1, c1, r2, c2,
      roffs1, coffs1, roffs2, coffs2,  // mm
      x, y, w, h);                     // mm;

    descr := ExtractFileName(FWorkbook.GetEmbeddedObj(img.index).Filename);
    if descr = '' then descr := 'image';

    AppendToStream(FSDrawings[FCurSheetNum],
      '<xdr:twoCellAnchor editAs="oneCell">');
    AppendToStream(FSDrawings[FCurSheetNum], Format(
        '<xdr:from>'+
          '<xdr:col>%d</xdr:col>' +
          '<xdr:colOff>%d</xdr:colOff>'+
          '<xdr:row>%d</xdr:row>'+
          '<xdr:rowOff>%d</xdr:rowOff>'+
        '</xdr:from>', [
        c1, mmToEMU(coffs1),
        r1, mmToEMU(roffs1)
    ]));
    AppendToStream(FSDrawings[FCurSheetNum], Format(
        '<xdr:to>'+
          '<xdr:col>%d</xdr:col>'+
          '<xdr:colOff>%d</xdr:colOff>'+
          '<xdr:row>%d</xdr:row>'+
          '<xdr:rowOff>%d</xdr:rowOff>'+
        '</xdr:to>', [
        c2, mmToEMU(coffs2),
        r2, mmToEMU(roffs2)
    ]));
    AppendToStream(FSDrawings[FCurSheetNum], Format(
        '<xdr:pic>'+
          '<xdr:nvPicPr>'+
            '<xdr:cNvPr id="%d" name="Grafik %d" descr="%s"/>'+  // 1, 2, orig file name
            '<xdr:cNvPicPr>'+
              '<a:picLocks noChangeAspect="1"/>'+
            '</xdr:cNvPicPr>'+
          '</xdr:nvPicPr>'+
          '<xdr:blipFill>'+
            '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId%d" cstate="print"/>'+  // 1
            '<a:stretch>'+
              '<a:fillRect/>'+
            '</a:stretch>'+
          '</xdr:blipFill>'+
          '<xdr:spPr>' +
            '<a:xfrm>'+
              '<a:off x="%d" y="%d"/>' +
              '<a:ext cx="%d" cy="%d"/>' +   // size in EMU
            '</a:xfrm>'+
            '<a:prstGeom prst="rect">'+
              '<a:avLst/>'+
            '</a:prstGeom>'+
          '</xdr:spPr>'+
        '</xdr:pic>' +
        '<xdr:clientData/>', [
        i+2, i+1, descr,
        i+1,
        mmToEMU(x), mmToEMU(y),
        mmToEMU(w), mmToEMU(h)
    ]));
    AppendToStream(FSDrawings[FCurSheetNum],
      '</xdr:twoCellAnchor>');
  end;
  AppendToStream(FSDrawings[FCurSheetNum],
    '</xdr:wsDr>');
end;

// For each sheet, writes a "drawingX.xml.rels" file to
// folder "../drawings/_rels". X matches the (1-base) sheet index.
// See also: WriteVmlDrawingRels
procedure TsSpreadOOXMLWriter.WriteDrawingRels(AWorksheet: TsWorksheet);
var
  i: Integer;
  ext: String;
  img: TsImage;
begin
  if (AWorksheet.GetImageCount = 0) then
    exit;

  SetLength(FSDrawingsRels, FCurSheetNum + 1);
  FSDrawingsRels[FCurSheetNum] := CreateTempStream(FWorkbook, Format('fpsDR%d', [FCurSheetNum]));

  // Header
  AppendToStream(FSDrawingsRels[FCurSheetNum],
    XML_HEADER + LineEnding,
    '<Relationships xmlns="' + SCHEMAS_RELS + '">' + LineEnding);

  // Repeat for each image
  for i:=0 to AWorksheet.GetImageCount - 1 do
  begin
    img := AWorksheet.GetImage(i);
    ext := GetImageTypeExt(FWorkbook.GetEmbeddedObj(img.Index).Imagetype);
    AppendToStream(FSDrawingsRels[FCurSheetNum], Format(
    '  <Relationship Id="rId%d" Type="%s" Target="../media/image%d.%s"/>' + LineEnding, [
       i+1, SCHEMAS_IMAGE, img.Index+1, ext
    ]));
  end;

  AppendToStream(FSDrawingsRels[FCurSheetNum],
    '</Relationships>');
end;
                             (*
procedure TsSpreadOOXMLWriter.WriteDrawingsOfSheet(AStream: TStream;
  AWorksheet: TsWorksheet; rId: Integer);
// Use stream FSDrawingS[sheetindex]
var
  i: Integer;
  AVLNode: TAVLTreeNode;
  hyperlink: PsHyperlink;
  target, bookmark: String;
begin
  // Keep in sync with WriteWorksheetRels !
  FNext_rID := IfThen(AWorksheet.Comments.Count = 0, 1, 3);

  AVLNode := AWorksheet.Hyperlinks.FindLowest;
  while AVLNode <> nil do begin
    inc(FNext_rID);
    AVLNode := AWorksheet.Hyperlinks.FindSuccessor(AVLNode);
  end;

  for i:=0 to AWorksheet.GetImageCount-1 do
  begin
    AppendToStream(AStream, Format(
      '<drawing r:id="rId%d" />', [FNext_rId]));
    inc(FNext_rId);
  end;
end;                           *)

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
procedure TsSpreadOOXMLWriter.WriteVmlDrawings(AWorksheet: TsWorksheet);
begin
  // At first write the VmlDrawings related to comments
  WriteVmlDrawings_Comments(AWorksheet);

  // Now write the vmlDrawings related to headers/footers
  WriteVmlDrawings_HeaderFooterImages(AWorksheet);
end;

procedure TsSpreadOOXMLWriter.WriteVMLDrawings_Comments(AWorksheet: TsWorksheet);
var
  comment: PsComment;
  fileindex: Integer;
  index: Integer;
  id: Integer;
begin
  if AWorksheet.Comments.Count = 0 then
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
  for comment in AWorksheet.Comments do
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
  AWorksheet: TsWorksheet);

  { AName = 'header' or 'footer'
    ATag  = 'L', 'C', 'R', 'x', or ' '
    AChar = 'H' or 'F' }
  procedure Process(AStream: TStream; AName: String; ATag, AChar: Char;
    AImage: TsHeaderFooterImage; var id, index: Integer);
  var
    fn: String;
  begin
    if AImage.Index = -1 then
      exit;
    if ATag = 'x' then
    begin
      FWorkbook.AddErrorMsg(rsOnlyOneHeaderFooterImageAllowed, [AName]);
      exit;
    end;
    if ATag = ' ' then
    begin
      FWorkbook.AddErrorMsg(rsIncorrectPositionOfImageInHeaderFooter, [AName]);
      exit;
    end;
    fn := ChangeFileExt(Workbook.GetEmbeddedObj(AImage.Index).FileName, '');
    if fn = '' then fn := 'image';
    AppendToStream(AStream, Format(
      ' <v:shape id="%s" o:spid="_x0000_s%d" type="#_x0000_t75"' + LineEnding +
      //    e.g.    "CH"         _x0000_s1025
      '   style=''position:absolute;margin-left:0;margin-top:0;width:12pt;height:12pt;z-index:%d''>' + LineEnding +
      //    e.g.                                                                      z-index:1
      '   <v:imagedata o:relid="rId%d" o:title="%s"/>' + LineEnding +
      //    e.g.               "rId1"          "arrow_down"
      '   <o:lock v:ext="edit" rotation="t" />' + LineEnding +
      ' </v:shape>' + LineEnding, [
      ATag + AChar, id, index, index, fn
    ]));
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
  if not AWorksheet.PageLayout.HasHeaderFooterImages then
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

  AWorksheet.PageLayout.GetImageSections(headerTags, footerTags);

  // Write the data for the image in each section of the header
  for sec in TsHeaderFooterSectionIndex do
  begin
    tagIndex := ord(sec) + 1;
    img := AWorksheet.PageLayout.HeaderImages[sec];
    Process(FSVmlDrawings[fileIndex], rsHeader, headerTags[tagIndex], 'H', img, id, index);
  end;
  // Repeat with footer
  for sec in TsHeaderFooterSectionIndex do
  begin
    img := AWorksheet.PageLayout.FooterImages[sec];
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
procedure TsSpreadOOXMLWriter.WriteVmlDrawingRels(AWorksheet: TsWorksheet);
var
  fileindex: Integer;
  sec: TsHeaderFooterSectionIndex;
  rId: Integer;
  img: TsHeaderFooterImage;
//  imgIdx: Integer;
//  imgName: String;
  ext: String;
begin
  if not AWorksheet.PageLayout.HasHeaderFooterImages then
    exit;

  fileIndex := Length(FSVmlDrawingsRels);
  if AWorksheet.Comments.Count > 0 then
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
    img := AWorksheet.PageLayout.HeaderImages[sec];
    if img.Index = -1 then
      continue;
//    imgName := FWorkbook.GetEmbeddedObj(img.Index).Name;
//    imgIdx := img.Index;
    ext := GetImageTypeExt(FWorkbook.GetEmbeddedObj(img.Index).ImageType);
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
    img := AWorksheet.PageLayout.FooterImages[sec];
    if img.Index = -1 then
      continue;
//    imgName := FWorkbook.GetEmbeddedObj(img.Index).Name;
//    imgIdx := img.Index;
//    imgIdx := FWorkbook.FindEmbeddedObj(imgName);
    ext := GetImageTypeExt(FWorkbook.GetEmbeddedObj(img.Index).Imagetype);
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

procedure TsSpreadOOXMLWriter.WriteWorksheetRels(AWorksheet: TsWorksheet);
var
  AVLNode: TAVLTreeNode;
  hyperlink: PsHyperlink;
  s: String;
  target, bookmark: String;
  rId_Comments, rId_Hyperlink, rId_Drawing, rId_DrawingHF: Integer;
begin
  // Extend stream array
  // NOTE: If no .rels file is written for this sheet at least an empty stream
  // must be provided to keep the numbering intact.
  SetLength(FSSheetRels, FCurSheetNum + 1);

  // Anything to write?
  if (AWorksheet.Comments.Count = 0) and (AWorksheet.Hyperlinks.Count = 0) and
     (AWorksheet.GetImageCount = 0) and not (AWorksheet.PageLayout.HasHeaderFooterImages)
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
  if AWorksheet.Comments.Count > 0 then
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
        s := Format('Id="rId%d" Target="%s" TargetMode="External" Type="%s"',
          [rId_Hyperlink, target, SCHEMAS_HYPERLINK]);
        AppendToStream(FSSheetRels[FCurSheetNum],
          '  <Relationship ' + s + ' />' + LineEnding);
        inc(rId_Hyperlink);
      end;
      AVLNode := AWorksheet.Hyperlinks.FindSuccessor(AVLNode);
    end;
  end;

  // Relationships for embedded images
  // relationship with to the ../drawings/drawingX.xml file containing all
  // image infos. X is the 1-base sheet index
  if AWorksheet.GetImageCount > 0 then
    AppendToStream(FSSheetRels[FCurSheetNum], Format(
      '  <Relationship Id="rId%d" Target="../drawings/drawing%d.xml" Type="%s" />' + LineEnding,
      [rId_Drawing, FCurSheetNum + 1, SCHEMAS_DRAWING]
    ));

  // Relationships for embedded header/footer images
  if AWorksheet.PageLayout.HasHeaderFooterImages then
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

  { --- _rels/.rels --- }
  AppendToStream(FSRelsRels,
    XML_HEADER + LineEnding);
  AppendToStream(FSRelsRels, Format(
    '<Relationships xmlns="%s">' + LineEnding, [SCHEMAS_RELS]));
  AppendToStream(FSRelsRels, Format(
    '  <Relationship Id="rId1" Target="xl/workbook.xml" Type="%s" />' + LineEnding,
      [SCHEMAS_DOCUMENT]));
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

{ Write folder "media" with embedded streams }
procedure TsSpreadOOXMLWriter.WriteMedia(AZip: TZipper);
var
  i: Integer;
  stream: TMemoryStream;
  embObj: TsEmbeddedObj;
  embName: String;
  ext: String;
begin
  for i:=0 to FWorkbook.GetEmbeddedObjCount-1 do
  begin
    embObj := FWorkbook.GetEmbeddedObj(i);
    stream := embObj.Stream;
    stream.Position := 0;
    ext := GetImageTypeExt(embObj.ImageType);
    embName := Format('image%d.%s', [i+1, ext]);
    AZip.Entries.AddFileEntry(stream, OOXML_PATH_XL_MEDIA + embname);
  end;
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
  for i := 0 to Workbook.GetWorksheetCount - 1 do
  begin
    FWorksheet := Workbook.GetWorksheetByIndex(i);
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

  { Workbook relations - Mark relation to all sheets }
  WriteWorkbookRels(FSWorkbookRels);
end;

procedure TsSpreadOOXMLWriter.WriteContentTypes;
var
  i,j: Integer;
  imgext: TStringList;
  ext: String;
  sheet: TsWorksheet;
begin
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

  if Workbook.GetEmbeddedObjCount > 0 then
  begin
    imgExt := TStringList.Create;
    try
      for i:=0 to Workbook.GetEmbeddedObjCount-1 do
      begin
        ext := GetImageTypeExt(Workbook.GetEmbeddedObj(i).ImageType);
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

  for i:=1 to Workbook.GetWorksheetCount do
  begin
    AppendToStream(FSContentTypes, Format(
      '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="%s" />' + LineEnding,
        [i, MIME_WORKSHEET]));
    sheet := Workbook.GetWorksheetByIndex(i-1);
    if sheet.GetImageCount > 0 then
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
  {
  AppendToStream(FSContentTypes,
      '<Override PartName="/docProps/core.xml" ContentType="' + MIME_CORE + '" />');
  }
  AppendToStream(FSContentTypes,
    '</Types>');
end;

procedure TsSpreadOOXMLWriter.WriteDefinedNames(AStream: TStream);
var
  sheet: TsWorksheet;
  stotal, srng, sheetname: String;
  i, j: Integer;
  prng: TsCellRange;
  firstIndex, lastIndex: Integer;
begin
  stotal := '';

  // Write print ranges and repeatedly printed rows and columns
  for i := 0 to Workbook.GetWorksheetCount-1 do
  begin
    sheet := Workbook.GetWorksheetByIndex(i);
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
    AppendtoStream(AStream,
      '<definedNames>' + stotal + '</definedNames>');
end;

procedure TsSpreadOOXMLWriter.WriteWorkbook(AStream: TStream);
var
  actTab: String;
  sheetName: String;
  counter: Integer;
begin
  actTab := IfThen(FWorkbook.ActiveWorksheet = nil, '',
    'activeTab="' + IntToStr(FWorkbook.GetWorksheetIndex(FWorkbook.ActiveWorksheet)) + '"');

  AppendToStream(AStream,
    XML_HEADER);
  AppendToStream(AStream, Format(
    '<workbook xmlns="%s" xmlns:r="%s">', [SCHEMAS_SPREADML, SCHEMAS_DOC_RELS]));
  AppendToStream(AStream,
      '<fileVersion appName="fpspreadsheet" />');
  AppendToStream(AStream,
      '<workbookPr defaultThemeVersion="124226" />');
  AppendToStream(AStream,
      '<bookViews>' +
        '<workbookView xWindow="480" yWindow="90" windowWidth="15195" windowHeight="12525" ' + actTab + '/>' +
      '</bookViews>');

  AppendToStream(AStream,
      '<sheets>');
  for counter:=1 to Workbook.GetWorksheetCount do
  begin
    sheetname := UTF8TextToXMLText(Workbook.GetWorksheetByIndex(counter-1).Name);
    AppendToStream(AStream, Format(
        '<sheet name="%s" sheetId="%d" r:id="rId%d" />',
        [sheetname, counter, counter]));
  end;
  AppendToStream(AStream,
      '</sheets>');

  WriteDefinedNames(AStream);

  AppendToStream(AStream,
      '<calcPr calcId="114210" />');
  AppendToStream(AStream,
    '</workbook>');
end;

procedure TsSpreadOOXMLWriter.WriteWorkbookRels(AStream: TStream);
var
  counter: Integer;
begin
  AppendToStream(AStream,
    XML_HEADER + LineEnding,
    '<Relationships xmlns="' + SCHEMAS_RELS + '">' + LineEnding);

  counter := 1;
  while counter <= Workbook.GetWorksheetCount do begin
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

procedure TsSpreadOOXMLWriter.WriteWorksheet(AWorksheet: TsWorksheet);
var
  rId_Comments: Integer;
  rId_FirstHyperlink: Integer;
  rId_Drawing, rId_DrawingHF: Integer;
begin
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
  WriteMergedCells(FSSheets[FCurSheetNum], AWorksheet);
  WriteHyperlinks(FSSheets[FCurSheetNum], AWorksheet, rId_FirstHyperlink);

  WritePrintOptions(FSSheets[FCurSheetNum], AWorksheet);
  WritePageMargins(FSSheets[FCurSheetNum], AWorksheet);
  WritePageSetup(FSSheets[FCurSheetNum], AWorksheet);
  WriteHeaderFooter(FSSheets[FCurSheetNum], AWorksheet);

  { This item is required for all embedded images.
    There must be a matching file in "drawingX.xml" file in "../drawings"
    which contains the image-related data of all images in this sheet.
    The file in turn requires an entry "drawingX.xml.rels" in the drawings rels
    folder }
  if AWorksheet.GetImageCount > 0 then
    AppendToStream(FSSheets[FCurSheetNum], Format(
      '<drawing r:id="rId%d" />', [rId_Drawing]));

  { This item is required for all comments of a worksheet.
    Comments have two entries in the sheet's .rels file, one for the
    "../comments.xml" file, and one for the "../drawings/vmlDrawingX.vml" file.
    The vmlDrawing file must have an entry "vmlDrawingX.vml.rels" in the drawings
    rels folder. }
  if AWorksheet.Comments.Count > 0 then
    AppendToStream(FSSheets[FCurSheetNum], Format(
      '<legacyDrawing r:id="rId%d" />', [rId_Comments]));

  { This item is required for all images embedded to a header/footer.
    There must be a corresponding "vmlDrawingX.vml" file in "../drawings". }
  if AWorksheet.PageLayout.HasHeaderFooterImages then
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
  DestroyTempStream(FSContentTypes);
  DestroyTempStream(FSRelsRels);
  DestroyTempStream(FSWorkbookRels);
  DestroyTempStream(FSWorkbook);
  DestroyTempStream(FSStyles);
  DestroyTempStream(FSSharedStrings);
  DestroyTempStream(FSSharedStrings_complete);
  for stream in FSSheets do DestroyTempStream(stream);
  SetLength(FSSheets, 0);
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
  txt := ResultingValue;
  if not ValidXMLText(txt) then
    Workbook.AddErrorMsg(
      rsInvalidCharacterInCell, [
      GetCellString(ARow, ACol)
    ]);

  { Write string to SharedString table }
  if Length(ACell^.RichTextParams) = 0 then
    // unformatted string
    AppendToStream(FSSharedStrings,
      '<si>' +
        '<t xml:space="preserve">' + txt + '</t>' +
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
      fnt := FWorkbook.GetFont(rtParam.FontIndex);
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
begin
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

