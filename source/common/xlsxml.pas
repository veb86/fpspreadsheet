{-------------------------------------------------------------------------------
Unit     : xlsxml

Implements a reader and writer for the SpreadsheetXML format.
This document was introduced by Microsoft for Excel XP and 2003.

REFERENCE: http://msdn.microsoft.com/en-us/library/aa140066%28v=office.15%29.aspx

AUTHOR   : Werner Pamler

LICENSE  : For details about the license, see the file
           COPYING.modifiedLGPL.txt included in the Lazarus distribution.
-------------------------------------------------------------------------------}

unit xlsxml;

{$ifdef fpc}
  {$mode objfpc}{$H+}
{$endif}

interface

uses
  Classes, SysUtils,
  laz2_xmlread, laz2_DOM,
  fpsTypes, fpsReaderWriter, fpsXMLCommon, xlsCommon;

type
  { TsSpreadExcelXMLReader }
  TsSpreadExcelXMLReader = class(TsSpreadXMLReader)
  private
    FDateMode: TDateMode;
    FPointSeparatorSettings: TFormatSettings;
    function ExtractDateTime(AText: String): TDateTime;

  protected
    FFirstNumFormatIndexInFile: Integer;
    procedure AddBuiltinNumFormats; override;

  protected
    procedure ReadAlignment(ANode: TDOMNode; var AFormat: TsCellFormat);
    procedure ReadBorder(ANode: TDOMNode; var AFormat: TsCellFormat);
    procedure ReadBorders(ANode: TDOMNode; var AFormat: TsCellFormat);
    procedure ReadCell(ANode: TDOMNode; AWorksheet: TsBasicWorksheet; ARow, ACol: Integer);
    procedure ReadCellProtection(ANode: TDOMNode; var AFormat: TsCellFormat);
    procedure ReadComment(ANode: TDOMNode; AWorksheet: TsBasicWorksheet; ACell: PCell);
    procedure ReadExcelWorkbook(ANode: TDOMNode);
    procedure ReadFont(ANode: TDOMNode; var AFormat: TsCellFormat);
    procedure ReadInterior(ANode: TDOMNode; var AFormat: TsCellFormat);
    procedure ReadNames(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadNumberFormat(ANode: TDOMNode; var AFormat: TsCellFormat);
    procedure ReadPageBreak(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadPageBreaks(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadPageSetup(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadPrint(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadRow(ANode: TDOMNode; AWorksheet: TsBasicWorksheet; ARow: Integer);
    procedure ReadStyle(ANode: TDOMNode);
    procedure ReadStyles(ANode: TDOMNode);
    procedure ReadTable(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadWorksheet(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadWorksheetOptions(ANode: TDOMNode; AWorksheet: TsBasicWorksheet);
    procedure ReadWorksheets(ANode: TDOMNode);

  public
    constructor Create(AWorkbook: TsBasicWorkbook); override;
    procedure ReadFromStream(AStream: TStream; APassword: String = '';
      AParams: TsStreamParams = []); override;
  end;


  { TsSpreadExcelXMLWriter }

  TsSpreadExcelXMLWriter = class(TsCustomSpreadWriter)
  private
    FDateMode: TDateMode;
    FPointSeparatorSettings: TFormatSettings;
    FFirstRow, FFirstCol: Cardinal;
    FlastRow, FLastCol: Cardinal;
    FPrevRow, FPrevCol: Cardinal;
    function GetCommentStr(ACell: PCell): String;
    function GetFormulaStr(ACell: PCell): String;
    function GetFrozenPanesStr(AWorksheet: TsBasicWorksheet; AIndent: String): String;
    function GetHyperlinkStr(ACell: PCell): String;
    function GetIndexStr(AIndex, APrevIndex: Integer): String;
    function GetLayoutStr(AWorksheet: TsBasicWorksheet): String;
    function GetMergeStr(ACell: PCell): String;
    function GetPageFooterStr(AWorksheet: TsBasicWorksheet): String;
    function GetPageHeaderStr(AWorksheet: TsBasicWorksheet): String;
    function GetPageMarginStr(AWorksheet: TsBasicWorksheet): String;
    function GetPrintStr(AWorksheet: TsBasicWorksheet): String;
    function GetStyleStr(AFormatIndex: Integer): String;
    procedure WriteCellNodes(AStream: TStream; AWorksheet: TsBasicWorksheet; ARow: Cardinal);
    procedure WriteColumns(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteExcelWorkbook(AStream: TStream);
    procedure WriteNames(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WritePageBreaks(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteRows(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteStyle(AStream: TStream; AIndex: Integer);
    procedure WriteStyles(AStream: TStream);
    procedure WriteTable(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteWorksheet(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteWorksheetOptions(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteWorksheets(AStream: TStream);

  protected
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: boolean; ACell: PCell); override;
    procedure WriteCellToStream(AStream: TStream; ACell: PCell); override;
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
    procedure WriteToStream(AStream: TStream; AParams: TsStreamParams = []); override;

  end;

  TExcelXmlSettings = record
    DateMode: TDateMode;
  end;

var
  { Default parameters for reading/writing }
  ExcelXmlSettings: TExcelXmlSettings = (
    DateMode: dm1900;
  );

  sfidExcelXML: TsSpreadFormatID;


implementation

uses
  StrUtils, DateUtils, Math, Variants,
  fpsStrings, fpsClasses, fpspreadsheet, fpsUtils, fpsNumFormat, fpsHTMLUtils,
  fpsExprParser;

const
  FMT_OFFSET   = 61;

  INDENT1      = '  ';
  INDENT2      = '    ';
  INDENT3      = '      ';
  INDENT4      = '        ';
  INDENT5      = '          ';
  NAMES_INDENT = INDENT2;
  NAME_INDENT  = INDENT3;
  TABLE_INDENT = INDENT2;
  ROW_INDENT   = INDENT3;
  COL_INDENT   = INDENT3;
  CELL_INDENT  = INDENT4;
  VALUE_INDENT = INDENT5;

  LF           = LineEnding;

const
  {TsFillStyle = (
    fsNoFill, fsSolidFill,
    fsGray75, fsGray50, fsGray25, fsGray12, fsGray6,
    fsStripeHor, fsStripeVert, fsStripeDiagUp, fsStripeDiagDown,
    fsThinStripeHor, fsThinStripeVert, fsThinStripeDiagUp, fsThinStripeDiagDown,
    fsHatchDiag, fsThinHatchDiag, fsThickHatchDiag, fsThinHatchHor) }
  FILL_NAMES: array[TsFillStyle] of string = (
    '', 'Solid',
//    'Solid', 'Solid', 'Solid', 'Solid', 'Solid',
    'Gray75', 'Gray50', 'Gray25', 'Gray125', 'Gray0625',
    'HorzStripe', 'VertStripe', 'DiagStripe', 'ReverseDiagStripe',
    'ThinHorzStripe', 'ThinVertStripe', 'ThinDiagStripe', 'ThinReverseDiagStripe',
    'DiagCross', 'ThinDiagCross', 'ThickDiagCross', 'ThinHorzCross'
  );

  {TsCellBorder = (cbNorth, cbWest, cbEast, cbSouth, cbDiagUp, cbDiagDown); }
  BORDER_NAMES: array[TsCellBorder] of string = (
    'Top', 'Left', 'Right', 'Bottom', 'DiagonalRight', 'DiagonalLeft'
  );

  {TsLineStyle = (
    lsThin, lsMedium, lsDashed, lsDotted, lsThick, lsDouble, lsHair,
    lsMediumDash, lsDashDot, lsMediumDashDot, lsDashDotDot, lsMediumDashDotDot,
    lsSlantDashDot) }
  LINE_STYLES: array[TsLineStyle] of string = (
    'Continuous', 'Continuous', 'Dash', 'Dot', 'Continuous', 'Double', 'Continuous',
    'Dash', 'DashDot', 'DashDot', 'DashDotDot', 'DashDotDot',
    'SlantDashDot'
  );
  LINE_WIDTHS: array[TsLineStyle] of Integer = (
    1, 2, 1, 1, 3, 3, 0,
    2, 1, 2, 1, 2,
    2
  );

  FALSE_TRUE: array[boolean] of string = ('False', 'True');

function GetCellContentTypeStr(ACell: PCell): String;
begin
  case ACell^.ContentType of
    cctNumber     : Result := 'Number';
    cctUTF8String : Result := 'String';
    cctDateTime   : Result := 'DateTime';
    cctBool       : Result := 'Boolean';
    cctError      : Result := 'Error';
  else
    raise EFPSpreadsheet.Create('Content type error in cell ' + GetCellString(ACell^.Row, ACell^.Col));
  end;
end;

{ Helper routine to rebuild the html content of the "ss:Data" nodes }
procedure RebuildChildNodes(ANode: TDOMNode; var AText: String);
var
  nodeName: String;
  s: String;
  i: Integer;
begin
  if ANode = nil then
    exit;
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = '#text' then
      AText := AText + ANode.NodeValue
    else begin
      s := '';
      for i := 0 to ANode.Attributes.Length-1 do
        s := Format('%s %s="%s"', [s, ANode.Attributes.Item[i].NodeName, ANode.Attributes.Item[i].NodeValue]);
      AText := Format('%s<%s%s>', [AText, nodeName, s]);
      s := '';
      RebuildChildNodes(ANode.FirstChild, s);
      if s <> '' then
        AText := Format('%s%s</%s>', [AText, s, nodeName]);
    end;
    ANode := ANode.NextSibling;
  end;
end;


{===============================================================================
                          TsSpreadExcelXMLReader
===============================================================================}

{@@ ----------------------------------------------------------------------------
  Constructor of the ExcelXML reader
-------------------------------------------------------------------------------}
constructor TsSpreadExcelXMLReader.Create(AWorkbook: TsBasicWorkbook);
begin
  inherited;

  // Cell formats (named "Styles" here).
  FCellFormatList := TsCellFormatList.Create(true);  // is destroyed by ancestor

  // Special version of FormatSettings using a point decimal separator for sure.
  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';
end;

procedure TsSpreadExcelXMLReader.AddBuiltinNumFormats;
begin
  FFirstNumFormatIndexInFile := 164;
  AddBuiltInBiffFormats(
    FNumFormatList, FWorkbook.FormatSettings, FFirstNumFormatIndexInFile-1
  );
end;

{@@ ----------------------------------------------------------------------------
  Extracts the date/time value from the given string.
  The string is formatted as 'yyyy-mm-dd"T"hh:nn:ss.zzz'
-------------------------------------------------------------------------------}
function TsSpreadExcelXMLReader.ExtractDateTime(AText: String): TDateTime;
var
  dateStr, timeStr: String;
begin
  dateStr := Copy(AText, 1, 10);
  timeStr := Copy(AText, 12, MaxInt);
  Result := ScanDateTime('yyyy-mm-dd', dateStr) + ScanDateTime('hh:nn:ss.zzz', timeStr);
end;

{@@ ----------------------------------------------------------------------------
  Reads the cell alignment from the given node attributes
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadAlignment(ANode: TDOMNode;
  var AFormat: TsCellFormat);
var
  s: String;
begin
  // Vertical alignment
  s := GetAttrValue(ANode, 'ss:Vertical');
  if s <> '' then
    with AFormat do begin
      Include(UsedFormattingFields, uffVertAlign);
      case s of
        'Top':
          VertAlignment := vaTop;
        'Center':
          VertAlignment := vaCenter;
        'Bottom':
          VertAlignment := vaBottom;
        else
          Exclude(UsedFormattingFields, uffVertAlign);
      end;
    end;

  // Horizontal alignment
  s := GetAttrValue(ANode, 'ss:Horizontal');
  if s <> '' then
    with AFormat do begin
      Include(UsedFormattingFields, uffHorAlign);
      case s of
        'Left':
          HorAlignment := haLeft;
        'Center':
          HorAlignment := haCenter;
        'Right':
          HorAlignment := haRight;
        else
          Exclude(UsedFormattingFields, uffHorAlign);
      end;
    end;

  // Vertical text
  s := GetAttrValue(ANode, 'ss:Rotate');
  if s = '90' then
    with AFormat do begin
      TextRotation := rt90DegreeCounterClockwiseRotation;
      Include(UsedFormattingFields, uffTextRotation);
    end
  else if s = '-90' then
    with AFormat do begin
      TextRotation := rt90DegreeClockwiseRotation;
      Include(UsedFormattingFields, uffTextRotation);
    end;
  s := GetAttrValue(ANode, 'ss:VerticalText');
  if s <> '' then
    with AFormat do begin
      TextRotation := rtStacked;
      Include(UsedFormattingFields, uffTextRotation);
    end;

  // Word wrap
  s := GetAttrValue(ANode, 'ss:WrapText');
  if s = '1' then
    with AFormat do
      Include(UsedFormattingFields, uffWordWrap);

  // BiDi
  s := GetAttrValue(ANode, 'ss:ReadingOrder');
  if s <> '' then
    with AFormat do begin
      case s of
        'RightToLeft': BiDiMode := bdRTL;
        'LeftToRight': BiDiMode := bdLTR;
      end;
      Include(UsedFormattingFields, uffBiDi);
    end;
end;

{@@ ----------------------------------------------------------------------------
  Read a "Style/Borders/Border" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadBorder(ANode: TDOMNode;
  var AFormat: TsCellFormat);
// <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="3" ss:Color="#ED7D31"/>
var
  s, sw: String;
  b: TsCellBorder;
begin
  AFormat.UsedFormattingFields := AFormat.UsedFormattingFields + [uffBorder];

  // Border position
  s := GetAttrValue(ANode, 'ss:Position');
  case s of
    'Left':
      b := cbWest;
    'Right':
      b := cbEast;
    'Top':
      b := cbNorth;
    'Bottom':
      b := cbSouth;
    'DiagonalRight':
      b := cbDiagUp;
    'DiagonalLeft':
      b := cbDiagDown;
  end;
  Include(AFormat.Border, b);

  // Border color
  s := GetAttrValue(ANode, 'ss:Color');
  if s = '' then
    AFormat.BorderStyles[b].Color := scBlack
  else
    AFormat.BorderStyles[b].Color := HTMLColorStrToColor(s);

  // Line style
  s := GetAttrValue(ANode, 'ss:LineStyle');
  sw := GetAttrValue(ANode, 'ss:Weight');
  case s of
    'Continuous':
      if sw = '1' then
        AFormat.BorderStyles[b].LineStyle := lsThin
      else if sw = '2' then
        AFormat.BorderStyles[b].LineStyle := lsMedium
      else if sw = '3' then
        AFormat.BorderStyles[b].LineStyle := lsThick
      else if sw = '' then
        AFormat.BorderStyles[b].LineStyle := lsHair;
    'Double':
      AFormat.BorderStyles[b].LineStyle := lsDouble;
    'Dot':
      AFormat.BorderStyles[b].LineStyle := lsDotted;
    'Dash':
      if sw = '2' then
        AFormat.BorderStyles[b].LineStyle := lsMediumDash
      else
        AFormat.BorderStyles[b].LineStyle := lsDashed;
    'DashDot':
      if sw = '2' then
        AFormat.BorderStyles[b].LineStyle := lsMediumDashDot
      else
        AFormat.BorderStyles[b].LineStyle := lsDashDot;
    'DashDotDot':
      if sw = '2' then
        AFormat.BorderStyles[b].LineStyle := lsMediumDashDotDot
      else
        AFormat.BorderStyles[b].LineStyle := lsDashDotDot;
    'SlantDashDot':
        AFormat.BorderStyles[b].LineStyle := lsSlantDashDot;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the "Styles/Style/Borders" nodes
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadBorders(ANode: TDOMNode;
  var AFormat: TsCellFormat);
var
  nodeName: String;
begin
  if ANode = nil then exit;
  ANode := ANode.FirstChild;
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'Border' then
      ReadBorder(ANode, AFormat);
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads a "Worksheet/Table/Row/Cell" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadCell(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARow, ACol: Integer);
var
  book: TsWorkbook;
  sheet: TsWorksheet absolute AWorksheet;
  nodeName: string;
  s, st, sv: String;
  txt: String;
  node: TDOMNode;
  err: TsErrorValue;
  cell: PCell;
  fmt: TsCellFormat;
  nfp: TsNumFormatParams;
  idx: Integer;
  mergedCols, mergedRows: Integer;
  font: TsFont;
  dt: TDateTime;
begin
  if ANode = nil then
    exit;

  nodeName := ANode.NodeName;
  if nodeName <> 'Cell' then
    raise Exception.Create('[ReadCell] "Cell" node expected.');

  book := TsWorkbook(FWorkbook);
  font := book.GetDefaultFont;

  if FIsVirtualMode then
  begin
    if not Assigned(book.OnReadCellData) then
      exit;
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := sheet.AddCell(ARow, ACol);

  s := GetAttrValue(ANode, 'ss:StyleID');
  if s <> '' then
  begin
    idx := FCellFormatList.FindIndexOfName(s);
    if idx <> -1 then begin
      fmt := FCellFormatList.Items[idx]^;
      cell^.FormatIndex := book.AddCellFormat(fmt);
      font := book.GetFont(fmt.FontIndex);
    end;
  end else
  begin
    InitFormatRecord(fmt);
    cell^.FormatIndex := 0;
  end;

  // Merged cells
  s := GetAttrValue(ANode, 'ss:MergeAcross');
  if not ((s <> '') and TryStrToInt(s, mergedCols)) then mergedCols := 0;
  s := GetAttrValue(ANode, 'ss:MergeDown');
  if not ((s <> '') and TryStrToint(s, mergedRows)) then mergedRows := 0;
  if (mergedCols > 0) or (mergedRows > 0) then
    sheet.MergeCells(ARow, ACol, ARow + mergedRows, ACol + mergedCols);

  // Formula
  s := GetAttrValue(ANode, 'ss:Formula');
  if s <> '' then begin
    try
      sheet.WriteFormula(cell, s, false, true);
    except
      on E:EExprParser do begin
        FWorkbook.AddErrorMsg(E.Message);
        if (boAbortReadOnFormulaError in FWorkbook.Options) then raise;
      end;
      on E:ECalcEngine do begin
        FWorkbook.AddErrorMsg(E.Message);
        if (boAbortReadOnFormulaError in FWorkbook.Options) then raise;
      end;
    end;
  end;

  // Hyperlink
  s := GetAttrValue(ANode, 'ss:HRef');
  if s <> '' then begin
    st := GetAttrValue(ANode, 'x:HRefScreenTip');
    sheet.WriteHyperlink(cell, s, st);
  end;

  // Cell data and comment
  node := ANode.FirstChild;
  if node = nil then
    sheet.WriteBlank(cell)
  else begin
    book.LockFormulas;  // Protect formulas from being deleted by the WriteXXXX calls
    try
      while node <> nil do begin
        nodeName := node.NodeName;
        if (nodeName = 'Data') or (nodeName = 'ss:Data') then begin
          sv := node.TextContent;
          st := GetAttrValue(node, 'ss:Type');
          case st of
            'String':
              sheet.WriteText(cell, sv);
            'Number':
              sheet.WriteNumber(cell, StrToFloat(sv, FPointSeparatorSettings));
            'DateTime':
              begin
                dt := ExtractDateTime(sv);
                if (cell^.FormatIndex > 0) then begin
                  nfp := TsWorkbook(FWorkbook).GetNumberFormat(fmt.NumberFormatIndex);
                  if not IsTimeIntervalFormat(nfp) then
                    dt := ConvertExcelDateTimeToDateTime(dt, FDateMode);
                end;
                sheet.WriteDateTime(cell, dt);
              end;
            'Boolean':
              if sv = '1' then
                sheet.WriteBoolValue(cell, true)
              else if sv = '0' then
                sheet.WriteBoolValue(cell, false);
            'Error':
              if TryStrToErrorValue(sv, err) then
                sheet.WriteErrorValue(cell, err);
          end;
          if nodeName = 'ss:Data' then begin
            txt := '';
            RebuildChildNodes(node, txt);
            HTMLToRichText(FWorkbook, font, txt, s, cell^.RichTextParams, 'html:');
          end;
        end
        else
        if (nodeName = 'Comment') then
          ReadComment(node, AWorksheet, cell);
        node := node.NextSibling;
      end;

      if FIsVirtualMode then
        book.OnReadCellData(book, ARow, ACol, cell);

    finally
      book.UnlockFormulas;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the "Styles/Style/Protection" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadCellProtection(ANode: TDOMNode;
  var AFormat: TsCellFormat);
var
  s: String;
begin
  if ANode = nil then
    exit;

  s := GetAttrValue(ANode, 'ss:Protected');
  if s = '0' then
    Exclude(AFormat.Protection, cpLockCell);

  s := GetAttrValue(ANode, 'x:HideFormula');
  if s = '1' then
    Include(AFormat.Protection, cpHideFormulas);

  if AFormat.Protection <> DEFAULT_CELL_PROTECTION then
    Include(AFormat.UsedFormattingFields, uffProtection);
end;

{@@ ----------------------------------------------------------------------------
  Reads the "Worksheet/Table/Row/Cell/Comment" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadComment(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ACell: PCell);
var
  txt: String;
begin
  txt := ANode.TextContent;
  TsWorksheet(AWorksheet).WriteComment(ACell, txt);
end;

{@@ ----------------------------------------------------------------------------
  Reads the "ExcelWorkbook" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadExcelWorkbook(ANode: TDOMNode);
var
  s: String;
  nodeName: String;
  n: Integer;
begin
  if ANode = nil then
    exit;

  ANode := ANode.FirstChild;
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'ActiveSheet' then begin
      s := ANode.TextContent;
      if (s <> '') and TryStrToInt(s, n) then
        with TsWorkbook(FWorkbook) do
          SelectWorksheet(GetWorksheetByIndex(n));
    end else
    if nodeName = 'ProtectStructure' then begin
      s := ANode.TextContent;
      if s = 'True' then
        FWorkbook.Protection := FWorkbook.Protection + [bpLockStructure];
    end else
    if nodeName = 'ProtectWindows' then begin
      s := ANode.TextContent;
      if s = 'True' then
        FWorkbook.Protection := FWorkbook.Protection + [bpLockWindows];
    end else
    if nodeName = 'Date1904' then
      FDateMode := dm1904;

    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the "Styles/Style/Font" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLreader.ReadFont(ANode: TDOMNode;
  var AFormat: TsCellFormat);
var
  book: TsWorkbook;
  fname: String;
  fsize: Single;
  fcolor: TsColor;
  fstyle: TsFontStyles;
  s: String;
begin
  if ANode = nil then
    exit;

  book := TsWorkbook(FWorkbook);

  fname := GetAttrValue(ANode, 'ss:FontName');
  if fname = '' then
    fname := book.GetDefaultFont.FontName;

  s := GetAttrValue(ANode, 'ss:Size');
  if (s = '') or not TryStrToFloat(s, fsize, FPointSeparatorSettings) then
    fsize := book.GetDefaultFont.Size;

  s := GetAttrValue(ANode, 'ss:Color');
  if s <> '' then
    fcolor := HTMLColorStrToColor(s)
  else
    fcolor := book.GetDefaultFont.Color;

  fstyle := [];
  s := GetAttrValue(ANode, 'ss:Bold');
  if s = '1' then
    Include(fstyle, fssBold);
  s := GetAttrValue(ANode, 'ss:Italic');
  if s = '1' then
    Include(fstyle, fssItalic);
  s := GetAttrValue(ANode, 'ss:Underline');
  if s <> '' then
    Include(fstyle, fssUnderline);
  s := GetAttrValue(ANode, 'ss:StrikeThrough');
  if s = '1' then
    Include(fstyle, fssStrikeout);

  AFormat.FontIndex := book.AddFont(fname, fsize, fstyle, fcolor);
  Include(AFormat.UsedFormattingFields, uffFont);
end;

{@@ ----------------------------------------------------------------------------
  Reads the "Styles/Style/Interior" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadInterior(ANode: TDOMNode;
  var AFormat: TsCellFormat);
var
  s, sfg, sbg: String;
  fs: TsFillStyle;
begin
  if ANode = nil then
    exit;

  // Pattern
  s := GetAttrValue(ANode, 'ss:Pattern');
  if s = '' then
    exit;

  for fs in TsFillStyle do
    if FILL_NAMES[fs] = s then begin
      AFormat.Background.Style := fs;
      break;
    end;

  // Foreground color (pattern color)
  sfg := GetAttrValue(ANode, 'ss:PatternColor');
  if sfg = '' then
    AFormat.Background.FgColor := scBlack
  else
    AFormat.Background.FgColor := HTMLColorStrToColor(sfg);

  // Background color
  sbg := GetAttrValue(ANode, 'ss:Color');
  if sbg = '' then
    AFormat.Background.BgColor := scWhite
  else
    AFormat.Background.BgColor := HTMLColorStrToColor(sbg);

  // Fix solid fill colors: make foreground and background color the same
  if AFormat.Background.Style = fsSolidFill then begin
    if (sfg <> '') then
      AFormat.Background.BgColor := AFormat.Background.FgColor  // Forground priority
    else if (sfg = '') and (sbg <> '') then
      AFormat.Background.FgColor := AFormat.Background.BgColor;
  end;

  Include(AFormat.UsedFormattingFields, uffBackground);
end;

{@@ ----------------------------------------------------------------------------
  Reads a "Worksheet/Names" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadNames(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);

  procedure DoProcess(AStr: String; var ARowIndex, AColIndex: Cardinal;
    out IsRow: Boolean);
  var
    p: Integer;
  begin
    p := pos('!', AStr);
    if p > 0 then AStr := Copy(AStr, p+1, MaxInt);
    IsRow := AStr[1] in ['R', 'r'];
    Delete(AStr, 1, 1);
    if IsRow then
      ARowIndex := StrToInt(AStr) - 1
    else
      AColIndex := StrToInt(AStr) - 1;
  end;

  procedure DoRepeatedRowsCols(AStr: String);
  var
    p: Integer;
    isRow: Boolean;
    r1: Cardinal = UNASSIGNED_ROW_COL_INDEX;
    c1: Cardinal = UNASSIGNED_ROW_COL_INDEX;
    r2: Cardinal = UNASSIGNED_ROW_COL_INDEX;
    c2: Cardinal = UNASSIGNED_ROW_COL_INDEX;
  begin
    p := pos(':', AStr);
    // No colon --> Single range, e.g. "=Sheet1!C1"
    if p = 0 then
    begin
      DoProcess(AStr, r1, c1, isRow);
      r2 := r1;
      c2 := c1;
    end else
    // Colon --> Range block, e.g. "Sheet1!R1:R2"
    begin
      DoProcess(copy(AStr, 1, p-1), r1, c1, isRow);
      DoProcess(copy(AStr, p+1, MaxInt), r2, c2, isRow);
    end;
    if isRow then
      TsWorksheet(AWorksheet).PageLayout.SetRepeatedRows(r1, r2)
    else
      TsWorksheet(AWorksheet).PageLayout.SetRepeatedCols(c1, c2);
  end;

var
  sheet: TsWorksheet absolute AWorksheet;
  s, sr: String;
  nodeName: String;
  sheet1, sheet2: String;
  r1, c1, r2, c2: Cardinal;
  flags: TsRelFlags;
  p: Integer;
  ok: Boolean;
begin
  ok := true;
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'NamedRange' then begin
      s := GetAttrValue(ANode, 'ss:Name');
      if s = 'Print_Area' then begin
        // <NamedRange ss:Name="Print_Area" ss:RefersTo="=Tabelle2!R2C2:R5C7"/>
        s := GetAttrValue(ANode, 'ss:RefersTo');
        if (s <> '') then begin
          p := pos(',', s);
          while p > 0 do begin
            sr := Copy(s, 1, p-1);
            if ParseCellRangeString_R1C1(sr, 0, 0, sheet1, sheet2, r1, c1, r2, c2, flags) then
              sheet.PageLayout.AddPrintRange(r1, c1, r2, c2)
            else begin
              FWorkbook.AddErrorMsg('Invalid print range.');
              ok := false;
              break;
            end;
            s := copy(s, p+1, MaxInt);
            p := pos(',', s);
          end;
          if ok then begin
            if ParseCellRangeString_R1C1(s, 0, 0, sheet1, sheet2, r1, c1, r2, c2, flags) then
              sheet.PageLayout.AddPrintRange(r1, c1, r2, c2)
            else
              FWorkbook.AddErrorMsg('Invalid print range.');
          end;
        end;
      end else
      if s = 'Print_Titles' then begin
        // <NamedRange ss:Name="Print_Titles" ss:RefersTo="=Tabelle2!C1,Tabelle2!R1:R2"/>
        s := GetAttrValue(ANode, 'ss:RefersTo');
        if s <> '' then begin
          p := pos(',', s);
          if p > 0 then begin
            DoRepeatedRowsCols(copy(s, 1, p-1));
            DoRepeatedRowsCols(copy(s, p+1, MaxInt));
          end else
            DoRepeatedRowsCols(s);
        end;
      end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads a "Styles/Style/NumberFormat" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadNumberFormat(ANode: TDOMNode;
  var AFormat: TsCellFormat);
var
  s: String;
  nf: TsNumberFormat = nfGeneral;
  nfs: String;
begin
  if ANode = nil then
    exit;

  s := GetAttrValue(ANode, 'ss:Format');
  case s of
    'General': Exit;
    'Standard':
      begin
        nf := nfFixedTh;
        nfs := BuildNumberFormatString(nf, FWorkbook.FormatSettings, 2);
      end;
    'Fixed':
      begin
        nf := nfFixed;
        nfs := BuildNumberFormatString(nf, FWorkbook.FormatSettings, 2);
      end;
    'Percent':
      begin
        nf := nfPercentage;
        nfs := BuildNumberFormatString(nf, FWorkbook.FormatSettings, 2);
      end;
    'Scientific':
      begin
        nf := nfExp;
        nfs := BuildNumberFormatString(nf, FWorkbook.FormatSettings);
      end;
    'Short Date':
      begin
        nf := nfShortDate;
        nfs := BuildDateTimeFormatString(nf, FWorkbook.FormatSettings);
      end;
    'Short Time':
      begin
        nf := nfShortTime;
        nfs := BuildDateTimeFormatString(nf, FWorkbook.FormatSettings);
      end;
    else
      nfs := s;
  end;
  if nfs = '' then
    exit;

  AFormat.NumberFormatIndex := TsWorkbook(FWorkbook).AddNumberFormat(nfs);
  AFormat.NumberFormatStr := nfs;
  AFormat.NumberFormat := nf;
  Include(AFormat.UsedFormattingFields, uffNumberFormat);
end;

{@@ ----------------------------------------------------------------------------
  Reads a "Worksheet / PageBreaks / RowBreaks / RowBreak" node
  or a "Worksheet / PageBreaks / ColBreaks / ColBreak" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadPageBreak(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  sheet: TsWorksheet absolute AWorksheet;
  node: TDOMNode;
  nodeName: String;
  s: String;
  n: Integer;
begin
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'Row' then begin
      s := ANode.TextContent;
      if (s <> '') and TryStrToInt(s, n) then
        sheet.AddPageBreakToRow(n);
    end else
    if nodeName = 'Column' then begin
      s := ANode.TextContent;
      if (s <> '') and TryStrToInt(s, n) then
        sheet.AddPageBreakToCol(n);
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the "Wrksheet / PageBreaks" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadPageBreaks(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  sheet: TsWorksheet absolute AWorksheet;
  nodeName: String;
  node: TDOMNode;
  child: TDOMNode;
  s: String;
begin
  while ANode <> nil do
  begin
    nodeName := ANode.NodeName;
    if nodeName = 'RowBreaks' then begin
      node := ANode.FirstChild;
      while node <> nil do begin
        nodeName := node.NodeName;
        if nodeName = 'RowBreak' then
          ReadPageBreak(node.FirstChild, AWorksheet);
        node := node.NextSibling;
      end;
    end else
    if nodeName = 'ColBreaks' then begin
      node := ANode.FirstChild;
      while node <> nil do begin
        nodeName := node.NodeName;
        if nodeName = 'ColBreak' then
          ReadPageBreak(node.FirstChild, AWorksheet);
        node := node.NextSibling;
      end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the "WorksheetOptions/PageSetup" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadPageSetup(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  sheet: TsWorksheet absolute AWorksheet;
  nodeName: String;
  s: String;
  n: Integer;
  x: Double;
begin
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'Layout' then begin
      s := GetAttrValue(ANode, 'x:Orientation');
      if s = 'Landscape' then
        sheet.PageLayout.Orientation := spoLandscape;
      s := GetAttrValue(ANode, 'x:CenterHorizontal');
      if s = '1' then
        sheet.PageLayout.Options := sheet.PageLayout.Options + [poHorCentered];
      s := GetAttrValue(ANode, 'x:CenterVertical');
      if s = '1' then
        sheet.PageLayout.Options := sheet.PageLayout.Options + [poVertCentered];
      s := GetAttrValue(ANode, 'x:StartPageNumber');
      if (s <> '') and TryStrToInt(s, n) then
        sheet.PageLayout.StartPageNumber := n;
    end
    else if nodeName = 'Header' then begin
      s := GetAttrValue(ANode, 'x:Margin');
      if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        sheet.PageLayout.HeaderMargin := InToMM(x);
      s := GetAttrValue(ANode, 'x:Data');
      sheet.PageLayout.Headers[0] := s;
      sheet.PageLayout.Headers[1] := s;
      sheet.PageLayout.Headers[2] := s;
    end
    else if nodeName = 'Footer' then begin
      s := GetAttrValue(ANode, 'x:Margin');
      if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        sheet.PageLayout.FooterMargin := InToMM(x);
      s := GetAttrValue(ANode, 'x:Data');
      sheet.PageLayout.Footers[0] := s;
      sheet.PageLayout.Footers[1] := s;
      sheet.PageLayout.Footers[2] := s;
    end
    else if nodeName = 'PageMargins' then begin
      s := GetAttrValue(ANode, 'x:Bottom');
      if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        sheet.PageLayout.BottomMargin := InToMM(x);
      s := GetAttrValue(ANode, 'x:Top');
      if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        sheet.PageLayout.TopMargin := InToMM(x);
      s := GetAttrValue(ANode, 'x:Left');
      if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        sheet.PageLayout.LeftMargin := InToMM(x);
      s := GetAttrValue(ANode, 'x:Right');
      if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        sheet.PageLayout.RightMargin := InToMM(x);
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the "WorksheetOptions/Print" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadPrint(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  sheet: TsWorksheet absolute AWorksheet;
  nodeName: String;
  s: String;
  n: Integer;
  x: Double;
begin
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'PaperSizeIndex' then begin
      s := ANode.TextContent;
      if (s <> '') and TryStrToInt(s, n) and (n < Length(PAPER_SIZES)) then begin
        sheet.PageLayout.PageWidth := PAPER_SIZES[n, 0];
        sheet.PageLayout.pageHeight := PAPER_SIZES[n, 1];
      end;
    end
    else if nodeName = 'FitHeight' then begin
      s := ANode.TextContent;
      if (s <> '') and TryStrToInt(s, n) then
        sheet.PageLayout.FitHeightToPages := n;
    end
    else if nodeName = 'FitWidth' then begin
      s := ANode.TextContent;
      if (s <> '') and TryStrToInt(s, n) then
        sheet.PageLayout.FitWidthToPages := n;
    end
    else if nodeName = 'Scale' then begin
      s := ANode.TextContent;
      if (s <> '') and TryStrToInt(s, n) then
        sheet.PageLayout.ScalingFactor := n;
    end
    else if nodeName = 'Gridlines' then
      sheet.PageLayout.Options := sheet.PageLayout.Options + [poPrintGridLines]
    else if nodeName = 'BlackAndWhite' then
      sheet.PageLayout.Options := sheet.PageLayout.Options + [poMonochrome]
    else if nodeName = 'DraftQuality' then
      sheet.PageLayout.Options := sheet.PageLayout.Options + [poDraftQuality]
    else if nodeName = 'LeftToRight' then
      sheet.PageLayout.Options := sheet.PageLayout.Options + [poPrintPagesByRows]
    else if nodeName = 'RowColHeadings' then
      sheet.PageLayout.Options := sheet.PageLayout.Options + [poPrintHeaders]
    else if nodeName = 'CommentsLayout' then begin
      s := ANode.TextContent;
      if s = 'SheetEnd' then
        sheet.PageLayout.Options := sheet.PageLayout.Options + [poCommentsAtEnd]
      else if s = 'InPlace' then
        sheet.PageLayout.Options := sheet.PageLayout.Options + [poPrintCellComments];
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads a "Worksheet/Table/Row" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadRow(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet; ARow: Integer);
var
  nodeName: String;
  s: String;
  c: Integer;
begin
  c := 0;
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'Cell' then begin
      s := GetAttrValue(ANode, 'ss:Index');
      if s <> '' then c := StrToInt(s) - 1;
      ReadCell(ANode, AWorksheet, ARow, c);
      inc(c);
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads a "Styles/Style" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadStyle(ANode: TDOMNode);
var
  nodeName: String;
  fmt: TsCellFormat;
  s: String;
  idx: Integer;
  childNode: TDOMNode;
begin
  // Respect ancestor of current style
  s := GetAttrValue(ANode, 'ss:Parent');
  if s <> '' then begin
    idx := FCellFormatList.FindIndexOfName(s);
    if idx > -1 then
      fmt := FCellFormatList.Items[idx]^;
  end else
    InitFormatRecord(fmt);

  // ID of current style. We store it in the "Name" field of the TsCellFormat
  // because it is a string while ID is an Integer (mostly "s<number>", but also
  // "Default").
  fmt.Name := GetAttrValue(ANode, 'ss:ID');

  if fmt.Name = 's125' then
    idx := 0;

  // Style elements
  childNode := ANode.FirstChild;
  while childNode <> nil do begin
    nodeName := childNode.NodeName;
    if nodeName = 'Alignment' then
      ReadAlignment(childNode, fmt)
    else if nodeName = 'Borders' then
      ReadBorders(childNode, fmt)
    else if nodeName = 'Interior' then
      ReadInterior(childNode, fmt)
    else if nodeName = 'Font' then
      ReadFont(childNode, fmt)
    else if nodeName = 'NumberFormat' then
      ReadNumberFormat(childnode, fmt)
    else if nodeName = 'Protection' then
      ReadCellProtection(childNode, fmt);
    childNode := childNode.NextSibling;
  end;

  FCellFormatList.Add(fmt);
end;

{@@ ----------------------------------------------------------------------------
  Reads the "Styles" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadStyles(ANode: TDOMNode);
var
  nodeName: String;
  styleNode: TDOMNode;
begin
  if ANode = nil then
    exit;
  styleNode := ANode.FirstChild;
  while styleNode <> nil do begin
    nodeName := styleNode.NodeName;
    if nodeName = 'Style' then
      ReadStyle(styleNode);
    styleNode := styleNode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the "Worksheet/Table" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadTable(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  sheet: TsWorksheet absolute AWorksheet;
  nodeName: String;
  s: String;
  r, c: Integer;
  x: Double;
  idx: Integer;
  fmt: TsCellFormat;
  rht: TsRowHeightType;
begin
  r := 0;
  c := 0;
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'Column' then begin
      // Default column width
      s := GetAttrValue(ANode, 'ss:DefaultColumnWidth');
      if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        sheet.WriteDefaultColWidth(x, suPoints);

      // Column index
      s := GetAttrValue(ANode, 'ss:Index');
      if (s <> '') and TryStrToInt(s, c) then
        dec(c);

      // Column width, in Points
      s := GetAttrValue(ANode, 'ss:Width');
      if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        sheet.WriteColWidth(c, x, suPoints);

      // Column format
      s := GetAttrValue(ANode, 'ss:StyleID');
      if s <> '' then begin
        idx := FCellFormatList.FindIndexOfName(s);
        if idx <> -1 then begin
          fmt := FCellFormatList.Items[idx]^;
          idx := TsWorkbook(FWorkbook).AddCellFormat(fmt);
          sheet.WriteColFormatIndex(c, idx);
        end;
      end;

      // Hidden
      s := GetAttrValue(ANode, 'ss:Hidden');
      if s = '1' then
        sheet.HideCol(c);

      inc(c);
    end
    else
    if nodeName = 'Row' then begin
      // Default row height
      s := GetAttrValue(ANode, 'ss:DefaultRowHeight');
      if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        sheet.WriteDefaultRowHeight(x, suPoints);

      // Index
      s := GetAttrValue(ANode, 'ss:Index');
      if s <> '' then r := StrToInt(s) - 1;

      // AutoFitHeight
      s := GetAttrValue(ANode, 'ss:AutoFitHeight');
      if s = '1' then
        rht := rhtAuto
      else
        rht := rhtCustom;

      // Height
      s := GetAttrValue(ANode, 'ss:Height');
      if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        sheet.WriteRowHeight(r, x, suPoints, rht);

      // Hidden
      s := GetAttrValue(ANode, 'ss:Hidden');
      if (s = '1') then
        sheet.HideRow(r);

      // Row format
      s := GetAttrValue(ANode, 'ss:StyleID');
      if s <> '' then begin
        idx := FCellFormatList.FindIndexOfName(s);
        if idx <> -1 then begin
          fmt := FCellFormatList.Items[idx]^;
          idx := TsWorkbook(FWorkbook).AddCellFormat(fmt);
          sheet.WriteRowFormatIndex(r, idx);
        end;
      end;

      // Cells in row
      ReadRow(ANode.FirstChild, AWorksheet, r);

      inc(r);
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the "Worksheet" node
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadWorksheet(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  nodeName: String;
  s: String;
begin
  if ANode = nil then
    exit;

  s := GetAttrValue(ANode, 'ss:Protected');
  if s ='1' then
    AWorksheet.Options := AWorksheet.Options + [soProtected];

  ANode := ANode.FirstChild;
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'Table' then
      ReadTable(ANode.FirstChild, AWorksheet)
    else if nodeName = 'WorksheetOptions' then
      ReadWorksheetOptions(ANode.FirstChild, AWorksheet)
    else if nodeName = 'Names' then
      ReadNames(ANode.FirstChild, AWorksheet)
    else if nodeName = 'PageBreaks' then
      ReadPageBreaks(ANode.FirstChild, AWorksheet);
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the "Worksheet/WorksheetOptions" nodes
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadWorksheetOptions(ANode: TDOMNode;
  AWorksheet: TsBasicWorksheet);
var
  sheet: TsWorksheet absolute AWorksheet;
  node, childnode: TDOMNode;
  nodeName: String;
  s: String;
  x: Double;
  n: Integer;
  hasFitToPage: Boolean = false;
  c, r: Cardinal;
begin
  if ANode = nil then
    exit;

  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'PageSetup' then
      ReadPageSetup(ANode.FirstChild, AWorksheet)
    else
    if nodeName = 'FitToPage' then begin
      hasFitToPage := true;
      sheet.PageLayout.Options := sheet.PageLayout.Options + [poFitPages];
    end else
    if nodeName = 'Print' then begin
      node := ANode.FirstChild;
      ReadPrint(ANode.FirstChild, AWorksheet);
    end else
    if nodeName = 'Selected' then
      TsWorkbook(FWorkbook).ActiveWorksheet := sheet
    else
    if nodeName = 'Panes' then begin
      c := sheet.ActiveCellCol;
      r := sheet.ActiveCellRow;
      node := ANode.FirstChild;
      while node <> nil do begin
        nodeName := node.NodeName;
        if nodeName = 'Pane' then begin
          childnode := node.FirstChild;
          while childnode <> nil do begin
            nodeName := childNode.NodeName;
            if nodeName = 'ActiveRow' then begin
              s := childNode.TextContent;
              if (s <> '') and TryStrToInt(s, n) then
                r := n;
            end else
            if nodeName = 'ActiveCol' then begin
              s := childNode.TextContent;
              if (s <> '') and TryStrToInt(s, n) then
                c := n;
            end;
            childnode := childNode.NextSibling;
          end;
        end;
        node := node.NextSibling;
      end;
      sheet.SelectCell(r, c);
    end else
    if nodeName = 'FreezePanes' then
      sheet.Options := sheet.Options + [soHasFrozenPanes]
    else
    if (nodeName = 'TopRowBottomPane') then begin
      s := ANode.TextContent;
      if (s <> '') and TryStrToInt(s, n) then
        sheet.TopPaneHeight := n;
    end else
    if (nodeName = 'LeftColumnRightPane') then begin
      s := ANode.TextContent;
      if (s <> '') and TryStrToInt(s, n) then
        sheet.LeftPaneWidth := n;
    end else
    if nodeName = 'DoNotDisplayGridlines' then
      sheet.Options := sheet.Options - [soShowGridLines]
    else
    if nodeName = 'DoNotDisplayHeadings' then
      sheet.Options := sheet.Options - [soShowHeaders]
    else
    if nodeName = 'Zoom' then begin
      s := ANode.TextContent;
      if (s <> '') and TryStrToFloat(s, x) then
        sheet.Zoomfactor := x * 0.01;
    end else
    if nodeName = 'Visible' then begin
      s := ANode.TextContent;
      if s = 'SheetHidden' then
        sheet.Options := sheet.Options + [soHidden];
    end else
    if nodeName = 'AllowFormatCells' then
      sheet.Protection := sheet.Protection - [spFormatCells]
    else
    if nodeName = 'AllowSizeCols' then
      sheet.Protection := sheet.Protection - [spFormatColumns]
    else
    if nodeName = 'AllowSizeRows' then
      sheet.Protection := sheet.Protection - [spFormatRows]
    else
    if nodeName = 'AllowInsertCols' then
      sheet.Protection := sheet.Protection - [spInsertColumns]
    else
    if nodeName = 'AllowInsertRows' then
      sheet.Protection := sheet.Protection - [spInsertRows]
    else
    if nodeName = 'AllowInsertHyperlinks' then
      sheet.Protection := sheet.Protection - [spInsertHyperLinks]
    else
    if nodeName = 'AllowDeleteCols' then
      sheet.Protection := sheet.Protection - [spDeleteColumns]
    else
    if nodeName = 'AllowDeleteRows' then
      sheet.Protection := sheet.Protection - [spDeleteRows]
    else
    if nodeName = 'AllowSort' then
      sheet.Protection := sheet.Protection - [spSort]
    else
    if nodeName = 'ProtectObjects' then
      sheet.Protection := sheet.Protection + [spObjects]
    else
      {
    if nodeName = 'ProtectScenarios' then
      sheet.Protection := sheet.Protection + [spScenarios];
    else
      }
    if nodeName = 'EnableSelection' then begin
      s := ANode.TextContent;
      if s = 'NoSelection' then
        sheet.Protection := sheet.Protection + [spSelectLockedCells, spSelectUnlockedCells]
      else
      if s = 'Unlocked' then
        sheet.Protection := sheet.Protection + [spSelectLockedCells];
    end;

    ANode := ANode.NextSibling;
  end;

  if hasFitToPage then begin
    // The ScalingFactor is always written to the xml file. This makes TsPageLayout
    // automatically remove the poFitPages option which is restored here.
    if (sheet.PageLayout.ScalingFactor <> 100) then begin
      sheet.PageLayout.ScalingFactor := 100;
      sheet.Pagelayout.Options := sheet.PageLayout.Options + [poFitPages];
    end;
    // When FitToPages is active, but FitWidthToPages and/or FitHeightToPages
    // are not specified, they should be set to 1
    if sheet.PageLayout.FitWidthToPages = 0 then
      sheet.PageLayout.FitWidthToPages := 1;
    if sheet.PageLayout.FitHeightToPages = 0 then
      sheet.PageLayout.FitHeightToPages := 1;
  end;
end;

          (*

  function TsSpreadExcelXMLWriter.GetLayoutStr(AWorksheet: TsBasicWorksheet): String;
  var
    sheet: TsWorksheet absolute AWorksheet;
  begin
    Result := '';
    if sheet.PageLayout.Orientation = spoLandscape then
      Result := Result + ' x:Orientation="Landscape"';
    if (poHorCentered in sheet.PageLayout.Options) then
      Result := Result + ' x:CenterHorizontal="1"';
    if (poVertCentered in sheet.PageLayout.Options) then
      Result := Result + ' x:CenterVertical="1"';
    if (poUseStartPageNumber in sheet.PageLayout.Options) then
      Result := Result + ' x:StartPageNumber="' + IntToStr(sheet.PageLayout.StartPageNumber) + '"';
    Result := '<Layout' + Result + '/>';
  end;
  *)

{@@ ----------------------------------------------------------------------------
  Reads the "Worksheet" nodes
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadWorksheets(ANode: TDOMNode);
var
  node: TDOMNode;
  nodeName: String;
  s: String;
begin
  node := ANode;
  // first iterate through all worksheets, get the name and add them to the
  // workbook. This is because 3D formulas may refer to sheets not yet loaded.
  while node <> nil do begin
    nodeName := node.NodeName;
    if nodeName = 'Worksheet' then begin
      s := GetAttrValue(node, 'ss:Name');
      if s <> '' then       // the case of '' should not happen...
        FWorksheet := TsWorkbook(FWorkbook).AddWorksheet(s);
    end;
    node := node.NextSibling;
  end;

  // Now iterate through the worksheets again and read their contents
  while ANode <> nil do begin
    nodeName := ANode.NodeName;
    if nodeName = 'Worksheet' then begin
      s := GetAttrValue(ANode, 'ss:Name');
      FWorksheet := TsWorkbook(FWorkbook).GetWorksheetByName(s);
      ReadWorksheet(ANode, FWorksheet);
    end;
    ANode := ANode.NextSibling;
  end;
end;


{@@ ----------------------------------------------------------------------------
  Reads the workbook from the specified stream
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLReader.ReadFromStream(AStream: TStream;
  APassword: String = ''; AParams: TsStreamParams = []);
var
  doc: TXMLDocument;
begin
  try
    ReadXMLStream(doc, AStream);

    // Read style list
    ReadStyles(doc.DocumentElement.FindNode('Styles'));

    // Read worksheets and their contents
    ReadWorksheets(doc.DocumentElement.FindNode('Worksheet'));

    // Read ExcelWorkbook node after worksheet nodes although before it is
    // found before the worksheet nodes in the file, because is requires
    // worksheets to be existing.
    ReadExcelWorkbook(doc.DocumentElement.FindNode('ExcelWorkbook'));
  finally
    doc.Free;
  end;
end;


{===============================================================================
                          TsSpreadExcelXMLWriter
===============================================================================}

{@@ ----------------------------------------------------------------------------
  Constructor of the ExcelXML writer

  Defines the date mode and the limitations of the file format.
  Initializes the format settings to be used when writing to xml.
-------------------------------------------------------------------------------}
constructor TsSpreadExcelXMLWriter.Create(AWorkbook: TsBasicWorkbook);
begin
  inherited Create(AWorkbook);

  // Initial base date in case it won't be set otherwise.
  // Use 1900 to get a bit more range between 1900..1904.
  FDateMode := ExcelXMLSettings.DateMode;

  // Special version of FormatSettings using a point decimal separator for sure.
  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';

  // http://en.wikipedia.org/wiki/List_of_spreadsheet_software#Specifications
  FLimitations.MaxColCount := 256;
  FLimitations.MaxRowCount := 65536;
end;

function TsSpreadExcelXMLWriter.GetCommentStr(ACell: PCell): String;
var
  comment: PsComment;
begin
  Result := '';
  comment := (FWorksheet as TsWorksheet).FindComment(ACell);
  if Assigned(comment) then
    Result := INDENT1 +
      '<Comment><Data>' +
        UTF8TextToXMLText(comment^.Text) +
      '</Data></Comment>' +
      LF + CELL_INDENT;
  // If there will be some rich-text-like formatting in the future, use
  //  Result := '<Comment><ss:Data xmlns="http://www.w3.org/TR/REC-html40">'+comment^.Text+'</ss:Data></Comment>':
end;

function TsSpreadExcelXMLWriter.GetFormulaStr(ACell: PCell): String;
begin
  if HasFormula(ACell) then
  begin
    Result := UTF8TextToXMLText((FWorksheet as TsWorksheet).ConvertFormulaDialect(ACell, fdExcelR1C1));
    Result := ' ss:Formula="=' + Result + '"';
  end else
    Result := '';
end;

function TsSpreadExcelXMLWriter.GetFrozenPanesStr(AWorksheet: TsBasicWorksheet;
  AIndent: String): String;
var
  activePane: Integer;
  sheet: TsWorksheet absolute AWorksheet;
begin
  if (soHasFrozenPanes in sheet.Options) then
  begin
    Result := AIndent +
        '<FreezePanes/>' + LF + AIndent +
        '<FrozenNoSplit/>' + LF;

    if sheet.LeftPaneWidth > 0 then
      Result := Result + AIndent +
        '<SplitVertical>1</SplitVertical>' + LF + AIndent +
        '<LeftColumnRightPane>' + IntToStr(sheet.LeftPaneWidth) + '</LeftColumnRightPane>' + LF;

    if sheet.TopPaneHeight > 0 then
      Result := Result + AIndent +
        '<SplitHorizontal>1</SplitHorizontal>' + LF + AIndent +
        '<TopRowBottomPane>' + IntToStr(sheet.TopPaneHeight) + '</TopRowBottomPane>' + LF;

    if (sheet.LeftPaneWidth = 0) and (sheet.TopPaneHeight = 0) then
      activePane := 3
    else
    if (sheet.LeftPaneWidth = 0) then
      activePane := 2
    else
    if (sheet.TopPaneHeight = 0) then
      activePane := 1
    else
      activePane := 0;
    Result := Result + AIndent +
      '<ActivePane>' + IntToStr(activePane) + '</ActivePane>' + LF;
  end else
    Result := '';
end;

function TsSpreadExcelXMLWriter.GetHyperlinkStr(ACell: PCell): String;
var
  hyperlink: PsHyperlink;
begin
  hyperlink := (FWorksheet as TsWorksheet).FindHyperlink(ACell);
  if Assigned(hyperlink) then begin
    Result := ' ss:HRef="' + hyperlink^.Target + '"';
    if hyperlink^.ToolTip <> '' then
      Result := Result + ' x:HRefScreenTip="' + UTF8TextToXMLText(hyperlink^.ToolTip) + '"';
  end else
    Result := '';
end;

function TsSpreadExcelXMLWriter.GetIndexStr(AIndex, APrevIndex: Integer): String;
begin
  if (APrevIndex = -1) and (AIndex = 0)  then
    Result := ''
  else
  if (APrevIndex >= 0) and (AIndex = APrevIndex + 1) then
    Result := ''
  else
    Result := Format(' ss:Index="%d"', [AIndex + 1]);
end;

function TsSpreadExcelXMLWriter.GetLayoutStr(AWorksheet: TsBasicWorksheet): String;
var
  sheet: TsWorksheet absolute AWorksheet;
begin
  Result := '';
  if sheet.PageLayout.Orientation = spoLandscape then
    Result := Result + ' x:Orientation="Landscape"';
  if (poHorCentered in sheet.PageLayout.Options) then
    Result := Result + ' x:CenterHorizontal="1"';
  if (poVertCentered in sheet.PageLayout.Options) then
    Result := Result + ' x:CenterVertical="1"';
  if (poUseStartPageNumber in sheet.PageLayout.Options) then
    Result := Result + ' x:StartPageNumber="' + IntToStr(sheet.PageLayout.StartPageNumber) + '"';
  Result := '<Layout' + Result + '/>';
end;

function TsSpreadExcelXMLWriter.GetMergeStr(ACell: PCell): String;
var
  r1, c1, r2, c2: Cardinal;
begin
  Result := '';
  if (FWorksheet as TsWorksheet).IsMerged(ACell) then begin
    (FWorksheet as TsWorksheet).FindMergedRange(ACell, r1, c1, r2, c2);
    if c2 > c1 then
      Result := Result + Format(' ss:MergeAcross="%d"', [c2-c1]);
    if r2 > r1 then
      Result := Result + Format(' ss:MergeDown="%d"', [r2-r1]);
  end;
end;

function TsSpreadExcelXMLWriter.GetPageFooterStr(
  AWorksheet: TsBasicWorksheet): String;
var
  sheet: TsWorksheet absolute AWorksheet;
begin
  Result := Format('x:Margin="%g"', [mmToIn(sheet.PageLayout.FooterMargin)], FPointSeparatorSettings);
  if (sheet.PageLayout.Footers[HEADER_FOOTER_INDEX_ALL] <> '') then
    Result := Result + ' x:Data="' + UTF8TextToXMLText(sheet.PageLayout.Footers[HEADER_FOOTER_INDEX_ALL], true) + '"';
  Result := '<Footer ' + result + '/>';
end;

function TsSpreadExcelXMLWriter.GetPageHeaderStr(
  AWorksheet: TsBasicWorksheet): String;
var
  sheet: TsWorksheet absolute AWorksheet;
begin
  Result := Format('x:Margin="%g"', [mmToIn(sheet.PageLayout.HeaderMargin)], FPointSeparatorSettings);
  if (sheet.PageLayout.Headers[HEADER_FOOTER_INDEX_ALL] <> '') then
    Result := Result + ' x:Data="' + UTF8TextToXMLText(sheet.PageLayout.Headers[HEADER_FOOTER_INDEX_ALL], true) + '"';
  Result := '<Header ' + Result + '/>';
end;

function TsSpreadExcelXMLWriter.GetPageMarginStr(
  AWorksheet: TsBasicWorksheet): String;
var
  sheet: TsWorksheet absolute AWorksheet;
begin
  Result := Format('x:Bottom="%g" x:Left="%g" x:Right="%g" x:Top="%g"', [
    mmToIn(sheet.PageLayout.BottomMargin),
    mmToIn(sheet.PageLayout.LeftMargin),
    mmToIn(sheet.PageLayout.RightMargin),
    mmToIn(sheet.PageLayout.TopMargin)
    ], FPointSeparatorSettings);
  Result := '<PageMargins ' + Result + '/>';
end;

{ Todo: When can the "Print" node be skipped? }
function TsSpreadExcelXMLWriter.GetPrintStr(AWorksheet: TsBasicWorksheet): String;
var
  sheet: TsWorksheet absolute AWorksheet;
  i, pgSizeIdx: Integer;
  scalestr: String;
begin
  Result := '';
  pgSizeIdx := -1;
  for i:=0 to High(PAPER_SIZES) do
    if (SameValue(PAPER_SIZES[i,0], sheet.PageLayout.PageHeight) and
        SameValue(PAPER_SIZES[i,1], sheet.PageLayout.PageWidth))
    or (SameValue(PAPER_SIZES[i,1], sheet.PageLayout.PageHeight) and
        SameValue(PAPER_SIZES[i,0], sheet.PageLayout.PageWidth))
    then begin
      pgSizeIdx := i;
      break;
    end;

  if pgSizeidx = -1 then
    exit;

  // Scaling factor
  if sheet.PageLayout.ScalingFactor <> 100 then
    scaleStr := INDENT4 + '<Scale>' + IntToStr(sheet.PageLayout.ScalingFactor) + '</Scale>' + LF
  else
    scaleStr := '';

  Result :=
    INDENT4 + '<ValidPrinterInfo/>' + LF +
    INDENT4 + '<PaperSizeIndex>' + IntToStr(pgSizeIdx) + '</PaperSizeIndex>' + LF +
    scaleStr +
    INDENT4 + '<VerticalResolution>0</VerticalResolution>';

  if sheet.PageLayout.FitHeightToPages > 1 then
    Result := Result + LF + INDENT4 +
      '<FitHeight>' + IntToStr(sheet.PageLayout.FitHeightToPages) + '</FitHeight>';

  if sheet.PageLayout.FitWidthToPages > 1 then
    Result := result + LF + INDENT4 +
      '<FitWidth>' + IntToStr(sheet.PageLayout.FitWidthToPages) + '</FitWidth>';
end;

function TsSpreadExcelXMLWriter.GetStyleStr(AFormatIndex: Integer): String;
begin
  Result := '';
  if AFormatIndex > 0 then
    Result := Format(' ss:StyleID="s%d"', [AFormatIndex + FMT_OFFSET]);
end;

procedure TsSpreadExcelXMLWriter.WriteBlank(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
begin
  Unused(ARow, ACol);
  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s>' +              // colIndex, style, hyperlink, merge
      '%s' +                        // Comment <Comment>...</Comment>
    '</Cell>' + LF, [
    GetIndexStr(ACol, FPrevCol), GetStyleStr(ACell^.FormatIndex), GetHyperlinkStr(ACell), GetMergeStr(ACell),
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteBool(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: boolean; ACell: PCell);
begin
  Unused(ARow, ACol);
  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s%s>' +         // colIndex, style, formula, hyperlink, merge
      '<Data ss:Type="%s">' +    // data type
        '%s' +                   // value string
      '</Data>' +
      '%s' +                     // Comment <Comment>...</Comment>
    '</Cell>' + LF, [
    GetIndexStr(ACol, FPrevCol), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
      GetHyperlinkStr(ACell), GetMergeStr(ACell),
    StrUtils.IfThen(HasFormula(ACell), GetCellContentTypeStr(ACell), 'Boolean'),
    StrUtils.IfThen(AValue, '1', '0'),
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteCellToStream(AStream: TStream; ACell: PCell);
begin
  case ACell^.ContentType of
    cctBool:
      WriteBool(AStream, ACell^.Row, ACell^.Col, ACell^.BoolValue, ACell);
    cctDateTime:
      WriteDateTime(AStream, ACell^.Row, ACell^.Col, ACell^.DateTimeValue, ACell);
    cctEmpty:
      WriteBlank(AStream, ACell^.Row, ACell^.Col, ACell);
    cctError:
      WriteError(AStream, ACell^.Row, ACell^.Col, ACell^.ErrorValue, ACell);
    cctNumber:
      WriteNumber(AStream, ACell^.Row, ACell^.Col, ACell^.NumberValue, ACell);
    cctUTF8String:
      WriteLabel(AStream, ACell^.Row, ACell^.Col, ACell^.UTF8StringValue, ACell);
    cctFormula:
      WriteFormula(AStream, ACell^.Row, ACell^.Col, ACell);
  end;

  if (FWorksheet as TsWorksheet).ReadComment(ACell) <> '' then
    WriteComment(AStream, ACell);
end;

procedure TsSpreadExcelXMLWriter.WriteCellNodes(AStream: TStream;
  AWorksheet: TsBasicWorksheet; ARow: Cardinal);
var
  c: Cardinal;
  cell: PCell;
  lCell: TCell;
  styleCell: PCell;
  value: variant;
  sheet: TsWorksheet absolute AWorksheet;
begin
  if (boVirtualMode in FWorkbook.Options) and (not Assigned(sheet.OnWriteCellData)) then
    exit;

  FPrevCol := UNASSIGNED_ROW_COL_INDEX;
  for c := 0 to FLastCol do
  begin
    if (boVirtualMode in FWorkbook.Options) then begin
      lCell.Row := ARow; // to silence a compiler hint
      InitCell(lCell);
      value := varNull;
      styleCell := nil;
      sheet.OnWriteCellData(sheet, ARow, c, value, styleCell);
      if styleCell <> nil then
        lCell := styleCell^;
      lCell.Row := ARow;
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
      FPrevCol := c;
    end else
    begin
      // Normal mode
      cell := sheet.Findcell(ARow, c);
      if cell <> nil then
      begin
        if sheet.IsMerged(cell) and not sheet.IsMergeBase(cell) then
          Continue;
        WriteCellToStream(AStream, cell);
        FPrevCol := c;
      end;
    end;
  end;
end;

procedure TsSpreadExcelXMLWriter.WriteColumns(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  c, c1, c2: Cardinal;
  colwidthStr: String;
  styleStr: String;
  hiddenStr: String;
  col: PCol;
begin
  c1 := 0;
  c2 := TsWorksheet(AWorksheet).GetLastColIndex;
  FPrevCol := UNASSIGNED_ROW_COL_INDEX;
  for c := c1 to c2 do
  begin
    col := TsWorksheet(AWorksheet).FindCol(c);
    styleStr := '';
    colWidthStr := '';
    hiddenStr := '';

    if Assigned(col) then
    begin
      // column width is needed in pts.
      if col^.ColWidthType = cwtCustom then
        colwidthStr := Format(' ss:Width="%0.2f" ss:AutoFitWidth="0"',
          [(FWorkbook as TsWorkbook).ConvertUnits(col^.Width, FWorkbook.Units, suPoints)],
          FPointSeparatorSettings);
      // column style
      if col^.FormatIndex > 0 then
        styleStr := GetStyleStr(col^.FormatIndex);
    end;

    if TsWorksheet(AWorksheet).ColHidden(c) then
      hiddenStr := ' ss:Hidden="1"';

    if (colWidthStr <> '') or (stylestr <> '') or (hiddenstr <> '') then begin
      AppendToStream(AStream, COL_INDENT + Format(
        '<Column%s%s%s%s />' + LF, [GetIndexStr(c, FPrevCol), colWidthStr, styleStr, hiddenStr]));
      FPrevCol := c;
    end;
  end;
end;

procedure TsSpreadExcelXMLWriter.WriteDateTime(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell);
var
  valueStr: String;
  ExcelDate: TDateTime;
  nfp: TsNumFormatParams;
  fmt: PsCellFormat;
begin
  Unused(ARow, ACol);
  ExcelDate := AValue;
  fmt := (FWorkbook as TsWorkbook).GetPointerToCellFormat(ACell^.FormatIndex);
  // Times have an offset of 1 day!
  if (fmt <> nil) and (uffNumberFormat in fmt^.UsedFormattingFields) then
  begin
    nfp := (FWorkbook as TsWorkbook).GetNumberFormat(fmt^.NumberFormatIndex);
    if IsTimeIntervalFormat(nfp) or IsTimeFormat(nfp) then
      case FDateMode of
        dm1900: ExcelDate := AValue + DATEMODE_1900_BASE;
        dm1904: ExcelDate := AValue + DATEMODE_1904_BASE;
      end;
  end;
  valueStr := FormatDateTime('yyyy-mm-dd"T"hh:nn:ss.zzz', ExcelDate);

  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s%s>' + LF + VALUE_INDENT + // colIndex, style, formula, hyperlink, merge
      '<Data ss:Type="%s">' +                // data type
        '%s' +                               // value string
      '</Data>' + LF + CELL_INDENT +
      '%s' +                                 // Comment <Comment>...</Comment>
    '</Cell>' + LF, [
    GetIndexStr(ACol, FPrevCol), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
      GetHyperlinkStr(ACell), GetMergeStr(ACell),
    StrUtils.IfThen(HasFormula(ACell), GetCellContentTypeStr(ACell), 'DateTime'),
    valueStr,
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
begin
  Unused(ARow, ACol);
  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s%s>' + LF + VALUE_INDENT + // colIndex, style, formula, hyperlink, merge
      '<Data ss:Type="%s">' +                // data type
        '%s' +                               // value string
      '</Data>' + LF + CELL_INDENT +
      '%s' +                                 // Comment <Comment>...</Comment>
    '</Cell>' + LF, [
    GetIndexStr(ACol, FPrevCol), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
      GetHyperlinkStr(ACell), GetMergeStr(ACell),
    StrUtils.IfThen(HasFormula(ACell), GetCellContentTypeStr(ACell), 'Error'),
    GetErrorValueStr(AValue),
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteExcelWorkbook(AStream: TStream);
var
  datemodeStr: String;
  protectStr: String;
begin
  if FDateMode = dm1904 then
    datemodeStr := INDENT2 + '<Date1904/>' + LF else
    datemodeStr := '';

  protectStr := Format(
    '<ProtectStructure>%s</ProtectStructure>' + LF + INDENT2 +
    '<ProtectWindows>%s</ProtectWindows>' + LF, [
    FALSE_TRUE[bpLockStructure in Workbook.Protection],
    FALSE_TRUE[bpLockWindows in Workbook.Protection]
  ]);

  AppendToStream(AStream, INDENT1 +
    '<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">' + LF +
      datemodeStr + INDENT2 +
      protectStr + INDENT1 +
    '</ExcelWorkbook>' + LF);
end;

procedure TsSpreadExcelXMLWriter.WriteFormula(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
var
  xmlnsStr: String;
  dataTagStr: String;
begin
  if ACell^.ContentType <> cctFormula then
    raise Exception.Create('WriteFormula called for calculated cell.');

  xmlnsStr := ' xmlns="http://www.w3.org/TR/REC-html40"';
  dataTagStr := '';  // or 'ss:' -- to do...

  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s%s>' + LF + VALUE_INDENT + // colIndex, style, formula, hyperlink, merge
      '<%sData%s>'+             // "ss:", data type, "xmlns=.."
      '</%sData>' + LF + CELL_INDENT +       // "ss:"
      '%s' +                                 // Comment
    '</Cell>' + LF, [
    GetIndexStr(ACell^.Col, FPrevCol), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
      GetHyperlinkStr(ACell), GetMergeStr(ACell),
    dataTagStr, xmlnsStr,
    dataTagStr,
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
const
  MAXBYTES = 32767;  // limit for this format
var
  valueStr: String;
  cctStr: String;
  xmlnsStr: String;
  dataTagStr: String;
  p: Integer;
  tmp: String;
  ResultingValue: String;
begin
  // Office 2007-2010 (at least) supports no more characters in a cell;
  if Length(AValue) > MAXBYTES then
  begin
    ResultingValue := Copy(AValue, 1, MAXBYTES); //may chop off multicodepoint UTF8 characters but well...
    Workbook.AddErrorMsg(rsTruncateTooLongCellText, [
      MAXBYTES, GetCellString(ARow, ACol)
    ]);
  end else
    resultingValue := AValue;

  { Check for invalid characters }
  if not ValidXMLText(ResultingValue) then
    Workbook.AddErrorMsg(
      rsInvalidCharacterInCell, [
      GetCellString(ARow, ACol)
    ]);

  if Length(ACell^.RichTextParams) > 0 then
  begin
    RichTextToHTML(
      FWorkbook as TsWorkbook,
      (FWorksheet as TsWorksheet).ReadCellFont(ACell),
      ResultingValue,
      ACell^.RichTextParams,
      valueStr,             // html-formatted rich text
      'html:', tcProperCase
    );
    xmlnsStr := ' xmlns="http://www.w3.org/TR/REC-html40"';
    dataTagStr := 'ss:';

    // Excel does not like units in font size specification...
    tmp := valueStr;
    p := pos('<Font html:Size="', valueStr);
    if p > 0 then begin
      valueStr := '';
      while p > 0 do begin
        inc(p, Length('<Font html:Size="'));
        valueStr := valueStr + copy(tmp, 1, p-1);
        while (tmp[p] <> '"') do begin
          if (tmp[p] in ['0'..'9', '.']) then valueStr := valueStr + tmp[p];
          inc(p);
        end;
        tmp := copy(tmp, p, MaxInt);
        p := pos('<Font html:Size="', tmp);
      end;
      valueStr := valuestr + tmp;
    end;
  end else
  begin
    valueStr := ResultingValue;
    if not ValidXMLText(valueStr, true, true) then
      Workbook.AddErrorMsg(
        rsInvalidCharacterInCell, [
        GetCellString(ARow, ACol)
      ]);
    xmlnsStr := '';
    dataTagStr := '';
  end;

  cctStr := 'String';
  if HasFormula(ACell) then
    cctStr := GetCellContentTypeStr(ACell) else
    cctStr := 'String';

  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s%s>' + LF + VALUE_INDENT + // colIndex, style, formula, hyperlink, merge
      '<%sData ss:Type="%s"%s>'+             // "ss:", data type, "xmlns=.."
        '%s' +                               // value string
      '</%sData>' + LF + CELL_INDENT +       // "ss:"
      '%s' +                                 // Comment
    '</Cell>' + LF, [
    GetIndexStr(ACol, FPrevCol), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
      GetHyperlinkStr(ACell), GetMergeStr(ACell),
    dataTagStr, cctStr, xmlnsStr,
    valueStr,
    dataTagStr,
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteNames(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  sheet: TsWorksheet absolute AWorksheet;
  print_titles_str: string = '';
  print_range_str: String = '';
  s: String;
  rng: TsCellRange;
  i: Integer;
begin
  with sheet.PageLayout do begin

    // Print ranges --> Name "Print_Area"
    for i:=0 to NumPrintRanges-1 do begin
      rng := GetPrintRange(i);
      s := GetCellRangeString_R1C1(sheet.Name, sheet.Name, rng.Row1, rng.Col1, rng.Row2, rng.Col2, []);
      if print_range_str = '' then
        print_range_str := s
      else
        print_range_str := print_range_str + ',' + s;
    end;
    if print_range_str <> '' then
      print_range_str := NAME_INDENT +
        '<NamedRange ss:Name="Print_Area" ss:RefersTo="' + print_range_str + '"/>' + LF;

    // Repeated columns  -->  Name "Print_Titles"
    if (RepeatedCols.FirstIndex <> UNASSIGNED_ROW_COL_INDEX) and
       (RepeatedCols.LastIndex <> UNASSIGNED_ROW_COL_INDEX)
    then begin
      s := 'C' + IntToStr(RepeatedCols.FirstIndex + 1);
      if RepeatedCols.FirstIndex <> RepeatedCols.LastIndex then
        s := s + ':C' + IntToStr(RepeatedCols.LastIndex + 1);
      s := sheet.Name + '!' + s;
      print_titles_str := s;
    end;

    // Repeated rows  -->  Name "Print_Titles"
    if (RepeatedRows.FirstIndex <> UNASSIGNED_ROW_COL_INDEX) and
       (RepeatedRows.LastIndex <> UNASSIGNED_ROW_COL_INDEX)
    then begin
      s := 'R' + IntToStr(RepeatedRows.FirstIndex + 1);
      if RepeatedRows.FirstIndex <> RepeatedRows.LastIndex then
        s := s + ':R' + IntToStr(RepeatedRows.LastIndex + 1);
      s := sheet.Name + '!' + s;
      if print_titles_str = '' then
        print_titles_str := s
      else
        print_titles_str := print_titles_str + ',' + s;
    end;
    if print_titles_str <> '' then
      print_titles_str := NAME_INDENT +
        '<NamedRange ss:Name="Print_Titles" ss:RefersTo="' + print_titles_str + '"/>' + LF;
  end;

  if (print_range_str = '') and (print_titles_str = '') then
    exit;

  AppendToStream(AStream, NAMES_INDENT +
    '<Names>' + LF +
      print_titles_str + NAMES_INDENT +
      print_range_str + NAMES_INDENT +
    '</Names>' + LF);
end;

procedure TsSpreadExcelXMLWriter.WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: double; ACell: PCell);
begin
  Unused(ARow, ACol);
  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s%s>' + LF + VALUE_INDENT +  // colIndex, style, formula, hyperlink, merge
      '<Data ss:Type="%s">' +                 // data type
        '%g' +                                // value
      '</Data>' + LF + CELL_INDENT +
      '%s' +                                  // Comment <Comment>...</Comment>
    '</Cell>' + LF, [
    GetIndexStr(ACol, FPrevCol), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
      GetHyperlinkStr(ACell), GetMergeStr(ACell),
    StrUtils.IfThen(HasFormula(ACell), GetCellContentTypeStr(ACell), 'Number'),
    AValue,
    GetCommentStr(ACell)], FPointSeparatorSettings)
  );
end;

procedure TsSpreadExcelXMLWriter.WritePageBreaks(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  i: Integer;
  nc, nr: Integer;
  sheet: TsWorksheet absolute AWorksheet;
  s: String;
  col: PCol;
  row: PRow;
begin
  nc := 0;
  for i := 0 to sheet.Cols.Count - 1 do
    if (croPageBreak in PCol(sheet.Cols[i])^.Options) then inc(nc);

  nr := 0;
  for i:= 0 to sheet.Rows.Count - 1 do
    if (croPageBreak in PRow(sheet.Rows[i])^.Options) then inc(nr);

  if (nc = 0) and (nr = 0) then
    exit;

  s := INDENT2 +
    '<PageBreaks xmlns="urn:schemas-microsoft-com:office:excel">' + LF;

  if nc > 0 then begin
    s := s + INDENT3 +
      '<ColBreaks>' + LF;
    for i := 0 to sheet.Cols.Count - 1 do begin
      col := PCol(sheet.Cols[i]);
      if (croPageBreak in col^.Options) then
        s := s + INDENT4 +
          '<ColBreak>' + LF + INDENT5 +
            '<Column>' + IntToStr(col^.Col) + '</Column>' + LF + INDENT4 +
          '</ColBreak>' + LF;
    end;
    s := s + INDENT3 +
        '</ColBreaks>' + LF;
  end;

  if nr > 0 then begin
    s := s + INDENT3 +
      '<RowBreaks>' + LF;
    for i := 0 to sheet.Rows.Count - 1 do begin
      row := PRow(sheet.Rows[i]);
      if (croPageBreak in row^.Options) then
        s := s + INDENT4 +
          '<RowBreak>' + LF + INDENT5 +
            '<Row>' + IntToStr(row^.Row) + '</Row>' + LF + INDENT4 +
          '</RowBreak>' + LF;
    end;
    s := s + INDENT3 +
      '</RowBreaks>' + LF;
  end;

  s := s + INDENT2 +
    '</PageBreaks>' + LF;

  AppendToStream(AStream, s);
end;

procedure TsSpreadExcelXMLWriter.WriteRows(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  c: Cardinal;
  r: Cardinal;
  rowheightStr: String;
  hiddenStr: String;
  styleStr: String;
  s: String;
  row: PRow;
  cell: PCell;
  hasCells: Boolean;
  sheet: TsWorksheet absolute AWorksheet;
begin
  FPrevRow := UNASSIGNED_ROW_COL_INDEX;
  for r := 0 to FLastRow do
  begin
    row := sheet.FindRow(r);
    styleStr := '';
    hiddenStr := '';
    // Row height is needed in pts.
    if Assigned(row) then
    begin
      rowheightStr := Format(' ss:Height="%.2f"',
        [(FWorkbook as TsWorkbook).ConvertUnits(row^.Height, FWorkbook.Units, suPoints)],
        FPointSeparatorSettings
      );
      if row^.RowHeightType = rhtCustom then
        rowHeightStr := ' ss:AutoFitHeight="0"' + rowHeightStr
      else
        rowHeightStr := ' ss:AutoFitHeight="1"' + rowHeightStr;
      if row^.FormatIndex > 0 then
        styleStr := GetStyleStr(row^.FormatIndex);
    end else
      rowheightStr := ' ss:AutoFitHeight="1"';

    if sheet.RowHidden(r) then
      hiddenStr := ' ss:Hidden="1"';

    if boVirtualMode in FWorkbook.Options then
      hasCells := true
    else begin
      hasCells := false;
      for c := 0 to FLastCol do begin
        cell := sheet.FindCell(r, c);
        if cell <> nil then begin
          hasCells := true;
          break;
        end;
      end;
    end;

    s := Format('%s%s%s%s', [GetIndexStr(r, FPrevRow), rowheightStr, styleStr, hiddenStr]);
    if hasCells then begin
      AppendToStream(AStream, ROW_INDENT + Format(
        '<Row%s>', [s]) + LF);
      WriteCellNodes(AStream, AWorksheet, r);
      AppendToStream(AStream, ROW_INDENT +
        '</Row>' + LF);
      FPrevRow := r;
    end else
    if (rowheightStr <> '') or (styleStr <> '') or (hiddenStr <> '') then begin
      AppendToStream(AStream, ROW_INDENT + Format(
        '<Row%s/>', [s]) + LF);
      FPrevRow := r;
    end;
  end;
end;

procedure TsSpreadExcelXMLWriter.WriteStyle(AStream: TStream; AIndex: Integer);
var
  fmt: PsCellFormat;
  deffnt, fnt: TsFont;
  s, fmtVert, fmtHor, fmtWrap, fmtRot: String;
  nfp: TsNumFormatParams;
  fill: TsFillPattern;
  cb: TsCellBorder;
  cbs: TsCellBorderStyle;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;
  deffnt := book.GetDefaultFont;
  if AIndex = 0 then
  begin
    AppendToStream(AStream, Format(INDENT2 +
      '<Style ss:ID="Default" ss:Name="Normal">' + LF + INDENT3 +
        '<Aligment ss:Vertical="Bottom" />' + LF + INDENT3 +
        '<Borders />' + LF + INDENT3 +
        '<Font ss:FontName="%s" x:Family="Swiss" ss:Size="%d" ss:Color="%s" />' + LF + INDENT3 +
        '<Interior />' + LF + INDENT3 +
        '<NumberFormat />' + LF + INDENT3 +
        '<Protection />' + LF + INDENT2 +
      '</Style>' + LF,
      [deffnt.FontName, round(deffnt.Size), ColorToHTMLColorStr(deffnt.Color)] )
    )
  end else
  begin
    AppendToStream(AStream, Format(INDENT2 +
      '<Style ss:ID="s%d">' + LF, [AIndex + FMT_OFFSET]));

    fmt := book.GetPointerToCellFormat(AIndex);

    // Horizontal alignment
    fmtHor := '';
    if uffHorAlign in fmt^.UsedFormattingFields then
      case fmt^.HorAlignment of
        haDefault: ;
        haLeft   : fmtHor := 'ss:Horizontal="Left" ';
        haCenter : fmtHor := 'ss:Horizontal="Center" ';
        haRight  : fmtHor := 'ss:Horizontal="Right" ';
        else
          raise EFPSpreadsheetWriter.Create('[TsSpreadXMLWriter.WriteStyle] Horizontal alignment cannot be handled.');
      end;

    // Vertical alignment
    fmtVert := '';
    if uffVertAlign in fmt^.UsedFormattingFields then
      case fmt^.VertAlignment of
        vaDefault: ;
        vaTop    : fmtVert := 'ss:Vertical="Top" ';
        vaCenter : fmtVert := 'ss:Vertical="Center" ';
        vaBottom : fmtVert := 'ss:Vertical="Bottom" ';
        else
          raise EFPSpreadsheetWriter.Create('[TsSpreadXMLWriter.WriteStyle] Vertical alignment cannot be handled.');
      end;

    // Wrap text
    if uffWordwrap in fmt^.UsedFormattingFields then
      fmtWrap := 'ss:WrapText="1" ' else
      fmtWrap := '';

    // Text rotation
    fmtRot := '';
    if uffTextRotation in fmt^.UsedFormattingFields then
      case fmt^.TextRotation of
        rt90DegreeClockwiseRotation        : fmtRot := 'ss:Rotate="-90" ';
        rt90DegreeCounterClockwiseRotation : fmtRot := 'ss:Rotate="90" ';
        rtStacked                          : fmtRot := 'ss:VerticalText="1" ';
      end;

    // Write all the alignment, text rotation and wordwrap attributes to stream
    AppendToStream(AStream, Format(INDENT3 +
      '<Alignment %s%s%s%s />' + LF,
      [fmtHor, fmtVert, fmtWrap, fmtRot])
    );

    // Font
    if (uffFont in fmt^.UsedFormattingFields) then
    begin
      fnt := book.GetFont(fmt^.FontIndex);
      s := '';
      if fnt.FontName <> deffnt.FontName then
        s := s + Format('ss:FontName="%s" ', [fnt.FontName]);
      if not SameValue(fnt.Size, deffnt.Size, 1E-3) then
        s := s + Format('ss:Size="%g" ', [fnt.Size], FPointSeparatorSettings);
      if fnt.Color <> deffnt.Color then
        s := s + Format('ss:Color="%s" ', [ColorToHTMLColorStr(fnt.Color)]);
      if fssBold in fnt.Style then
        s := s + 'ss:Bold="1" ';
      if fssItalic in fnt.Style then
        s := s + 'ss:Italic="1" ';
      if fssUnderline in fnt.Style then
        s := s + 'ss:Underline="Single" ';    // or "Double", not supported by fps
      if fssStrikeout in fnt.Style then
        s := s + 'ss:StrikeThrough="1" ';
      if s <> '' then
        AppendToStream(AStream, INDENT3 +
          '<Font ' + s + '/>' + LF);
    end;

    // Number Format
    if (uffNumberFormat in fmt^.UsedFormattingFields) then
    begin
      nfp := book.GetNumberFormat(fmt^.NumberFormatIndex);
      nfp.AllowLocalizedAMPM := false;    // Replace "AMPM" by "AM/PM"
      AppendToStream(AStream, Format(INDENT3 +
        '<NumberFormat ss:Format="%s"/>' + LF, [UTF8TextToXMLText(nfp.NumFormatStr)]));
    end;

    // Background
    if (uffBackground in fmt^.UsedFormattingFields) then
    begin
      fill := fmt^.Background;
      if fill.Style = fsNoFill then
        AppendToStream(AStream, INDENT3 + '<Interior />' + LF)
      else begin
        if fill.Style = fsSolidFill then
          s := 'ss:Color="' + ColorToHtmlColorStr(fill.FgColor) + '" '
        else
          s := Format('ss:Color="%s" ss:PatternColor="%s" ', [
            ColorToHTMLColorStr(fill.BgColor),
            ColorToHTMLColorStr(fill.FgColor)
          ]);
        s := s + 'ss:Pattern="' + FILL_NAMES[fill.Style] + '" ';
        AppendToStream(AStream, INDENT3 +
          '<Interior ' + s + '/>' + LF)
      end;
    end;

    // Borders
    if (uffBorder in fmt^.UsedFormattingFields) then
    begin
      s := '';
      for cb in TsCellBorder do
        if cb in fmt^.Border then begin
          cbs := fmt^.BorderStyles[cb];
          s := s + INDENT4 + Format('<Border ss:Position="%s" ss:LineStyle="%s"', [
            BORDER_NAMES[cb], LINE_STYLES[cbs.LineStyle]]);
          if fmt^.BorderStyles[cb].LineStyle <> lsHair then
            s := Format('%s ss:Weight="%d"', [s, LINE_WIDTHS[cbs.LineStyle]]);
          s := Format('%s ss:Color="%s"', [s, ColorToHTMLColorStr(cbs.Color)]);
          s := s + '/>' + LF;
        end;
      if s <> '' then
        AppendToStream(AStream, INDENT3 +
          '<Borders>' + LF + s + INDENT3 +
          '</Borders>' + LF);
    end;

    // Protection
    s := '';
    if not (cpLockCell in fmt^.Protection) then
      s := s + 'ss:Protected="0" ';
    if cpHideFormulas in fmt^.Protection then
      s := s + 'x:HideFormula="1" ';
    if s <> '' then
      AppendToStream(AStream, INDENT3 +
        '<Protection ' + s + '/>' + LF);

    AppendToStream(AStream, INDENT2 +
      '</Style>' + LF);
  end;
end;

procedure TsSpreadExcelXMLWriter.WriteStyles(AStream: TStream);
var
  i: Integer;
begin
  AppendToStream(AStream, INDENT1 +
    '<Styles>' + LF);
  for i:=0 to (FWorkbook as TsWorkbook).GetNumCellFormats-1 do
    WriteStyle(AStream, i);
  AppendToStream(AStream, INDENT1 +
    '</Styles>' + LF);
end;

procedure TsSpreadExcelXMLWriter.WriteTable(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  sheet: TsWorksheet absolute AWorksheet;
begin
  AppendToStream(AStream, TABLE_INDENT + Format(
    '<Table ss:ExpandedColumnCount="%d" ss:ExpandedRowCount="%d" ' +
      'x:FullColumns="1" x:FullRows="1" ' +
      'ss:DefaultColumnWidth="%.2f" ' +
      'ss:DefaultRowHeight="%.2f">' + LF,
      [
      FLastCol + 1, FLastRow + 1,
      sheet.ReadDefaultColWidth(suPoints),
      sheet.ReadDefaultRowHeight(suPoints)
      ],
      FPointSeparatorSettings
    ));

  WriteColumns(AStream, AWorksheet);
  WriteRows(AStream, AWorksheet);

  AppendToStream(AStream, TABLE_INDENT +
    '</Table>' + LF);
end;

{@@ ----------------------------------------------------------------------------
  Writes an ExcelXML document to a stream
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLWriter.WriteToStream(AStream: TStream;
  AParams: TsStreamParams = []);
begin
  Unused(AParams);

  AppendToStream(AStream,
    '<?xml version="1.0"?>' + LF +
    '<?mso-application progid="Excel.Sheet"?>' + LF
  );
  AppendToStream(AStream,
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"' + LF +
    '          xmlns:o="urn:schemas-microsoft-com:office:office"' + LF +
    '          xmlns:x="urn:schemas-microsoft-com:office:excel"' + LF +
    '          xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"' + LF +
    '          xmlns:html="http://www.w3.org/TR/REC-html40">' + LF);

  WriteExcelWorkbook(AStream);
  WriteStyles(AStream);
  WriteWorksheets(AStream);

  AppendToStream(AStream,
    '</Workbook>');
end;

procedure TsSpreadExcelXMLWriter.WriteWorksheet(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  protectedStr: String;
begin
  FWorksheet := AWorksheet;
  GetSheetDimensions(FWorksheet, FFirstRow, FLastRow, FFirstCol, FLastCol);

  if FWorksheet.IsProtected then
    protectedStr := ' ss:Protected="1"'
  else
    protectedStr := '';

  AppendToStream(AStream, Format(
    '  <Worksheet ss:Name="%s"%s>' + LF, [
    UTF8TextToXMLText(AWorksheet.Name),
    protectedStr
  ]) );
  WriteNames(AStream, AWorksheet);
  WriteTable(AStream, AWorksheet);
  WriteWorksheetOptions(AStream, AWorksheet);
  WritePageBreaks(AStream, AWorksheet);
  AppendToStream(AStream,
    '  </Worksheet>' + LF
  );
end;

procedure TsSpreadExcelXMLWriter.WriteWorksheetOptions(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  footerStr, headerStr: String;
  hideGridStr: String;
  hideHeadersStr: String;
  frozenStr: String;
  layoutStr: String;
  marginStr: String;
  selectedStr: String;
  protectStr: String;
  visibleStr: String;
  printStr: String;
  fitToPageStr: String;
  enableSelectionStr: String;
  sheet: TsWorksheet absolute AWorksheet;
begin
  // Orientation, some PageLayout.Options
  layoutStr := GetLayoutStr(AWorksheet);
  if layoutStr <> '' then layoutStr := INDENT4 + layoutStr + LF;

  // Header
  headerStr := GetPageHeaderStr(AWorksheet);
  if headerStr <> '' then headerStr := INDENT4 + headerStr + LF;

  // Footer
  footerStr := GetPageFooterStr(AWorksheet);
  if footerStr <> '' then footerStr := INDENT4 + footerStr + LF;

  // Page margins
  marginStr := GetPageMarginStr(AWorksheet);
  if marginStr <> '' then marginStr := INDENT4 + marginStr + LF;

  // Show/hide grid lines
  if not (soShowGridLines in AWorksheet.Options) then
    hideGridStr := INDENT3 + '<DoNotDisplayGridlines/>' + LF
  else
    hideGridStr := '';

  // Show/hide column/row headers
  if not (soShowHeaders in AWorksheet.Options) then
    hideHeadersStr := INDENT3 + '<DoNotDisplayHeadings/>' + LF
  else
    hideHeadersStr := '';

  if (FWorkbook as TsWorkbook).ActiveWorksheet = AWorksheet then
    selectedStr := INDENT3 + '<Selected/>' + LF
  else
    selectedStr := '';

  // FitToPage node
  if poFitPages in sheet.PageLayout.Options then
    fitToPageStr := INDENT3 + '<FitToPage/>' + LF
  else
    fitToPageStr := '';

  // Print node
  printStr := GetPrintStr(AWorksheet);

  // Visible
  if (soHidden in AWorksheet.Options) then
    visibleStr := INDENT3 + '<Visible>SheetHidden</Visible>' + LF
  else
    visibleStr := '';

  // Frozen panes
  frozenStr := GetFrozenPanesStr(AWorksheet, INDENT3);

  // Protection
  protectStr := Format(INDENT3 + '<ProtectObjects>%s</ProtectObjects>' + LF +
                       INDENT3 + '<ProtectScenarios>%s</ProtectScenarios>' + LF, [
    StrUtils.IfThen(spObjects in AWorksheet.Protection, 'True', 'False'),
    StrUtils.IfThen(AWorksheet.IsProtected {and [spScenarios in AWorksheet.Protection])}, 'True', 'False')
  ]);

  // Enable selection
  enableSelectionStr := '';
  if (sheet.Protection * [spSelectLockedCells, spSelectUnlockedCells] <> []) then begin
    enableSelectionStr := INDENT3 + '<EnableSelection>' + LF;
    if spSelectUnlockedCells in sheet.Protection then
      enableSelectionStr := enableSelectionStr + INDENT4 + '<NoSelection/>' + LF;
    if (sheet.Protection * [spSelectLockedCells, spSelectUnlockedCells] = [spSelectLockedCells]) then
      enableSelectionStr := enableSelectionStr + INDENT4 + '<Unlocked/>' + LF;
    enableSelectionStr := INDENT3 + '</EnableSelection>' + LF;
  end;

  // todo - Several protection options

  // Put it all together...
  AppendToStream(AStream, INDENT2 +
    '<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">' + LF + INDENT3 +
      '<PageSetup>' + LF +
        layoutStr +
        headerStr +
        footerStr +
        marginStr + INDENT3 +
      '</PageSetup>' + LF +
      fitToPageStr + INDENT3 +
      '<Print>' + LF +
        printStr + LF + INDENT3 +
      '</Print>' + LF +
      visibleStr +
      selectedStr +
      IfThen(not (spFormatCells in sheet.Protection), INDENT4 + '<AllowFormatCells/>' + LF) +
      IfThen(not (spFormatColumns in sheet.Protection), INDENT4 + '<AllowSizeCols/>' + LF) +
      IfThen(not (spFormatRows in sheet.Protection), INDENT4 + '<AllowSizeRows/>' + LF) +
      IfThen(not (spDeleteColumns in sheet.Protection), INDENT4 + '<AllowDeleteCols/>' + LF) +
      IfThen(not (spDeleteRows in sheet.Protection), INDENT4 + '<AllowDeleteRows/>' + LF) +
      IfThen(not (spInsertColumns in sheet.Protection), INDENT4 + '<AllowInsertCols/>' + LF) +
      IfThen(not (spInsertHyperlinks in sheet.Protection), INDENT4 + '<AllowInsertHyperlinks/>' + LF) +
      IfThen(not (spInsertRows in sheet.Protection), INDENT4 + '<AllowInsertRows/>' + LF) +
      IfThen(not (spSort in sheet.Protection), INDENT4 + '<AllowSort/>' + LF) +
      enableSelectionStr +
      protectStr +
      frozenStr +
      hideGridStr +
      hideHeadersStr + INDENT2 +
    '</WorksheetOptions>' + LF
  );
end;

procedure TsSpreadExcelXMLWriter.WriteWorksheets(AStream: TStream);
var
  i: Integer;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;
  for i:=0 to book.GetWorksheetCount-1 do
    WriteWorksheet(AStream, book.GetWorksheetByIndex(i));
end;


initialization

  // Registers this reader / writer in fpSpreadsheet
  sfidExcelXML := RegisterSpreadFormat(sfExcelXML,
    TsSpreadExcelXMLReader, TsSpreadExcelXMLWriter,
    STR_FILEFORMAT_EXCEL_XML, 'ExcelXML', [STR_XML_EXCEL_EXTENSION]
  );

end.
