{-------------------------------------------------------------------------------
Unit     : xlsxml

Implements a reader and writer for the SpreadsheetXML format.
This document was introduced by Microsoft for Excel XP and 2003.

REFERENCE: https://msdn.microsoft.com/en-us/library/aa140066%28v=office.15%29.aspx

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
  fpsTypes, fpspreadsheet, fpsReaderWriter, xlsCommon;

type

  { TsSpreadExcelXMLWriter }

  TsSpreadExcelXMLWriter = class(TsCustomSpreadWriter)
  private
    FDateMode: TDateMode;
    FPointSeparatorSettings: TFormatSettings;
    function GetCommentStr(ACell: PCell): String;
    function GetFormulaStr(ACell: PCell): String;
    function GetHyperlinkStr(ACell: PCell): String;
    function GetIndexStr(AIndex: Integer): String;
    function GetMergeStr(ACell: PCell): String;
    function GetStyleStr(ACell: PCell): String;
    procedure WriteCells(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteExcelWorkbook(AStream: TStream);
    procedure WriteStyle(AStream: TStream; AIndex: Integer);
    procedure WriteStyles(AStream: TStream);
    procedure WriteWorksheet(AStream: TStream; AWorksheet: TsWorksheet);
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
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    procedure WriteToFile(const AFileName: string; const AOverwriteExisting: Boolean = False); override;
    procedure WriteToStream(AStream: TStream); override;

  end;

  TExcelXmlSettings = record
    DateMode: TDateMode;
  end;

var
  ExcelXmlSettings: TExcelXmlSettings = (
    DateMode: dm1900;
  );


implementation

uses
  StrUtils, Math,
  fpsStrings, fpsUtils, fpsStreams, fpsNumFormat, fpsHTMLUtils;

const
  FMT_OFFSET   = 61;
  INDENT1      = '  ';
  INDENT2      = '    ';
  INDENT3      = '      ';
  INDENT4      = '        ';
  INDENT5      = '          ';
  VALUE_INDENT = INDENT5;
  CELL_INDENT  = INDENT4;
  ROW_INDENT   = INDENT3;
  COL_INDENT   = INDENT3;
  TABLE_INDENT = INDENT2;
  LF           = LineEnding;

const
  { TsFillStyle = (
    fsNoFill, fsSolidFill,
    fsGray75, fsGray50, fsGray25, fsGray12, fsGray6,
    fsStripeHor, fsStripeVert, fsStripeDiagUp, fsStripeDiagDown,
    fsThinStripeHor, fsThinStripeVert, fsThinStripeDiagUp, fsThinStripeDiagDown,
    fsHatchDiag, fsThinHatchDiag, fsThickHatchDiag, fsThinHatchHor) }
  FILL_NAMES: array[TsFillStyle] of string = (
    '', 'Solid',
    'Gray75', 'Gray50', 'Gray25', 'Gray12', 'Gray0625',
    'HorzStripe', 'VertStripe', 'DiagStripe', 'ReverseDiagStripe',
    'ThinHorzStripe', 'ThinVertStripe', 'ThinDiagStripe', 'ThinReverseDiagStripe',
    'DiagCross', 'ThinDiagCross', 'ThickDiagCross', 'ThinHorzCross'
  );

  {TsCellBorder = (cbNorth, cbWest, cbEast, cbSouth, cbDiagUp, cbDiagDown); }
  BORDER_NAMES: array[TsCellBorder] of string = (
    'Top', 'Left', 'Right', 'Bottom', 'DiagonalRight', 'DiagonalLeft'
  );

  { TsLineStyle = (
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

function GetCellContentTypeStr(ACell: PCell): String;
begin
  case ACell^.ContentType of
    cctNumber     : Result := 'Number';
    cctUTF8String : Result := 'String';
    cctDateTime   : Result := 'DateTime';
    cctBool       : Result := 'Boolean';
    cctError      : Result := 'Error';
  else
    raise Exception.Create('Content type error in cell ' + GetCellString(ACell^.Row, ACell^.Col));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Constructor of the ExcelXML writer

  Defines the date mode and the limitations of the file format.
  Initializes the format settings to be used when writing to xml.
-------------------------------------------------------------------------------}
constructor TsSpreadExcelXMLWriter.Create(AWorkbook: TsWorkbook);
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
  comment := FWorksheet.FindComment(ACell);
  if Assigned(comment) then
    Result := INDENT1 + '<Comment><Data>' + comment^.Text + '</Data></Comment>' + LF + CELL_INDENT;
  // If there will be some rich-text-like formatting in the future, use
  //  Result := '<Comment><ss:Data xmlns="http://www.w3.org/TR/REC-html40">'+comment^.Text+'</ss:Data></Comment>':
end;

function TsSpreadExcelXMLWriter.GetFormulaStr(ACell: PCell): String;
begin
  if HasFormula(ACell) then
  begin
    Result := UTF8TextToXMLText(FWorksheet.ConvertFormulaDialect(ACell, fdExcelR1C1));
    Result := ' ss:Formula="=' + Result + '"';
  end else
    Result := '';
end;

function TsSpreadExcelXMLWriter.GetHyperlinkStr(ACell: PCell): String;
var
  hyperlink: PsHyperlink;
begin
  Result := '';
  hyperlink := FWorksheet.FindHyperlink(ACell);
  if Assigned(hyperlink) then
    Result := ' ss:HRef="' + hyperlink^.Target + '"';
end;

function TsSpreadExcelXMLWriter.GetIndexStr(AIndex: Integer): String;
begin
  Result := Format(' ss:Index="%d"', [AIndex]);
end;

function TsSpreadExcelXMLWriter.GetMergeStr(ACell: PCell): String;
var
  r1, c1, r2, c2: Cardinal;
begin
  Result := '';
  if FWorksheet.IsMerged(ACell) then begin
    FWorksheet.FindMergedRange(ACell, r1, c1, r2, c2);
    if c2 > c1 then
      Result := Result + Format(' ss:MergeAcross="%d"', [c2-c1]);
    if r2 > r1 then
      Result := Result + Format(' ss:MergeDown="%d"', [r2-r1]);
  end;
end;

function TsSpreadExcelXMLWriter.GetStyleStr(ACell: PCell): String;
begin
  Result := '';
  if ACell^.FormatIndex > 0 then
    Result := Format(' ss:StyleID="s%d"', [ACell^.FormatIndex + FMT_OFFSET]);
end;

procedure TsSpreadExcelXMLWriter.WriteBlank(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
begin
  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s>' +              // colIndex, style, hyperlink, merge
      '%s' +                        // Comment <Comment>...</Comment>
    '</Cell>' + LF, [
    GetIndexStr(ACol+1), GetStyleStr(ACell), GetHyperlinkStr(ACell), GetMergeStr(ACell),
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteBool(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: boolean; ACell: PCell);
begin
  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s%s>' +         // colIndex, style, formula, hyperlink, merge
      '<Data ss:Type="%s">' +    // data type
        '%s' +                   // value string
      '</Data>' +
      '%s' +                     // Comment <Comment>...</Comment>
    '</Cell>' + LF, [
    GetIndexStr(ACol+1), GetStyleStr(ACell), GetFormulaStr(ACell), GetHyperlinkStr(ACell), GetMergeStr(ACell),
    StrUtils.IfThen(HasFormula(ACell), GetCellContentTypeStr(ACell), 'Boolean'),
    StrUtils.IfThen(AValue, '1', '0'),
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteCells(AStream: TStream; AWorksheet: TsWorksheet);
var
  c, c1, c2: Cardinal;
  r, r1, r2: Cardinal;
  cell: PCell;
  rowheightStr: String;
  colwidthStr: String;
  defFnt: TsFont;
  col: PCol;
  row: PRow;
  cw_fact, rh_fact: Double;
begin
  defFnt := FWorkbook.GetDefaultFont;
  cw_fact := defFnt.Size * 0.5;  // ColWidthFactor = Approx width of "0" character in pts
  rh_fact := defFnt.Size;        // RowHeightFactor = Height of a single line

  r1 := 0;
  c1 := 0;
  r2 := AWorksheet.GetLastRowIndex;
  c2 := AWorksheet.GetLastColIndex;
  AppendToStream(AStream, TABLE_INDENT + Format(
    '<Table ss:ExpandedColumnCount="%d" ss:ExpandedRowCount="%d" ' +
      'x:FullColumns="1" x:FullRows="1" ' +
      'ss:DefaultColumnWidth="%.2f" ' +
      'ss:DefaultRowHeight="%.2f">' + LF,
      [
      AWorksheet.GetLastColIndex + 1, AWorksheet.GetLastRowIndex + 1,
      FWorksheet.DefaultColWidth * cw_fact,
      (FWorksheet.DefaultRowHeight + ROW_HEIGHT_CORRECTION) * rh_fact
      ],
      FPointSeparatorSettings
    ));

  for c := c1 to c2 do
  begin
    col := FWorksheet.FindCol(c);
    // column width in the worksheet is in multiples of the "0" character width.
    // In the xml file, it is needed in pts.
    if Assigned(col) then
      colwidthStr := Format(' ss:Width="%0.2f"',
        [col^.Width * cw_fact],
        FPointSeparatorSettings)
    else
      colwidthStr := '';
    AppendToStream(AStream, COL_INDENT + Format(
      '<Column ss:Index="%d" ss:AutoFitWidth="0"%s />' + LF, [c+1, colWidthStr]));
  end;

  for r := r1 to r2 do
  begin
    row := FWorksheet.FindRow(r);
    // Row height in the worksheet is in multiples of the default font height
    // In the xml file, it is needed in pts.
    if Assigned(row) then
      rowheightStr := Format(' ss:Height="%.2f"',
        [(row^.Height + ROW_HEIGHT_CORRECTION) * rh_fact],
        FPointSeparatorSettings
      )
    else
      rowheightStr := '';
    AppendToStream(AStream, ROW_INDENT + Format(
      '<Row ss:AutoFitHeight="1"%s>' + LF, [rowheightStr]));
    for c := c1 to c2 do
    begin
      cell := AWorksheet.FindCell(r, c);
      if cell <> nil then
      begin
        if FWorksheet.IsMerged(cell) and not FWorksheet.IsMergeBase(cell) then
          Continue;
        WriteCellToStream(AStream, cell);
      end;
    end;
    AppendToStream(AStream, ROW_INDENT +
      '</Row>' + LF);
  end;

  AppendToStream(AStream, TABLE_INDENT +
    '</Table>' + LF);
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
  end;

  if FWorksheet.ReadComment(ACell) <> '' then
    WriteComment(AStream, ACell);
end;

procedure TsSpreadExcelXMLWriter.WriteDateTime(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TDateTime; ACell: PCell);
var
  valueStr: String;
  ExcelDate: TDateTime;
  nfp: TsNumFormatParams;
  fmt: PsCellFormat;
begin
  ExcelDate := AValue;
  fmt := FWorkbook.GetPointerToCellFormat(ACell^.FormatIndex);
  // Times have an offset of 1 day!
  if (fmt <> nil) and (uffNumberFormat in fmt^.UsedFormattingFields) then
  begin
    nfp := FWorkbook.GetNumberFormat(fmt^.NumberFormatIndex);
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
    GetIndexStr(ACol+1), GetStyleStr(ACell), GetFormulaStr(ACell), GetHyperlinkStr(ACell), GetMergeStr(ACell),
    StrUtils.IfThen(HasFormula(ACell), GetCellContentTypeStr(ACell), 'DateTime'),
    valueStr,
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
begin
  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s%s>' + LF + VALUE_INDENT + // colIndex, style, formula, hyperlink, merge
      '<Data ss:Type="%s">' +                // data type
        '%s' +                               // value string
      '</Data>' + LF + CELL_INDENT +
      '%s' +                                 // Comment <Comment>...</Comment>
    '</Cell>' + LF, [
    GetIndexStr(ACol+1), GetStyleStr(ACell), GetFormulaStr(ACell), GetHyperlinkStr(ACell), GetMergeStr(ACell),
    StrUtils.IfThen(HasFormula(ACell), GetCellContentTypeStr(ACell), 'Error'),
    GetErrorValueStr(AValue),
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteExcelWorkbook(AStream: TStream);
var
  datemodeStr: String;
begin
  if FDateMode = dm1904 then
    datemodeStr := INDENT2 + '<Date1904/>' + LF else
    datemodeStr := '';

  AppendToStream(AStream, INDENT1 +
    '<ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">' + LF +
      datemodeStr + INDENT2 +
      '<ProtectStructure>False</ProtectStructure>' + LF + INDENT2 +
      '<ProtectWindows>False</ProtectWindows>' + LF + INDENT1 +
    '</ExcelWorkbook>' + LF);
end;

procedure TsSpreadExcelXMLWriter.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
var
  valueStr: String;
  cctStr: String;
  xmlnsStr: String;
  dataTagStr: String;
begin
  if Length(ACell^.RichTextParams) > 0 then
  begin
    RichTextToHTML(
      FWorkbook,
      FWorksheet.ReadCellFont(ACell),
      AValue,
      ACell^.RichTextParams,
      valueStr,             // html-formatted rich text
      'html:', tcProperCase
    );
    xmlnsStr := ' xmlns="http://www.w3.org/TR/REC-html40"';
    dataTagStr := 'ss:';
  end else
  begin
    valueStr := AValue;
    if not ValidXMLText(valueStr) then
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
    GetIndexStr(ACol+1), GetStyleStr(ACell), GetFormulaStr(ACell), GetHyperlinkStr(ACell), GetMergeStr(ACell),
    dataTagStr, cctStr, xmlnsStr,
    valueStr,
    dataTagStr,
    GetCommentStr(ACell)
  ]));
end;

procedure TsSpreadExcelXMLWriter.WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: double; ACell: PCell);
begin
  AppendToStream(AStream, Format(CELL_INDENT +
    '<Cell%s%s%s%s%s>' + LF + VALUE_INDENT +  // colIndex, style, formula, hyperlink, merge
      '<Data ss:Type="%s">' +                 // data type
        '%g' +                                // value
      '</Data>' + LF + CELL_INDENT +
      '%s' +                                  // Comment <Comment>...</Comment>
    '</Cell>' + LF, [
    GetIndexStr(ACol+1), GetStyleStr(ACell), GetFormulaStr(ACell), GetHyperlinkStr(ACell), GetMergeStr(ACell),
    StrUtils.IfThen(HasFormula(ACell), GetCellContentTypeStr(ACell), 'Number'),
    AValue,
    GetCommentStr(ACell)], FPointSeparatorSettings)
  );
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
begin
  deffnt := FWorkbook.GetDefaultFont;
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

    fmt := FWorkbook.GetPointerToCellFormat(AIndex);

    // Horizontal alignment
    fmtHor := '';
    if uffHorAlign in fmt^.UsedFormattingFields then
      case fmt^.HorAlignment of
        haDefault: ;
        haLeft   : fmtHor := 'ss:Horizontal="Left" ';
        haCenter : fmtHor := 'ss:Horizontal="Center" ';
        haRight  : fmtHor := 'ss:Horizontal="Right" ';
        else
          raise Exception.Create('[TsSpreadXMLWriter.WriteStyle] Horizontal alignment cannot be handled.');
      end;

    // Vertical alignment
    fmtVert := 'ss:Vertical="Bottom" ';
    if uffVertAlign in fmt^.UsedFormattingFields then
      case fmt^.VertAlignment of
        vaDefault: ;
        vaTop    : fmtVert := 'ss:Vertical="Top" ';
        vaCenter : fmtVert := 'ss:Vertical="Center" ';
        vaBottom : ;
        else
          raise Exception.Create('[TsSpreadXMLWriter.WriteStyle] Vertical alignment cannot be handled.');
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
      fnt := FWorkbook.GetFont(fmt^.FontIndex);
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
      nfp := FWorkbook.GetNumberFormat(fmt^.NumberFormatIndex);
      AppendToStream(AStream, Format(INDENT3 +
        '<NumberFormat ss:Format="%s"/>' + LF, [UTF8TextToXMLText(nfp.NumFormatStr)]));
    end;

    // Background
    if (uffBackground in fmt^.UsedFormattingFields) then
    begin
      fill := fmt^.Background;
      s := 'ss:Color="' + ColorToHTMLColorStr(fill.BgColor) + '" ';
      if not (fill.Style in [fsNoFill, fsSolidFill]) then
        s := s + 'ss:PatternColor="' + ColorToHTMLColorStr(fill.FgColor) + '" ';
      s := s + 'ss:Pattern="' + FILL_NAMES[fill.Style] + '"';
      AppendToStream(AStream, INDENT3 +
        '<Interior ' + s + '/>' + LF)
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
          if fmt^.BorderStyles[cb].Color <> scBlack then
            s := Format('%s ss:Color="%s"', [s, ColorToHTMLColorStr(cbs.Color)]);
          s := s + '/>' + LF;
        end;
      if s <> '' then
        AppendToStream(AStream, INDENT3 +
          '   <Borders>' + LF + s +
          '   </Borders>' + LF);
    end;

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
  for i:=0 to FWorkbook.GetNumCellFormats-1 do WriteStyle(AStream, i);
  AppendToStream(AStream, INDENT1 +
    '</Styles>' + LF);
end;

{@@ ----------------------------------------------------------------------------
  Writes an ExcelXML document to the file
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLWriter.WriteToFile(const AFileName: string;
  const AOverwriteExisting: Boolean);
var
  stream: TStream;
  mode: word;
begin
  mode := fmCreate or fmShareDenyNone;
  if AOverwriteExisting
    then mode := mode or fmOpenWrite;

  if (boBufStream in Workbook.Options) then
    stream := TBufStream.Create(AFileName, mode)
  else
    stream := TFileStream.Create(AFileName, mode);

  try
    WriteToStream(stream);
  finally
    FreeAndNil(stream);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an ExcelXML document to a stream
-------------------------------------------------------------------------------}
procedure TsSpreadExcelXMLWriter.WriteToStream(AStream: TStream);
begin
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
  AWorksheet: TsWorksheet);
begin
  FWorksheet := AWorksheet;
  AppendToStream(AStream, Format(
    '  <Worksheet ss:Name="%s">' + LF, [AWorksheet.Name]) );
  WriteCells(AStream, AWorksheet);
  AppendToStream(AStream,
    '  </Worksheet>' + LF
  );
end;

procedure TsSpreadExcelXMLWriter.WriteWorksheets(AStream: TStream);
var
  i: Integer;
begin
  for i:=0 to FWorkbook.GetWorksheetCount-1 do
    WriteWorksheet(AStream, FWorkbook.GetWorksheetByIndex(i));
end;


initialization

  // Registers this reader / writer in fpSpreadsheet
  RegisterSpreadFormat(nil, TsSpreadExcelXMLWriter, sfExcelXML);

end.
