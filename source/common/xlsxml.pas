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
  fpsTypes, fpspreadsheet, fpsReaderWriter, xlsCommon;

type

  { TsSpreadExcelXMLWriter }

  TsSpreadExcelXMLWriter = class(TsCustomSpreadWriter)
  private
    FDateMode: TDateMode;
    FPointSeparatorSettings: TFormatSettings;
    function GetCommentStr(ACell: PCell): String;
    function GetFormulaStr(ACell: PCell): String;
    function GetFrozenPanesStr(AWorksheet: TsWorksheet; AIndent: String): String;
    function GetHyperlinkStr(ACell: PCell): String;
    function GetIndexStr(AIndex: Integer): String;
    function GetLayoutStr(AWorksheet: TsWorksheet): String;
    function GetMergeStr(ACell: PCell): String;
    function GetPageFooterStr(AWorksheet: TsWorksheet): String;
    function GetPageHeaderStr(AWorksheet: TsWorksheet): String;
    function GetPageMarginStr(AWorksheet: TsWorksheet): String;
    function GetStyleStr(AFormatIndex: Integer): String;
    procedure WriteExcelWorkbook(AStream: TStream);
    procedure WriteStyle(AStream: TStream; AIndex: Integer);
    procedure WriteStyles(AStream: TStream);
    procedure WriteTable(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteWorksheet(AStream: TStream; AWorksheet: TsWorksheet);
    procedure WriteWorksheetOptions(AStream: TStream; AWorksheet: TsWorksheet);
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
  StrUtils, Math,
  fpsStrings, fpsUtils, fpsNumFormat, fpsXmlCommon, fpsHTMLUtils;

const
  FMT_OFFSET   = 61;

  INDENT1      = '  ';
  INDENT2      = '    ';
  INDENT3      = '      ';
  INDENT4      = '        ';
  INDENT5      = '          ';
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
    'Gray75', 'Gray50', 'Gray25', 'Gray12', 'Gray0625',
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

function TsSpreadExcelXMLWriter.GetFrozenPanesStr(AWorksheet: TsWorksheet;
  AIndent: String): String;
var
  activePane: Integer;
begin
  if (soHasFrozenPanes in AWorksheet.Options) then
  begin
    Result := AIndent +
        '<FreezePanes/>' + LF + AIndent +
        '<FrozenNoSplit/>' + LF;

    if FWorksheet.LeftPaneWidth > 0 then
      Result := Result + AIndent +
        '<SplitVertical>1</SplitVertical>' + LF + AIndent +
        '<LeftColumnRightPane>' + IntToStr(FWorksheet.LeftPaneWidth) + '</LeftColumnRightPane>' + LF;

    if FWorksheet.TopPaneHeight > 0 then
      Result := Result + AIndent +
        '<SplitHorizontal>1</SplitHorizontal>' + LF + AIndent +
        '<TopRowBottomPane>' + IntToStr(FWorksheet.TopPaneHeight) + '</TopRowBottomPane>' + LF;

    if (FWorksheet.LeftPaneWidth = 0) and (FWorkSheet.TopPaneHeight = 0) then
      activePane := 3
    else
    if (FWorksheet.LeftPaneWidth = 0) then
      activePane := 2
    else
    if (FWorksheet.TopPaneHeight = 0) then
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
  Result := '';
  hyperlink := FWorksheet.FindHyperlink(ACell);
  if Assigned(hyperlink) then
    Result := ' ss:HRef="' + hyperlink^.Target + '"';
end;

function TsSpreadExcelXMLWriter.GetIndexStr(AIndex: Integer): String;
begin
  Result := Format(' ss:Index="%d"', [AIndex]);
end;

function TsSpreadExcelXMLWriter.GetLayoutStr(AWorksheet: TsWorksheet): String;
begin
  Result := '';
  if AWorksheet.PageLayout.Orientation = spoLandscape then
    Result := Result + ' x:Orientation="Landscape"';
  if (poHorCentered in AWorksheet.PageLayout.Options) then
    Result := Result + ' x:CenterHorizontal="1"';
  if (poVertCentered in AWorksheet.PageLayout.Options) then
    Result := Result + ' x:CenterVertical="1"';
  if (poUseStartPageNumber in AWorksheet.PageLayout.Options) then
    Result := Result + ' x:StartPageNumber="' + IntToStr(AWorksheet.PageLayout.StartPageNumber) + '"';
  Result := '<Layout' + Result + '/>';
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

function TsSpreadExcelXMLWriter.GetPageFooterStr(AWorksheet: TsWorksheet): String;
begin
  Result := Format('x:Margin="%g"', [mmToIn(AWorksheet.PageLayout.FooterMargin)], FPointSeparatorSettings);
  if (AWorksheet.PageLayout.Footers[HEADER_FOOTER_INDEX_ALL] <> '') then
    Result := Result + ' x:Data="' + UTF8TextToXMLText(AWorksheet.PageLayout.Footers[HEADER_FOOTER_INDEX_ALL], true) + '"';
  Result := '<Footer ' + result + '/>';
end;

function TsSpreadExcelXMLWriter.GetPageHeaderStr(AWorksheet: TsWorksheet): String;
begin
  Result := Format('x:Margin="%g"', [mmToIn(AWorksheet.PageLayout.HeaderMargin)], FPointSeparatorSettings);
  if (AWorksheet.PageLayout.Headers[HEADER_FOOTER_INDEX_ALL] <> '') then
    Result := Result + ' x:Data="' + UTF8TextToXMLText(AWorksheet.PageLayout.Headers[HEADER_FOOTER_INDEX_ALL], true) + '"';
  Result := '<Header ' + Result + '/>';
end;

function TsSpreadExcelXMLWriter.GetPageMarginStr(AWorksheet: TsWorksheet): String;
begin
  Result := Format('x:Bottom="%g" x:Left="%g" x:Right="%g" x:Top="%g"', [
    mmToIn(AWorksheet.PageLayout.BottomMargin),
    mmToIn(AWorksheet.PageLayout.LeftMargin),
    mmToIn(AWorksheet.PageLayout.RightMargin),
    mmToIn(AWorksheet.PageLayout.TopMargin)
    ], FPointSeparatorSettings);
  Result := '<PageMargins ' + Result + '/>';
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
    GetIndexStr(ACol+1), GetStyleStr(ACell^.FormatIndex), GetHyperlinkStr(ACell), GetMergeStr(ACell),
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
    GetIndexStr(ACol+1), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
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
  Unused(ARow, ACol);
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
    GetIndexStr(ACol+1), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
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
    GetIndexStr(ACol+1), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
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
    GetIndexStr(ACol+1), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
      GetHyperlinkStr(ACell), GetMergeStr(ACell),
    dataTagStr, cctStr, xmlnsStr,
    valueStr,
    dataTagStr,
    GetCommentStr(ACell)
  ]));
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
    GetIndexStr(ACol+1), GetStyleStr(ACell^.FormatIndex), GetFormulaStr(ACell),
      GetHyperlinkStr(ACell), GetMergeStr(ACell),
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
          raise EFPSpreadsheetWriter.Create('[TsSpreadXMLWriter.WriteStyle] Horizontal alignment cannot be handled.');
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
          '<Borders>' + LF + s + INDENT3 +
          '</Borders>' + LF);
    end;

    // Protection
    s := '';
    if FWorkbook.IsProtected then begin
      if not (cpLockCell in fmt^.Protection) then
        s := s + 'ss:Protected="0" ';
      if cpHideFormulas in fmt^.Protection then
        s := s + 'x:HideFormula="1" ';
    end;
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
  for i:=0 to FWorkbook.GetNumCellFormats-1 do WriteStyle(AStream, i);
  AppendToStream(AStream, INDENT1 +
    '</Styles>' + LF);
end;

procedure TsSpreadExcelXMLWriter.WriteTable(AStream: TStream; AWorksheet: TsWorksheet);
var
  c, c1, c2: Cardinal;
  r, r1, r2: Cardinal;
  cell: PCell;
  rowheightStr: String;
  colwidthStr: String;
  styleStr: String;
  col: PCol;
  row: PRow;
begin
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
      AWorksheet.ReadDefaultColWidth(suPoints),
      AWorksheet.ReadDefaultRowHeight(suPoints)
      ],
      FPointSeparatorSettings
    ));

  for c := c1 to c2 do
  begin
    col := FWorksheet.FindCol(c);
    styleStr := '';
    colWidthStr := '';
    if Assigned(col) then
    begin
      // column width is needed in pts.
      if col^.ColWidthType = cwtCustom then
        colwidthStr := Format(' ss:Width="%0.2f" ss:AutoFitWidth="0"',
          [FWorkbook.ConvertUnits(col^.Width, FWorkbook.Units, suPoints)],
          FPointSeparatorSettings);
      // column style
      if col^.FormatIndex > 0 then
        styleStr := GetStyleStr(col^.FormatIndex);
    end;
    AppendToStream(AStream, COL_INDENT + Format(
      '<Column ss:Index="%d" %s%s />' + LF, [c+1, colWidthStr, styleStr]));
  end;

  for r := r1 to r2 do
  begin
    row := FWorksheet.FindRow(r);
    styleStr := '';
    // Row height is needed in pts.
    if Assigned(row) then
    begin
      rowheightStr := Format(' ss:Height="%.2f"',
        [FWorkbook.ConvertUnits(row^.Height, FWorkbook.Units, suPoints)],
        FPointSeparatorSettings
      );
      if row^.RowHeightType = rhtCustom then
        rowHeightStr := 'ss:AutoFitHeight="0"' + rowHeightStr else
        rowHeightStr := 'ss:AutoFitHeight="1"' + rowHeightStr;
      if row^.FormatIndex > 0 then
        styleStr := GetStyleStr(row^.FormatIndex);
    end else
      rowheightStr := 'ss:AutoFitHeight="1"';
    AppendToStream(AStream, ROW_INDENT + Format(
      '<Row %s%s>' + LF, [rowheightStr, styleStr]));
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
  AWorksheet: TsWorksheet);
var
  protectedStr: String;
begin
  FWorksheet := AWorksheet;

  if FWorksheet.IsProtected then
    protectedStr := ' ss:Protected="1"' else
    protectedStr := '';

  AppendToStream(AStream, Format(
    '  <Worksheet ss:Name="%s"%s>' + LF, [
    UTF8TextToXMLText(AWorksheet.Name),
    protectedStr
  ]) );
  WriteTable(AStream, AWorksheet);
  WriteWorksheetOptions(AStream, AWorksheet);
  AppendToStream(AStream,
    '  </Worksheet>' + LF
  );
end;

procedure TsSpreadExcelXMLWriter.WriteWorksheetOptions(AStream: TStream;
  AWorksheet: TsWorksheet);
var
  footerStr, headerStr: String;
  hideGridStr: String;
  hideHeadersStr: String;
  frozenStr: String;
  layoutStr: String;
  marginStr: String;
  selectedStr: String;
  protectStr: String;
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
    hideGridStr := INDENT3 + '<DoNotDisplayGridlines/>' + LF else
    hideGridStr := '';

  // Show/hide column/row headers
  if not (soShowHeaders in AWorksheet.Options) then
    hideHeadersStr := INDENT3 + '<DoNotDisplayHeadings/>' + LF else
    hideHeadersStr := '';

  if FWorkbook.ActiveWorksheet = AWorksheet then
    selectedStr := INDENT3 + '<Selected/>' + LF else
    selectedStr := '';

  // Frozen panes
  frozenStr := GetFrozenPanesStr(AWorksheet, INDENT3);

  // Protection
  protectStr := Format(INDENT3 + '<ProtectObjects>%s</ProtectObjects>' + LF +
                       INDENT3 + '<ProtectScenarios>%s</ProtectScenarios>' + LF, [
    AWorksheet.IsProtected and (spObjects in AWorksheet.Protection),
    AWorksheet.IsProtected {and [spScenarios in AWorksheet.Protection])}
  ]);

  // Put it all together...
  AppendToStream(AStream, INDENT2 +
    '<WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">' + LF + INDENT3 +
      '<PageSetup>' + LF +
        layoutStr +
        headerStr +
        footerStr +
        marginStr + INDENT3 +
      '</PageSetup>' + LF +
      selectedStr +
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
begin
  for i:=0 to FWorkbook.GetWorksheetCount-1 do
    WriteWorksheet(AStream, FWorkbook.GetWorksheetByIndex(i));
end;


initialization

  // Registers this reader / writer in fpSpreadsheet
  sfidExcelXML := RegisterSpreadFormat(sfExcelXML,
    nil, TsSpreadExcelXMLWriter,
    STR_FILEFORMAT_EXCEL_XML, 'ExcelXML', [STR_XML_EXCEL_EXTENSION]
  );

end.
