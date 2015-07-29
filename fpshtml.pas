unit fpsHTML;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fasthtmlparser,
  fpstypes, fpspreadsheet, fpsReaderWriter;

type            (*
  TsHTMLReader = class(TsCustomSpreadReader)
  private
    FWorksheetName: String;
    FFormatSettings: TFormatSettings;
    function IsBool(AText: String; out AValue: Boolean): Boolean;
    function IsDateTime(AText: String; out ADateTime: TDateTime;
      out ANumFormat: TsNumberFormat): Boolean;
    function IsNumber(AText: String; out ANumber: Double; out ANumFormat: TsNumberFormat;
      out ADecimals: Integer; out ACurrencySymbol, AWarning: String): Boolean;
    function IsQuotedText(var AText: String): Boolean;
    procedure ReadCellValue(ARow, ACol: Cardinal; AText: String);
  protected
    procedure ReadBlank(AStream: TStream); override;
    procedure ReadFormula(AStream: TStream); override;
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadNumber(AStream: TStream); override;
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    procedure ReadFromFile(AFileName: String); override;
    procedure ReadFromStream(AStream: TStream); override;
    procedure ReadFromStrings(AStrings: TStrings); override;
  end;
              *)
  TsHTMLWriter = class(TsCustomSpreadWriter)
  private
    FPointSeparatorSettings: TFormatSettings;
//    function CellFormatAsString(ACell: PCell; ForThisTag: String): String;
    function CellFormatAsString(AFormat: PsCellFormat; ATagName: String): String;
    function GetBackgroundAsStyle(AFill: TsFillPattern): String;
    function GetBorderAsStyle(ABorder: TsCellBorders; const ABorderStyles: TsCellBorderStyles): String;
    function GetColWidthAsAttr(AColIndex: Integer): String;
    function GetDefaultHorAlignAsStyle(ACell: PCell): String;
    function GetFontAsStyle(AFontIndex: Integer): String;
    function GetGridBorderAsStyle: String;
    function GetHorAlignAsStyle(AHorAlign: TsHorAlignment): String;
    function GetMergedRangeAsStyle(AMergeBase: PCell): String;
    function GetRowHeightAsAttr(ARowIndex: Integer): String;
    function GetTextRotationAsStyle(ATextRot: TsTextRotation): String;
    function GetVertAlignAsStyle(AVertAlign: TsVertAlignment): String;
    function GetWordWrapAsStyle(AWordWrap: Boolean): String;
    function IsHyperlinkTarget(ACell: PCell; out ABookmark: String): Boolean;
    procedure WriteBody(AStream: TStream);
    procedure WriteStyles(AStream: TStream);
    procedure WriteWorksheet(AStream: TStream; ASheet: TsWorksheet);

  protected
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
    destructor Destroy; override;
    procedure WriteToStream(AStream: TStream); override;
    procedure WriteToStrings(AStrings: TStrings); override;
  end;

  TsHTMLParams = record
    SheetIndex: Integer;             // W: Index of the sheet to be written
    ShowRowColHeaders: Boolean;      // RW: Show row/column headers
    TrueText: String;                // RW: String for boolean TRUE
    FalseText: String;               // RW: String for boolean FALSE
  end;

var
  HTMLParams: TsHTMLParams = (
    SheetIndex: -1;                  // -1 = active sheet, MaxInt = all sheets
    ShowRowColHeaders: false;
    TrueText: 'TRUE';
    FalseText: 'FALSE';
  );

implementation

uses
  LazUTF8, URIParser, Math, StrUtils,
  fpsUtils;

constructor TsHTMLWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FPointSeparatorSettings := DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator := '.';

  // No design limiations in table size
  // http://stackoverflow.com/questions/4311283/max-columns-in-html-table
  FLimitations.MaxColCount := MaxInt;
  FLimitations.MaxRowCount := MaxInt;
end;

destructor TsHTMLWriter.Destroy;
begin
  inherited Destroy;
end;
               (*
function TsHTMLWriter.CellFormatAsString(ACell: PCell; ForThisTag: String): String;
var
  fmt: PsCellFormat;
begin
  Result := '';
  if ACell <> nil then
    fmt := FWorkbook.GetPointerToCellFormat(ACell^.FormatIndex)
  else
    fmt := nil;
  case ForThisTag of
    'td':
      if ACell = nil then
      begin
        Result := 'border-collapse:collapse;';
        if soShowGridLines in FWorksheet.Options then
          Result := Result + GetGridBorderAsStyle;
      end else
      begin
        if (uffBackground in fmt^.UsedFormattingFields) then
          Result := Result + GetBackgroundAsStyle(fmt^.Background);
        if (uffFont in fmt^.UsedFormattingFields) then
          Result := Result + GetFontAsStyle(fmt^.FontIndex);
        if (uffTextRotation in fmt^.UsedFormattingFields) then
          Result := Result + GetTextRotationAsStyle(fmt^.TextRotation);
        if (uffHorAlign in fmt^.UsedFormattingFields) and (fmt^.HorAlignment <> haDefault) then
          Result := Result + GetHorAlignAsStyle(fmt^.HorAlignment)
        else
          case ACell^.ContentType of
            cctNumber    : Result := Result + GetHorAlignAsStyle(haRight);
            cctDateTime  : Result := Result + GetHorAlignAsStyle(haLeft);
            cctBool      : Result := Result + GetHorAlignAsStyle(haCenter);
            else           Result := Result + GetHorAlignAsStyle(haLeft);
          end;
        if (uffVertAlign in fmt^.UsedFormattingFields) then
          Result := Result + GetVertAlignAsStyle(fmt^.VertAlignment);
        if (uffBorder in fmt^.UsedFormattingFields) then
          Result := Result + GetBorderAsStyle(fmt^.Border, fmt^.BorderStyles)
        else begin
          if soShowGridLines in FWorksheet.Options then
            Result := Result + GetGridBorderAsStyle;
        end;
        if (uffFont in fmt^.UsedFormattingFields) then
          Result := Result + GetFontAsStyle(fmt^.FontIndex);   {
        if (uffTextRotation in fmt^.UsedFormattingFields) then
          Result := Result + GetTextRotation(fmt^.TextRotation);}
        Result := Result + GetWordwrapAsStyle(uffWordwrap in fmt^.UsedFormattingFields);
      end;
    'div', 'p':
      begin
        if fmt = nil then
          exit;
        {
        if (uffHorAlign in fmt^.UsedFormattingFields) and (fmt^.HorAlignment <> haDefault) then
          Result := Result + GetHorAlignAsStyle(fmt^.HorAlignment)
        else
          case ACell^.ContentType of
            cctNumber    : Result := Result + GetHorAlignAsStyle(haRight);
            cctDateTime  : Result := Result + GetHorAlignAsStyle(haLeft);
            cctBool      : Result := Result + GetHorAlignAsStyle(haCenter);
            else           Result := Result + GetHorAlignAsStyle(haLeft);
          end;
        if (uffFont in fmt^.UsedFormattingFields) then
          Result := Result + GetFontAsStyle(fmt^.FontIndex);   {
        if (uffTextRotation in fmt^.UsedFormattingFields) then
          Result := Result + GetTextRotation(fmt^.TextRotation);}
        Result := Result + GetWordwrapAsStyle(uffWordwrap in fmt^.UsedFormattingFields);
        }
      end;
  end;
  if Result <> '' then
    Result := ' style="' + Result +'"';
end;
         *)
function TsHTMLWriter.CellFormatAsString(AFormat: PsCellFormat; ATagName: String): String;
begin
  Result := '';

  if (uffBackground in AFormat^.UsedFormattingFields) then
    Result := Result + GetBackgroundAsStyle(AFormat^.Background);

  if (uffFont in AFormat^.UsedFormattingFields) then
    Result := Result + GetFontAsStyle(AFormat^.FontIndex);

  if (uffTextRotation in AFormat^.UsedFormattingFields) then
    Result := Result + GetTextRotationAsStyle(AFormat^.TextRotation);

  if (uffHorAlign in AFormat^.UsedFormattingFields) and (AFormat^.HorAlignment <> haDefault) then
    Result := Result + GetHorAlignAsStyle(AFormat^.HorAlignment);

  if (uffVertAlign in AFormat^.UsedFormattingFields) then
    Result := Result + GetVertAlignAsStyle(AFormat^.VertAlignment);

  if (uffBorder in AFormat^.UsedFormattingFields) then
    Result := Result + GetBorderAsStyle(AFormat^.Border, AFormat^.BorderStyles);
  {
  else begin
    if soShowGridLines in FWorksheet.Options then
      Result := Result + GetGridBorderAsStyle;
  end;
  }

  Result := Result + GetWordwrapAsStyle(uffWordwrap in AFormat^.UsedFormattingFields);
end;

function TsHTMLWriter.GetBackgroundAsStyle(AFill: TsFillPattern): String;
begin
  Result := '';
  if AFill.Style = fsSolidFill then
    Result := 'background-color:' + ColorToHTMLColorStr(AFill.FgColor) + ';';
  // other fills not supported
end;

function TsHTMLWriter.GetBorderAsStyle(ABorder: TsCellBorders;
  const ABorderStyles: TsCellBorderStyles): String;
const
  BORDER_NAMES: array[TsCellBorder] of string = (
    'border-top',    // cbNorth
    'border-left',   // cbWest
    'border-right',  // cbEast
    'border-bottom', // cbSouth
    '',              // cbDiagUp
    ''               // cbDiagDown
  );
  LINESTYLE_NAMES: array[TsLineStyle] of string = (
    'thin solid',    // lsThin
    'medium solid',  // lsMedium
    'thin dashed',   // lsDashed
    'thin dotted',   // lsDotted
    'thick solid',   // lsThick,
    'double',        // lsDouble,
    '1px solid'      // lsHair
  );
var
  cb: TsCellBorder;
  allEqual: Boolean;
  bs: TsCellBorderStyle;
begin
  Result := 'border-collape:collapse;';
  if ABorder = [cbNorth, cbEast, cbWest, cbSouth] then
  begin
    allEqual := true;
    bs := ABorderStyles[cbNorth];
    for cb in TsCellBorder do
    begin
      if bs.LineStyle <> ABorderStyles[cb].LineStyle then
      begin
        allEqual := false;
        break;
      end;
      if bs.Color <> ABorderStyles[cb].Color then
      begin
        allEqual := false;
        break;
      end;
    end;
    if allEqual then
    begin
      Result := 'border:' +
        LINESTYLE_NAMES[bs.LineStyle] + ' ' +
        ColorToHTMLColorStr(bs.Color) + ';';
      exit;
    end;
  end;

  for cb in TsCellBorder do
  begin
    if BORDER_NAMES[cb] = '' then
      continue;
    if cb in ABorder then
      Result := Result + BORDER_NAMES[cb] + ':' +
        LINESTYLE_NAMES[ABorderStyles[cb].LineStyle] + ' ' +
        ColorToHTMLColorStr(ABorderStyles[cb].Color) + ';';
  end;
end;

function TsHTMLWriter.GetColWidthAsAttr(AColIndex: Integer): String;
var
  col: PCol;
  w: Single;
  rLast: Cardinal;
begin
  if AColIndex < 0 then  // Row header column
  begin
    rLast := FWorksheet.GetLastRowIndex;
    w := Length(IntToStr(rLast)) + 2;
  end else
  begin
    w := FWorksheet.DefaultColWidth;
    col := FWorksheet.FindCol(AColIndex);
    if col <> nil then
      w := col^.Width;
  end;
  w := w * FWorkbook.GetDefaultFont.Size;
  Result:= Format(' width="%.1fpt"', [w], FPointSeparatorSettings);
end;

function TsHTMLWriter.GetDefaultHorAlignAsStyle(ACell: PCell): String;
begin
  Result := '';
  if ACell = nil then
    exit;
  case ACell^.ContentType of
    cctNumber  : Result := GetHorAlignAsStyle(haRight);
    cctDateTime: Result := GetHorAlignAsStyle(haRight);
    cctBool    : Result := GetHorAlignAsStyle(haCenter);
  end;
end;

function TsHTMLWriter.GetFontAsStyle(AFontIndex: Integer): String;
var
  font: TsFont;
begin
  font := FWorkbook.GetFont(AFontIndex);
  Result := Format('font-family:''%s'';font-size:%.1fpt;color:%s;', [
    font.FontName, font.Size, ColorToHTMLColorStr(font.Color)], FPointSeparatorSettings);
  if fssBold in font.Style then
    Result := Result + 'font-weight:700;';
  if fssItalic in font.Style then
    Result := Result + 'font-style:italic;';
  if [fssUnderline, fssStrikeout] * font.Style = [fssUnderline, fssStrikeout] then
    Result := Result + 'text-decoration:underline,line-through;'
  else
  if [fssUnderline, fssStrikeout] * font.Style = [fssUnderline] then
    Result := Result + 'text-decoration:underline;'
  else
  if [fssUnderline, fssStrikeout] * font.Style = [fssStrikeout] then
    Result := Result + 'text-decoration:line-through;';
end;

function TsHTMLWriter.GetGridBorderAsStyle: String;
begin
  if (soShowGridLines in FWorksheet.Options) then
    Result := 'border:1px solid lightgrey;'
  else
    Result := '';
end;

function TsHTMLWriter.GetHorAlignAsStyle(AHorAlign: TsHorAlignment): String;
begin
  case AHorAlign of
    haLeft   : Result := 'text-align:left;';
    haCenter : Result := 'text-align:center;';
    haRight  : Result := 'text-align:right;';
  end;
end;

function TsHTMLWriter.GetMergedRangeAsStyle(AMergeBase: PCell): String;
var
  r1, r2, c1, c2: Cardinal;
begin
  Result := '';
  FWorksheet.FindMergedRange(AMergeBase, r1, c1, r2, c2);
  if c1 <> c2 then
    Result := Result + ' colspan="' + IntToStr(c2-c1+1) + '"';
  if r1 <> r2 then
    Result := Result + ' rowspan="' + IntToStr(r2-r1+1) + '"';
end;

function TsHTMLWriter.GetRowHeightAsAttr(ARowIndex: Integer): String;
var
  h: Single;
  row: PRow;
begin
  h := FWorksheet.DefaultRowHeight;
  row := FWorksheet.FindRow(ARowIndex);
  if row <> nil then
    h := row^.Height;
  h := (h + ROW_HEIGHT_CORRECTION) * FWorkbook.GetDefaultFont.Size;
  Result := Format(' height="%.1fpt"', [h], FPointSeparatorSettings);
end;


function TsHTMLWriter.GetTextRotationAsStyle(ATextRot: TsTextRotation): String;
begin
  Result := '';
  case ATextRot of
    trHorizontal: ;
    rt90DegreeClockwiseRotation:
      Result := 'writing-mode:vertical-rl;transform:rotate(90deg);'; //-moz-transform: rotate(90deg);';
//      Result := 'writing-mode:vertical-rl;text-orientation:sideways-right;-moz-transform: rotate(-90deg);';
    rt90DegreeCounterClockwiseRotation:
      Result := 'writing-mode:vertical-rt;transform:rotate(-90deg);'; //-moz-transform: rotate(-90deg);';
//    Result := 'writing-mode:vertical-rt;text-orientation:sideways-left;-moz-transform: rotate(-90deg);';
    rtStacked:
      Result := 'writing-mode:vertical-rt;text-orientation:upright;';
  end;
end;

function TsHTMLWriter.GetVertAlignAsStyle(AVertAlign: TsVertAlignment): String;
begin
  case AVertAlign of
    vaTop    : Result := 'vertical-align:top;';
    vaCenter : Result := 'vertical-align:middle;';
    vaBottom : Result := 'vertical-align:bottom;';
  end;
end;

function TsHTMLWriter.GetWordwrapAsStyle(AWordwrap: Boolean): String;
begin
  if AWordwrap then
    Result := 'word-wrap:break-word;'
  else
    Result := 'white-space:nowrap;';
end;

function TsHTMLWriter.IsHyperlinkTarget(ACell: PCell; out ABookmark: String): Boolean;
var
  sheet: TsWorksheet;
  hyperlink: PsHyperlink;
  target, sh: String;
  i, r, c: Cardinal;
begin
  Result := false;
  if ACell = nil then
    exit;

  for i:=0 to FWorkbook.GetWorksheetCount-1 do
  begin
    sheet := FWorkbook.GetWorksheetByIndex(i);
    for hyperlink in sheet.Hyperlinks do
    begin
      SplitHyperlink(hyperlink^.Target, target, ABookmark);
      if (target <> '') or (ABookmark = '') then
        continue;
      if ParseSheetCellString(ABookmark, sh, r, c) then
        if (sh = TsWorksheet(ACell^.Worksheet).Name) and
           (r = ACell^.Row) and (c = ACell^.Col)
        then
          exit(true);
      if (sheet = FWorksheet) and  ParseCellString(ABookmark, r, c) then
        if (r = ACell^.Row) and (c = ACell^.Col) then
          exit(true);
    end;
  end;
end;

procedure TsHTMLWriter.WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  // nothing to do
end;

procedure TsHTMLWriter.WriteBody(AStream: TStream);
var
  i: Integer;
begin
  AppendToStream(AStream,
    '<body>');
  if HTMLParams.SheetIndex < 0 then      // active sheet
  begin
    if FWorkbook.ActiveWorksheet = nil then
      FWorkbook.SelectWorksheet(FWorkbook.GetWorksheetByIndex(0));
    WriteWorksheet(AStream, FWorkbook.ActiveWorksheet)
  end else
  if HTMLParams.SheetIndex = MaxInt then  // all sheets
    for i:=0 to FWorkbook.GetWorksheetCount-1 do
      WriteWorksheet(AStream, FWorkbook.GetWorksheetByIndex(i))
  else                                    // specific sheet
    WriteWorksheet(AStream, FWorkbook.GetWorksheetbyIndex(HTMLParams.SheetIndex));
  AppendToStream(AStream,
    '</body>');
end;

{ Write boolean cell to stream formatted as string }
procedure TsHTMLWriter.WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: Boolean; ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  AppendToStream(AStream,
    '<div>' + IfThen(AValue, HTMLParams.TrueText, HTMLParams.FalseText) + '</div>');
end;

{ Write date/time values in the same way they are displayed in the sheet }
procedure TsHTMLWriter.WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: TDateTime; ACell: PCell);
var
  s: String;
begin
  s := FWorksheet.ReadAsUTF8Text(ACell);
  AppendToStream(AStream,
    '<div>' + s + '</div>');
end;

procedure TsHTMLWriter.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
var
  s: String;
begin
  s := FWOrksheet.ReadAsUTF8Text(ACell);
  AppendToStream(AStream,
    '<div>' + s + '</div>');
end;

{ HTML does not support formulas, but we can write the formula results to
  to stream. }
procedure TsHTMLWriter.WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  if ACell = nil then
    exit;
  case ACell^.ContentType of
    cctBool      : WriteBool(AStream, ARow, ACol, ACell^.BoolValue, ACell);
    cctEmpty     : ;
    cctDateTime  : WriteDateTime(AStream, ARow, ACol, ACell^.DateTimeValue, ACell);
    cctNumber    : WriteNumber(AStream, ARow, ACol, ACell^.NumberValue, ACell);
    cctUTF8String: WriteLabel(AStream, ARow, ACol, ACell^.UTF8StringValue, ACell);
    cctError     : ;
  end;
end;

{ Writes a LABEL cell to the stream. }
procedure TsHTMLWriter.WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: string; ACell: PCell);
const
  ESCAPEMENT_TAG: Array[TsFontPosition] of String = ('', 'sup', 'sub');
var
  style: String;
  i, n, len: Integer;
  txt, textp, target, bookmark: String;
  rtParam: TsRichTextParam;
  fnt, cellfnt: TsFont;
  escapement: String;
  hyperlink: PsHyperlink;
  isTargetCell: Boolean;
  u: TUri;
begin
  Unused(ARow, ACol, AValue);

  txt := ACell^.UTF8StringValue;
  if txt = '' then
    exit;

  style := ''; //CellFormatAsString(ACell, 'div');
  cellfnt := FWorksheet.ReadCellFont(ACell);

  // Hyperlink
  target := '';
  if FWorksheet.HasHyperlink(ACell) then
  begin
    hyperlink := FWorksheet.FindHyperlink(ACell);
    SplitHyperlink(hyperlink^.Target, target, bookmark);

    n := Length(hyperlink^.Target);
    i := Length(target);
    len := Length(bookmark);

    if (target <> '') and (pos('file:', target) = 0) then
    begin
      u := ParseURI(target);
      if u.Protocol = '' then
        target := '../' + target;
    end;

    // ods absolutely wants "/" path delimiters in the file uri!
    FixHyperlinkPathdelims(target);

    if (bookmark <> '') then
      target := target + '#' + bookmark;
  end;

  // Activate hyperlink target if it is within the same file
  isTargetCell := IsHyperlinkTarget(ACell, bookmark);
  if isTargetCell then bookmark := ' id="' + bookmark + '"' else bookmark := '';

  // No hyperlink, normal text only
  if Length(ACell^.RichTextParams) = 0 then
  begin
    // Standard text formatting
    ValidXMLText(txt);
    if target <> '' then
      txt := Format('<a href="%s">%s</a>', [target, txt]);
    if cellFnt.Position <> fpNormal then
      txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[cellFnt.Position], txt]);
    AppendToStream(AStream,
      '<div' + bookmark + style + '>' + txt + '</div>')
  end else
  begin
    // "Rich-text" formatted string
    len := UTF8Length(AValue);
    textp := '<div' + bookmark + style + '>';
    if target <> '' then
      textp := textp + '<a href="' + target + '">';
    rtParam := ACell^.RichTextParams[0];
    // Part before first formatted section (has cell fnt)
    if rtParam.StartIndex > 0 then
    begin
      txt := UTF8Copy(AValue, 1, rtParam.StartIndex);
      ValidXMLText(txt);
      if cellfnt.Position <> fpNormal then
        txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[cellFnt.Position], txt]);
      textp := textp + txt;
    end;
    for i := 0 to High(ACell^.RichTextParams) do
    begin
      // formatted section
      rtParam := ACell^.RichTextParams[i];
      fnt := FWorkbook.GetFont(rtParam.FontIndex);
      style := GetFontAsStyle(rtParam.FontIndex);
      if style <> '' then
        style := ' style="' + style +'"';
      n := rtParam.EndIndex - rtParam.StartIndex;
      txt := UTF8Copy(AValue, rtParam.StartIndex+1, n);
      ValidXMLText(txt);
      if fnt.Position <> fpNormal then
        txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[fnt.Position], txt]);
      textp := textp + '<span' + style +'>' + txt + '</span>';
      // unformatted section before end
      if (rtParam.EndIndex < len) and (i = High(ACell^.RichTextParams)) then
      begin
        txt := UTF8Copy(AValue, rtParam.EndIndex+1, MaxInt);
        ValidXMLText(txt);
        if cellFnt.Position <> fpNormal then
          txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[cellFnt.Position], txt]);
        textp := textp + txt;
      end else
      // unformatted section between two formatted sections
      if (i < High(ACell^.RichTextParams)) and (rtParam.EndIndex < ACell^.RichTextParams[i+1].StartIndex)
      then begin
        n := ACell^.RichTextParams[i+1].StartIndex - rtParam.EndIndex;
        txt := UTF8Copy(AValue, rtParam.EndIndex+1, n);
        ValidXMLText(txt);
        if cellFnt.Position <> fpNormal then
          txt := Format('<%0:s>%1:s</%0:s>', [ESCAPEMENT_TAG[cellFnt.Position], txt]);
        textp := textp + txt;
      end;
    end;
    if target <> '' then
      textp := textp + '</a></div>' else
      textp := textp + '</div>';
    AppendToStream(AStream, textp);
  end;
end;

{ Writes a number cell to the stream. }
procedure TsHTMLWriter.WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: double; ACell: PCell);
var
  s: String;
begin
  Unused(ARow, ACol);
  s := FWorksheet.ReadAsUTF8Text(ACell, FWorkbook.FormatSettings);
  AppendToStream(AStream,
    '<div>' + s + '</div>');
end;

procedure TsHTMLWriter.WriteToStream(AStream: TStream);
begin
  FWorkbook.UpdateCaches;
  AppendToStream(AStream,
    '<!DOCTYPE html>' +
    '<html>' +
      '<head>'+
        '<meta charset="utf-8">');
  WriteStyles(AStream);
  AppendToStream(AStream,
      '</head>');
      WriteBody(AStream);
  AppendToStream(AStream,
    '</html>');
end;

procedure TsHTMLWriter.WriteStyles(AStream: TStream);
var
  i: Integer;
  fmt: PsCellFormat;
  fmtStr: String;
begin
  AppendToStream(AStream,
    '<style>');
  for i:=0 to FWorkbook.GetNumCellFormats-1 do begin
    fmt := FWorkbook.GetPointerToCellFormat(i);
    fmtStr := CellFormatAsString(fmt, 'td');
    if fmtStr <> '' then
      fmtStr := Format('td.style%d {%s}', [i+1, fmtStr]);
    AppendToStream(AStream, fmtStr);
  end;
  AppendToStream(AStream,
    '</style>');
end;

procedure TsHTMLWriter.WriteToStrings(AStrings: TStrings);
var
  Stream: TStream;
begin
  Stream := TStringStream.Create('');
  try
    WriteToStream(Stream);
    Stream.Position := 0;
    AStrings.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;

procedure TsHTMLWriter.WriteWorksheet(AStream: TStream; ASheet: TsWorksheet);
var
  r, rFirst, rLast: Cardinal;
  c, cFirst, cLast: Cardinal;
  cell: PCell;
  style, s: String;
  fixedLayout: Boolean;
  col: PCol;
  fmt: PsCellFormat;
begin
  FWorksheet := ASheet;

  rFirst := FWorksheet.GetFirstRowIndex;
  cFirst := FWorksheet.GetFirstColIndex;
  rLast := FWorksheet.GetLastOccupiedRowIndex;
  cLast := FWorksheet.GetLastOccupiedColIndex;

  fixedLayout := false;
  for c:=cFirst to cLast do
  begin
    col := FWorksheet.GetCol(c);
    if col <> nil then
    begin
      fixedLayout := true;
      break;
    end;
  end;

  style := GetFontAsStyle(DEFAULT_FONTINDEX);

  style := style + 'border-collapse:collapse; ';
  if soShowGridLines in FWorksheet.Options then
    style := style + GetGridBorderAsStyle;

  if fixedLayout then
    style := style + 'table-layout:fixed; '
  else
    style := style + 'table-layout:auto; width:100%; ';

  AppendToStream(AStream,
    '<div>' +
      '<table style="' + style + '">');

  if HTMLParams.ShowRowColHeaders then
  begin
    // width of row-header column
    style := '';
    if soShowGridLines in FWorksheet.Options then
      style := style + GetGridBorderAsStyle;
    if style <> '' then
      style := ' style="' + style + '"';
    style := style + GetColWidthAsAttr(-1);
    AppendToStream(AStream,
        '<th' + style + '/>');
    // Column headers
    for c := cFirst to cLast do
    begin
      style := '';
      if soShowGridLines in FWorksheet.Options then
        style := style + GetGridBorderAsStyle;
      if style <> '' then
        style := ' style="' + style + '"';
      if fixedLayout then
        style := style + GetColWidthAsAttr(c);
      AppendToStream(AStream,
        '<th' + style + '>' + GetColString(c) + '</th>');
    end;
  end;

  for r := rFirst to rLast do begin
    AppendToStream(AStream,
        '<tr>');

    // Row headers
    if HTMLParams.ShowRowColHeaders then begin
      style := '';
      if soShowGridLines in FWorksheet.Options then
        style := style + GetGridBorderAsStyle;
      if style <> '' then
        style := ' style="' + style + '"';
      style := style + GetRowHeightAsAttr(r);
      AppendToStream(AStream,
          '<th' + style + '>' + IntToStr(r+1) + '</th>');
    end;

    for c := cFirst to cLast do begin
      // Pointer to current cell in loop
      cell := FWorksheet.FindCell(r, c);

      // Cell formatting via predefined styles ("class")
      style := '';
      fmt := nil;
      if cell <> nil then
      begin
        style := Format(' class="style%d"', [cell^.FormatIndex+1]);
        fmt := FWorkbook.GetPointerToCellFormat(cell^.FormatIndex);
      end;

      // Overriding differences between html and fps formatting
      s := '';
      if (fmt = nil) then
        s := s + GetGridBorderAsStyle
      else begin
        if ((not (uffBorder in fmt^.UsedFormattingFields)) or (fmt^.Border = [])) then
          s := s + GetGridBorderAsStyle;
        if ((not (uffHorAlign in fmt^.UsedFormattingFields)) or (fmt^.HorAlignment = haDefault)) then
          s := s + GetDefaultHorAlignAsStyle(cell);
        if ((not (uffVertAlign in fmt^.UsedFormattingFields)) or (fmt^.VertAlignment = vaDefault)) then
          s := s + GetVertAlignAsStyle(vaBottom);
      end;
      if s <> '' then
        style := style + ' style="' + s + '"';

      if not HTMLParams.ShowRowColHeaders then
      begin
        // Column width
        if fixedLayout then
          style := GetColWidthAsAttr(c) + style;

        // Row heights (should be in "tr", but does not work there)
        style := GetRowHeightAsAttr(r) + style;
      end;

      // Merged cells
      if FWorksheet.IsMerged(cell) then
      begin
        if FWorksheet.IsMergeBase(cell) then
          style := style + GetMergedRangeAsStyle(cell)
        else
          Continue;
      end;

      if (cell = nil) or (cell^.ContentType = cctEmpty) then
        // Empty cell
        AppendToStream(AStream,
          '<td' + style + ' />')
      else
      begin
        // Cell with data
        AppendToStream(AStream,
          '<td' + style + '>');
        WriteCellToStream(AStream, cell);
        AppendToStream(AStream,
          '</td>');
      end;
    end;
    AppendToStream(AStream,
        '</tr>');
  end;
  AppendToStream(AStream,
      '</table>' +
    '</div>');
end;

initialization
  RegisterSpreadFormat(nil, TsHTMLWriter, sfHTML);

end.

