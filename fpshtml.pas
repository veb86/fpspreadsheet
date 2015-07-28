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
    FFormatSettings: TFormatSettings;
    function GetBackgroundAsStyle(AFill: TsFillPattern): String;
    function GetBorderAsStyle(ABorder: TsCellBorders; const ABorderStyles: TsCellBorderStyles): String;
    function GetFontAsStyle(AFontIndex: Integer): String;
    function GetHorAlignAsStyle(AHorAlign: TsHorAlignment): String;
    function GetTextRotation(ATextRot: TsTextRotation): String;
    function GetVertAlignAsStyle(AVertAlign: TsVertAlignment): String;
    function GetWordWrapAsStyle(AWordWrap: Boolean): String;
    procedure WriteBody(AStream: TStream);
    procedure WriteWorksheet(AStream: TStream; ASheet: TsWorksheet);

  protected
    function CellFormatAsString(ACell: PCell; ForThisTag: String): String;
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
    procedure WriteToStream(AStream: TStream); override;
    procedure WriteToStrings(AStrings: TStrings); override;
  end;

  TsHTMLParams = record
    SheetIndex: Integer;             // W: Index of the sheet to be written
    TrueText: String;                // RW: String for boolean TRUE
    FalseText: String;               // RW: String for boolean FALSE
  end;

var
  HTMLParams: TsHTMLParams = (
    SheetIndex: -1;       // -1 = active sheet, MaxInt = all sheets
    TrueText: 'TRUE';
    FalseText: 'FALSE';
  );

implementation

uses
  LazUTF8, fpsUtils;

constructor TsHTMLWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
end;

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
        Result := 'border-collapse:collapse; ';
        if soShowGridLines in FWorksheet.Options then
          Result := Result + 'border:1px solid lightgrey; '
      end else
      begin
        if (uffVertAlign in fmt^.UsedFormattingFields) then
          Result := Result + GetVertAlignAsStyle(fmt^.VertAlignment);
        if (uffBorder in fmt^.UsedFormattingFields) then
          Result := Result + GetBorderAsStyle(fmt^.Border, fmt^.BorderStyles)
        else begin
          Result := Result + 'border-collapse:collapse; ';
          if soShowGridLines in FWorksheet.Options then
            Result := Result + 'border:1px solid lightgrey; ';
        end;
        if (uffBackground in fmt^.UsedFormattingFields) then
          Result := Result + GetBackgroundAsStyle(fmt^.Background);
        if (uffFont in fmt^.UsedFormattingFields) then
          Result := Result + GetFontAsStyle(fmt^.FontIndex);
        if (uffTextRotation in fmt^.UsedFormattingFields) then
          Result := Result + GetTextRotation(fmt^.TextRotation);
      end;
    'div', 'p':
      begin
        if fmt = nil then
          exit;
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
      end;
  end;
  if Result <> '' then
    Result := ' style="' + Result +'"';
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
    'border-top', 'border-left', 'border-right', 'border-bottom', '', ''
  );
  LINESTYLE_NAMES: array[TsLineStyle] of string = (
    'thin solid',    // lsThin
    'medium solid',  // lsMedium
    'thin dashed',   // lsDashed
    'thin dotted',   // lsDotted
    'thick solid',   // lsThick,
    'thin double',   // lsDouble,
    '1px solid'      // lsHair
  );
var
  cb: TsCellBorder;
begin
  Result := 'border-collape:collapse';
  for cb in TsCellBorder do
  begin
    if BORDER_NAMES[cb] = '' then
      continue;
    Result := Result + BORDER_NAMES[cb] + ':' +
      LINESTYLE_NAMES[ABorderStyles[cb].LineStyle] + ' ' +
      ColorToHTMLColorStr(ABorderStyles[cb].Color) + ';';
  end;
end;

function TsHTMLWriter.GetFontAsStyle(AFontIndex: Integer): String;
var
  fs: TFormatSettings;
  font: TsFont;
begin
  fs := DefaultFormatSettings;
  fs.DecimalSeparator := '.';
  font := FWorkbook.GetFont(AFontIndex);
  Result := Format('font-family:''%s'';font-size:%.1fpt;color:%s;', [
    font.FontName, font.Size, ColorToHTMLColorStr(font.Color)], fs);
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

function TsHTMLWriter.GetHorAlignAsStyle(AHorAlign: TsHorAlignment): String;
begin
  case AHorAlign of
    haLeft   : Result := 'text-align:left;';
    haCenter : Result := 'text-align:center;';
    haRight  : Result := 'text-align:right;';
  end;
end;

function TsHTMLWriter.GetTextRotation(ATextRot: TsTextRotation): String;
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
    Result := 'white-space:nowrap'; //-moz-pre-wrap -o-pre-wrap pre-wrap;';
                          { Firefox      Opera        Chrome  }
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
var
  s: String;
  style: String;
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  if AValue then
    s := HTMLParams.TrueText
  else
    s := HTMLParams.FalseText;
  AppendToStream(AStream,
    '<div' + style + '>' + s + '</div>');
end;

{ Write date/time values in the same way they are displayed in the sheet }
procedure TsHTMLWriter.WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: TDateTime; ACell: PCell);
var
  style: String;
  s: String;
begin
  style := CellFormatAsString(ACell, 'div');
  s := FWorksheet.ReadAsUTF8Text(ACell);
  AppendToStream(AStream,
    '<div' + style + '>' + s + '</div>');
end;

procedure TsHTMLWriter.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
var
  style: String;
  s: String;
begin
  style := CellFormatAsString(ACell, 'div');
  s := FWOrksheet.ReadAsUTF8Text(ACell);
  AppendToStream(AStream,
    '<div' + style + '>' + s + '</div>');
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
  L: TStringList;
  style: String;
  i, n, len: Integer;
  txt, textp: String;
  rtParam: TsRichTextParam;
  fnt, cellfnt: TsFont;
  escapement: String;
begin
  Unused(ARow, ACol, AValue);

  txt := ACell^.UTF8StringValue;
  if txt = '' then
    exit;

  style := CellFormatAsString(ACell, 'div');

  // No hyperlink, normal text only
  if Length(ACell^.RichTextParams) = 0 then
  begin
    // Standard text formatting
    ValidXMLText(txt);
    AppendToStream(AStream,
      '<div' + style + '>' + txt + '</div>')
  end else
  begin
    // "Rich-text" formatting
    cellfnt := FWorksheet.ReadCellFont(ACell);
    len := UTF8Length(AValue);
    textp := '<div' + style + '>';
    rtParam := ACell^.RichTextParams[0];
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
      if (rtParam.EndIndex < len) and (i = High(ACell^.RichTextParams)) then
      begin
        txt := UTF8Copy(AValue, rtParam.EndIndex+1, MaxInt);
        ValidXMLText(txt);
        textp := textp + txt;
      end else
      if (i < High(ACell^.RichTextParams)) and (rtParam.EndIndex < ACell^.RichTextParams[i+1].StartIndex)
      then begin
        n := ACell^.RichTextParams[i+1].StartIndex - rtParam.EndIndex;
        txt := UTF8Copy(AValue, rtParam.EndIndex+1, n);
        ValidXMLText(txt);
        textp := textp + txt;
      end;
    end;
    textp := textp + '</div>';
    AppendToStream(AStream, textp);
  end;

{
  L := TStringList.Create;
  try
    L.Text := ACell^.UTF8StringValue;
    if L.Count = 1 then
      AppendToStream(AStream,
        '<div' + style + '>' + s + '</div>')
    else
    for i := 0 to L.Count-1 do
      AppendToStream(AStream, '<p><div'+ style + '>' + L[i] + '</div></p>');
  finally
    L.Free;
  end;
  }
end;

{ Writes a number cell to the stream. }
procedure TsHTMLWriter.WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: double; ACell: PCell);
var
  s: String;
  style: String;
begin
  Unused(AStream);
  Unused(ARow, ACol);

  style := CellFormatAsString(ACell, 'div');

  {
  if HTMLParams.NumberFormat <> '' then
    s := Format(HTMLParams.NumberFormat, [AValue], FFormatSettings)
  else
  }
  s := FWorksheet.ReadAsUTF8Text(ACell, FFormatSettings);
  AppendToStream(AStream,
    '<div' + style + '>' + s + '</div>');
end;

procedure TsHTMLWriter.WriteToStream(AStream: TStream);
begin
  FWorkbook.UpdateCaches;
  AppendToStream(AStream,
    '<!DOCTYPE html>' +
    '<html>' +
      '<head>'+
    //    '<title>Written by FPSpreadsheet</title>' +
        '<meta charset="utf-8">' +
      '</head>');
      WriteBody(AStream);
  AppendToStream(AStream,
    '</html>');
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
  txt: String;
  cell: PCell;
  style: String;
  fixedLayout: Boolean;
  col: PCol;
  w: Single;
  fs: TFormatSettings;
begin
  FWorksheet := ASheet;

  rFirst := FWorksheet.GetFirstRowIndex;
  cFirst := FWorksheet.GetFirstColIndex;
  rLast := FWorksheet.GetLastOccupiedRowIndex;
  cLast := FWorksheet.GetLastOccupiedColIndex;

  fs := DefaultFormatSettings;
  fs.DecimalSeparator := '.';

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
    style := style + 'border:1px solid lightgrey; ';

  if fixedLayout then
    style := style + 'table-layout:fixed; '
  else
    style := style + 'table-layout:auto; width:100%; ';

  AppendToStream(AStream,
    '<div>' +
      '<table style="' + style + '">');
  for r := rFirst to rLast do begin
    AppendToStream(AStream,
        '<tr>');
      for c := cFirst to cLast do begin
        cell := FWorksheet.FindCell(r, c);
        style := CellFormatAsString(cell, 'td');

        if (c = cFirst) then
        begin
          w := FWorksheet.DefaultColWidth;
          if fixedLayout then
          begin
            col := FWorksheet.GetCol(c);
            if col <> nil then
              w := col^.Width;
            style := Format(' width="%.1fpt"', [w*FWorkbook.GetDefaultFont.Size], fs) + style;
          end;
        end;

        if (cell = nil) or (cell^.ContentType = cctEmpty) then
          AppendToStream(AStream,
            '<td' + style + ' />')
        else
        begin
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

