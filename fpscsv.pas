unit fpscsv;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  fpspreadsheet;

type
  TsCSVReader = class(TsCustomSpreadReader)
  private
    FWorksheetName: String;
    function IsDateTime(AText: String; out ADateTime: TDateTime): Boolean;
    function IsNumber(AText: String; out ANumber: Double): Boolean;
    function IsQuotedText(var AText: String): Boolean;
    procedure ReadCellValue(ARow, ACol: Cardinal; AText: String);
  protected
    procedure ReadBlank(AStream: TStream); override;
    procedure ReadFormula(AStream: TStream); override;
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadNumber(AStream: TStream); override;
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    procedure ReadFromFile(AFileName: String; AData: TsWorkbook); override;
    procedure ReadFromStream(AStream: TStream; AData: TsWorkbook); override;
    procedure ReadFromStrings(AStrings: TStrings; AData: TsWorkbook); override;
  end;

  TsCSVWriter = class(TsCustomSpreadWriter)
  private
    FLineEnding: String;
  protected
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TDateTime; ACell: PCell); override;
    procedure WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;
    procedure WriteSheet(AStream: TStream; AWorksheet: TsWorksheet);

  public
    constructor Create(AWorkbook: TsWorkbook); override;
    procedure WriteToStream(AStream: TStream); override;
    procedure WriteToStrings(AStrings: TStrings); override;
  end;

  TsCSVLineEnding = (leSystem, leCRLF, leCR, leLF);

  TsCSVParams = record
    SheetIndex: Integer;
    LineEnding: TsCSVLineEnding;
    Delimiter: Char;
    QuoteChar: Char;
    NumberFormat: String;
    FormatSettings: TFormatSettings;
  end;

var
  CSVParams: TsCSVParams = (
    SheetIndex: 0;              // Store sheet #0 by default
    LineEnding: leSystem;       // Write system lineending, read any
    Delimiter: ';';             // Column delimiter
    QuoteChar: '"';             // for quoted strings
    NumberFormat: '';           // if empty write numbers like in sheet, otherwise use this format
  );


implementation

uses
  StrUtils, DateUtils, fpsutils;

{ Initializes the FormatSettings of the CSVParams to default values which
  can be replaced by the FormatSettings of the workbook's FormatSettings }
procedure InitCSVFormatSettings;
var
  i: Integer;
begin
  with CSVParams.FormatSettings do
  begin
    CurrencyFormat := Byte(-1);
    NegCurrFormat := Byte(-1);
    ThousandSeparator := #0;
    DecimalSeparator := #0;
    CurrencyDecimals := Byte(-1);
    DateSeparator := #0;
    TimeSeparator := #0;
    ListSeparator := #0;
    CurrencyString := '';
    ShortDateFormat := '';
    LongDateFormat := '';
    TimeAMString := '';
    TimePMString := '';
    ShortTimeFormat := '';
    LongTimeFormat := '';
    for i:=1 to 12 do
    begin
      ShortMonthNames[i] := '';
      LongMonthNames[i] := '';
    end;
    for i:=1 to 7 do
    begin
      ShortDayNames[i] := '';
      LongDayNames[i] := '';
    end;
    TwoDigitYearCenturyWindow := Word(-1);
  end;
end;

procedure ReplaceFormatSettings(var AFormatSettings: TFormatSettings;
  const ADefaultFormats: TFormatSettings);
var
  i: Integer;
begin
  if AFormatSettings.CurrencyFormat = Byte(-1) then
    AFormatSettings.CurrencyFormat := ADefaultFormats.CurrencyFormat;
  if AFormatSettings.NegCurrFormat = Byte(-1) then
    AFormatSettings.NegCurrFormat := ADefaultFormats.NegCurrFormat;
  if AFormatSettings.ThousandSeparator = #0 then
    AFormatSettings.ThousandSeparator := ADefaultFormats.ThousandSeparator;
  if AFormatSettings.DecimalSeparator = #0 then
    AFormatSettings.DecimalSeparator := ADefaultFormats.DecimalSeparator;
  if AFormatSettings.CurrencyDecimals = Byte(-1) then
    AFormatSettings.CurrencyDecimals := ADefaultFormats.CurrencyDecimals;
  if AFormatSettings.DateSeparator = #0 then
    AFormatSettings.DateSeparator := ADefaultFormats.DateSeparator;
  if AFormatSettings.TimeSeparator = #0 then
    AFormatSettings.TimeSeparator := ADefaultFormats.TimeSeparator;
  if AFormatSettings.ListSeparator = #0 then
    AFormatSettings.ListSeparator := ADefaultFormats.ListSeparator;
  if AFormatSettings.CurrencyString = '' then
    AFormatSettings.CurrencyString := ADefaultFormats.CurrencyString;
  if AFormatSettings.ShortDateFormat = '' then
    AFormatSettings.ShortDateFormat := ADefaultFormats.ShortDateFormat;
  if AFormatSettings.LongDateFormat = '' then
    AFormatSettings.LongDateFormat := ADefaultFormats.LongDateFormat;
  if AFormatSettings.ShortTimeFormat = '' then
    AFormatSettings.ShortTimeFormat := ADefaultFormats.ShortTimeFormat;
  if AFormatSettings.LongTimeFormat = '' then
    AFormatSettings.LongTimeFormat := ADefaultFormats.LongTimeFormat;
  for i:=1 to 12 do
  begin
    if AFormatSettings.ShortMonthNames[i] = '' then
      AFormatSettings.ShortMonthNames[i] := ADefaultFormats.ShortMonthNames[i];
    if AFormatSettings.LongMonthNames[i] = '' then
      AFormatSettings.LongMonthNames[i] := ADefaultFormats.LongMonthNames[i];
  end;
  for i:=1 to 7 do
  begin
    if AFormatSettings.ShortDayNames[i] = '' then
      AFormatSettings.ShortDayNames[i] := ADefaultFormats.ShortDayNames[i];
    if AFormatSettings.LongDayNames[i] = '' then
      AFormatSettings.LongDayNames[i] := ADefaultFormats.LongDayNames[i];
  end;
  if AFormatSettings.TwoDigitYearCenturyWindow = Word(-1) then
    AFormatSettings.TwoDigitYearCenturyWindow := ADefaultFormats.TwoDigitYearCenturyWindow;
end;


{ -----------------------------------------------------------------------------}
{                              TsCSVReader                                     }
{------------------------------------------------------------------------------}

constructor TsCSVReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  ReplaceFormatSettings(CSVParams.FormatSettings, AWorkbook.FormatSettings);
  FWorksheetName := 'Sheet1';  // will be replaced by filename
end;

function TsCSVReader.IsDateTime(AText: String; out ADateTime: TDateTime): Boolean;
begin
  Result := TryStrToDateTime(AText, ADateTime, CSVParams.FormatSettings);
end;

function TsCSVReader.IsNumber(AText: String; out ANumber: Double): Boolean;
begin
  Result := TryStrToFloat(AText, ANumber, CSVParams.FormatSettings);
end;

function TsCSVReader.IsQuotedText(var AText: String): Boolean;
begin
  if (Length(AText) > 1) and (CSVParams.QuoteChar <> #0) and
   (AText[1] = CSVParams.QuoteChar) and
   (AText[Length(AText)] = CSVParams.QuoteChar) then
  begin
    Delete(AText, 1, 1);
    Delete(AText, Length(AText), 1);
    Result := true;
  end else
    Result := false;
end;

procedure TsCSVReader.ReadBlank(AStream: TStream);
begin
  Unused(AStream);
end;

procedure TsCSVReader.ReadCellValue(ARow, ACol: Cardinal; AText: String);
var
  dbl: Double;
  dt: TDateTime;
begin
  // Empty strings are blank cells -- nothing to do
  if AText = '' then
    exit;

  // Quoted text is a TEXT cell
  if IsQuotedText(AText) then
  begin
    FWorksheet.WriteUTF8Text(ARow, ACol, AText);
    exit;
  end;

  // Check for a NUMBER cell
  if IsNumber(AText, dbl) then
  begin
    FWorksheet.WriteNumber(ARow, ACol, dbl);
    exit;
  end;

  // Check for a DATE/TIME cell
  if IsDateTime(AText, dt) then
  begin
    FWorksheet.WriteDateTime(ARow, ACol, dt);
    exit;
  end;

  // What is left is handled as a TEXT cell
  FWorksheet.WriteUTF8Text(ARow, ACol, AText);
end;

procedure TsCSVReader.ReadFormula(AStream: TStream);
begin
  Unused(AStream);
end;

procedure TsCSVReader.ReadFromFile(AFileName: String; AData: TsWorkbook);
begin
  FWorksheetName := ChangeFileExt(ExtractFileName(AFileName), '');
  inherited;
end;

procedure TsCSVReader.ReadFromStream(AStream: TStream; AData: TsWorkbook);
var
  n: Int64;
  ch: Char;
  nextch: Char;
  cellValue: String;
  r, c: Cardinal;
begin
  FWorkbook := AData;
  FWorksheet := AData.AddWorksheet(FWorksheetName);
  n := AStream.Size;
  cellValue := '';
  r := 0;
  c := 0;
  while AStream.Position < n do begin
    ch := char(AStream.ReadByte);
    if ch = CSVParams.Delimiter then begin
      // End of column reached
      ReadCellValue(r, c, cellValue);
      inc(c);
      cellValue := '';
    end else
    if (ch = #13) or (ch = #10) then begin
      // End of row reached
      ReadCellValue(r, c, cellValue);
      inc(r);
      c := 0;
      cellValue := '';

      // look for CR+LF: if true, skip next byte
      if AStream.Position+1 < n then begin
        nextch := char(AStream.ReadByte);
        if ((ch = #13) and (nextch <> #10)) then
          AStream.Position := AStream.Position - 1;  // re-read nextchar in next loop
      end;
    end else
      cellValue := cellValue + ch;
  end;
end;

procedure TsCSVReader.ReadFromStrings(AStrings: TStrings; AData: TsWorkbook);
var
  stream: TStringStream;
begin
  stream := TStringStream.Create(AStrings.Text);
  try
    ReadFromStream(stream, AData);
  finally
    stream.Free;
  end;
end;

procedure TsCSVReader.ReadLabel(AStream: TStream);
begin
  Unused(AStream);
end;

procedure TsCSVReader.ReadNumber(AStream: TStream);
begin
  Unused(AStream);
end;


{ -----------------------------------------------------------------------------}
{                              TsCSVWriter                                     }
{------------------------------------------------------------------------------}
constructor TsCSVWriter.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  ReplaceFormatSettings(CSVParams.FormatSettings, FWorkbook.FormatSettings);
  case CSVParams.LineEnding of
    leSystem : FLineEnding := LineEnding;
    leCRLF   : FLineEnding := #13#10;
    leCR     : FLineEnding := #13;
    leLF     : FLineEnding := #10;
  end;
end;

procedure TsCSVWriter.WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  // nothing to do
end;

{ Write date/time values in the same way they are displayed in the sheet }
procedure TsCSVWriter.WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: TDateTime; ACell: PCell);
begin
  Unused(ARow, ACol);
  AppendToStream(AStream, FWorksheet.ReadAsUTF8Text(ACell));
end;

procedure TsCSVWriter.WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  // no formulas in CSV
  Unused(AStream);
  Unused(ARow, ACol, AStream);
end;

procedure TsCSVWriter.WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: string; ACell: PCell);
var
  s: String;
begin
  Unused(ARow, ACol);
  if ACell = nil then
    exit;
  s := ACell^.UTF8StringValue;
  if CSVParams.QuoteChar <> #0 then
    s := CSVParams.QuoteChar + s + CSVParams.QuoteChar;
  AppendToStream(AStream, s);
end;

procedure TsCSVWriter.WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: double; ACell: PCell);
var
  s: String;
  mask: String;
begin
  Unused(ARow, ACol);
  if ACell = nil then
    exit;
  if CSVParams.NumberFormat <> '' then
    s := Format(CSVParams.NumberFormat, [AValue], CSVParams.FormatSettings)
  else
    s := FWorksheet.ReadAsUTF8Text(ACell);
  AppendToStream(AStream, s);
end;

procedure TsCSVWriter.WriteSheet(AStream: TStream; AWorksheet: TsWorksheet);
var
  r, c: Cardinal;
  lastRow, lastCol: Cardinal;
  cell: PCell;
begin
  FWorksheet := AWorksheet;
  lastRow := FWorksheet.GetLastOccupiedRowIndex;
  lastCol := FWorksheet.GetLastOccupiedColIndex;
  for r := 0 to lastRow do
    for c := 0 to lastCol do begin
      cell := FWorksheet.FindCell(r, c);
      if cell <> nil then
        WriteCellCallback(cell, AStream);
      if c = lastCol then
        AppendToStream(AStream, FLineEnding)
      else
        AppendToStream(AStream, CSVParams.Delimiter);
    end;
end;

procedure TsCSVWriter.WriteToStream(AStream: TStream);
var
  n: Integer;
begin
  if (CSVParams.SheetIndex >= 0) and (CSVParams.SheetIndex < FWorkbook.GetWorksheetCount)
    then n := CSVParams.SheetIndex
    else n := 0;
  WriteSheet(AStream, FWorkbook.GetWorksheetByIndex(n));
end;

procedure TsCSVWriter.WriteToStrings(AStrings: TStrings);
var
  stream: TStream;
begin
  stream := TStringStream.Create('');
  try
    WriteToStream(stream);
    stream.Position := 0;
    AStrings.LoadFromStream(stream);
  finally
    stream.Free;
  end;
end;


initialization
  InitCSVFormatSettings;
  RegisterSpreadFormat(TsCSVReader, TsCSVWriter, sfCSV);

end.

