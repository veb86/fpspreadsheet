unit fpscsv;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils,
  fpstypes, fpspreadsheet, fpsReaderWriter, fpsCsvDocument;

type
  TsCSVReader = class(TsCustomSpreadReader)
  private
    FWorksheetName: String;
    FFormatSettings: TFormatSettings;
    function IsQuotedText(var AText: String): Boolean;
    procedure ReadCellValue(ARow, ACol: Cardinal; AText: String);
  protected
    procedure ReadBlank(AStream: TStream); override;
    procedure ReadFormula(AStream: TStream); override;
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadNumber(AStream: TStream); override;
  public
    constructor Create(AWorkbook: TsWorkbook); override;
    procedure ReadFromFile(AFileName: String; AParams: TsStreamParams = []); override;
    procedure ReadFromStream(AStream: TStream; AParams: TsStreamParams = []); override;
    procedure ReadFromStrings(AStrings: TStrings; AParams: TsStreamParams = []); override;
  end;

  TsCSVWriter = class(TsCustomSpreadWriter)
  private
    FCSVBuilder: TCSVBuilder;
    FEncoding: String;
    FFormatSettings: TFormatSettings;
    FClipboardMode: Boolean;
  protected
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: Boolean; ACell: PCell); override;
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
    procedure WriteToStream(AStream: TStream; AParams: TsStreamParams = []); override;
    procedure WriteToStrings(AStrings: TStrings; AParams: TsStreamParams = []); override;
  end;

  TsCSVLineEnding = (leSystem, leCRLF, leCR, leLF);

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

  sfidCSV: TsSpreadFormatID;

function LineEndingAsString(ALineEnding: TsCSVLineEnding): String;


implementation

uses
  DateUtils, LConvEncoding, Math,
  fpsRegFileFormats, fpsUtils, fpsNumFormat;

function LineEndingAsString(ALineEnding: TsCSVLineEnding): String;
begin
  case ALineEnding of
    leSystem: Result := LineEnding;
    leCR    : Result := #13;
    leLF    : Result := #10;
    leCRLF  : Result := #13#10;
  end;
end;


{ -----------------------------------------------------------------------------}
{                              TsCSVReader                                     }
{------------------------------------------------------------------------------}

constructor TsCSVReader.Create(AWorkbook: TsWorkbook);
begin
  inherited Create(AWorkbook);
  FWorksheetName := 'Sheet1';  // will be replaced by filename
  FFormatSettings := CSVParams.FormatSettings;
  ReplaceFormatSettings(FFormatSettings, FWorkbook.FormatSettings);
end;

{ Checks if text is quoted; strips any starting and ending quotes }
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

{ Determines content types from/for the text read from the csv file and writes
  the corresponding data to the worksheet. }
procedure TsCSVReader.ReadCellValue(ARow, ACol: Cardinal; AText: String);
var
  dblValue: Double;
  dtValue: TDateTime;
  boolValue: Boolean;
  currSym: string;
  warning: String;
  nf: TsNumberFormat;
  decs: Integer;
  cell: PCell;
begin
  // Empty strings are blank cells -- nothing to do
  if AText = '' then
    exit;

  cell := FWorksheet.AddCell(ARow, ACol);

  // Do not try to interpret the strings. --> everything is a LABEL cell.
  if not CSVParams.DetectContentType then
  begin
    FWorksheet.WriteText(cell, AText);
    exit;
  end;

  // Check for a NUMBER or CURRENCY cell
  if IsNumberValue(AText, CSVParams.AutoDetectNumberFormat, FFormatSettings,
    dblValue, nf, decs, currSym, warning) then
  begin
    if currSym <> '' then
      FWorksheet.WriteCurrency(cell, dblValue, nfCurrency, decs, currSym)
    else
      FWorksheet.WriteNumber(cell, dblValue, nf, decs);
    if warning <> '' then
      FWorkbook.AddErrorMsg('Cell %s: %s', [GetCellString(ARow, ACol), warning]);
    exit;
  end;

  // Check for a DATE/TIME cell
  // No idea how to apply the date/time formatsettings here...
  if IsDateTimeValue(AText, FFormatSettings, dtValue, nf) then
  begin
    FWorksheet.WriteDateTime(cell, dtValue, nf);
    exit;
  end;

  // Check for a BOOLEAN cell
  if IsBoolValue(AText, CSVParams.TrueText, CSVParams.FalseText, boolValue) then
  begin
    FWorksheet.WriteBoolValue(cell, boolValue);
    exit;
  end;

  // What is left is handled as a TEXT cell
  FWorksheet.WriteText(cell, AText);
end;

procedure TsCSVReader.ReadFormula(AStream: TStream);
begin
  Unused(AStream);
end;

procedure TsCSVReader.ReadFromFile(AFileName: String;
  AParams: TsStreamParams = []);
begin
  FWorksheetName := ChangeFileExt(ExtractFileName(AFileName), '');
  inherited ReadFromFile(AFilename, AParams);
end;

procedure TsCSVReader.ReadFromStream(AStream: TStream;
  AParams: TsStreamParams = []);
var
  Parser: TCSVParser;
  s: String;
  encoding: String;
begin
  Unused(AParams);

  // Try to determine encoding of the input file
  SetLength(s, Min(1000, AStream.Size));
  AStream.ReadBuffer(s[1], Length(s));
  if CSVParams.Encoding = '' then
    encoding := GuessEncoding(s)
  else
    encoding := CSVParams.Encoding;

  // Create worksheet
  FWorksheet := FWorkbook.AddWorksheet(FWorksheetName, true);

  // Create csv parser, read file and store in worksheet
  Parser := TCSVParser.Create;
  try
    Parser.Delimiter := CSVParams.Delimiter;
    Parser.LineEnding := LineEndingAsString(CSVParams.LineEnding);
    Parser.QuoteChar := CSVParams.QuoteChar;
    // Indicate column counts between rows may differ:
    Parser.EqualColCountPerRow := false;
    Parser.SetSource(AStream);
    while Parser.ParseNextCell do begin
      // Convert string to UTF8
      s := Parser.CurrentCellText;
      s := ConvertEncoding(s, encoding, EncodingUTF8);
      ReadCellValue(Parser.CurrentRow, Parser.CurrentCol, s);
    end;
  finally
    Parser.Free;
  end;
end;

procedure TsCSVReader.ReadFromStrings(AStrings: TStrings;
  AParams: TsStreamParams = []);
var
  Stream: TStringStream;
begin
  Stream := TStringStream.Create(AStrings.Text);
  try
    ReadFromStream(Stream, AParams);
  finally
    Stream.Free;
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
  FFormatSettings := CSVParams.FormatSettings;
  ReplaceFormatSettings(FFormatSettings, FWorkbook.FormatSettings);
  if CSVParams.Encoding = '' then
    FEncoding := 'utf8'
  else
    FEncoding := CSVParams.Encoding;
end;

procedure TsCSVWriter.WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
  ACell: PCell);
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  // nothing to do
end;

{ Write boolean cell to stream formatted as string }
procedure TsCSVWriter.WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: Boolean; ACell: PCell);
var
  s: String;
begin
  Unused(AStream);
  Unused(ARow, ACol, ACell);
  if AValue then
    s := CSVParams.TrueText
  else
    s := CSVParams.FalseText;
  s := ConvertEncoding(s, EncodingUTF8, FEncoding);
  FCSVBuilder.AppendCell(s);
end;

{ Write date/time values in the same way they are displayed in the sheet }
procedure TsCSVWriter.WriteDateTime(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: TDateTime; ACell: PCell);
var
  s: String;
begin
  Unused(AStream);
  Unused(ARow, ACol, AValue);
  s := FWorksheet.ReadAsText(ACell);
  s := ConvertEncoding(s, EncodingUTF8, FEncoding);
  FCSVBuilder.AppendCell(s);
end;

{ CSV does not support formulas, but we can write the formula results to
  to stream. }
procedure TsCSVWriter.WriteFormula(AStream: TStream; const ARow, ACol: Cardinal;
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
procedure TsCSVWriter.WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: string; ACell: PCell);
var
  s: String;
begin
  Unused(AStream);
  Unused(ARow, ACol, AValue);
  if ACell = nil then
    exit;
  s := ACell^.UTF8StringValue;
  s := ConvertEncoding(s, EncodingUTF8, FEncoding);
  // No need to quote; csvdocument will do that for us...
  FCSVBuilder.AppendCell(s);
end;

{ Writes a number cell to the stream. }
procedure TsCSVWriter.WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
  const AValue: double; ACell: PCell);
var
  s: String;
begin
  Unused(AStream);
  Unused(ARow, ACol);
  if ACell = nil then
    exit;
  if CSVParams.NumberFormat <> '' then
    s := Format(CSVParams.NumberFormat, [AValue], FFormatSettings)
  else
    s := FWorksheet.ReadAsText(ACell, FFormatSettings);
  s := ConvertEncoding(s, EncodingUTF8, FEncoding);
  FCSVBuilder.AppendCell(s);
end;

procedure TsCSVWriter.WriteSheet(AStream: TStream; AWorksheet: TsWorksheet);
var
  r: Cardinal;
  firstRow: Cardinal;
  lastRow: Cardinal;
  cell: PCell;
  n: Integer;
begin
  FWorksheet := AWorksheet;

  FCSVBuilder := TCSVBuilder.Create;
  try
    FCSVBuilder.Delimiter := CSVParams.Delimiter;
    FCSVBuilder.LineEnding := LineEndingAsString(CSVParams.LineEnding);
    FCSVBuilder.QuoteChar := CSVParams.QuoteChar;
    FCSVBuilder.SetOutput(AStream);

    n := FWorksheet.GetCellCount;
    if FClipboardMode and (n = 1) then
    begin
      cell := FWorksheet.Cells.GetFirstCell;
      WriteCellToStream(AStream, cell);
    end else
    begin
      if FClipboardMode then
        firstRow := FWorksheet.GetFirstRowIndex else
        firstRow := 0;
      lastRow := FWorksheet.GetLastOccupiedRowIndex;
      for r := firstRow to lastRow do
      begin
        for cell in FWorksheet.Cells.GetRowEnumerator(r) do
          WriteCellToStream(AStream, cell);
        FCSVBuilder.AppendRow;
      end;
    end;
  finally
    FreeAndNil(FCSVBuilder);
  end;
end;

procedure TsCSVWriter.WriteToStream(AStream: TStream;
  AParams: TsStreamParams = []);
var
  n: Integer;
begin
  FClipboardMode := (spClipboard in AParams);
  if (CSVParams.SheetIndex >= 0) and (CSVParams.SheetIndex < FWorkbook.GetWorksheetCount)
    then n := CSVParams.SheetIndex
    else n := 0;
  WriteSheet(AStream, FWorkbook.GetWorksheetByIndex(n));
end;

procedure TsCSVWriter.WriteToStrings(AStrings: TStrings;
  AParams: TsStreamParams = []);
var
  Stream: TStream;
begin
  Stream := TStringStream.Create('');
  try
    WriteToStream(Stream, AParams);
    Stream.Position := 0;
    AStrings.LoadFromStream(Stream);
  finally
    Stream.Free;
  end;
end;


initialization
  InitFormatSettings(CSVParams.FormatSettings);

  // Registers this reader / writer in fpSpreadsheet
  sfidCSV := RegisterSpreadFormat(sfCSV,
    TsCSVReader, TsCSVWriter,
    STR_FILEFORMAT_CSV, 'CSV', [STR_COMMA_SEPARATED_EXTENSION, '.txt']
  );

end.

