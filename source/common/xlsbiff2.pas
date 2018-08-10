{
xlsbiff2.pas

Writes an Excel 2.x file

Excel 2.x files support only one Worksheet per Workbook, so only the first
will be written.

An Excel file consists of a number of subsequent records.
To ensure a properly formed file, the following order must be respected:

1st record:        BOF
2nd to Nth record: Any record
Last record:       EOF

The row and column numbering in BIFF files is zero-based.

Excel file format specification obtained from:

http://sc.openoffice.org/excelfileformat.pdf

Encoding information: ISO_8859_1 is used, to have support to
other characters, please use a format which support unicode

AUTHORS: Felipe Monteiro de Carvalho
}
unit xlsbiff2;

{$ifdef fpc}
  {$mode delphi}{$H+}
{$endif}

interface

uses
  Classes, SysUtils, lconvencoding,
  fpsTypes, fpsUtils, xlscommon;

const
  BIFF2_MAX_PALETTE_SIZE = 8;
  // There are more colors but they do not seem to be controlled by a palette.

type

  { TsSpreadBIFF2Reader }

  TsSpreadBIFF2Reader = class(TsSpreadBIFFReader)
  private
    FFont: TsFont;
    FPendingXFIndex: Word;
  protected
    procedure AddBuiltinNumFormats; override;
    procedure ReadBlank(AStream: TStream); override;
    procedure ReadBool(AStream: TStream); override;
    procedure ReadColumnDefault(AStream: TStream);
    procedure ReadColWidth(AStream: TStream);
    procedure ReadDefRowHeight(AStream: TStream);
    procedure ReadFONT(AStream: TStream);
    procedure ReadFONTCOLOR(AStream: TStream);
    procedure ReadFORMAT(AStream: TStream); override;
    procedure ReadFormula(AStream: TStream); override;
    procedure ReadInteger(AStream: TStream);
    procedure ReadIXFE(AStream: TStream);
    procedure ReadLabel(AStream: TStream); override;
    procedure ReadNumber(AStream: TStream); override;
    procedure ReadPASSWORD(AStream: TStream);
    procedure ReadPROTECT(AStream: TStream);
    procedure ReadRowColXF(AStream: TStream; out ARow, ACol: Cardinal; out AXF: Word); override;
    procedure ReadRowInfo(AStream: TStream); override;
    procedure ReadRPNAttr(AStream: TStream; AIdentifier: Byte); override;
    function ReadRPNFunc(AStream: TStream): Word; override;
    procedure ReadRPNSharedFormulaBase(AStream: TStream; out ARow, ACol: Cardinal); override;
    function ReadRPNTokenArraySize(AStream: TStream): Word; override;
    procedure ReadStringRecord(AStream: TStream); override;
    procedure ReadWindow2(AStream: TStream); override;
    procedure ReadXF(AStream: TStream);
  public
    constructor Create(AWorkbook: TsBasicWorkbook); override;
    { General reading methods }
    procedure ReadFromStream(AStream: TStream; APassword: String = '';
      AParams: TsStreamParams = []); override;
  end;


  { TsSpreadBIFF2Writer }

  TsSpreadBIFF2Writer = class(TsSpreadBIFFWriter)
  private
    FSheetIndex: Integer;  // Index of worksheet to be written
    procedure GetAttributes(AFormatIndex: Integer; XFIndex: Word;
      out Attrib1, Attrib2, Attrib3: Byte);
    procedure GetFormatAndFontIndex(AFormatRecord: PsCellFormat;
      out AFormatIndex, AFontIndex: Integer);
    { Record writing methods }
    procedure WriteBOF(AStream: TStream);
    procedure WriteCellAttributes(AStream: TStream; AFormatIndex: Integer; XFIndex: Word);
    procedure WriteColWidth(AStream: TStream; ACol: PCol);
    procedure WriteColWidths(AStream: TStream);
//    procedure WriteColumnDefault(AStream: TStream; ACol: PCol);
    procedure WriteColumnDefault(AStream: TStream;
        AFirstColIndex, ALastColIndex: Word; AFormatIndex: Integer);
    procedure WriteColumnDefaults(AStream: TStream);
    procedure WriteDefaultRowHeight(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteDimensions(AStream: TStream; AWorksheet: TsBasicWorksheet);
    procedure WriteEOF(AStream: TStream);
    procedure WriteFont(AStream: TStream; AFontIndex: Integer);
    procedure WriteFonts(AStream: TStream);
    procedure WriteFORMATCOUNT(AStream: TStream);
    procedure WriteIXFE(AStream: TStream; XFIndex: Word);
  protected
    procedure AddBuiltinNumFormats; override;
    function FunctionSupported(AExcelCode: Integer;
      const AFuncName: String): Boolean; override;
    procedure PopulatePalette(AWorkbook: TsbasicWorkbook); override;
    procedure WriteBlank(AStream: TStream; const ARow, ACol: Cardinal;
      ACell: PCell); override;
    procedure WriteBool(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: Boolean; ACell: PCell); override;
    procedure WriteCodePage(AStream: TStream; ACodePage: String); override;
    procedure WriteError(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: TsErrorValue; ACell: PCell); override;
    procedure WriteFORMAT(AStream: TStream; ANumFormatStr: String;
      AFormatIndex: Integer); override;
    procedure WriteLabel(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: string; ACell: PCell); override;
    procedure WriteNumber(AStream: TStream; const ARow, ACol: Cardinal;
      const AValue: double; ACell: PCell); override;
    procedure WritePASSWORD(AStream: TStream);
    procedure WriteRow(AStream: TStream; ASheet: TsBasicWorksheet;
      ARowIndex, AFirstColIndex, ALastColIndex: Cardinal; ARow: PRow); override;
    procedure WriteRPNFormula(AStream: TStream; const ARow, ACol: Cardinal;
      AFormula: TsRPNFormula; ACell: PCell); override;
    function WriteRPNFunc(AStream: TStream; AIdentifier: Word): Word; override;
    procedure WriteRPNTokenArraySize(AStream: TStream; ASize: Word); override;
    procedure WriteStringRecord(AStream: TStream; AString: String); override;
    procedure WriteWindow1(AStream: TStream); override;
    procedure WriteWindow2(AStream: TStream; ASheet: TsBasicWorksheet);
    procedure WriteXF(AStream: TStream; AFormatRecord: PsCellFormat;
      XFType_Prot: Byte = 0); override;
  public
    constructor Create(AWorkbook: TsbasicWorkbook); override;
    procedure WriteToStream(AStream: TStream; AParams: TsStreamParams = []); override;
  end;

  TExcel2Settings = record
    // Settings used when writing to file
    DateMode: TDateMode;
    CodePage: String;
    SheetIndex: Integer;
  end;

var
  Excel2Settings: TExcel2Settings = (
    DateMode: dm1900;
    CodePage: 'cp1252';   // on Windows, will be replaced --> see initalization
    SheetIndex: 0;
  );

  { the palette of the default BIFF2 colors as "big-endian color" values }
  PALETTE_BIFF2: array[$0..$07] of TsColor = (
    $000000,  // $00: black
    $FFFFFF,  // $01: white
    $FF0000,  // $02: red
    $00FF00,  // $03: green
    $0000FF,  // $04: blue
    $FFFF00,  // $05: yellow
    $FF00FF,  // $06: magenta
    $00FFFF   // $07: cyan
  );

  sfidExcel2: TsSpreadFormatID;

procedure InitBiff2Limitations(out ALimitations: TsSpreadsheetFormatLimitations);


implementation

uses
 {$IFDEF FPSpreadDebug}
  LazLogger,
 {$ENDIF}
  Math,
  fpsStrings, fpspreadsheet, fpsReaderWriter, fpsPalette, fpsNumFormat;

const
  { Excel record IDs }
  INT_EXCEL_ID_DIMENSIONS    = $0000;
  INT_EXCEL_ID_BLANK         = $0001;
  INT_EXCEL_ID_INTEGER       = $0002;
  INT_EXCEL_ID_NUMBER        = $0003;
  INT_EXCEL_ID_LABEL         = $0004;
  INT_EXCEL_ID_BOOLERROR     = $0005;
  INT_EXCEL_ID_ROW           = $0008;
  INT_EXCEL_ID_BOF           = $0009;
  {%H-}INT_EXCEL_ID_INDEX    = $000B;
  INT_EXCEL_ID_FORMAT        = $001E;
  INT_EXCEL_ID_FORMATCOUNT   = $001F;
  INT_EXCEL_ID_COLUMNDEFAULT = $0020;
  INT_EXCEL_ID_COLWIDTH      = $0024;
  INT_EXCEL_ID_DEFROWHEIGHT  = $0025;
  INT_EXCEL_ID_WINDOW2       = $003E;
  INT_EXCEL_ID_XF            = $0043;
  INT_EXCEL_ID_IXFE          = $0044;
  INT_EXCEL_ID_FONTCOLOR     = $0045;

  { BOF record constants }
  INT_EXCEL_SHEET            = $0010;
  {%H-}INT_EXCEL_CHART       = $0020;
  {%H-}INT_EXCEL_MACRO_SHEET = $0040;

  MASK_XF_TYPE_PROT_LOCKED_BIFF2          = $40;
  MASK_XF_TYPE_PROT_FORMULA_HIDDEN_BIFF2  = $80;

type
  TBIFF2_BoolErrRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    BoolErrValue: Byte;
    ValueType: Byte;
  end;

  TBIFF2_DimensionsRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FirstRow: Word;
    LastRowPlus1: Word;
    FirstCol: Word;
    LastColPlus1: Word;
  end;

  TBIFF2_LabelRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    TextLen: Byte;
  end;

  TBIFF2_NumberRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    Value: Double;
  end;

  TBIFF2_IntegerRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1: Byte;
    Attrib2: Byte;
    Attrib3: Byte;
    Value: Word;
  end;

  TBIFF2_XFRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FontIndex: Byte;
    NotUsed: Byte;
    NumFormat_Prot: Byte;
    HorAlign_Border_BkGr: Byte;
  end;


procedure InitBiff2Limitations(out ALimitations: TsSpreadsheetFormatLimitations);
begin
  InitBiffLimitations(ALimitations);
  ALimitations.MaxPaletteSize := BIFF2_MAX_PALETTE_SIZE;
end;

procedure InternalAddBuiltinNumFormats(AList: TStringList; AFormatSettings: TFormatSettings);
var
  fs: TFormatSettings absolute AFormatSettings;
  cs: String;
begin
  cs := fs.CurrencyString;
  with AList do
  begin
    Clear;
    Add('');          // 0
    Add('0');         // 1
    Add('0.00');      // 2
    Add('#,##0');     // 3
    Add('#,##0.00');  // 4
    Add(BuildCurrencyFormatString(nfCurrency, fs, 0, fs.CurrencyFormat, fs.NegCurrFormat, cs));     // 5
    Add(BuildCurrencyFormatString(nfCurrencyRed, fs, 0, fs.CurrencyFormat, fs.NegCurrFormat, cs));  // 6
    Add(BuildCurrencyFormatString(nfCurrency, fs, 2, fs.CurrencyFormat, fs.NegCurrFormat, cs));     // 7
    Add(BuildCurrencyFormatString(nfCurrencyRed, fs, 2, fs.CurrencyFormat, fs.NegCurrFormat, cs));  // 8
    Add('0%');        // 9
    Add('0.00%');     // 10
    Add('0.00E+00');  // 11
    Add(BuildDateTimeFormatString(nfShortDate, fs));     // 12
    Add(BuildDateTimeFormatString(nfLongDate, fs));      // 13
    Add(BuildDateTimeFormatString(nfDayMonth, fs));      // 14: 'd/mmm'
    Add(BuildDateTimeFormatString(nfMonthYear, fs));     // 15: 'mmm/yy'
    Add(BuildDateTimeFormatString(nfShortTimeAM, fs));   // 16;
    Add(BuildDateTimeFormatString(nfLongTimeAM, fs));    // 17
    Add(BuildDateTimeFormatString(nfShortTime, fs));     // 18
    Add(BuildDateTimeFormatString(nfLongTime, fs));      // 19
    Add(BuildDateTimeFormatString(nfShortDateTime, fs)); // 20
  end;
end;


{------------------------------------------------------------------------------}
{                             TsSpreadBIFF2Reader                              }
{------------------------------------------------------------------------------}

constructor TsSpreadBIFF2Reader.Create(AWorkbook: TsBasicWorkbook);
begin
  inherited Create(AWorkbook);
  InitBiff2Limitations(FLimitations);
end;

procedure TsSpreadBIFF2Reader.AddBuiltInNumFormats;
begin
  FFirstNumFormatIndexInFile := 0;
end;

procedure TsSpreadBIFF2Reader.ReadBlank(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
  cell: PCell;
begin
  ReadRowColXF(AStream, ARow, ACol, XF);
  if FIsVirtualMode then begin
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := TsWorksheet(FWorksheet).AddCell(ARow, ACol);
  ApplyCellFormatting(cell, XF);
  if FIsVirtualMode then
    TsWorkbook(FWorkbook).OnReadCellData(Workbook, ARow, ACol, cell);
end;

{@@ ----------------------------------------------------------------------------
  The name of this method is misleading - it reads a BOOLEAN cell value,
  but also an ERROR value; BIFF stores them in the same record.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadBool(AStream: TStream);
var
  rec: TBIFF2_BoolErrRecord;
  r, c: Cardinal;
  xf: Word;
  cell: PCell;
  sheet: TsWorksheet;
begin
  sheet := FWorksheet as TsWorksheet;

  { Read entire record, starting at Row }
  rec.Row := 0;  // to silence the compiler...
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2_BoolErrRecord) - 2*SizeOf(Word));
  r := WordLEToN(rec.Row);
  c := WordLEToN(rec.Col);
  xf := rec.Attrib1 and $3F;
  if xf = 63 then xf := FPendingXFIndex;

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(FWorksheet, r, c, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := sheet.AddCell(r, c);

  { Retrieve boolean or error value depending on the "ValueType" }
  case rec.ValueType of
    0: sheet.WriteBoolValue(cell, boolean(rec.BoolErrValue));
    1: sheet.WriteErrorValue(cell, ConvertFromExcelError(rec.BoolErrValue));
  end;

  { Apply formatting }
  ApplyCellFormatting(cell, xf);

  if FIsVirtualMode then
    TsWorkbook(FWorkbook).OnReadCellData(Workbook, r, c, cell);
end;

procedure TsSpreadBIFF2Reader.ReadColumnDefault(AStream: TStream);
var
  c, col1, col2: Word;
  attr2, attr3: Byte;
  fmt: TsCellFormat;
  fmtIndex: Integer;
  fontIndex: Integer;
  fnt: TsFont;
  nf: TsNumFormatParams;
  nfs: String;
  b: Byte;
  book: TsWorkbook;
begin
  book := TsWorkbook(FWorkbook);

  { Index of first column }
  col1 := WordLEToN(AStream.ReadWord);

  { Index of last column - note: the file value is incremented by 1 }
  col2 := WordLEToN(AStream.ReadWord) - 1;

  { Attributes }
  {attr1 := }AStream.ReadByte;     // Avoid compiler warning of unused attr1
  attr2 := AStream.ReadByte;
  attr3 := AStream.ReadByte;

  InitFormatRecord(fmt);
  fmt.ID := FCellFormatList.Count;

  // Font index
  fontIndex := (attr2 and $C0) shr 6;
  if fontIndex > 4 then dec(fontIndex);  // Watch out for the nasty missing font #4...
  fnt := TsFont(FFontList[fontIndex]);
  fmt.FontIndex := book.FindFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color);
  if fmt.FontIndex = -1 then
    fmt.FontIndex := book.AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color);
  if fmt.FontIndex > 0 then
    Include(fmt.UsedFormattingFields, uffFont);

  // Number format index
  b := attr2 and $3F;
  nfs := NumFormatList[b];
  if nfs <> '' then begin
    fmt.NumberFormatIndex := book.AddNumberFormat(nfs);
    nf := book.GetNumberFormat(fmt.NumberFormatIndex);
    fmt.NumberFormat := nf.NumFormat;
    fmt.NumberFormatStr := nf.NumFormatStr;
    if fmt.NumberFormat <> nfGeneral then
      Include(fmt.UsedFormattingfields, uffNumberFormat);
  end;

  // Horizontal alignment
  b := attr3 and MASK_XF_HOR_ALIGN;
  if (b <= ord(High(TsHorAlignment))) then
  begin
    fmt.HorAlignment := TsHorAlignment(b);
    if fmt.HorAlignment <> haDefault then
      Include(fmt.UsedFormattingFields, uffHorAlign);
  end;

  // Vertical alignment - not used in BIFF2
  fmt.VertAlignment := vaDefault;

  // Word wrap - not used in BIFF2
  // -- nothing to do here

  // Text rotation - not used in BIFF2
  // -- nothing to do here

  // Borders
  fmt.Border := [];
  if attr3 and $08 <> 0 then
    Include(fmt.Border, cbWest);
  if attr3 and $10 <> 0 then
    Include(fmt.Border, cbEast);
  if attr3 and $20 <> 0 then
    Include(fmt.Border, cbNorth);
  if attr3 and $40 <> 0 then
    Include(fmt.Border, cbSouth);
  if fmt.Border <> [] then
    Include(fmt.UsedFormattingFields, uffBorder);

  // Background color not supported, only shaded background
  if attr3 and $80 <> 0 then
  begin
    fmt.Background.Style := fsGray50;
    fmt.Background.FgColor := scBlack;
    fmt.Background.BgColor := scTransparent;
    Include(fmt.UsedFormattingFields, uffBackground);
  end;

  // Add the decoded data to the format list
  FCellFormatList.Add(fmt);
  fmtIndex := book.AddCellFormat(fmt);

  for c := col1 to col2 do
    TsWorksheet(FWorksheet).WriteColFormatIndex(c, fmtIndex);
end;

procedure TsSpreadBIFF2Reader.ReadColWidth(AStream: TStream);
const
  EPS = 1E-3;
var
  c, c1, c2: Cardinal;
  w: Word;
  colwidth: Single;
  sheet: TsWorksheet;
begin
  sheet := TsWorksheet(FWorksheet);

  // read column start and end index of column range
  c1 := AStream.ReadByte;
  c2 := AStream.ReadByte;
  // read col width in 1/256 of the width of "0" character
  w := WordLEToN(AStream.ReadWord);
  // calculate width in units of "characters"
  colwidth := (FWorkbook as TsWorkbook).ConvertUnits(w / 256, suChars, FWorkbook.Units);
  // assign width to columns, but only if different from default column width.
  if not SameValue(colwidth, sheet.ReadDefaultColWidth(FWorkbook.Units), EPS) then
    for c := c1 to c2 do
      sheet.WriteColWidth(c, colwidth, FWorkbook.Units);
end;

procedure TsSpreadBIFF2Reader.ReadDefRowHeight(AStream: TStream);
var
  hw: word;
  h: Single;
begin
  hw := WordLEToN(AStream.ReadWord);
  h := TwipsToPts(hw and $7FFF);
  (FWorksheet as TsWorksheet).WriteDefaultRowHeight(h, suPoints);
end;

procedure TsSpreadBIFF2Reader.ReadFONT(AStream: TStream);
var
  lHeight: Word;
  lOptions: Word;
  Len: Byte;
  lFontName: UTF8String;
  isDefaultFont: Boolean;
begin
  FFont := TsFont.Create;

  { Height of the font in twips = 1/20 of a point }
  lHeight := WordLEToN(AStream.ReadWord);
  FFont.Size := lHeight/20;

  { Option flags }
  lOptions := WordLEToN(AStream.ReadWord);
  FFont.Style := [];
  if lOptions and $0001 <> 0 then Include(FFont.Style, fssBold);
  if lOptions and $0002 <> 0 then Include(FFont.Style, fssItalic);
  if lOptions and $0004 <> 0 then Include(FFont.Style, fssUnderline);
  if lOptions and $0008 <> 0 then Include(FFont.Style, fssStrikeout);

  { Font name: Unicodestring, char count in 1 byte }
  Len := AStream.ReadByte();
  SetLength(lFontName, Len);
  AStream.ReadBuffer(lFontName[1], Len);
  FFont.FontName := lFontName;

  isDefaultFont := FFontList.Count = 0;

  { Add font to internal font list }
  FFontList.Add(FFont);

  if isDefaultFont then
    TsWorkbook(FWorkbook).SetDefaultFont(FFont.FontName, FFont.Size);
end;

procedure TsSpreadBIFF2Reader.ReadFONTCOLOR(AStream: TStream);
var
  lColor: Word;
begin
  lColor := WordLEToN(AStream.ReadWord);   // Palette index
  FFont.Color := IfThen(lColor = SYS_DEFAULT_WINDOW_TEXT_COLOR,
    scBlack, FPalette[lColor]);
end;

{@@ ----------------------------------------------------------------------------
  Reads the FORMAT record required for formatting numerical data
-------------------------------------------------------------------------------}
(*
procedure TsSpreadBIFF2Reader.ReadFORMAT(AStream: TStream);
begin
  Unused(AStream);
  // We ignore the formats in the file, everything is known
  // (Using the formats in the file would require de-localizing them).
end;*)
procedure TsSpreadBIFF2Reader.ReadFormat(AStream: TStream);
var
  len: byte;
  fmtString: AnsiString;
  nfs: String;
begin
  // number format string
  len := AStream.ReadByte;
  SetLength(fmtString, len);
  AStream.ReadBuffer(fmtString[1], len);

  // We need the format string as utf8 and non-localized
  nfs := ConvertEncoding(fmtString, FCodePage, encodingUTF8);

  // Add to the end of the list.
  NumFormatList.Add(nfs);
end;


procedure TsSpreadBIFF2Reader.ReadFromStream(AStream: TStream;
  APassword: String = ''; AParams: TsStreamParams = []);
var
  BIFF2EOF: Boolean;
  RecordType: Word;
  CurStreamPos: Int64;
  BOFFound: Boolean;
begin
  Unused(APassword, AParams);
  BIFF2EOF := False;

  { In BIFF2 files there is only one worksheet, let's create it }
  FWorksheet := TsWorkbook(FWorkbook).AddWorksheet('Sheet', true);

  { Read all records in a loop }
  BOFFound := false;
  while not BIFF2EOF do
  begin
    { Read the record header }
    RecordType := WordLEToN(AStream.ReadWord);
    RecordSize := WordLEToN(AStream.ReadWord);

    CurStreamPos := AStream.Position;

    case RecordType of
      INT_EXCEL_ID_BLANK         : ReadBlank(AStream);
      INT_EXCEL_ID_BOF           : BOFFound := true;
      INT_EXCEL_ID_BOOLERROR     : ReadBool(AStream);
      INT_EXCEL_ID_BOTTOMMARGIN  : ReadMargin(AStream, 3);
      INT_EXCEL_ID_CODEPAGE      : ReadCodePage(AStream);
      INT_EXCEL_ID_COLUMNDEFAULT : ReadColumnDefault(AStream);
      INT_EXCEL_ID_COLWIDTH      : ReadColWidth(AStream);
      INT_EXCEL_ID_DEFCOLWIDTH   : ReadDefColWidth(AStream);
      INT_EXCEL_ID_EOF           : BIFF2EOF := True;
      INT_EXCEL_ID_FONT          : ReadFont(AStream);
      INT_EXCEL_ID_FONTCOLOR     : ReadFontColor(AStream);
      INT_EXCEL_ID_FOOTER        : ReadHeaderFooter(AStream, false);
      INT_EXCEL_ID_FORMAT        : ReadFormat(AStream);
      INT_EXCEL_ID_FORMULA       : ReadFormula(AStream);
      INT_EXCEL_ID_HEADER        : ReadHeaderFooter(AStream, true);
      INT_EXCEL_ID_INTEGER       : ReadInteger(AStream);
      INT_EXCEL_ID_IXFE          : ReadIXFE(AStream);
      INT_EXCEL_ID_LABEL         : ReadLabel(AStream);
      INT_EXCEL_ID_LEFTMARGIN    : ReadMargin(AStream, 0);
      INT_EXCEL_ID_NOTE          : ReadComment(AStream);
      INT_EXCEL_ID_NUMBER        : ReadNumber(AStream);
      INT_EXCEL_ID_PANE          : ReadPane(AStream);
      INT_EXCEL_ID_OBJECTPROTECT : ReadObjectProtect(AStream);
      INT_EXCEL_ID_PASSWORD      : ReadPASSWORD(AStream);
      INT_EXCEL_ID_PRINTGRID     : ReadPrintGridLines(AStream);
      INT_EXCEL_ID_PRINTHEADERS  : ReadPrintHeaders(AStream);
      INT_EXCEL_ID_PROTECT       : ReadPROTECT(AStream);
      INT_EXCEL_ID_RIGHTMARGIN   : ReadMargin(AStream, 1);
      INT_EXCEL_ID_ROW           : ReadRowInfo(AStream);
      INT_EXCEL_ID_SELECTION     : ReadSELECTION(AStream);
      INT_EXCEL_ID_STRING        : ReadStringRecord(AStream);
      INT_EXCEL_ID_TOPMARGIN     : ReadMargin(AStream, 2);
      INT_EXCEL_ID_DEFROWHEIGHT  : ReadDefRowHeight(AStream);
      INT_EXCEL_ID_WINDOW2       : ReadWindow2(AStream);
      INT_EXCEL_ID_WINDOWPROTECT : ReadWindowProtect(AStream);
      INT_EXCEL_ID_XF            : ReadXF(AStream);
    else
      // nothing
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    if AStream.Position >= AStream.Size then
      BIFF2EOF := True;

    if not BOFFound then
      raise EFPSpreadsheet.Create('BOF record not found.');
  end;

  FixCols(FWorksheet);
  FixRows(FWorksheet);
end;

procedure TsSpreadBIFF2Reader.ReadFormula(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: Word;
  ok: Boolean;
  formulaResult: Double = 0.0;
  Data: array [0..7] of byte;
  dt: TDateTime;
  nf: TsNumberFormat;
  nfs: String;
  err: TsErrorValue;
  cell: PCell;
  sheet: TsWorksheet;
begin
  sheet := TsWorksheet(FWorksheet);

  { BIFF Record row/column/style }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Result of the formula result in IEEE 754 floating-point value }
  Data[0] := 0;   // to silence the compiler...
  AStream.ReadBuffer(Data, Sizeof(Data));

  { Recalculation byte - currently not used }
  AStream.ReadByte;

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := sheet.AddCell(ARow, ACol);

  // Now determine the type of the formula result
  if (Data[6] = $FF) and (Data[7] = $FF) then
    case Data[0] of
      0: // String -> Value is found in next record (STRING)
         FIncompleteCell := cell;
      1: // Boolean value
         sheet.WriteBoolValue(cell, Data[2] = 1);
      2: begin  // Error value
           case Data[2] of
             ERR_INTERSECTION_EMPTY   : err := errEmptyIntersection;
             ERR_DIVIDE_BY_ZERO       : err := errDivideByZero;
             ERR_WRONG_TYPE_OF_OPERAND: err := errWrongType;
             ERR_ILLEGAL_REFERENCE    : err := errIllegalRef;
             ERR_WRONG_NAME           : err := errWrongName;
             ERR_OVERFLOW             : err := errOverflow;
             ERR_ARG_ERROR            : err := errArgError;
           end;
           sheet.WriteErrorValue(cell, err);
         end;
      3: // Empty cell
         sheet.WriteBlank(cell);
    end
  else
  begin
    // Result is a number or a date/time
    Move(Data[0], formulaResult, SizeOf(Data));

    {Find out what cell type, set content type and value}
    ExtractNumberFormat(XF, nf, nfs);
    if IsDateTime(formulaResult, nf, nfs, dt) then
      sheet.WriteDateTime(cell, dt, nf, nfs)
    else
      sheet.WriteNumber(cell, formulaResult, nf, nfs);
  end;

  { Formula token array }
  if (boReadFormulas in FWorkbook.Options) then
  begin
    ok := ReadRPNTokenArray(AStream, cell);
    if not ok then sheet.WriteErrorValue(cell, errFormulaNotSupported);
  end;

  { Apply formatting to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    TsWorkbook(FWorkbook).OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF2Reader.ReadLabel(AStream: TStream);
var
  rec: TBIFF2_LabelRecord;
  L: Byte;
  ARow, ACol: Cardinal;
  XF: Word;
  ansiStr: ansistring;
  valueStr: UTF8String;
  cell: PCell;
  sheet: TsWorksheet;
begin
  sheet := FWorksheet as TsWorksheet;

  { Read entire record, starting at Row, except for string data }
  rec.Row := 0;  // to silence the compiler...
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2_LabelRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := rec.Attrib1 and $3F;
  if XF = 63 then XF := FPendingXFIndex;

  { String with 8-bit size }
  L := rec.TextLen;
  SetLength(ansiStr, L);
  AStream.ReadBuffer(ansiStr[1], L);

  { Save the data }
  valueStr := ConvertEncoding(ansiStr, FCodePage, encodingUTF8);

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := sheet.AddCell(ARow, ACol);
  sheet.WriteText(cell, valueStr);

  { Apply formatting to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    TsWorkbook(FWorkbook).OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF2Reader.ReadNumber(AStream: TStream);
var
  rec: TBIFF2_NumberRecord;
  ARow, ACol: Cardinal;
  XF: Word;
  value: Double = 0.0;
  dt: TDateTime;
  nf: TsNumberFormat;
  nfs: String;
  cell: PCell;
  sheet: TsWorksheet;
begin
  sheet := FWorksheet as TsWorksheet;

  { Read entire record, starting at Row }
  rec.Row := 0;  // to silence the compiler...
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2_NumberRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := rec.Attrib1 and $3F;
  if XF = 63 then XF := FPendingXFIndex;
  value := rec.Value;

  {Create cell}
  if FIsVirtualMode then begin
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := sheet.AddCell(ARow, ACol);

  {Find out what cell type, set content type and value}
  ExtractNumberFormat(XF, nf, nfs);
  if IsDateTime(value, nf, nfs, dt) then
    sheet.WriteDateTime(cell, dt, nf, nfs)
  else
    sheet.WriteNumber(cell, value, nf, nfs);

  { Apply formatting to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    TsWorkbook(FWorkbook).OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF2Reader.ReadInteger(AStream: TStream);
var
  sheet: TsWorksheet;
  ARow, ACol: Cardinal;
  XF: Word;
  AWord  : Word = 0;
  cell: PCell;
  rec: TBIFF2_IntegerRecord;
begin
  sheet := FWorksheet as TsWorksheet;

  { Read record into buffer }
  rec.Row := 0;   // to silence the comiler...
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF2_NumberRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := rec.Attrib1 and $3F;
  if XF = 63 then XF := FPendingXFIndex;
  AWord := WordLEToN(rec.Value);

  { Create cell }
  if FIsVirtualMode then
  begin
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := sheet.AddCell(ARow, ACol);

  { Save the data }
  sheet.WriteNumber(cell, AWord);

  { Apply formatting to cell }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    TsWorkbook(FWorkbook).OnReadCellData(Workbook, ARow, ACol, cell);
end;

{@@ ----------------------------------------------------------------------------
  Reads an IXFE record. This record contains the "true" XF index of a cell. It
  is used if there are more than 62 XF records (XF field is only 6-bit). The
  IXFE record is used in front of the cell record using it
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadIXFE(AStream: TStream);
begin
  FPendingXFIndex := WordLEToN(AStream.ReadWord);
end;

{@@ ----------------------------------------------------------------------------
  Reads a PASSWORD record. Since BIFF2 does not have multiple worksheets the
  same password is stored in the workbook and worksheet cryptoinfo records.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadPASSWORD(AStream: TStream);
var
  hash: Word;
  cinfo: TsCryptoInfo;
begin
  hash := WordLEToN(AStream.ReadWord);
  if hash = 0 then
    exit;  // no password

  InitCryptoInfo(cinfo);
  cinfo.PasswordHash := Format('%.4x', [hash]);
  cinfo.Algorithm := caExcel;

  // Use the same password for workbook and worksheet protection because
  // BIFF2 can have only a single sheet.
  (FWorkbook as TsWorkbook).CryptoInfo := cinfo;
  (FWorksheet as TsWorksheet).CryptoInfo := cinfo;
end;

procedure TsSpreadBIFF2Reader.ReadPROTECT(AStream: TStream);
begin
  inherited ReadPROTECT(AStream);
  (FWorksheet as TsWorksheet).Protect(Workbook.IsProtected);
end;

{@@ ----------------------------------------------------------------------------
  Reads the row, column and xf index from the stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadRowColXF(AStream: TStream;
  out ARow, ACol: Cardinal; out AXF: WORD);
begin
  { BIFF Record data for row and column}
  ARow := WordLEToN(AStream.ReadWord);
  ACol := WordLEToN(AStream.ReadWord);

  { Index to XF record }
  AXF := AStream.ReadByte and $3F;
  // If AXF = $3F = 63 then there is an IXFE record containing the true XF index!
  if AXF = $3F then
    AXF := FPendingXFIndex;

  { Index to format and font record, cell style - ignored because contained in XF
    Must read to keep the record in sync. }
  AStream.ReadWord;
end;

procedure TsSpreadBIFF2Reader.ReadRowInfo(AStream: TStream);
type
  TRowRecord = packed record
    RowIndex: Word;
    Col1: Word;
    Col2: Word;
    Height: Word;
    NotUsed: Word;
    ContainsXF: Byte;
    OffsetToCell: Word;
    Attributes1: Byte;
    Attributes2: Byte;
    Attributes3: Byte;
    XFIndex: Word;
  end;
var
  rowrec: TRowRecord;
  lRow: PRow;
  h: word;
  auto: Boolean;
  rowheight: Single;
  defRowHeight: Single;
  containsXF: Boolean;
  xf: Word;
  book: TsWorkbook;
  sheet: TsWorksheet;
begin
  book := FWorkbook as TsWorkbook;
  sheet := FWorksheet as TsWorksheet;

  rowRec.RowIndex := 0;  // to silence the compiler...
  AStream.ReadBuffer(rowrec, SizeOf(TRowRecord));
  h := WordLEToN(rowrec.Height);
  auto := h and $8000 <> 0;
  rowheight := book.ConvertUnits(TwipsToPts(h and $7FFF), suPoints, FWorkbook.Units);
  defRowHeight := sheet.ReadDefaultRowHeight(FWorkbook.Units);
  containsXF := rowRec.ContainsXF = 1;
  xf := WordLEToN(rowRec.XFIndex);

  // No row record if rowheight in file is the same as the default rowheight and
  // if there is no formatting record.
  if SameValue(rowheight, defRowHeight, ROWHEIGHT_EPS) and (not containsXF) then
    exit;

  // Otherwise: create a row record
  lRow := sheet.GetRow(WordLEToN(rowrec.RowIndex));
  lRow^.Height := rowHeight;
  if auto then
    lRow^.RowHeightType := rhtAuto else
    lRow^.RowHeightType := rhtCustom;
  lRow^.FormatIndex := XFToFormatIndex(xf);
end;

{ ------------------------------------------------------------------------------
  Reads the RPN attribute. Most attributes are not handled, but the
  data associated with it must be read to keep the stream in sync with the
  data structure.
  This version of the method is valid for BIFF2 only.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadRPNAttr(AStream: TStream; AIdentifier: Byte);
var
  nc: Byte;
begin
  case AIdentifier of
    $01: AStream.ReadByte;   // tAttrVolatile, data not used
    $02: AStream.ReadByte;   // tAttrIF, not explicitely used
    $04: begin               // tAttrChoose, not supported
           nc := AStream.ReadByte;
           AStream.Position := AStream.Position + 2*nc + 1;
         end;
    $08: AStream.ReadByte;   // tAttrSkip, data not used
    $10: AStream.ReadByte;   // tAttrSum, to be processed by ReadRPNTokenArray, byte not used
    $20: AStream.ReadByte;   // tAttrAssign, not used
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the identifier for an RPN function with fixed argument count from the
  stream.
  Valid for BIFF2-BIFF3.
-------------------------------------------------------------------------------}
function TsSpreadBIFF2Reader.ReadRPNFunc(AStream: TStream): Word;
var
  b: Byte;
begin
  b := AStream.ReadByte;
  Result := b;
end;

{@@ ----------------------------------------------------------------------------
  Reads the cell coordiantes of the top/left cell of a range using a
  shared formula.
  This cell contains the rpn token sequence of the formula.
  Is overridden because BIFF2 has 1 byte for column.
  Code is not called for shared formulas (which are not supported by BIFF2), but
  maybe for array formulas.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadRPNSharedFormulaBase(AStream: TStream;
  out ARow, ACol: Cardinal);
begin
  // 2 bytes for row of first cell in shared formula
  ARow := WordLEToN(AStream.ReadWord);
  // 1 byte for column of first cell in shared formula
  ACol := AStream.ReadByte;
end;

{@@ ----------------------------------------------------------------------------
  Helper funtion for reading of the size of the token array of an RPN formula.
  Is overridden because BIFF2 uses 1 byte only.
-------------------------------------------------------------------------------}
function TsSpreadBIFF2Reader.ReadRPNTokenArraySize(AStream: TStream): Word;
begin
  Result := AStream.ReadByte;
end;

{@@ ----------------------------------------------------------------------------
  Reads a STRING record which contains the result of string formula.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadStringRecord(AStream: TStream);
var
  len: Byte;
  s: ansistring;
begin
  // The string is a byte-string with 8 bit length
  len := AStream.ReadByte;
  if len > 0 then
  begin
    SetLength(s, Len);
    AStream.ReadBuffer(s[1], len);
    if (FIncompleteCell <> nil) and (s <> '') then
    begin
      // The "IncompleteCell" has been identified in the sheet when reading
      // the FORMULA record which precedes the String record.
//      FIncompleteCell^.UTF8StringValue := AnsiToUTF8(s);
      FIncompleteCell^.UTF8StringValue := ConvertEncoding(s, FCodePage, encodingUTF8);
      FIncompleteCell^.ContentType := cctUTF8String;
      if FIsVirtualMode then
        TsWorkbook(FWorkbook).OnReadCellData(FWorkbook,
          FIncompleteCell^.Row, FIncompleteCell^.Col, FIncompleteCell
        );
    end;
  end;
  FIncompleteCell := nil;
end;

{@@ ----------------------------------------------------------------------------
  Reads the WINDOW2 record containing information like "show grid lines",
  "show sheet headers", "panes are frozen", etc.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Reader.ReadWindow2(AStream: TStream);
begin
  // Show formulas, not results
  AStream.ReadByte;

  // Show grid lines
  if AStream.ReadByte <> 0 then
    FWorksheet.Options := FWorksheet.Options + [soShowGridLines]
  else
    FWorksheet.Options := FWorksheet.Options - [soShowGridLines];

  // Show sheet headers
  if AStream.ReadByte <> 0 then
    FWorksheet.Options := FWorksheet.Options + [soShowHeaders]
  else
    FWorksheet.Options := FWorksheet.Options - [soShowHeaders];

  // Panes are frozen
  if AStream.ReadByte <> 0 then
    FWorksheet.Options := FWorksheet.Options + [soHasFrozenPanes]
  else
    FWorksheet.Options := FWorksheet.Options - [soHasFrozenPanes];

  // Show zero values
  AStream.ReadByte;

  // Index to first visible row
  WordLEToN(AStream.ReadWord);

  // Indoex to first visible column
  WordLEToN(AStream.ReadWord);

  // Use automatic grid line color (0= manual)
  AStream.ReadByte;

  // Manual grid line line color (rgb)
  DWordToLE(AStream.ReadDWord);
end;

procedure TsSpreadBIFF2Reader.ReadXF(AStream: TStream);
var
  rec: TBIFF2_XFRecord;
  fmt: TsCellFormat;
  b: Byte;
  nf: TsNumFormatParams;
  nfs: String;
  i: Integer;
  fnt: TsFont;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;

  // Read entire xf record into buffer
  InitFormatRecord(fmt);
  fmt.ID := FCellFormatList.Count;

  rec.FontIndex := 0;  // to silence the compiler...
  AStream.ReadBuffer(rec.FontIndex, SizeOf(rec) - 2*SizeOf(word));

  // Font index
  i := rec.FontIndex;
  if i > 4 then dec(i);  // Watch out for the nasty missing font #4...
  fnt := TsFont(FFontList[i]);
  fmt.FontIndex := book.FindFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color);
  if fmt.FontIndex = -1 then
    fmt.FontIndex := book.AddFont(fnt.FontName, fnt.Size, fnt.Style, fnt.Color);
  if fmt.FontIndex > 0 then
    Include(fmt.UsedFormattingFields, uffFont);

  // Number format index
  b := rec.NumFormat_Prot and $3F;
  nfs := NumFormatList[b];
  if nfs <> '' then
  begin
    fmt.NumberFormatIndex := book.AddNumberFormat(nfs);
    nf := book.GetNumberFormat(fmt.NumberFormatIndex);
    fmt.NumberFormat := nf.NumFormat;
    fmt.NumberFormatStr := nf.NumFormatStr;
    if fmt.NumberFormat <> nfGeneral then
      Include(fmt.UsedFormattingFields, uffNumberFormat);
  end;

  // Horizontal alignment
  b := rec.HorAlign_Border_BkGr and MASK_XF_HOR_ALIGN;
  if (b <= ord(High(TsHorAlignment))) then
  begin
    fmt.HorAlignment := TsHorAlignment(b);
    if fmt.HorAlignment <> haDefault then
      Include(fmt.UsedFormattingFields, uffHorAlign);
  end;

  // Vertical alignment - not used in BIFF2
  fmt.VertAlignment := vaDefault;

  // Word wrap - not used in BIFF2
  // -- nothing to do here

  // Text rotation - not used in BIFF2
  // -- nothing to do here

  // Borders
  fmt.Border := [];
  if rec.HorAlign_Border_BkGr and $08 <> 0 then
    Include(fmt.Border, cbWest);
  if rec.HorAlign_Border_BkGr and $10 <> 0 then
    Include(fmt.Border, cbEast);
  if rec.HorAlign_Border_BkGr and $20 <> 0 then
    Include(fmt.Border, cbNorth);
  if rec.HorAlign_Border_BkGr and $40 <> 0 then
    Include(fmt.Border, cbSouth);
  if fmt.Border <> [] then
    Include(fmt.UsedFormattingFields, uffBorder);

  // Background color not supported, only shaded background
  if rec.HorAlign_Border_BkGr and $80 <> 0 then
  begin
    fmt.Background.Style := fsGray50;
    fmt.Background.FgColor := scBlack;
    fmt.Background.BgColor := scTransparent;
    Include(fmt.UsedFormattingFields, uffBackground);
  end;

  // Protection
  b := rec.NumFormat_Prot and $C0;
  case b of
    $00: fmt.Protection := [];
    $40: fmt.Protection := [cpLockCell];
    $80: fmt.Protection := [cpHideFormulas];
    $C0: fmt.Protection := [cpLockCell, cpHideFormulas];
  end;
  if fmt.Protection <> DEFAULT_CELL_PROTECTION then
    Include(fmt.UsedFormattingFields, uffProtection);

  // Add the decoded data to the format list
  FCellFormatList.Add(fmt);
end;


{------------------------------------------------------------------------------}
{                           TsSpreadBIFF2Writer                                }
{------------------------------------------------------------------------------}

constructor TsSpreadBIFF2Writer.Create(AWorkbook: TsBasicWorkbook);
begin
  inherited Create(AWorkbook);

  InitBiff2Limitations(FLimitations);

  FDateMode := Excel2Settings.DateMode;
  FCodePage := Excel2Settings.CodePage;
  FSheetIndex := Excel2Settings.SheetIndex;
end;

{@@ ----------------------------------------------------------------------------
  Adds the built-in number formats to the NumFormatList.
  Inherited method overridden for BIFF2 specialties.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.AddBuiltInNumFormats;
begin
  FFirstNumFormatIndexInFile := 0;
  InternalAddBuiltInNumFormats(FNumFormatList, Workbook.FormatSettings);
end;

function TsSpreadBIFF2Writer.FunctionSupported(AExcelCode: Integer;
  const AFuncName: String): Boolean;
begin
  Result := inherited and (AExcelCode < 200);
end;

{@@ ----------------------------------------------------------------------------
  Determines the formatting attributes of a cell, row or column. This is needed,
  for example, for writing a cell content record, such as WriteLabel,
  WriteNumber, etc.

  The attributes contain, in bit masks, xf record index, font index,
  borders, etc.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.GetAttributes(AFormatIndex: Integer;
  XFIndex: Word; out Attrib1, Attrib2, Attrib3: Byte);
var
  fmt: PsCellFormat;
  fontIdx, formatIdx: Integer;
begin
  fmt := (Workbook as TsWorkbook).GetPointerToCellFormat(AFormatIndex);

  if fmt^.UsedFormattingFields = [] then begin
    Attrib1 := 15 + MASK_XF_TYPE_PROT_LOCKED_BIFF2;  // $40
    Attrib2 := 0;
    Attrib3 := 0;
    exit;
  end;

  // 1st byte:
  //   Mask $3F: Index to XF record
  //   Mask $40: 1 = Cell is locked
  //   Mask $80: 1 = Formula is hidden
  Attrib1 := Min(XFIndex, $3F) and $3F;
  if cpLockCell in fmt^.Protection then
    Attrib1 := Attrib1 or MASK_XF_TYPE_PROT_LOCKED_BIFF2;
  if cpHideFormulas in fmt^.Protection then
    Attrib1 := Attrib1 or MASK_XF_TYPE_PROT_FORMULA_HIDDEN_BIFF2;

  // 2nd byte:
  //   Mask $3F: Index to FORMAT record ("FORMAT" = number format!)
  //   Mask $C0: Index to FONT record
  GetFormatAndFontIndex(fmt, formatIdx, fontIdx);
  Attrib2 := formatIdx + fontIdx shr 6;
//  Attrib2 := fmt^.FontIndex shr 6;

  // 3rd byte
  //   Mask $07: horizontal alignment
  //   Mask $08: Cell has left border
  //   Mask $10: Cell has right border
  //   Mask $20: Cell has top border
  //   Mask $40: Cell has bottom border
  //   Mask $80: Cell has shaded background
  Attrib3 := 0;
  if uffHorAlign in fmt^.UsedFormattingFields then
    Attrib3 := ord (fmt^.HorAlignment);
  if uffBorder in fmt^.UsedFormattingFields then begin
    if cbNorth in fmt^.Border then Attrib3 := Attrib3 or $20;
    if cbWest in fmt^.Border then Attrib3 := Attrib3 or $08;
    if cbEast in fmt^.Border then Attrib3 := Attrib3 or $10;
    if cbSouth in fmt^.Border then Attrib3 := Attrib3 or $40;
  end;
  if (uffBackground in fmt^.UsedFormattingFields) then
    Attrib3 := Attrib3 or $80;
end;

procedure TsSpreadBIFF2Writer.GetFormatAndFontIndex(AFormatRecord: PsCellFormat;
  out AFormatIndex, AFontIndex: Integer);
var
  nfparams: TsNumFormatParams;
  nfs: String;
begin
  { Index to FORMAT record }
  AFormatIndex := 0;
  if (AFormatRecord <> nil) and (uffNumberFormat in AFormatRecord^.UsedFormattingFields) then
  begin
    nfParams := TsWorkbook(FWorkbook).GetNumberFormat(AFormatRecord^.NumberFormatIndex);
    nfs := nfParams.NumFormatStr;
    AFormatIndex := NumFormatList.IndexOf(nfs);
    if AFormatIndex = -1 then AFormatIndex := 0;
  end;

  { Index to FONT record }
  AFontIndex := 0;
  if (AFormatRecord <> nil) and (uffFont in AFormatRecord^.UsedFormattingFields) then
  begin
    AFontIndex := AFormatRecord^.FontIndex;
    if AFontIndex >= 4 then inc(AFontIndex);  // Font #4 does not exist in BIFF
  end;
end;

procedure TsSpreadBIFF2Writer.PopulatePalette(AWorkbook: TsBasicWorkbook);
begin
  FPalette.Clear;
  FPalette.AddBuiltinColors(false);
  { The next instruction creates an error log entry in CheckLimitations
    if the workbook contains more colors than the default 8.
    This is because BIFF2 can only have a palette with 8 colors. }
  FPalette.CollectFromWorkbook(AWorkbook);
end;

{@@ ----------------------------------------------------------------------------
  Attaches cell formatting data for the given cell to the current record.
  Is called from all writing methods of cell contents and rows

  @param  AFormatIndex  Index into the workbook's FCellFormatList
  @param  XFIndex       Index of the XF record used here
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteCellAttributes(AStream: TStream;
  AFormatIndex: Integer; XFIndex: Word);
type
  TCellFmtRecord = packed record
    XFIndex_Locked_Hidden: Byte;
    Format_Font: Byte;
    Align_Border_BkGr: Byte;
  end;
var
  rec: TCellFmtRecord;
  fmt: PsCellFormat;
  w: Word;
begin
  fmt := TsWorkbook(FWorkbook).GetPointerToCellFormat(AFormatIndex);
  rec.XFIndex_Locked_Hidden := 0;  // to silence the compiler...
  FillChar(rec, SizeOf(rec), 0);

  if fmt^.UsedFormattingFields <> [] then
  begin
    // 1st byte:
    //   Mask $3F: Index to XF record
    //   Mask $40: 1 = Cell is locked
    //   Mask $80: 1 = Formula is hidden
    rec.XFIndex_Locked_Hidden := Min(XFIndex, $3F) and $3F;
    if cpLockCell in fmt^.Protection then
      rec.XFIndex_Locked_Hidden := rec.XFIndex_Locked_Hidden or MASK_XF_TYPE_PROT_LOCKED_BIFF2;
    if cpHideFormulas in fmt^.Protection then
      rec.XFIndex_Locked_Hidden := rec.XFIndex_Locked_Hidden or MASK_XF_TYPE_PROT_FORMULA_HIDDEN_BIFF2;

    // 2nd byte:
    //   Mask $3F: Index to FORMAT record
    //   Mask $C0: Index to FONT record
    w := fmt^.FontIndex shr 6;   // was shl --> MUST BE shr!   // ??????????????????????
    rec.Format_Font := Lo(w);

    // 3rd byte
    //   Mask $07: horizontal alignment
    //   Mask $08: Cell has left border
    //   Mask $10: Cell has right border
    //   Mask $20: Cell has top border
    //   Mask $40: Cell has bottom border
    //   Mask $80: Cell has shaded background
    if uffHorAlign in fmt^.UsedFormattingFields then
      rec.Align_Border_BkGr := ord(fmt^.HorAlignment);
    if uffBorder in fmt^.UsedFormattingFields then begin
      if cbNorth in fmt^.Border then
        rec.Align_Border_BkGr := rec.Align_Border_BkGr or $20;
      if cbWest in fmt^.Border then
        rec.Align_Border_BkGr := rec.Align_Border_BkGr or $08;
      if cbEast in fmt^.Border then
        rec.Align_Border_BkGr := rec.Align_Border_BkGr or $10;
      if cbSouth in fmt^.Border then
        rec.Align_Border_BkGr := rec.Align_Border_BkGr or $40;
    end;
    if uffBackground in fmt^.UsedFormattingFields then
      rec.Align_Border_BkGr := rec.Align_Border_BkGr or $80;
  end;
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 CODEPAGE record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteCodePage(AStream: TStream; ACodePage: String);
//  AEncoding: TsEncoding);
begin
  if ACodePage = 'cp1251' then begin
    AStream.WriteWord(WordToLE(INT_EXCEL_ID_CODEPAGE));
    AStream.WriteWord(WordToLE(2));
    AStream.WriteWord(WordToLE(WORD_CP_1258_Latin1_BIFF2_3));
    FCodePage := ACodePage;
  end else
    inherited;
              (*
  if AEncoding = seLatin1 then begin
    cp := WORD_CP_1258_Latin1_BIFF2_3;
    FCodePage := 'cp1252';

    { BIFF Record header }
    AStream.WriteWord(WordToLE(INT_EXCEL_ID_CODEPAGE));
    AStream.WriteWord(WordToLE(2));
    AStream.WriteWord(WordToLE(cp));
  end else
    inherited; *)
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 COLUMNDEFAULT record containing default column formatting of
  specified columns
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteColumnDefault(AStream: TStream;
  AFirstColIndex, ALastColIndex: Word; AFormatIndex: Integer);
//ACol: PCol);
var
  attr1, attr2, attr3: Byte;
  xf: Word;
begin
  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_COLUMNDEFAULT, 2+2+3+2);

  { Index to first column }
  AStream.WriteWord(WordToLE(AFirstColIndex));

  { Index to last column }
  AStream.WriteWord(WordToLE(ALastColIndex + 1));
  // Unlike specified in the excelfileformat.pdf, Excel 2 wants to have the
  // last column index incremented by 1!

  { Attributes }
  xf := FindXFIndex(AFormatIndex);
  GetAttributes(AFormatIndex, xf, attr1, attr2, attr3);
  AStream.WriteByte(attr1);
  AStream.WriteByte(attr2);
  AStream.WriteByte(attr3);

  { Not used }
  AStream.WriteWord(0);
end;

procedure TsSpreadBIFF2Writer.WriteColumnDefaults(AStream: TStream);
var
  j, j1: Integer;
  sheet: TsWorksheet;
  lCol, lCol1: PCol;
  lastcol: Integer;
begin
  sheet := TsWorkbook(FWorkbook).GetFirstWorksheet;
  j := 0;
  while (j < sheet.Cols.Count) do begin
    lCol := PCol(sheet.Cols[j]);
    j1 := j;
    lastcol := lCol^.Col;
    while (j1 < sheet.Cols.Count) do begin
      lCol1 := PCol(sheet.Cols[j1]);
      if lCol1^.FormatIndex <> lCol^.FormatIndex then
        break;
      lastCol := lCol1^.Col;
      inc(j1);
    end;
    WriteColumnDefault(AStream, lCol^.Col, lastCol, lCol^.FormatIndex);
    j := j1;
  end;
{
  for j := 0 to sheet.Cols.Count-1 do begin
    lCol := PCol(sheet.Cols[j]);
    if lCol^.FormatIndex > 0 then
      WriteColumnDefault(AStream, lCol);
  end;
  }
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 COLWIDTH record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteColWidth(AStream: TStream; ACol: PCol);
type
  TColRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    StartCol: Byte;
    EndCol: Byte;
    ColWidth: Word;
  end;
var
  rec: TColRecord;
  w: Single;
begin
  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_COLWIDTH);
  rec.RecordSize := WordToLE(SizeOf(TColRecord) - 4);

  { Start and end column }
  rec.StartCol := ACol^.Col;
  rec.EndCol := ACol^.Col;

  { Column width }
  { calculate width to be in units of 1/256 of pixel width of character "0" }
  if ACol^.ColWidthType = cwtDefault then
    w := TsWorksheet(FWorksheet).ReadDefaultColWidth(suChars)
  else
    w := tsWorkbook(FWorkbook).ConvertUnits(ACol^.Width, FWorkbook.Units, suChars);
  rec.ColWidth := WordToLE(round(w*256));

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Write COLWIDTH records for all columns
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteColWidths(AStream: TStream);
var
  j: Integer;
  sheet: TsWorksheet;
  col: PCol;
begin
  sheet := TsWorkbook(FWorkbook).GetFirstWorksheet;
  for j := 0 to sheet.Cols.Count-1 do begin
    col := PCol(sheet.Cols[j]);
    WriteColWidth(AStream, col);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 DIMENSIONS record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteDimensions(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  firstRow, lastRow, firstCol, lastCol: Cardinal;
  rec: TBIFF2_DimensionsRecord;
begin
  { Determine sheet size }
  GetSheetDimensions(AWorksheet, firstRow, lastRow, firstCol, lastCol);

  { Populate BIFF record }
  rec.RecordID := WordToLE(INT_EXCEL_ID_DIMENSIONS);
  rec.RecordSize := WordToLE(8);
  rec.FirstRow := WordToLE(firstRow);
  if lastRow < $FFFF then             // avoid WORD overflow when adding 1
    rec.LastRowPlus1 := WordToLE(lastRow+1)
  else
    rec.LastRowPlus1 := $FFFF;
  rec.FirstCol := WordToLE(firstCol);
  rec.LastColPlus1 := WordToLE(lastCol+1);

  { Write BIFF record to stream }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{ ------------------------------------------------------------------------------
  Writes an Excel 2 IXFE record
  This record contains the "real" XF index if it is > 62.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteIXFE(AStream: TStream; XFIndex: Word);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_IXFE));
  AStream.WriteWord(WordToLE(2));
  AStream.WriteWord(WordToLE(XFIndex));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 file to a stream

  Excel 2.x files support only one Worksheet per Workbook,
  so only the first one will be written.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteToStream(AStream: TStream;
  AParams: TsStreamParams = []);
var
  pane: Byte;
begin
  Unused(AParams);

  FWorksheet := (FWorkbook as TsWorkbook).GetWorksheetByIndex(FSheetIndex);
  if FWorksheet = nil then
    raise EFPSpreadsheetWriter.Create(rsWorksheetNotFound1);

  WriteBOF(AStream);
    WriteCodePage(AStream, FCodePage);
    WritePrintHeaders(AStream);
    WritePrintGridLines(AStream);
    WriteDefaultRowHeight(AStream, FWorksheet);
    WriteFonts(AStream);

    // Page settings block
    WriteHeaderFooter(AStream, true);    // true = header
    WriteHeaderFooter(AStream, false);   // false = footer
    WriteMargin(AStream, 0);             // 0 = left margin
    WriteMargin(AStream, 1);             // 1 = right margin
    WriteMargin(AStream, 2);             // 2 = top margin
    WriteMargin(AStream, 3);             // 3 = bottom margin

    WriteFormatCount(AStream);
    WriteNumFormats(AStream);

    if (bpLockStructure in Workbook.Protection) or FWorksheet.IsProtected then
      WritePROTECT(AStream, true);
    WriteWindowProtect(AStream, bpLockWindows in Workbook.Protection);
    WriteObjectProtect(AStream, FWorksheet);
    WritePASSWORD(AStream);

    WriteXFRecords(AStream);
    WriteDefaultColWidth(AStream, FWorksheet);
    WriteColWidths(AStream);
    WriteDimensions(AStream, FWorksheet);
    WriteColumnDefaults(AStream);
    WriteRows(AStream, FWorksheet);

    if (boVirtualMode in Workbook.Options) then
      WriteVirtualCells(AStream, FWorksheet)
    else
      WriteCellsToStream(AStream, TsWorksheet(FWorksheet).Cells);

    WriteWindow1(AStream);
    //  { -- currently not working
    WriteWindow2(AStream, FWorksheet);
    WritePane(AStream, FWorksheet, false, pane);  // false = "is not BIFF5 or BIFF8"
    WriteSelections(AStream, FWorksheet);
      //}
  WriteEOF(AStream);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 WINDOW1 record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteWindow1(AStream: TStream);
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_WINDOW1));
  AStream.WriteWord(WordToLE(9));

  { Horizontal position of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE(0));

  { Vertical position of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE($0069));

  { Width of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE($339F));

  { Height of the document window, in twips = 1 / 20 of a point }
  AStream.WriteWord(WordToLE($1B5D));

  { Window is visible (1) / hidden (0) }
  AStream.WriteByte(WordToLE(0));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 WINDOW2 record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteWindow2(AStream: TStream;
  ASheet: TsBasicWorksheet);
var
  b: Byte;
  sheet: TsWorksheet absolute ASheet;
begin
  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_WINDOW2));
  AStream.WriteWord(WordToLE(14));

  { Show formulas, not results }
  AStream.WriteByte(0);

  { Show grid lines }
  b := IfThen(soShowGridLines in sheet.Options, 1, 0);
  AStream.WriteByte(b);

  { Show sheet headers }
  b := IfThen(soShowHeaders in sheet.Options, 1, 0);
  AStream.WriteByte(b);

  { Panes are frozen? }
  b := 0;
  if (soHasFrozenPanes in sheet.Options) and
     ((sheet.LeftPaneWidth > 0) or (sheet.TopPaneHeight > 0))
  then
    b := 1;
  AStream.WriteByte(b);

  { Show zero values as zeros, not empty cells }
  AStream.WriteByte(1);

  { Index to first visible row }
  AStream.WriteWord(0);

  { Index to first visible column }
  AStream.WriteWord(0);

  { Use automatic grid line color }
  AStream.WriteByte(1);

  { RGB of manual grid line color }
  AStream.WriteDWord(0);
end;

procedure TsSpreadBIFF2Writer.WriteXF(AStream: TStream;
 AFormatRecord: PsCellFormat; XFType_Prot: Byte = 0);
var
  rec: TBIFF2_XFRecord;
  b: Byte;
  formatIdx, fontIdx: Integer;
  fmtProt: byte;
begin
  Unused(XFType_Prot);
  GetFormatAndFontIndex(AFormatRecord, formatIdx, fontIdx);

  { BIFF Record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_XF);
  rec.RecordSize := WordToLE(SizeOf(TBIFF2_XFRecord) - 2*SizeOf(word));

  { Index to FONT record }
  rec.FontIndex := WordToLE(fontIdx);

  { Not used byte }
  rec.NotUsed := 0;

  { Number format index and cell flags
      Bit   Mask  Contents
      ----- ----  --------------------------------
      5-0   $3F   Index to (number) FORMAT record
       6    $40   1 = Cell is locked
       7    $80   1 = Formula is hidden }
  fmtProt := formatIdx + MASK_XF_TYPE_PROT_LOCKED_BIFF2;
  if AFormatRecord <> nil then
  begin
    if not (cpLockCell in AFormatRecord^.Protection) then
      fmtProt := fmtProt and not MASK_XF_TYPE_PROT_LOCKED_BIFF2;
    if (cpHideFormulas in AFormatRecord^.Protection) then
      fmtProt := fmtProt or MASK_XF_TYPE_PROT_FORMULA_HIDDEN_BIFF2;
  end;
  rec.NumFormat_Prot := WordToLE(fmtProt);

  {Horizontal alignment, border style, and background
  Bit  Mask  Contents
  ---  ----  ------------------------------------------------
  2-0  $07   XF_HOR_ALIGN – Horizontal alignment (0=General, 1=Left, 2=Centered, 3=Right)
   3   $08   1 = Cell has left black border
   4   $10   1 = Cell has right black border
   5   $20   1 = Cell has top black border
   6   $40   1 = Cell has bottom black border
   7   $80   1 = Cell has shaded background }
  b := 0;
  if (AFormatRecord <> nil) then
  begin
    if (uffHorAlign in AFormatRecord^.UsedFormattingFields) then
      b := b + byte(AFormatRecord^.HorAlignment);
    if (uffBorder in AFormatRecord^.UsedFormattingFields) then
    begin
      if cbWest in AFormatRecord^.Border then b := b or $08;
      if cbEast in AFormatRecord^.Border then b := b or $10;
      if cbNorth in AFormatRecord^.Border then b := b or $20;
      if cbSouth in AFormatRecord^.Border then b := b or $40;
    end;
    if (uffBackground in AFormatRecord^.UsedFormattingFields) then
      b := b or $80;
  end;
  rec.HorAlign_Border_BkGr:= b;

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 BOF record
  This must be the first record in an Excel 2 stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteBOF(AStream: TStream);
begin
  { BIFF Record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_BOF, 4);

  { Unused }
  AStream.WriteWord($0000);

  { Data type }
  AStream.WriteWord(WordToLE(INT_EXCEL_SHEET));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 EOF record
  This must be the last record in an Excel 2 stream
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteEOF(AStream: TStream);
begin
  { BIFF Record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_EOF, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 font record
  The font data is passed as font index.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteFont(AStream: TStream; AFontIndex: Integer);
var
  Len: Byte;
  lFontName: AnsiString;
  optn: Word;
  font: TsFont;
begin
  font := TsWorkbook(FWorkbook).GetFont(AFontIndex);
  if font = nil then  // this happens for FONT4 in case of BIFF
    exit;

  if font.FontName = '' then
    raise EFPSpreadsheetWriter.Create('Font name not specified.');
  if font.Size <= 0.0 then
    raise EFPSpreadsheetWriter.Create('Font size not specified.');

  lFontName := font.FontName;
  Len := Length(lFontName);

  { BIFF Record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_FONT, 4 + 1 + Len * SizeOf(AnsiChar));

  { Height of the font in twips = 1/20 of a point }
  AStream.WriteWord(WordToLE(round(font.Size*20)));

  { Option flags }
  optn := 0;
  if fssBold in font.Style then optn := optn or $0001;
  if fssItalic in font.Style then optn := optn or $0002;
  if fssUnderline in font.Style then optn := optn or $0004;
  if fssStrikeout in font.Style then optn := optn or $0008;
  AStream.WriteWord(WordToLE(optn));

  { Font name: Unicodestring, char count in 1 byte }
  AStream.WriteByte(Len);
  AStream.WriteBuffer(lFontName[1], Len * Sizeof(AnsiChar));

  { Font color: goes into next record! }

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FONTCOLOR));
  AStream.WriteWord(WordToLE(2));

  { Font color index, only first 8 palette entries allowed! }
  AStream.WriteWord(WordToLE(PaletteIndex(font.Color)));
end;

{@@ ----------------------------------------------------------------------------
  Writes all font records to the stream
  @see WriteFont
-------------------------------------------------------------------------------}
procedure TsSpreadBiff2Writer.WriteFonts(AStream: TStream);
var
  i: Integer;
begin
  for i:=0 to TsWorkbook(FWorkbook).GetFontCount-1 do
    WriteFont(AStream, i);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 FORMAT record which describes formatting of numerical data.
-------------------------------------------------------------------------------}
procedure TsSpreadBiff2Writer.WriteFORMAT(AStream: TStream;
  ANumFormatStr: String; AFormatIndex: Integer);
type
  TNumFormatRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FormatLen: Byte;
  end;
var
  len: Integer;
  s: string; //ansistring;
  rec: TNumFormatRecord;
  buf: array of byte;
begin
  Unused(AFormatIndex);

  { Convert format string to code page used by the writer }
  s := ConvertEncoding(ANumFormatStr, encodingUTF8, FCodePage);
  len := Length(s);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_FORMAT);
  rec.RecordSize := WordToLE(1 + len);

  { Length byte of format string }
  rec.FormatLen := len;

  { Copy the format string characters into a buffer immediately after rec }
  SetLength(buf, SizeOf(rec) + SizeOf(ansiChar)*len);
  Move(rec, buf[0], SizeOf(rec));
  Move(s[1], buf[SizeOf(rec)], len*SizeOf(ansiChar));

  { Write out }
  AStream.WriteBuffer(buf[0], SizeOf(rec) + SizeOf(ansiChar)*len);

  { Clean up }
  SetLength(buf, 0);
end;

{@@ ----------------------------------------------------------------------------
  Writes the number of FORMAT records contained in the file.

  There are 21 built-in formats. The file may contain more, but Excel
  expects a "21" here...
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteFORMATCOUNT(AStream: TStream);
begin
  WriteBiffHeader(AStream, INT_EXCEL_ID_FORMATCOUNT, 2);
  AStream.WriteWord(WordToLE(21));
//  AStream.WriteWord(WordToLE(NumFormatList.Count));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 FORMULA record
  The formula is an RPN formula that was converted from usual user-readable
  string to an RPN array by the calling method.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteRPNFormula(AStream: TStream;
  const ARow, ACol: Cardinal; AFormula: TsRPNFormula; ACell: PCell);
var
  RPNLength: Word;
  RecordSizePos, FinalPos: Cardinal;
  xf: Word;
  isSupported: Boolean;
  unsupportedFormulas: String;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  { Check if formula is supported by this file format. If not, write only
    the result }
  isSupported := FormulaSupported(AFormula, unsupportedFormulas);
  if not IsSupported then
    Workbook.AddErrorMsg(rsFormulaNotSupported, [
      GetCellString(ARow, ACol), unsupportedformulas
    ]);

  RPNLength := 0;

  xf := FindXFIndex(ACell^.FormatIndex);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF Record header }
  AStream.WriteWord(WordToLE(INT_EXCEL_ID_FORMULA));
  RecordSizePos := AStream.Position;
  AStream.WriteWord(0); // We don't know the record size yet. It will be replaced at end.

  { Row and column }
  AStream.WriteWord(WordToLE(ARow));
  AStream.WriteWord(WordToLE(ACol));

  { BIFF2 Attributes }
  WriteCellAttributes(AStream, ACell^.FormatIndex, xf);

  { Encoded result of RPN formula }
  WriteRPNResult(AStream, ACell);

  { 0 = Do not recalculate
    1 = Always recalculate }
  AStream.WriteByte(1);

  { Formula data (RPN token array) }
  WriteRPNTokenArray(AStream, ACell, AFormula, false, IsSupported, RPNLength);

  { Finally write sizes after we know them }
  FinalPos := AStream.Position;
  AStream.Position := RecordSizePos;
  AStream.WriteWord(WordToLE(17 + RPNLength));
  AStream.Position := FinalPos;

  { Write following STRING record if formula result is a non-empty string }
  if (ACell^.ContentType = cctUTF8String) and (ACell^.UTF8StringValue <> '') then
    WriteSTRINGRecord(AStream, ACell^.UTF8StringValue);
end;

{@@ ----------------------------------------------------------------------------
  Writes the identifier for an RPN function with fixed argument count and
  returns the number of bytes written.
-------------------------------------------------------------------------------}
function TsSpreadBIFF2Writer.WriteRPNFunc(AStream: TStream;
  AIdentifier: Word): Word;
begin
  AStream.WriteByte(Lo(AIdentifier));
  Result := 1;
end;

{@@ ----------------------------------------------------------------------------
  Writes the size of the RPN token array. Called from WriteRPNFormula.
  Overrides xlscommon.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteRPNTokenArraySize(AStream: TStream;
  ASize: Word);
begin
  AStream.WriteByte(ASize);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 STRING record which immediately follows a FORMULA record
  when the formula result is a string.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteStringRecord(AStream: TStream;
  AString: String);
var
  s: ansistring;
  len: Integer;
begin
  s := ConvertEncoding(AString, encodingUTF8, FCodePage);
  len := Length(s);

  { BIFF Record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_STRING, 1 + len*SizeOf(ansichar));

  { Write string length }
  AStream.WriteByte(len);
  { Write characters }
  AStream.WriteBuffer(s[1], len * SizeOf(ansichar));
end;

{@@ ----------------------------------------------------------------------------
  Writes a Excel 2 BOOLEAN cell record.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteBool(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: Boolean; ACell: PCell);
var
  rec: TBIFF2_BoolErrRecord;
  xf: Integer;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  xf := FindXFIndex(ACell^.FormatIndex);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_BOOLERROR);
  rec.RecordSize := WordToLE(9);

  { Row and column index }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { BIFF2 attributes }
  GetAttributes(ACell^.FormatIndex, xf, rec.Attrib1, rec.Attrib2, rec.Attrib3);

  { Cell value }
  rec.BoolErrValue := ord(AValue);
  rec.ValueType := 0;  // 0 = boolean value, 1 = error value

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 DEFAULTROWHEIGHT record
  Specifies the default height and default flags for rows that do not have a
  corresponding ROW record
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteDefaultRowHeight(AStream: TStream;
  AWorksheet: TsBasicWorksheet);
var
  h: Single;
begin
  { BIFF record header }
  WriteBIFFHeader(AStream, INT_EXCEL_ID_DEFROWHEIGHT, 2);

  { Default height for unused rows, in twips = 1/20 of a point
    Bits 0-14: Default height for unused rows, in twips
    Bit 15 = 1: Row height not changed manually }
  h := TsWorksheet(AWorksheet).ReadDefaultRowHeight(suPoints);  // h is in points
  AStream.WriteWord(WordToLE(PtsToTwips(h)));      // write as twips
end;


{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 ERROR cell record.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteError(AStream: TStream;
  const ARow, ACol: Cardinal; const AValue: TsErrorValue; ACell: PCell);
var
  rec: TBIFF2_BoolErrRecord;
  xf: Integer;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  xf := FindXFIndex(ACell^.FormatIndex);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_BOOLERROR);
  rec.RecordSize := WordToLE(9);

  { Row and column index }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { BIFF2 attributes }
  GetAttributes(ACell^.FormatIndex, xf, rec.Attrib1, rec.Attrib2, rec.Attrib3);

  { Cell value }
  rec.BoolErrValue := ConvertToExcelError(AValue);
  rec.ValueType := 1;  // 0 = boolean value, 1 = error value

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 record for an empty cell
  Required if this cell should contain formatting, but no data.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteBlank(AStream: TStream;
  const ARow, ACol: Cardinal; ACell: PCell);
type
  TBlankRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    Attrib1, Attrib2, Attrib3: Byte;
  end;
var
  xf: Word;
  rec: TBlankRecord;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  xf := FindXFIndex(ACell^.FormatIndex);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_BLANK);
  rec.RecordSize := WordToLE(7);

  { BIFF record data }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { BIFF2 attributes }
  GetAttributes(ACell^.FormatIndex, xf, rec.Attrib1, rec.Attrib2, rec.Attrib3);

  { Write out }
  AStream.WriteBuffer(rec, Sizeof(rec));
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 LABEL record
  If the string length exceeds 255 bytes, the string will be truncated and an
  error message will be logged.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteLabel(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: string; ACell: PCell);
const
  MAXBYTES = 255; //limit for this format
var
  L: Byte;
  AnsiText: ansistring;
  rec: TBIFF2_LabelRecord;
  buf: array of byte;
var
  xf: Word;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  if AValue = '' then Exit; // Writing an empty text doesn't work

  AnsiText := UTF8ToISO_8859_1(FixLineEnding(AValue));

  if Length(AnsiText) > MAXBYTES then begin
    // BIFF 5 does not support labels/text bigger than 255 chars,
    // so BIFF2 won't either
    // Rather than lose data when reading it, let the application programmer deal
    // with the problem or purposefully ignore it.
    AnsiText := Copy(AnsiText, 1, MAXBYTES);
    Workbook.AddErrorMsg(rsTruncateTooLongCellText, [
      MAXBYTES, GetCellString(ARow, ACol)
    ]);
  end;
  L := Length(AnsiText);

  xf := FindXFIndex(ACell^.FormatIndex);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_LABEL);
  rec.RecordSize := WordToLE(8 + L);

  { BIFF record data }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { BIFF2 attributes }
  GetAttributes(ACell^.FormatIndex, xf, rec.Attrib1, rec.Attrib2, rec.Attrib3);

  { Text length: 8 bit }
  rec.TextLen := L;

  { Copy the text characters into a buffer immediately after rec }
  SetLength(buf, SizeOf(rec) + SizeOf(ansiChar)*L);
  Move(rec, buf[0], SizeOf(rec));
  Move(AnsiText[1], buf[SizeOf(rec)], L*SizeOf(ansiChar));

  { Write out }
  AStream.WriteBuffer(buf[0], SizeOf(Rec) + SizeOf(ansiChar)*L);
end;

{@@ ----------------------------------------------------------------------------
  Writes an Excel 2 NUMBER record
  A "number" is a 64-bit IEE 754 floating point.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF2Writer.WriteNumber(AStream: TStream; const ARow,
  ACol: Cardinal; const AValue: double; ACell: PCell);
var
  xf: Word;
  rec: TBIFF2_NumberRecord;
begin
  if (ARow >= FLimitations.MaxRowCount) or (ACol >= FLimitations.MaxColCount) then
    exit;

  xf := FindXFIndex(ACell^.FormatIndex);
  if xf >= 63 then
    WriteIXFE(AStream, xf);

  { BIFF record header }
  rec.RecordID := WordToLE(INT_EXCEL_ID_NUMBER);
  rec.RecordSize := WordToLE(15);

  { BIFF record data }
  rec.Row := WordToLE(ARow);
  rec.Col := WordToLE(ACol);

  { BIFF2 attributes }
  GetAttributes(ACell^.FormatIndex, xf, rec.Attrib1, rec.Attrib2, rec.Attrib3);

  { Number value }
  rec.Value := AValue;

  { Write out }
  AStream.WriteBuffer(rec, SizeOf(Rec));
end;

procedure TsSpreadBIFF2Writer.WritePassword(AStream: TStream);
var
  hash: Word;
  hb, hs: LongInt;
  book: TsWorkbook;
  sheet: TsWorksheet;
begin
  book := FWorkbook as TsWorkbook;
  sheet := FWorksheet as TsWorksheet;

  hb := 0;
  if (book.CryptoInfo.PasswordHash <> '') and
     not TryStrToInt('$' + book.CryptoInfo.PasswordHash, hb) then
  begin
    book.AddErrorMsg(rsPasswordRemoved_NotValid);
    exit;
  end;

  hs := 0;
  if (sheet.CryptoInfo.PasswordHash <> '') and
     not TryStrToInt('$' + sheet.CryptoInfo.PasswordHash, hs) then
  begin
    book.AddErrorMsg(rsPasswordRemoved_NotValid);
    exit;
  end;

  // Neither workbook nor worksheet password set
  if (hb = 0) and (hs = 0) then
    exit;

  // Only workbook password set. Check for Excel algorithm.
  if (hb <> 0) and (hs = 0) then begin
    if book.CryptoInfo.Algorithm <> caExcel then begin
      book.AddErrorMsg(rsPasswordRemoved_Excel);
      exit;
    end;
    hash := hb;
  end else
  // Only worksheet password set, check for Excel algorithm
  if (hs <> 0) and (hb = 0) then begin
    if sheet.CryptoInfo.Algorithm <> caExcel then begin
      book.AddErrorMsg(rsPasswordRemoved_Excel);
      exit;
    end;
    hash := hs;
  end else
  if (hs <> hb) then begin
    book.AddErrorMsg(rsPasswordRemoved_BIFF2);
    exit;
  end else
  if (book.CryptoInfo.Algorithm <> caExcel) or
     (sheet.CryptoInfo.Algorithm <> caExcel) then
  begin
    book.AddErrorMsg(rsPasswordRemoved_Excel);
    exit;
  end else
    hash := hs;  // or hb -- they are equal here.

  // Write out record
  WriteBIFFHeader(AStream, INT_EXCEL_ID_PASSWORD, 2);
  AStream.WriteWord(WordToLE(hash));
end;

procedure TsSpreadBIFF2Writer.WriteRow(AStream: TStream; ASheet: TsBasicWorksheet;
  ARowIndex, AFirstColIndex, ALastColIndex: Cardinal; ARow: PRow);
var
  containsXF: Boolean;
  rowheight: Word;
  auto: Boolean;
  w: Word;
  xf: Word;
  book: TsWorkbook;
  sheet: TsWorksheet;
begin
  if (ARowIndex >= FLimitations.MaxRowCount) or
     (AFirstColIndex >= FLimitations.MaxColCount) or
     (ALastColIndex >= FLimitations.MaxColCount)
  then
    exit;

  book := FWorkbook as TsWorkbook;
  sheet := ASheet as TsWorksheet;

  containsXF := (ARow <> nil) and (ARow^.FormatIndex > 0);

  { BIFF record header }
  WriteBiffHeader(AStream, INT_EXCEL_ID_ROW, IfThen(containsXF, 18, 13));

  { Index of row }
  AStream.WriteWord(WordToLE(Word(ARowIndex)));

  { Index to column of the first cell which is described by a cell record }
  AStream.WriteWord(WordToLE(Word(AFirstColIndex)));

  { Index to column of the last cell which is described by a cell record, increased by 1 }
  AStream.WriteWord(WordToLE(Word(ALastColIndex) + 1));

  auto := true;
  { Row height (in twips, 1/20 point) and info on custom row height }
  if (ARow = nil) or (ARow^.RowHeightType = rhtDefault) then
    rowheight := PtsToTwips(sheet.ReadDefaultRowHeight(suPoints))
  else
  if (ARow^.Height = 0) then
    rowheight := 0
  else begin
    rowheight := PtsToTwips(book.ConvertUnits(ARow^.Height, book.Units, suPoints));
    auto := ARow^.RowHeightType <> rhtCustom;
  end;
  w := rowheight and $7FFF;
  if auto then
    w := w or $8000;
  AStream.WriteWord(WordToLE(w));

  { not used }
  AStream.WriteWord(0);

  { Does the record contain row attribute field and XF index? }
  AStream.WriteByte(ord(containsXF));

  { Relative offset to calculate stream position of the first cell record for this row }
  AStream.WriteWord(0);

  if containsXF then begin
    xf := FindXFIndex(ARow^.FormatIndex);

  { Default row attributes }
    WriteCellAttributes(AStream, ARow^.FormatIndex, xf);

    { Index to XF record }
    AStream.WriteWord(WordToLE(xf));
  end;
end;

{*******************************************************************
*  Initialization section
*
*  Registers this reader / writer to fpspreadsheet
*  Converts the palette to litte-endian
*
*******************************************************************}

initialization

 {$IFDEF MSWINDOWS}
  Excel2Settings.CodePage := GetDefaultTextEncoding;
 {$ENDIF}

  sfidExcel2 := RegisterSpreadFormat(sfExcel2,
    TsSpreadBIFF2Reader, TsSpreadBIFF2Writer,
    STR_FILEFORMAT_EXCEL_2, 'BIFF2', [STR_EXCEL_EXTENSION]
  );

  MakeLEPalette(PALETTE_BIFF2);

end.
