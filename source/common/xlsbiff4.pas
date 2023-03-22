unit xlsBIFF4;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils,
  fpsTypes, fpsUtils, xlsCommon;

type

  { TsSpreadBIFF4Reader }

  TsSpreadBIFF4Reader = class(TsSpreadBIFFReader)
  protected
    procedure AddBuiltInNumFormats; override;
    procedure PopulatePalette; override;
    procedure ReadDEFINEDNAME(AStream: TStream);
    procedure ReadFONT(const AStream: TStream);
    procedure ReadFORMAT(AStream: TStream); override;
    procedure ReadFORMULA(AStream: TStream); override;
    procedure ReadLABEL(AStream: TStream); override;
    procedure ReadStandardWidth(AStream: TStream; ASheet: TsBasicWorksheet);
    procedure ReadStringRecord(AStream: TStream); override;
    procedure ReadXF(AStream: TStream);
  public
    constructor Create(AWorkbook: TsBasicWorkbook); override;
    { General reading methods }
    procedure ReadFromStream(AStream: TStream; APassword: String = '';
      AParams: TsStreamParams = []); override;
    { File format detection }
    class function CheckfileFormat(AStream: TStream): Boolean; override;
  end;

var
  sfidExcel4: TsSpreadFormatID;

const
  {@@ palette of the default BIFF4 colors as "big-endian color" values }
  PALETTE_BIFF4: array[$00..$17] of TsColor = (
    $000000,  // $00: black
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
    $808080   // $17: EGA gray
  );

implementation

uses
  LConvEncoding, Math,
  fpSpreadsheet, fpsStrings, fpsReaderWriter, fpsPalette, fpsNumFormat;

const
  BIFF4_MAX_PALETTE_SIZE = 8 + 16;
  SYS_DEFAULT_FOREGROUND_COLOR  = $18;
  SYS_DEFAULT_BACKGROUND_COLOR  = $19;

  { Excel record IDs }
  INT_EXCEL_ID_BLANK         = $0201;
  INT_EXCEL_ID_NUMBER        = $0203;
  INT_EXCEL_ID_LABEL         = $0204;
  INT_EXCEL_ID_BOOLERROR     = $0205;
  INT_EXCEL_ID_BOF           = $0409;
  INT_EXCEL_ID_FONT          = $0231;
  INT_EXCEL_ID_FORMULA       = $0406;
  INT_EXCEL_ID_STANDARDWIDTH = $0099;
  INT_EXCEL_ID_XF            = $0443;

  // XF Text orientation
  MASK_XF_ORIENTATION     = $C0;
  XF_ROTATION_HORIZONTAL  = 0;
  XF_ROTATION_STACKED     = 1;
  XF_ROTATION_90DEG_CCW   = 2;
  XF_ROTATION_90DEG_CW    = 3;

  // XF cell background
  MASK_XF_BKGR_FILLPATTERN      = $003F;
  MASK_XF_BKGR_PATTERN_COLOR    = $07C0;
  MASK_XF_BKGR_BACKGROUND_COLOR = $F800;

  // XF cell border
  MASK_XF_BORDER_TOP_STYLE    = $00000007;
  MASK_XF_BORDER_TOP_COLOR    = $000000F8;        // shr 3
  MASK_XF_BORDER_LEFT_STYLE   = $00000700;        // shr 8
  MASK_XF_BORDER_LEFT_COLOR   = $0000F800;        // shr 11
  MASK_XF_BORDER_BOTTOM_STYLE = $00070000;        // shr 16
  MASK_XF_BORDER_BOTTOM_COLOR = $00F80000;        // shr 19
  MASK_XF_BORDER_RIGHT_STYLE  = $07000000;        // shr 24
  MASK_XF_BORDER_RIGHT_COLOR  = $F8000000;        // shr 27

type
  TBIFF4_LabelRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    Row: Word;
    Col: Word;
    XFIndex: Word;
    TextLen: Word;
  end;

  TBIFF4_XFRecord = packed record
    RecordID: Word;
    RecordSize: Word;
    FontIndex: byte;
    NumFormatIndex: byte;
    XFType_Prot_ParentXF: Word;
    Align_TextBreak_Orientation: Byte;
    UsedAttribGroups: Byte;
    BackGround: Word;
    Border: DWord;
  end;

procedure InitBiff4Limitations(out ALimitations: TsSpreadsheetFormatLimitations);
begin
  InitBiffLimitations(ALimitations);
  ALimitations.MaxPaletteSize := BIFF4_MAX_PALETTE_SIZE;
end;

{ ------------------------------------------------------------------------------
                         TsSpreadBIFF4Reader
-------------------------------------------------------------------------------}
constructor TsSpreadBIFF4Reader.Create(AWorkbook: TsBasicWorkbook);
begin
  inherited Create(AWorkbook);
  InitBiff4Limitations(FLimitations);
end;

procedure TsSpreadBIFF4Reader.AddBuiltInNumFormats;
begin
  FFirstNumFormatIndexInFile := 0;
end;

{@@ ----------------------------------------------------------------------------
  Checks the header of the stream for the signature of BIFF2 files
-------------------------------------------------------------------------------}
class function TsSpreadBIFF4Reader.CheckFileFormat(AStream: TStream): Boolean;
const
  BIFF4_HEADER: packed array[0..1] of byte = (
    $09, $04);
var
  P: Int64;
  buf: packed array[0..1] of byte = (0, 0);
  n: Integer;
begin
  Result := false;
  P := AStream.Position;
  try
    AStream.Position := 0;
    n := AStream.Read(buf, SizeOf(buf));
    if n < Length(BIFF4_HEADER) then
      exit;
    Result := CompareMem(@buf[0], @BIFF4_HEADER, 2);
  finally
    AStream.Position := P;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Populates the reader's default palette using the BIFF4 default colors.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF4Reader.PopulatePalette;
begin
  FPalette.Clear;
  FPalette.UseColors(PALETTE_BIFF4, true);
end;

{@@ ----------------------------------------------------------------------------
  Reads a DEFINEDNAME record. Currently only extracts print ranges and titles.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF4Reader.ReadDEFINEDNAME(AStream: TStream);
{
var
  options: Word;
  len: byte;
  formulaSize: Word;
  ansistr: ansiString = '';
  defName: String;
  rpnformula: TsRPNFormula;
  {%H-}extsheetIndex: Integer;
  sheetIndex: Integer;
  }
begin
  (*     FIX ME  --  this is the code of BIFF5. Must be adapted.

  // Options
  options := WordLEToN(AStream.ReadWord);
  if options and $0020 = 0 then   // only support built-in names at the moment!
    exit;

  // Keyboard shortcut  --> ignore
  AStream.ReadByte;

  // Length of name (character count)
  len := AStream.ReadByte;

  // Size of formula data
  formulasize := WordLEToN(AStream.ReadWord);

  // EXTERNSHEET index (1-base), or 0 if global name
  extsheetIndex := SmallInt(WordLEToN(AStream.ReadWord)) - 1;  // now 0-based!

  // Sheet index (1-based) on which the name is valid (0 = global)
  sheetIndex := SmallInt(WordLEToN(AStream.ReadWord)) - 1;  // now 0-based!

  // Length of Menu text (ignore)
  AStream.ReadByte;

  // Length of description text(ignore)
  AStream.ReadByte;

  // Length of help topic text (ignore)
  AStream.ReadByte;

  // Length of status bar text (ignore)
  AStream.ReadByte;

  // Name
  SetLength(ansistr, len);
  AStream.ReadBuffer(ansistr[1], len);
  defName := ConvertEncoding(ansistr, FCodepage, encodingUTF8);

  // Formula
  if not ReadRPNTokenArray(AStream, formulaSize, rpnFormula) then
    exit;
  // Store defined name in internal list
  FDefinedNames.Add(TsBIFFDefinedName.Create(defName, rpnFormula, sheetIndex));

  // Skip rest...
  *)
end;

procedure TsSpreadBIFF4Reader.ReadFont(const AStream: TStream);
var
  {%H-}lCodePage: Word;
  lHeight: Word;
  lOptions: Word;
  lColor: Word;
  lWeight: Word;
  lEsc: Word;
  Len: Byte;
  fontname: ansistring = '';
  font: TsFont;
  isDefaultFont: Boolean;
begin
  font := TsFont.Create;

  { Height of the font in twips = 1/20 of a point }
  lHeight := WordLEToN(AStream.ReadWord); // WordToLE(200)
  font.Size := lHeight/20;

  { Option flags }
  lOptions := WordLEToN(AStream.ReadWord);
  font.Style := [];
  if lOptions and $0001 <> 0 then Include(font.Style, fssBold);
  if lOptions and $0002 <> 0 then Include(font.Style, fssItalic);
  if lOptions and $0004 <> 0 then Include(font.Style, fssUnderline);
  if lOptions and $0008 <> 0 then Include(font.Style, fssStrikeout);

  { Color index }
  // The problem is that the palette is loaded after the font list; therefore
  // we do not know the rgb color of the font here. We store the palette index
  // ("SetAsPaletteIndex") and replace it by the rgb color after reading of the
  // palette and after reading the workbook globals records. As an indicator
  // that the font does not yet contain an rgb color a control bit is set in
  // the high-byte of the TsColor.
  lColor := WordLEToN(AStream.ReadWord);
  if lColor < 8 then
    // Use built-in colors directly otherwise the Workbook's FindFont would not find the font in ReadXF
    font.Color := FPalette[lColor]
  else
  if lColor = SYS_DEFAULT_WINDOW_TEXT_COLOR then
    font.Color := scBlack
  else
    font.Color := SetAsPaletteIndex(lColor);

  { Font name: Ansistring, char count in 1 byte }
  Len := AStream.ReadByte();
  SetLength(fontname, Len);
  AStream.ReadBuffer(fontname[1], Len);
  font.FontName := ConvertEncoding(fontname, FCodePage, encodingUTF8);

  isDefaultFont := FFontList.Count = 0;

  { Add font to internal font list. Will be copied to workbook's font list later
    as the font index in the internal list may be different from the index in
    the workbook's list. }
  FFontList.Add(font);

  { Excel does not have zero-based font #4! }
  if FFontList.Count = 4 then FFontList.Add(nil);

  if isDefaultFont then
    (FWorkbook as TsWorkbook).SetDefaultFont(font.FontName, font.Size);
end;

// Read the FORMAT record for formatting numerical data
procedure TsSpreadBIFF4Reader.ReadFormat(AStream: TStream);
var
  len: byte;
  fmtString: AnsiString = '';
  nfs: String;
begin
  // Record FORMAT, BIFF4 (5.49):
  // Offset Size Contents
  // 0      2     BIFF4 not used
  // 2      var   Number format string (byte string, 8-bit string length)

  // not used
  AStream.ReadWord;

  // number format string
  len := AStream.ReadByte;
  SetLength(fmtString, len);
  AStream.ReadBuffer(fmtString[1], len);

  // We need the format string as utf8 and non-localized
  nfs := ConvertEncoding(fmtString, FCodePage, encodingUTF8);

  // Add to the end of the list.
  NumFormatList.Add(nfs);
end;

{@@ ----------------------------------------------------------------------------
  Reads a FORMULA record, retrieves the RPN formula and puts the result in the
  corresponding field. The formula is not recalculated here!
  Valid for BIFF4.
-------------------------------------------------------------------------------}
procedure TsSpreadBIFF4Reader.ReadFormula(AStream: TStream);
var
  ARow, ACol: Cardinal;
  XF: WORD;
  ResultFormula: Double = 0.0;
  Data: array [0..7] of byte;
  dt: TDateTime;
  nf: TsNumberFormat;
  nfs: String;
  err: TsErrorValue;
  ok: Boolean;
  cell: PCell;
  sheet: TsWorksheet;
  msg: String;
begin
  sheet := TsWorksheet(FWorksheet);

  { Index to XF Record }
  ReadRowColXF(AStream, ARow, ACol, XF);

  { Result of the formula result in IEEE 754 floating-point value }
  Data[0] := 0;  // to silence the compiler...
  AStream.ReadBuffer(Data, Sizeof(Data));

  { Options flags }
  WordLEtoN(AStream.ReadWord);

  { Create cell }
  if FIsVirtualMode then                       // "Virtual" cell
  begin
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := sheet.GetCell(ARow, ACol);    // "Real" cell
    // Don't call "AddCell" because, if the cell belongs to a shared formula, it
    // already has been created before, and then would exist in the tree twice.

  // Prevent shared formulas (which already may have been written at this time)
  // being erased when cell content is written
  TsWorkbook(sheet.Workbook).LockFormulas;
  try
    // Now determine the type of the formula result
    if (Data[6] = $FF) and (Data[7] = $FF) then
      case Data[0] of
        0: // String -> Value is found in next record (STRING)
           FIncompleteCell := cell;

        1: // Boolean value
           sheet.WriteBoolValue(cell, Data[2] = 1);

        2: begin  // Error value
             err := ConvertFromExcelError(Data[2]);
             sheet.WriteErrorValue(cell, err);
           end;

        3: sheet.WriteBlank(cell);
      end
    else
    begin
      // Result is a number or a date/time
      Move(Data[0], ResultFormula, SizeOf(Data));

      {Find out what cell type, set content type and value}
      ExtractNumberFormat(XF, nf, nfs);
      if IsDateTime(ResultFormula, nf, nfs, dt) then
        sheet.WriteDateTime(cell, dt) //, nf, nfs)
      else
        sheet.WriteNumber(cell, ResultFormula);
    end;
  finally
    TsWorkbook(sheet.Workbook).UnlockFormulas;
  end;

  { Formula token array }
  if boReadFormulas in FWorkbook.Options then
  begin
    ok := ReadRPNTokenArray(AStream, cell);
    if not ok then
    begin
      msg := Format(rsFormulaNotSupported, [
        GetCellString(ARow, ACol), '.xls'
      ]);
      if (boAbortReadOnFormulaError in Workbook.Options) then
        raise Exception.Create(msg)
      else begin
        sheet.WriteErrorValue(cell, errFormulaNotSupported);
        FWorkbook.AddErrorMsg(msg);
      end;
    end;
  end;

  {Add attributes}
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode and (cell <> FIncompleteCell) then
    (FWorkbook as TsWorkbook).OnReadCellData(Workbook, ARow, ACol, cell);
end;

procedure TsSpreadBIFF4Reader.ReadFromStream(AStream: TStream;
  APassword: String = ''; AParams: TsStreamParams = []);
var
  BIFF4EOF: Boolean;
  RecordType: Word;
  CurStreamPos: Int64;
  BOFFound: Boolean;
begin
  Unused(APassword, AParams);
  BIFF4EOF := False;

  { In BIFF2 files there is only one worksheet, let's create it }
  FWorksheet := TsWorkbook(FWorkbook).AddWorksheet('Sheet', true);

  { Read all records in a loop }
  BOFFound := false;
  while not BIFF4EOF do
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
      INT_EXCEL_ID_COLINFO       : ReadColInfo(AStream);
      INT_EXCEL_ID_DATEMODE      : ReadDateMode(AStream);
      INT_EXCEL_ID_DEFCOLWIDTH   : ReadDefColWidth(AStream);
      INT_EXCEL_ID_DEFINEDNAME   : ReadDefinedName(AStream);
      INT_EXCEL_ID_DEFROWHEIGHT  : ReadDefRowHeight(AStream);
      INT_EXCEL_ID_EOF           : BIFF4EOF := True;
      INT_EXCEL_ID_EXTERNCOUNT   : ReadEXTERNCOUNT(AStream, FWorksheet);
      INT_EXCEL_ID_EXTERNSHEET   : ReadEXTERNSHEET(AStream, FWorksheet);
      INT_EXCEL_ID_FONT          : ReadFont(AStream);
      INT_EXCEL_ID_FOOTER        : ReadHeaderFooter(AStream, false);
      INT_EXCEL_ID_FORMAT        : ReadFormat(AStream);
      INT_EXCEL_ID_FORMULA       : ReadFormula(AStream);
      INT_EXCEL_ID_HEADER        : ReadHeaderFooter(AStream, true);
      INT_EXCEL_ID_HCENTER       : ReadHCENTER(AStream);
      INT_EXCEL_ID_HORZPAGEBREAK : ReadHorizontalPageBreaks(AStream, FWorksheet);
      INT_EXCEL_ID_LABEL         : ReadLabel(AStream);
      INT_EXCEL_ID_LEFTMARGIN    : ReadMargin(AStream, 0);
      INT_EXCEL_ID_NOTE          : ReadComment(AStream);
      INT_EXCEL_ID_NUMBER        : ReadNumber(AStream);
      INT_EXCEL_ID_OBJECTPROTECT : ReadObjectProtect(AStream);
      INT_EXCEL_ID_PAGESETUP     : ReadPageSetup(AStream);
      INT_EXCEL_ID_PALETTE       : ReadPALETTE(AStream);
      INT_EXCEL_ID_PANE          : ReadPane(AStream);
      INT_EXCEL_ID_PASSWORD      : ReadPASSWORD(AStream);
      INT_EXCEL_ID_PRINTGRID     : ReadPrintGridLines(AStream);
      INT_EXCEL_ID_PRINTHEADERS  : ReadPrintHeaders(AStream);
      INT_EXCEL_ID_PROTECT       : ReadPROTECT(AStream);
      INT_EXCEL_ID_RIGHTMARGIN   : ReadMargin(AStream, 1);
      INT_EXCEL_ID_RK            : ReadRKValue(AStream); //(RK) This record represents a cell that contains an RK value (encoded integer or floating-point value). If a floating-point value cannot be encoded to an RK value, a NUMBER record will be written. This record replaces the record INTEGER written in BIFF2.
      INT_EXCEL_ID_ROW           : ReadRowInfo(AStream);
      INT_EXCEL_ID_SCL           : ReadSCLRecord(AStream);
      INT_EXCEL_ID_SELECTION     : ReadSELECTION(AStream);
      INT_EXCEL_ID_SHEETPR       : ReadSHEETPR(AStream);
      INT_EXCEL_ID_STANDARDWIDTH : ReadStandardWidth(AStream, FWorksheet);
      INT_EXCEL_ID_STRING        : ReadStringRecord(AStream);
      INT_EXCEL_ID_TOPMARGIN     : ReadMargin(AStream, 2);
      INT_EXCEL_ID_VCENTER       : ReadVCENTER(AStream);
      INT_EXCEL_ID_VERTPAGEBREAK : ReadVerticalPageBreaks(AStream, FWorksheet);
      INT_EXCEL_ID_WINDOW2       : ReadWindow2(AStream);
      INT_EXCEL_ID_WINDOWPROTECT : ReadWindowProtect(AStream);
      INT_EXCEL_ID_XF            : ReadXF(AStream);
    else
      // nothing
    end;

    // Make sure we are in the right position for the next record
    AStream.Seek(CurStreamPos + RecordSize, soFromBeginning);

    if AStream.Position >= AStream.Size then
      BIFF4EOF := True;

    if not BOFFound then
      raise EFPSpreadsheetReader.Create('BOF record not found.');
  end;

  // Convert palette indexes to rgb colors
  FixColors;

  FixCols(FWorksheet);
  FixRows(FWorksheet);
end;

procedure TsSpreadBIFF4Reader.ReadLabel(AStream: TStream);
var
  rec: TBIFF4_LabelRecord;
  L: Word;
  ARow, ACol: Cardinal;
  XF: WORD;
  cell: PCell;
  ansistr: ansistring = '';
  valuestr: String;
begin
  rec.Row := 0;  // to silence the compiler...

  { Read entire record, starting at Row, except for string data }
  AStream.ReadBuffer(rec.Row, SizeOf(TBIFF4_LabelRecord) - 2*SizeOf(Word));
  ARow := WordLEToN(rec.Row);
  ACol := WordLEToN(rec.Col);
  XF := WordLEToN(rec.XFIndex);

  { Byte String with 16-bit size }
  L := WordLEToN(rec.TextLen);
  SetLength(ansistr, L);
  AStream.ReadBuffer(ansistr[1], L);

  { Create cell }
  if FIsVirtualMode then begin
    InitCell(FWorksheet, ARow, ACol, FVirtualCell);
    cell := @FVirtualCell;
  end else
    cell := (FWorksheet as TsWorksheet).AddCell(ARow, ACol);

  { Save the data }
  valueStr := ConvertEncoding(ansistr, FCodePage, encodingUTF8);
  (FWorksheet as TsWorksheet).WriteText(cell, valueStr); //ISO_8859_1ToUTF8(ansistr));

  { Add attributes }
  ApplyCellFormatting(cell, XF);

  if FIsVirtualMode then
    (Workbook as TsWorkbook).OnReadCellData(Workbook, ARow, ACol, cell);
end;

{ Reads the default column width that is used when a bit in the GCW bit structure
  is set for the corresponding column. The GCW is ignored here. The column
  width read from the STANDARDWIDTH record overrides the one from the
  DEFCOLWIDTH record. }
procedure TsSpreadBIFF4Reader.ReadStandardWidth(AStream: TStream;
  ASheet: TsBasicWorksheet);
var
  w: Word;
begin
  // read width in 1/256 of the width of "0" character
  w := WordLEToN(AStream.ReadWord);
  (ASheet as TsWorksheet).WriteDefaultRowHeight(w / 256, suChars);
end;

{ Reads a STRING record which contains the result of string formula. }
procedure TsSpreadBIFF4Reader.ReadStringRecord(AStream: TStream);
var
  len: Word;
  s: ansistring = '';
begin
  // The string is a byte-string with 16 bit length
  len := WordLEToN(AStream.ReadWord);
  if len > 0 then begin
    SetLength(s, Len);
    AStream.ReadBuffer(s[1], len);
    if (FIncompleteCell <> nil) and (s <> '') then begin
      FIncompletecell^.UTF8StringValue := ConvertEncoding(s, FCodePage, encodingUTF8);
      FIncompleteCell^.ContentType := cctUTF8String;
      if FIsVirtualMode then
        (Workbook as TsWorkbook).OnReadCellData(
          Workbook, FIncompleteCell^.Row, FIncompleteCell^.Col, FIncompleteCell
        );
    end;
  end;
  FIncompleteCell := nil;
end;

procedure TsSpreadBIFF4Reader.ReadXF(AStream: TStream);
var
  rec: TBIFF4_XFRecord;
  fmt: TsCellFormat;
  cidx: Integer;
  nfparams: TsNumFormatParams;
  nfs: String;
  b: Byte;
  dw: DWord;
  fill: Word;
  fs: TsFillStyle;
  book: TsWorkbook;
begin
  book := FWorkbook as TsWorkbook;

  InitFormatRecord(fmt);
  fmt.ID := FCellFormatList.Count;

  // Read the complete XF record into a buffer
  rec.FontIndex := 0;  // to silence the compiler...
  AStream.ReadBuffer(rec.FontIndex, SizeOf(rec) - 2*SizeOf(Word));

  // Font index
  fmt.FontIndex := FixFontIndex(rec.FontIndex);
  if fmt.FontIndex > 1 then
    Include(fmt.UsedFormattingFields, uffFont);

  // Number format index
  if rec.NumFormatIndex <> 0 then begin
    nfs := NumFormatList[rec.NumFormatIndex];
    // "General" (NumFormatIndex = 0) not stored in workbook's NumFormatList
    if (rec.NumFormatIndex > 0) and not SameText(nfs, 'General') then
    begin
      fmt.NumberFormatIndex := book.AddNumberFormat(nfs);
      nfParams := book.GetNumberFormat(fmt.NumberFormatIndex);
      fmt.NumberFormat := nfParams.NumFormat;
      fmt.NumberFormatStr := nfs;
      Include(fmt.UsedFormattingFields, uffNumberFormat);
    end;
  end;

  // Horizontal text alignment
  b := rec.Align_TextBreak_Orientation AND MASK_XF_HOR_ALIGN;
  if (b <= ord(High(TsHorAlignment))) then
  begin
    fmt.HorAlignment := TsHorAlignment(b);
    if fmt.HorAlignment <> haDefault then
      Include(fmt.UsedFormattingFields, uffHorAlign);
  end;

  // Vertical text alignment
  b := (rec.Align_TextBreak_Orientation AND MASK_XF_VERT_ALIGN) shr 4;
  if (b + 1 <= ord(high(TsVertAlignment))) then
  begin
    fmt.VertAlignment := TsVertAlignment(b + 1);      // + 1 due to vaDefault
    // Unfortunately BIFF does not provide a "default" vertical alignment code.
    // Without the following correction "non-formatted" cells would always have
    // the uffVertAlign FormattingField set which contradicts the statement of
    // not being formatted.
    if fmt.VertAlignment = vaBottom then
      fmt.VertAlignment := vaDefault;
    if fmt.VertAlignment <> vaDefault then
      Include(fmt.UsedFormattingFields, uffVertAlign);
  end;

  // Word wrap
  if (rec.Align_TextBreak_Orientation and MASK_XF_TEXTWRAP) <> 0 then
    Include(fmt.UsedFormattingFields, uffWordwrap);

  // Text rotation
  case (rec.Align_TextBreak_Orientation and MASK_XF_ORIENTATION) shr 6 of
    XF_ROTATION_HORIZONTAL : fmt.TextRotation := trHorizontal;
    XF_ROTATION_90DEG_CCW  : fmt.TextRotation := rt90DegreeCounterClockwiseRotation;
    XF_ROTATION_90DEG_CW   : fmt.TextRotation := rt90DegreeClockwiseRotation;
    XF_ROTATION_STACKED    : fmt.TextRotation := rtStacked;
  end;
  if fmt.TextRotation <> trHorizontal then
    Include(fmt.UsedFormattingFields, uffTextRotation);

  // Cell borders and background
  rec.Background := WordLEToN(rec.Background);
  rec.Border := DWordLEToN(rec.Border);
  // The 4 masked bits encode the line style of the border line. 0 = no line.
  // The case of "no line" is not included in the TsLineStyle enumeration.
  // --> correct by subtracting 1!
  dw := rec.Border and MASK_XF_BORDER_BOTTOM_STYLE;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbSouth);
    fmt.BorderStyles[cbSouth].LineStyle := TsLineStyle(dw shr 16 - 1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border and MASK_XF_BORDER_LEFT_STYLE;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbWest);
    fmt.BorderStyles[cbWest].LineStyle := TsLineStyle(dw shr 8 - 1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border and MASK_XF_BORDER_RIGHT_STYLE;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbEast);
    fmt.BorderStyles[cbEast].LineStyle := TsLineStyle(dw shr 24 - 1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;
  dw := rec.Border and MASK_XF_BORDER_TOP_STYLE;
  if dw <> 0 then
  begin
    Include(fmt.Border, cbNorth);
    fmt.BorderStyles[cbNorth].LineStyle := TsLineStyle(dw - 1);
    Include(fmt.UsedFormattingFields, uffBorder);
  end;

  // Border line colors
  // NOTE: It is possible that the palette is not yet known at this moment.
  // Therefore we store the palette index encoded into the colors.
  // They will be converted to rgb in "FixColors".
  cidx := (rec.Border and MASK_XF_BORDER_LEFT_COLOR) shr 11;
  fmt.BorderStyles[cbWest].Color := IfThen(cidx >= BIFF4_MAX_PALETTE_SIZE, scBlack, SetAsPaletteIndex(cidx));
  cidx := (rec.Border and MASK_XF_BORDER_RIGHT_COLOR) shr 27;
  fmt.BorderStyles[cbEast].Color := IfThen(cidx >= BIFF4_MAX_PALETTE_SIZE, scBlack, SetAsPaletteIndex(cidx));
  cidx := (rec.Border and MASK_XF_BORDER_TOP_COLOR) shr 3;
  fmt.BorderStyles[cbNorth].Color := IfThen(cidx >= BIFF4_MAX_PALETTE_SIZE, scBlack, SetAsPaletteIndex(cidx));
  cidx := (rec.Border and MASK_XF_BORDER_BOTTOM_COLOR) shr 19;
  fmt.BorderStyles[cbSouth].Color := IfThen(cidx >= BIFF4_MAX_PALETTE_SIZE, scBlack, SetAsPaletteIndex(cidx));

  // Background
  fill := rec.Background and MASK_XF_BKGR_FILLPATTERN;
  for fs in TsFillStyle do
  begin
    if fs = fsNoFill then
      Continue;
    if fill = MASK_XF_FILL_PATT[fs] then
    begin
      // Fill style
      fmt.Background.Style := fs;
      // Pattern color
      cidx := (rec.Background and MASK_XF_BKGR_PATTERN_COLOR) shr 6;  // Palette index
      fmt.Background.FgColor := IfThen(cidx = SYS_DEFAULT_FOREGROUND_COLOR,
        scBlack, SetAsPaletteIndex(cidx));
      cidx := (rec.Background and MASK_XF_BKGR_BACKGROUND_COLOR) shr 11;
      fmt.Background.BgColor := IfThen(cidx = SYS_DEFAULT_BACKGROUND_COLOR,
        scTransparent, SetAsPaletteIndex(cidx));
      Include(fmt.UsedFormattingFields, uffBackground);
      break;
    end;
  end;

  // Protection
  case WordLEToN(rec.XFType_Prot_ParentXF) and MASK_XF_TYPE_PROTECTION of
    0:
      fmt.Protection := [];
    MASK_XF_TYPE_PROT_LOCKED:
      fmt.Protection := [cpLockCell];
    MASK_XF_TYPE_PROT_FORMULA_HIDDEN:
      fmt.Protection := [cpHideFormulas];
    MASK_XF_TYPE_PROT_LOCKED + MASK_XF_TYPE_PROT_FORMULA_HIDDEN:
      fmt.Protection := [cpLockCell, cpHideFormulas];
  end;
  if fmt.Protection <> DEFAULT_CELL_PROTECTION then
    Include(fmt.UsedFormattingFields, uffProtection);

  // Add the XF to the list
  FCellFormatList.Add(fmt);
end;

initialization
  sfidExcel4 := RegisterSpreadFormat(sfUser,
    TsSpreadBIFF4Reader, nil, 'Excel 4', 'BIFF4', ['.xls']
  );

end.

