{ 
  Test related to BIFF8 shared string table
  This unit tests are writing out to and reading back from files.
}

unit ssttests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadWriteReadColorTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadSSTTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    // General test procedure
    procedure TestWriteRead_SST_General(ATestCase: Integer);

  published
    { 1 ASCII string in SST, entirely in SST record }
    procedure TestWriteRead_SST_1ASCII;
    { 1 ASCII wide in SST, entirely in SST record }
    procedure TestWriteRead_SST_1Wide;
    { 3 string in SST, all entirely in SST record }
    procedure TestWriteRead_SST_3ASCII;
    { 3 string in SST, widestring case, all entirely in SST record }
    procedure TestWriteRead_SST_3Wide;
    { 1 long ASCII string in SST, fills SST record completely, no CONTINUE record needed }
    procedure TestWriteRead_SST_1LongASCII;
    { 1 long wide string in SST, fills SST record completely, no CONTINUE record needed }
    procedure TestWriteRead_SST_1LongWide;
    { ASCII string 2 character longer than SST record max --> CONTINUE record needed }
    procedure TestWriteRead_SST_1CONTINUE_1ASCII;
    { wide string 2 character longer than SST record max --> CONTINUE record needed }
    procedure TestWriteRead_SST_1CONTINUE_1Wide;
    { short ASCII string, then long ASCII string, 1 CONTINUE record needed }
    procedure TestWriteRead_SST_1CONTINUE_ShortASCII_LongASCII;
    { short widestring, then long widestring, 1 CONTINUE record needed }
    procedure TestWriteRead_SST_1CONTINUE_ShortWide_LongWide;
    { long ASCII string, then short ASCII string, 1 CONTINUE record needed }
    procedure TestWriteRead_SST_1CONTINUE_LongASCII_ShortASCII;
    { long widestring, then short wide string into CONTINUE record }
    procedure TestWriteRead_SST_1CONTINUE_LongWide_ShortWide;
    { very long ASCII string needing two CONTINUE records }
    procedure TestWriteRead_SST_2CONTINUE_VeryLongASCII;
    { very long widestring needing two CONTINUE records }
    procedure TestWriteRead_SST_2CONTINUE_VeryLongWide;
    { three long ASCII strings needing two CONTINUE records }
    procedure TestWriteRead_SST_2CONTINUE_3LongASCII;
    { three long widestrings needing two CONTINUE records }
    procedure TestWriteRead_SST_2CONTINUE_3LongWide;
    { 1 ASCII string in SST, entirely in SST record, font alternating from char to char }
    procedure TestWriteRead_SST_1ASCII_RichText;
    { 1 widestring in SST, entirely in SST record, font alternating from char to char }
    procedure TestWriteRead_SST_1Wide_RichText;
    { long ASCII string which reaches beyond SST into CONTINUE. Short Rich-Text
      staying within the same CONTINUE record}
    procedure TestWriteRead_SST_CONTINUE_LongASCII_ShortRichText;
    { long widestring which reaches beyond SST into CONTINUE. Short Rich-Text
      staying within the same CONTINUE record}
    procedure TestWriteRead_SST_CONTINUE_LongWide_ShortRichText;
    { long ASCII string with rich-text formatting. The string stays within SST
      but rich-text parameters reach into CONTINUE record. }
    procedure TestWriteRead_SST_CONTINUE_ShortASCII_LongRichText;
    { long widestring with rich-text formatting. The string stays within SST
      but rich-text parameters reach into CONTINUE record. }
    procedure TestWriteRead_SST_CONTINUE_ShortWide_LongRichText;
    { long ASCII string with rich-text formatting. The string stays within SST
      but long rich-text parameters flow into 2 CONTINUE records. }
    procedure TestWriteRead_SST_2CONTINUE_ASCII_LongRichText;
    { long widestring with rich-text formatting. The string stays within SST
      but long rich-text parameters flow into 2 CONTINUE records. }
    procedure TestWriteRead_SST_2CONTINUE_Wide_LongRichText;
  end;


implementation

uses
  Math, LazUTF8;

const
  SST_Sheet = 'SST';
  MAX_BYTES_PER_RECORD = 8224;

{ TSpreadWriteReadSSTTests }

procedure TSpreadWriteReadSSTTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadSSTTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_General(ATestCase: Integer);
const
  // Every record can contain 8224 data bytes (without BIFF header).
  // The SST record needs 2x4 bytes for the string counts.
  // The rest (8224-8) is for the string wbich has a header of 3 bytes (2 bytes
  // string length + 1 byte flags). fpspreadsheet writes string as widestring,
  // i.2. 2 bytes per character.
  maxLenSST = MAX_BYTES_PER_RECORD - 3 - 8;
  maxLenCONTINUE = MAX_BYTES_PER_RECORD - 1;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  MyCell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  currentText: string;
  currentRtParams: TsRichTextParams;
  currentFont: TsFont;
  expectedText: array of string;
  expectedRtParams: array of TsRichTextParams;
  expectedFont: Array[0..1] of TsFont;
  expectedFontIndex: Array[0..1] of Integer;
  i, j: Integer;
  col, row: Cardinal;
  fnt: TsFont;

  function CreateString(ALen: Integer): String;
  var
    i: Integer;
  begin
    SetLength(Result, ALen);
    for i:=1 to ALen do
      Result[i] := char((i-1) mod 26 + ord('A'));
  end;

  function AlternatingFont(AStrLen: Integer): TsRichTextParams;
  var
    i: Integer;
  begin
    SetLength(Result, AStrLen div 2);
    for i := 0 to High(Result) do begin
      Result[i].FirstIndex := i*2 + 1;
        // character index is 1-based in fps
      Result[i].FontIndex := expectedFontIndex[i mod 2];
       // Avoid using the default font here, it makes counting too complex.
    end;
  end;

begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    fnt := MyWorkbook.GetDefaultFont;

    expectedFontIndex[0] := 1;
    expectedFontIndex[1] := 2;
    for j:=0 to 1 do
      expectedFont[j] := MyWorkbook.GetFont(expectedFontIndex[j]);

    case ATestCase of
      0: begin
           // 1 short ASCII string, easily fits within SST record
           SetLength(expectedtext, 1);
           expectedText[0] := 'ABC';
         end;
      1: begin
           // 1 short wide string, easily fits within SST record
           SetLength(expectedtext, 1);
           expectedText[0] := 'äöü';
         end;
      2: begin
           // 3 short ASCII strings, easily fit within SST record
           SetLength(expectedtext, 3);
           expectedText[0] := 'ABC';
           expectedText[1] := 'DEF';
           expectedText[2] := 'GHI';
         end;
      3: begin
           // 3 short strings, widestring case, easily fit within SST record
           SetLength(expectedtext, 3);
           expectedText[0] := 'äöü';
           expectedText[1] := 'DEF';
           expectedText[2] := 'GHI';
         end;
      4: begin
           // 1 long ASCII string, max length for SST record
           SetLength(expectedtext, 1);
           expectedText[0] := CreateString(maxLenSST);
         end;
      5: begin
           // 1 long widestring, max length for SST record
           SetLength(expectedtext, 1);
           expectedText[0] := 'ä' + CreateString(maxLenSST div 2 - 1);
         end;
      6: begin
           // 1 long ASCII string, 2 characters more than max SST length --> CONTINUE needed
           SetLength(expectedtext, 1);
           expectedText[0] := CreateString(maxLenSST + 2);
         end;
      7: begin
           // 1 long widestring, 2 characters more than max SST length --> CONTINUE needed
           SetLength(expectedtext, 1);
           expectedText[0] := 'ä' + CreateString(maxLenSST div 2 + 1);
         end;
      8: begin
           // a short ASCII string, plus 1 long ASCII string reaching into CONTINUE record
           SetLength(expectedtext, 2);
           expectedText[0] := 'ABC';
           expectedText[1] := CreateString(maxLenSST);
         end;
      9: begin
           // a short widestring, plus 1 long widestring reaching into CONTINUE record
           SetLength(expectedtext, 2);
           expectedText[0] := 'äöü';
           expectedText[1] := 'äöü' + CreateString(maxLenSST div 2);
         end;
     10: begin
           // 1 long ASCII string staying inside SST, 1 short ASCII string into CONTINUE
           // The header of the short string does no longer fit in the SST record.
           // The short string must bo into CONTINUE completely.
           SetLength(expectedtext, 2);
           expectedText[0] := CreateString(maxLenSST-2);
           expectedText[1] := 'ABCDEF';
         end;
     11: begin
           // 1 long widestring staying inside SST, 1 short widestring into CONTINUE
           SetLength(expectedtext, 2);
           expectedText[0] := 'ä' + CreateString(maxLenSST div 2 - 2);
           expectedText[1] := 'ÄÖÜabc';
         end;
     12: begin
           // a very long ASCII string needing two CONTINUE records
           SetLength(expectedtext, 1);
           expectedText[0] := CreateString(maxLenSST + maxLenCONTINUE + 3);
         end;
     13: begin
           // a very long wide string needing two CONTINUE records
           SetLength(expectedtext, 1);
           expectedText[0] := 'äöü' + CreateString(maxLenSST div 2  + maxLenCONTINUE div 2);
         end;
     14: begin
           // three long ASCII strings needing two CONTINUE records
           SetLength(expectedtext, 3);
           expectedText[0] := CreateString(maxLenSST - 3);
           expectedText[1] := CreateString(maxLenSST - 3 + maxLenCONTINUE - 3);
           expectedText[2] := CreateString(maxLenSST - 3 + maxLenCONTINUE - 3);
         end;
     15: begin
           // three long wide strings needing two CONTINUE records
           SetLength(expectedtext, 3);
           expectedText[0] := CreateString(maxLenSST div 2 - 3);
           expectedText[1] := CreateString(maxLenSST div 2 - 3 + maxLenCONTINUE div 2 - 3);
           expectedText[2] := CreateString(maxLenSST div 2 - 3 + maxLenCONTINUE div 2 - 3);
         end;
     16: begin
           // 1 short ASCII string, easily fits within SST record, with Rich-Text
           SetLength(expectedtext, 1);
           expectedText[0] := 'ABCD';
           SetLength(expectedRtParams, 1);
           expectedRtParams[0] := AlternatingFont(Length(expectedText[0]));
         end;
     17: begin
           // 1 short widestring, easily fits within SST record, with Rich-Text
           SetLength(expectedtext, 1);
           expectedText[0] := 'äöüa';
           SetLength(expectedRtParams, 1);
           expectedRtParams[0] := AlternatingFont(4);
         end;
     18: begin
           // 1 long ASCII string, reaches into CONTINUE record, short Rich-Text
           SetLength(expectedtext, 1);
           expectedText[0] := CreateString(maxLenSST+5);
           SetLength(expectedRtParams, 1);
           expectedRtParams[0] := AlternatingFont(10);
         end;
     19: begin
           // 1 long wide string, reaches into CONTINUE record, short Rich-Text
           SetLength(expectedtext, 1);
           expectedText[0] := 'äöü' + CreateString(maxLenSST div 2 + 5);
           SetLength(expectedRtParams, 1);
           expectedRtParams[0] := AlternatingFont(10);
         end;
     20: begin
           // ASCII string staying within SST. But has Rich-Text parameters
           // overflowing into the CONTINUE record
           SetLength(expectedtext, 1);
           expectedText[0] := CreateString(maxLenSST - 10);
           SetLength(expectedRtParams, 1);
           expectedRtParams[0] := AlternatingFont(100);
         end;
     21: begin
           // wide string staying within SST. But has Rich-Text parameters
           // overflowing into the CONTINUE record
           SetLength(expectedtext, 1);
           expectedText[0] := 'äöü' + CreateString(maxLenSST div 2 - 13);
           SetLength(expectedRtParams, 1);
           expectedRtParams[0] := AlternatingFont(100);
         end;
     22: begin
           // Long ASCII string staying within SST. But has long Rich-Text
           // parameters overflowing into two CONTINUE records
           SetLength(expectedtext, 1);
           expectedText[0] := CreateString(maxLenSST - 10);
           SetLength(expectedRtParams, 1);
           expectedRtParams[0] := AlternatingFont(Length(expectedText[0]));
         end;
     23: begin
           // Long widestring staying within SST. But has long Rich-Text
           // parameters overflowing into two CONTINUE records
           SetLength(expectedtext, 1);
           expectedText[0] := 'äöü' + CreateString(maxLenSST div 2 - 13);
           SetLength(expectedRtParams, 1);
           expectedRtParams[0] := AlternatingFont(UTF8Length(expectedText[0]) div 2);
         end;
    end;

    { Create spreadsheet and write to file }
    MyWorkSheet:= MyWorkBook.AddWorksheet(SST_Sheet);
    col := 0;
    for row := 0 to High(expectedText) do
      if row < Length(expectedRtParams) then
        MyCell := MyWorksheet.WriteText(row, col, expectedText[row], expectedRtParams[row])
      else
        MyCell := MyWorksheet.WriteText(row, col, expectedText[row]);
    MyWorkBook.WriteToFile(TempFile, sfExcel8, true);
  finally
    MyWorkbook.Free;
  end;

  { Read the spreadsheet }
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, sfExcel8);
    MyWorksheet := MyWorkbook.GetWorksheetByIndex(0);
    col := 0;
    for row := 0 to High(expectedText) do begin
      myCell := MyWorksheet.FindCell(row, col);
      if MyCell = nil then
        fail('Error in test code. Failed to get cell.');

      currentText := MyWorksheet.ReadAsText(MyCell);
      CheckEquals(expectedText[row], currentText,
        'Saved cell text mismatch, cell '+CellNotation(MyWorksheet, row, col));

      if row < Length(expectedRtParams) then
      begin
        currentRtParams := MyCell^.RichTextParams;
        CheckEquals(Length(expectedRtParams[row]), Length(currentRtParams),
          'Number of rich-text parameters mismatch, cell '+CellNotation(MyWorksheet, row, col));

        for i:=0 to High(currentRtParams) do
        begin
          CheckEquals(expectedRtParams[row][i].FirstIndex, currentRtParams[i].FirstIndex,
            'Character index mismatch in rich-text parameter #' + IntToStr(i) +
            ', cell ' + CellNotation(MyWorksheet, row, col));

          currentFont := MyWorkbook.GetFont(currentRtParams[i].FontIndex);
          CheckEquals(currentFont.Fontname, expectedFont[i mod 2].FontName,
            'Font name mismatch in rich-text parameter #' + IntToStr(i) +
            ', cell ' + CellNotation(MyWorksheet, row, col));
          CheckEquals(currentFont.Size, expectedFont[i mod 2].Size,
            'Font size mismatch in rich-text parameter #' + IntToStr(i) +
            ', cell ' + CellNotation(MyWorksheet, row, col));
          CheckEquals(integer(currentFont.Style), integer(expectedFont[i mod 2].Style),
            'Font style mismatch in rich-text parameter #' + IntToStr(i) +
            ', cell ' + CellNotation(MyWorksheet, row, col));
          CheckEquals(currentFont.Color, expectedFont[i mod 2].Color,
            'Font color mismatch in rich-text parameter #' + IntToStr(i) +
            ', cell ' + CellNotation(MyWorksheet, row, col));
        end;
      end;
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

{ Writes/reads one string ASCII only. The string fits in the SST record }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1ASCII;
begin
  TestWriteRead_SST_General(0);
end;

{ Writes/reads one wide string only. The string fits in the SST record }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1Wide;
begin
  TestWriteRead_SST_General(1);
end;

{ Writes/reads 3 strings, all entirely in SST record }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_3ASCII;
begin
  TestWriteRead_SST_General(2);
end;

{ Writes/reads 3 strings, widestring case, all entirely in SST record }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_3Wide;
begin
  TestWriteRead_SST_General(3);
end;

{ 1 long ASCII string in SST, fills SST record exactly, no CONTINUE record needed }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1LongASCII;
begin
  TestWriteRead_SST_General(4);
end;

{ 1 long widestring in SST, fills SST record exactly, no CONTINUE record needed }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1LongWide;
begin
  TestWriteRead_SST_General(5);
end;

{ 1 ASCII string, 2 characters longer than in SST record max
  --> CONTINUE record needed }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1CONTINUE_1ASCII;
begin
  TestWriteRead_SST_General(6);
end;

{ 1 widestring, 2 characters longer than in SST record max
   --> CONTINUE record needed }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1CONTINUE_1Wide;
begin
  TestWriteRead_SST_General(7);
end;

{ short ASCII string, then long ASCII string, 1 CONTINUE record needed }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1CONTINUE_ShortASCII_LongASCII;
begin
  TestWriteRead_SST_General(8);
end;

{ short widestring, then long widestring, 1 CONTINUE record needed }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1CONTINUE_ShortWide_LongWide;
begin
  TestWriteRead_SST_General(9);
end;

{ long ASCII string, then short ACII string into CONTINUE record }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1CONTINUE_LongASCII_ShortASCII;
begin
  TestWriteRead_SST_General(10);
end;

{ long widestring, then short widestring into CONTINUE record }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1CONTINUE_LongWide_ShortWide;
begin
  TestWriteRead_SST_General(11);
end;

{ very long ASCII string, needing two CONTINUE records }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_2CONTINUE_VeryLongASCII;
begin
  TestWriteRead_SST_General(12);
end;

{ very long widestring, needing two CONTINUE records }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_2CONTINUE_VeryLongWide;
begin
  TestWriteRead_SST_General(13);
end;

{ three long ASCII strings, needing two CONTINUE records }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_2CONTINUE_3LongASCII;
begin
  TestWriteRead_SST_General(14);
end;

{ three long widestrings, needing two CONTINUE records }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_2CONTINUE_3LongWide;
begin
  TestWriteRead_SST_General(15);
end;

{ Writes/reads one ASCII string only. The string fits in the SST record.
  Uses rich-text formatting toggling font every second character. }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1ASCII_RichText;
begin
  TestWriteRead_SST_General(16);
end;

{ Writes/reads one wide string only. The string fits in the SST record.
  Uses rich-text formatting toggling font every second character. }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_1Wide_RichText;
begin
  TestWriteRead_SST_General(17);
end;

{ Writes/reads one long ASCII string which reaches beyond SST into CONTINUE.
  Uses short rich-text formatting staying within this CONTINUE record. }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_CONTINUE_LongASCII_ShortRichText;
begin
  TestWriteRead_SST_General(18);
end;

{ Writes/reads one long wide string which reaches beyond SST into CONTINUE.
Uses short rich-text formatting staying within this CONTINUE record. }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_CONTINUE_LongWide_ShortRichText;
begin
  TestWriteRead_SST_General(19);
end;

{ Writes/reads one short ASCII string with rich-text formatting. The string
  stay within SST, but rich-text parameters reach into CONTINUE record. }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_CONTINUE_ShortASCII_LongRichText;
begin
  TestWriteRead_SST_General(20);
end;

{ Writes/reads one long widestring with rich-text formatting. The string
  stay within SST, but rich-text parameters reach into CONTINUE record. }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_CONTINUE_ShortWide_LongRichText;
begin
  TestWriteRead_SST_General(21);
end;

{ long ASCII string with rich-text formatting. The string stays within SST
  but long rich-text parameters flow into 2 CONTINUE records. }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_2CONTINUE_ASCII_LongRichText;
begin
  TestWriteRead_SST_General(22);
end;

{ long widestring with rich-text formatting. The string stays within SST
  but long rich-text parameters flow into 2 CONTINUE records. }
procedure TSpreadWriteReadSSTTests.TestWriteRead_SST_2CONTINUE_Wide_LongRichText;
begin
  TestWriteRead_SST_General(23);
end;


initialization
  RegisterTest(TSpreadWriteReadSSTTests);

end.

