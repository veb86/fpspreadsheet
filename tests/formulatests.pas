unit formulatests;

{$mode objfpc}{$H+}

{ Deactivate this define in order to bypass tests which will raise an exception
  when the corresponding rpn formula is calculated. }
{.$DEFINE ENABLE_CALC_RPN_EXCEPTIONS}

{ Deactivate this define to include errors in the structure of the rpn formulas.
  Note that Excel report a corrupted file when trying to read this file }
{.DEFINE ENABLE_DEFECTIVE_FORMULAS }

{ Activate the project define FORMULADEBUG to log the formulas written }


interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpstypes, fpsallformats, fpspreadsheet, fpsexprparser,
  xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadWriteReadFormula }
  //Write to xls/xml file and read back
  TSpreadWriteReadFormulaTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    // Test reconstruction of formula strings
    procedure Test_Write_Read_FormulaStrings(AFormat: TsSpreadsheetFormat;
      UseRPNFormula: Boolean);
    procedure Test_Write_Read_CalcFormulas(AFormat: TsSpreadsheetformat;
      UseRPNFormula: Boolean);
    procedure Test_Write_Read_Calc3DFormulas(AFormat: TsSpreadsheetFormat);

    procedure Test_OverwriteFormulaTest(ATest: Integer; AFormat: TsSpreadsheetFormat);

  published
    // Writes out formulas & reads them back.

    { BIFF2 Tests }
    procedure Test_Write_Read_FormulaStrings_BIFF2;
    { BIFF5 Tests }
    procedure Test_Write_Read_FormulaStrings_BIFF5;
    { BIFF8 Tests }
    procedure Test_Write_Read_FormulaStrings_BIFF8;
    { OOXML Tests }
    procedure Test_Write_Read_FormulaStrings_OOXML;
    { Excel2003/XML Tests }
    procedure Test_Write_Read_FormulaStrings_XML;
    { ODS Tests }
    procedure Test_Write_Read_FormulaStrings_ODS;

    // Writes out and calculates rpn formulas, read back
    { BIFF2 Tests }
    procedure Test_Write_Read_CalcRPNFormula_BIFF2;
    { BIFF5 Tests }
    procedure Test_Write_Read_CalcRPNFormula_BIFF5;
    { BIFF8 Tests }
    procedure Test_Write_Read_CalcRPNFormula_BIFF8;
    { OOXML Tests }
    procedure Test_Write_Read_CalcRPNFormula_OOXML;
    { Excel2003/XML Tests }
    procedure Test_Write_Read_CalcRPNFormula_XML;
    { ODSL Tests }
    procedure Test_Write_Read_CalcRPNFormula_ODS;

    // Writes out and calculates string formulas, read back
    { BIFF2 Tests }
    procedure Test_Write_Read_CalcStringFormula_BIFF2;
    { BIFF5 Tests }
    procedure Test_Write_Read_CalcStringFormula_BIFF5;
    { BIFF8 Tests }
    procedure Test_Write_Read_CalcStringFormula_BIFF8;
    { OOXML Tests }
    procedure Test_Write_Read_CalcStringFormula_OOXML;
    { Excel2003/XML Tests }
    procedure Test_Write_Read_CalcStringFormula_XML;
    { ODS Tests }
    procedure Test_Write_Read_CalcStringFormula_ODS;

    { Formulas with 3D references to other sheets }
    procedure Test_Write_Read_Calc3DFormula_BIFF5;
    procedure Test_Write_Read_Calc3DFormula_BIFF8;
    procedure Test_Write_Read_Calc3DFormula_OOXML;
//    procedure Test_Write_Read_Calc3DFormula_XML;
    procedure Test_Write_Read_Calc3DFormula_ODS;

    { Overwrite formula with other content }
    procedure Test_OverwriteFormula_Number_BIFF2;
    procedure Test_OverwriteFormula_Number_BIFF5;
    procedure Test_OverwriteFormula_Number_BIFF8;
    procedure Test_OverwriteFormula_Number_OOXML;
    procedure Test_OverwriteFormula_Number_XML;
    procedure Test_OverwriteFormula_Number_ODS;

    procedure Test_OverwriteFormula_Text_BIFF2;
    procedure Test_OverwriteFormula_Text_BIFF5;
    procedure Test_OverwriteFormula_Text_BIFF8;
    procedure Test_OverwriteFormula_Text_OOXML;
    procedure Test_OverwriteFormula_Text_XML;
    procedure Test_OverwriteFormula_Text_ODS;

    procedure Test_OverwriteFormula_Bool_BIFF2;
    procedure Test_OverwriteFormula_Bool_BIFF5;
    procedure Test_OverwriteFormula_Bool_BIFF8;
    procedure Test_OverwriteFormula_Bool_OOXML;
    procedure Test_OverwriteFormula_Bool_XML;
    procedure Test_OverwriteFormula_Bool_ODS;

    procedure Test_OverwriteFormula_Error_BIFF2;
    procedure Test_OverwriteFormula_Error_BIFF5;
    procedure Test_OverwriteFormula_Error_BIFF8;
    procedure Test_OverwriteFormula_Error_OOXML;
    procedure Test_OverwriteFormula_Error_XML;
    procedure Test_OverwriteFormula_Error_ODS;

  end;

implementation

uses
 {$IFDEF FORMULADEBUG}
  LazLogger,
 {$ENDIF}
  math, typinfo, lazUTF8, fpsUtils, fpsRPN, rpnFormulaUnit;

var
  // Array containing the "true" results of the formulas, for comparison
  SollValues: array of TsExpressionResult;

// Helper for statistics tests
const
  STATS_NUMBERS: Array[0..4] of Double = (1.0, 1.1, 1.2, 0.9, 0.8);
var
  numberArray: array[0..4] of Double;



{ TSpreadWriteReadFormatTests }

procedure TSpreadWriteReadFormulaTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadFormulaTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings(
  AFormat: TsSpreadsheetFormat; UseRPNFormula: Boolean);
{ If UseRPNFormula is true the test formulas are generated from RPN formulas.
  Otherwise they are generated from string formulas. }
const
  SHEET = 'Sheet1';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
  formula: String;
  expected: String;
  actual: String;
  cell: PCell;
  cellB1: Double;
  cellB2: Double;
  number: Double;
  s: String;
  hr, min, sec, msec: Word;
  k: Integer;
begin
  TempFile := GetTempFileName;

  // Create test workbook
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];
    MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);

    // Write out all test formulas
    // All formulas are in column B
    {$I testcases_calcrpnformula.inc}
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];

    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, SHEET);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');
    for Row := 0 to MyWorksheet.GetLastRowIndex do
    begin
      cell := MyWorksheet.FindCell(Row, 1);
      if HasFormula(cell) then begin
        actual := MyWorksheet.ReadFormulaAsString(cell);
        expected := MyWorksheet.ReadAsUTF8Text(Row, 0);
        CheckEquals(expected, actual, 'Test read formula mismatch, cell '+CellNotation(MyWorkSheet,Row,1));
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_BIFF2;
begin
  Test_Write_Read_FormulaStrings(sfExcel2, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_BIFF5;
begin
  Test_Write_Read_FormulaStrings(sfExcel5, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_BIFF8;
begin
  Test_Write_Read_FormulaStrings(sfExcel8, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_OOXML;
begin
  Test_Write_Read_FormulaStrings(sfOOXML, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_XML;
begin
  Test_Write_Read_FormulaStrings(sfExcelXML, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_FormulaStrings_ODS;
begin
  Test_Write_Read_FormulaStrings(sfOpenDocument, true);
end;


{ Test writing and reading (i.e. reconstruction) of shared formula strings }
                                                (*
procedure TSpreadWriteReadFormulaTests.Test_Write_Read_SharedFormulaStrings(
  AFormat: TsSpreadsheetFormat);
const
  SHEET = 'SharedFormulaSheet';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  cell: PCell;
  row, col: Cardinal;
  TempFile: String;
  actual, expected: String;
  sollValues: array[1..4, 0..4] of string;
begin
  TempFile := GetTempFileName;

  // Create test workbook
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];
    MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);

    // Write out test values
    MyWorksheet.WriteNumber(0, 0, 1.0);  // A1
    MyWorksheet.WriteNumber(0, 1, 2.0);
    MyWorksheet.WriteNumber(0, 2, 3.0);
    MyWorksheet.WriteNumber(0, 3, 4.0);
    MyWorksheet.WriteNumber(0, 4, 5.0);  // E1

    // Write out all test formulas
    // sollValues contains the expected formula as seen from each cell in the
    // shared formula block.
    MyWorksheet.WriteSharedFormula('A2:E2', 'A1');
      sollValues[1, 0] := 'A1';
      sollValues[1, 1] := 'B1';
      sollValues[1, 2] := 'C1';
      sollValues[1, 3] := 'D1';
      sollValues[1, 4] := 'E1';
    MyWorksheet.WriteSharedFormula('A3:E3', '$A1');
      sollValues[2, 0] := '$A1';
      sollValues[2, 1] := '$A1';
      sollValues[2, 2] := '$A1';
      sollValues[2, 3] := '$A1';
      sollValues[2, 4] := '$A1';
    MyWorksheet.WriteSharedFormula('A4:E4', 'A$1');
      sollValues[3, 0] := 'A$1';
      sollValues[3, 1] := 'B$1';
      sollValues[3, 2] := 'C$1';
      sollValues[3, 3] := 'D$1';
      sollValues[3, 4] := 'E$1';
    MyWorksheet.WriteSharedFormula('A5:E5', '$A$1');
      sollValues[4, 0] := '$A$1';
      sollValues[4, 1] := '$A$1';
      sollValues[4, 2] := '$A$1';
      sollValues[4, 3] := '$A$1';
      sollValues[4, 4] := '$A$1';

    MyWorksheet.WriteSharedFormula('A6:E6', 'SIN(A1)');
    MyWorksheet.WriteSharedFormula('A7:E7', 'SIN($A1)');
    MyWorksheet.WriteSharedFormula('A8:E8', 'SIN(A$1)');
    MyWorksheet.WriteSharedFormula('A9:E9', 'SIN($A$1)');

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas, boAutoCalc];

    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, SHEET);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    for row := 1 to 8 do begin
      for col := 0 to MyWorksheet.GetLastColIndex do begin
        cell := Myworksheet.FindCell(row, col);
        if HasFormula(cell) then begin
          actual := MyWorksheet.ReadFormulaAsString(cell);
          if row <= 4 then
            expected := SollValues[row, col]
          else
            expected := 'SIN(' + SollValues[row-4, col] + ')';
          CheckEquals(expected, actual, 'Test read formula mismatch, cell '+CellNotation(MyWorkSheet,Row,Col));
        end else
          fail('No formula found in cell ' + CellNotation(MyWorksheet, Row, Col));
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_SharedFormulaStrings_BIFF2;
begin
  Test_Write_Read_SharedFormulaStrings(sfExcel2);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_SharedFormulaStrings_BIFF5;
begin
  Test_Write_Read_SharedFormulaStrings(sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_SharedFormulaStrings_BIFF8;
begin
  Test_Write_Read_SharedFormulaStrings(sfExcel8);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_SharedFormulaStrings_OOXML;
begin
  Test_Write_Read_SharedFormulaStrings(sfOOXML);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_SharedFormulaStrings_ODS;
begin
  Test_Write_Read_SharedFormulaStrings(sfOpenDocument);
end;
                              *)
{ Test calculation of formulas }

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcFormulas(
  AFormat: TsSpreadsheetFormat; UseRPNFormula: Boolean);
{ If UseRPNFormula is TRUE, the test formulas are generated from RPN syntax,
  otherwise string formulas are used. }
const
  SHEET = 'Sheet1';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row: Integer;
  TempFile: string;    //write xls/xml to this file and read back from it
  actual: TsExpressionResult;
  expected: TsExpressionResult;
  cell: PCell;
  sollValues: array of TsExpressionResult;
  formula: String;
  s: String;
  hr,min,sec,msec: Word;
  ErrorMargin: double;
  k: Integer;
  { When comparing soll and formula values we must make sure that the soll
    values are calculated from double precision numbers, they are used in
    the formula calculation as well. The next variables, along with STATS_NUMBERS
    above, hold the arguments for the direction function calls. }
  number: Double;
  cellB1: Double;
  cellB2: Double;
begin
  ErrorMargin:=0; //1.44E-7;
  //1.44E-7 for SUMSQ formula
  //6.0E-8 for SUM formula
  //4.8E-8 for MAX formula
  //2.4E-8 for now formula
  //about 1E-15 is needed for some trig functions

  // Create test workbook
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];
    // Calculation of rpn formulas must be activated explicitly!

    MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);
    { Write out test formulas.
      This include file creates various rpn formulas and stores the expected
      results in array "sollValues".
      The test file contains the text representation in column A, and the
      formula in column B. }
    Row := 0;
    TempFile := GetTempFileName;
    {$I testcases_calcrpnformula.inc}
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the workbook
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := Myworkbook.Options + [boReadFormulas];
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    MyWorkbook.CalcFormulas;
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, SHEET);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    for Row := 0 to MyWorksheet.GetLastRowIndex do
    begin
      formula := MyWorksheet.ReadAsUTF8Text(Row, 0);
      cell := MyWorksheet.FindCell(Row, 1);
      if (cell = nil) then
        fail('Error in test code: Failed to get cell ' + CellNotation(MyWorksheet, Row, 1));

      case cell^.ContentType of
        cctBool       : actual := BooleanResult(cell^.BoolValue);
        cctNumber     : actual := FloatResult(cell^.NumberValue);
        cctDateTime   : actual := DateTimeResult(cell^.DateTimeValue);
        cctUTF8String : actual := StringResult(cell^.UTF8StringValue);
        cctError      : actual := ErrorResult(cell^.ErrorValue);
        cctEmpty      : actual := EmptyResult;
        else            fail('ContentType not supported');
      end;

      expected := SollValues[row];
      // Cell does not store integers!
      if expected.ResultType = rtInteger then expected := FloatResult(expected.ResInteger);

      CheckEquals(
        GetEnumName(TypeInfo(TsExpressionResult), ord(expected.ResultType)),
        GetEnumName(TypeInfo(TsExpressionResult), ord(actual.ResultType)),
        'Test read calculated formula data type mismatch, formula "' + formula +
        '", cell '+CellNotation(MyWorkSheet,Row,1));

      // The now function result is volatile, i.e. changes continuously. The
      // time for the soll value was created such that we can expect to have
      // the file value in the same second. Therefore we neglect the milliseconds.
      if formula = '=NOW()' then begin
        // Round soll value to seconds
        DecodeTime(expected.ResDateTime, hr,min,sec,msec);
        expected.ResDateTime := EncodeTime(hr, min, sec, 0);
        // Round formula value to seconds
        DecodeTime(actual.ResDateTime, hr,min,sec,msec);
        actual.ResDateTime := EncodeTime(hr,min,sec,0);
      end;

      case actual.ResultType of
        rtBoolean:
          CheckEquals(BoolToStr(expected.ResBoolean), BoolToStr(actual.ResBoolean),
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
        rtFloat:
          {$if (defined(mswindows)) or (FPC_FULLVERSION>=20701)}
          // FPC 2.6.x and trunk on Windows need this, also FPC trunk on Linux x64
          CheckEquals(expected.ResFloat, actual.ResFloat, ErrorMargin,
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
          {$else}
          // Non-Windows: test without error margin
          CheckEquals(expected.ResFloat, actual.ResFloat,
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
          {$endif}
        rtString:
          CheckEquals(expected.ResString, actual.ResString,
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
        rtError:
          CheckEquals(
            GetEnumName(TypeInfo(TsErrorValue), ord(expected.ResError)),
            GetEnumname(TypeInfo(TsErrorValue), ord(actual.ResError)),
            'Test read calculated formula error value mismatch, formula ' + formula +
            ', cell '+CellNotation(MyWorkSheet,Row,1));
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_BIFF2;
begin
  Test_Write_Read_CalcFormulas(sfExcel2, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_BIFF5;
begin
  Test_Write_Read_CalcFormulas(sfExcel5, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_BIFF8;
begin
  Test_Write_Read_CalcFormulas(sfExcel8, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_OOXML;
begin
  Test_Write_Read_CalcFormulas(sfOOXML, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_XML;
begin
  Test_Write_Read_CalcFormulas(sfExcelXML, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcRPNFormula_ODS;
begin
  Test_Write_Read_CalcFormulas(sfOpenDocument, true);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_BIFF2;
begin
  Test_Write_Read_CalcFormulas(sfExcel2, false);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_BIFF5;
begin
  Test_Write_Read_CalcFormulas(sfExcel5, false);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_BIFF8;
begin
  Test_Write_Read_CalcFormulas(sfExcel8, false);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_OOXML;
begin
  Test_Write_Read_CalcFormulas(sfOOXML, false);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_XML;
begin
  Test_Write_Read_CalcFormulas(sfExcelXML, false);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcStringFormula_ODS;
begin
  Test_Write_Read_CalcFormulas(sfOpenDocument, false);
end;

//------------------------------------------------------------------------------
//                   Calculation of shared formulas
//------------------------------------------------------------------------------
(*
procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcSharedFormulas(
  AFormat: TsSpreadsheetFormat);
const
  SHEET = 'SharedFormulaSheet';
  vA1 = 1.0;
  vB1 = 2.0;
  vC1 = 3.0;
  vD1 = 4.0;
  vE1 = 5.0;
  vF1 = 'A';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  cell: PCell;
  row, col: Cardinal;
  TempFile: String;
  actual, expected: String;
  sollValues: array[1..8, 0..5] of String;
begin
  TempFile := GetTempFileName;

  // Create test workbook
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];
    MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);

    // Write out test values
    MyWorksheet.WriteNumber(0, 0, vA1);    // A1
    MyWorksheet.WriteNumber(0, 1, vB1);    // B1
    MyWorksheet.WriteNumber(0, 2, vC1);    // C1
    MyWorksheet.WriteNumber(0, 3, vD1);    // D1
    MyWorksheet.WriteNumber(0, 4, vE1);    // E1
    MyWorksheet.WriteUTF8Text(0, 5, vF1);  // F1

    // Write out all test formulas
    // sollValues contains the expected formula as seen from each cell in the
    // shared formula block.
    MyWorksheet.WriteSharedFormula('A2:F2', 'A1');
      sollValues[1, 0] := Format('%g', [vA1]);
      sollValues[1, 1] := Format('%g', [vB1]);
      sollValues[1, 2] := Format('%g', [vC1]);
      sollValues[1, 3] := Format('%g', [vD1]);
      sollValues[1, 4] := Format('%g', [vE1]);
      sollValues[1, 5] := vF1;  // is a string
    MyWorksheet.WriteSharedFormula('A3:F3', '$A1');
      sollValues[2, 0] := Format('%g', [vA1]);
      sollValues[2, 1] := Format('%g', [vA1]);
      sollValues[2, 2] := Format('%g', [vA1]);
      sollValues[2, 3] := Format('%g', [vA1]);
      sollValues[2, 4] := Format('%g', [vA1]);
      sollValues[2, 5] := Format('%g', [vA1]);
    MyWorksheet.WriteSharedFormula('A4:F4', 'A$1');
      sollValues[3, 0] := Format('%g', [vA1]);
      sollValues[3, 1] := Format('%g', [vB1]);
      sollValues[3, 2] := Format('%g', [vC1]);
      sollValues[3, 3] := Format('%g', [vD1]);
      sollValues[3, 4] := Format('%g', [vE1]);
      sollValues[3, 5] := vF1;  // is a string
    MyWorksheet.WriteSharedFormula('A5:F5', '$A$1');
      sollValues[4, 0] := Format('%g', [vA1]);
      sollValues[4, 1] := Format('%g', [vA1]);
      sollValues[4, 2] := Format('%g', [vA1]);
      sollValues[4, 3] := Format('%g', [vA1]);
      sollValues[4, 4] := Format('%g', [vA1]);
      sollValues[4, 5] := Format('%g', [vA1]);

    MyWorksheet.WriteSharedFormula('A6:F6', 'SIN(A1)');
      sollValues[5, 0] := FloatToStr(sin(vA1));  // Using "FloatToStr" here like in ReadAsUTF8Text
      sollValues[5, 1] := FloatToStr(sin(vB1));
      sollValues[5, 2] := FloatToStr(sin(vC1));
      sollValues[5, 3] := FloatToStr(sin(vD1));
      sollValues[5, 4] := FloatToStr(sin(vE1));
      sollValues[5, 5] := FloatToStr(sin(0.0));  // vF1 is a string
    MyWorksheet.WriteSharedFormula('A7:F7', 'SIN($A1)');
      sollValues[6, 0] := FloatToStr(sin(vA1));
      sollValues[6, 1] := FloatToStr(sin(vA1));
      sollValues[6, 2] := FloatToStr(sin(vA1));
      sollValues[6, 3] := FloatToStr(sin(vA1));
      sollValues[6, 4] := FloatToStr(sin(vA1));
      sollValues[6, 5] := FloatToStr(sin(vA1));
    MyWorksheet.WriteSharedFormula('A8:F8', 'SIN(A$1)');
      sollValues[7, 0] := FloatToStr(sin(vA1));
      sollValues[7, 1] := FloatToStr(sin(vB1));
      sollValues[7, 2] := FloatToStr(sin(vC1));
      sollValues[7, 3] := FloatToStr(sin(vD1));
      sollValues[7, 4] := FloatToStr(sin(vE1));
      sollValues[7, 5] := FloatToStr(sin(0.0));  // vF1 is a string
    MyWorksheet.WriteSharedFormula('A9:F9', 'SIN($A$1)');
      sollValues[8, 0] := FloatToStr(sin(vA1));
      sollValues[8, 1] := FloatToStr(sin(vA1));
      sollValues[8, 2] := FloatToStr(sin(vA1));
      sollValues[8, 3] := FloatToStr(sin(vA1));
      sollValues[8, 4] := FloatToStr(sin(vA1));
      sollValues[8, 5] := FloatToStr(sin(vA1));

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas, boAutoCalc];

    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook, SHEET);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    for row := 1 to 8 do begin
      for col := 0 to MyWorksheet.GetLastColIndex do begin
        cell := Myworksheet.FindCell(row, col);
        if HasFormula(cell) then begin
          actual := copy(MyWorksheet.ReadAsUTF8Text(cell), 1, 6);   // cutting converted numbers off after some digits, certainly not always correct
          expected := copy(SollValues[row, col], 1, 6);
          CheckEquals(expected, actual, 'Test read formula mismatch, cell '+CellNotation(MyWorkSheet,Row,Col));
        end else
          fail('No formula found in cell ' + CellNotation(MyWorksheet, Row, Col));
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;


procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcSharedFormula_BIFF2;
begin
  Test_Write_Read_CalcSharedFormulas(sfExcel2);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcSharedFormula_BIFF5;
begin
  Test_Write_Read_CalcSharedFormulas(sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcSharedFormula_BIFF8;
begin
  Test_Write_Read_CalcSharedFormulas(sfExcel8);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcSharedFormula_OOXML;
begin
  Test_Write_Read_CalcSharedFormulas(sfOOXML);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_CalcSharedFormula_ODS;
begin
  Test_Write_Read_CalcSharedFormulas(sfOpenDocument);
end;
             *)


procedure TSpreadWriteReadFormulaTests.Test_Write_Read_Calc3DFormulas(
  AFormat: TsSpreadsheetFormat);
{ If UseRPNFormula is TRUE, the test formulas are generated from RPN syntax,
  otherwise string formulas are used. }
var
  sheet1, sheet2, sheet3: TsWorksheet;
  workbook: TsWorkbook;
  row: Integer;
  tempFile: string;    //write xls/xml to this file and read back from it
  actual, expected: TsExpressionResult;
  cell: PCell;
  sollValues: array of TsExpressionResult;
  formula, actualformula: String;
begin
  TempFile := GetTempFileName;
  try
    // Create test workbook
    workbook := TsWorkbook.Create;
    try
      workbook.Options := workbook.Options + [boCalcBeforeSaving];

      sheet1 := workBook.AddWorksheet('Sheet1');
      sheet2 := workbook.AddWorksheet('Sheet2');
      sheet3 := workbook.AddWorksheet('Sheet3');

      { Write out test formulas.
        This include file creates various formulas in column A and stores
        the expected results in the array SollValues. }
      Row := 0;
      {$I testcases_calc3dformula.inc}
      workbook.WriteToFile(TempFile, AFormat, true);
    finally
      workbook.Free;
    end;

    // Open the workbook
    workbook := TsWorkbook.Create;
    try
      workbook.Options := workbook.Options + [boReadFormulas];
      workbook.ReadFromFile(TempFile, AFormat);
      workbook.CalcFormulas;

      if AFormat = sfExcel2 then
        Fail('This test should not be executed')
      else
        sheet1 := workbook.GetWorksheetByName('Sheet1');
      if sheet1=nil then
        Fail('Error in test code. Failed to get named worksheet');

      for row := 0 to sheet1.GetLastRowIndex do
      begin
        cell := sheet1.FindCell(Row, 0);
        if (Cell = nil) then
          Fail('Error in test code: failed to get cell ' + CellNotation(sheet1, Row, 0));
        formula := sheet1.ReadAsText(cell);

        cell := sheet1.FindCell(Row, 1);
        if (cell = nil) then
          fail('Error in test code: Failed to get cell ' + CellNotation(sheet1, Row, 1));
        case cell^.ContentType of
          cctBool       : actual := BooleanResult(cell^.BoolValue);
          cctNumber     : actual := FloatResult(cell^.NumberValue);
          cctDateTime   : actual := DateTimeResult(cell^.DateTimeValue);
          cctUTF8String : actual := StringResult(cell^.UTF8StringValue);
          cctError      : actual := ErrorResult(cell^.ErrorValue);
          cctEmpty      : actual := EmptyResult;
          else            fail('ContentType not supported');
        end;
        actualformula := sheet1.Formulas.FindFormula(cell)^.Text; //cell^.FormulaValue;

        expected := SollValues[row];
        // Cell does not store integers!
        if expected.ResultType = rtInteger then expected := FloatResult(expected.ResInteger);

        {
        // The NOW() function result is volatile, i.e. changes continuously. The
        // time for the soll value was created such that we can expect to have
        // the file value in the same second. Therefore we neglect the milliseconds.
        if formula = '=NOW()' then begin
          // Round soll value to seconds
          DecodeTime(expected.ResDateTime, hr,min,sec,msec);
          expected.ResDateTime := EncodeTime(hr, min, sec, 0);
          // Round formula value to seconds
          DecodeTime(actual.ResDateTime, hr,min,sec,msec);
          actual.ResDateTime := EncodeTime(hr,min,sec,0);
        end;                                   }

        case actual.ResultType of
          rtBoolean:
            CheckEquals(BoolToStr(expected.ResBoolean), BoolToStr(actual.ResBoolean),
              'Test read calculated formula result mismatch, cell '+CellNotation(sheet1, Row, 1));
          rtFloat:
            {$if (defined(mswindows)) or (FPC_FULLVERSION>=20701)}
            // FPC 2.6.x and trunk on Windows need this, also FPC trunk on Linux x64
            CheckEquals(expected.ResFloat, actual.ResFloat,
              'Test read calculated formula result mismatch, cell '+CellNotation(sheet1, Row, 1));
            {$else}
            // Non-Windows: test without error margin
            CheckEquals(expected.ResFloat, actual.ResFloat,
              'Test read calculated formula result mismatch, cell '+CellNotation(sheet1, Row, 1));
            {$endif}
          rtString:
            CheckEquals(expected.ResString, actual.ResString,
              'Test read calculated formula result mismatch, cell '+CellNotation(sheet1, Row, 1));
          rtError:
            CheckEquals(
              GetEnumName(TypeInfo(TsErrorValue), ord(expected.ResError)),
              GetEnumname(TypeInfo(TsErrorValue), ord(actual.ResError)),
              'Test read calculated formula error value mismatch, cell '+CellNotation(sheet1, Row, 1));
        end;

        CheckEquals(formula, actualformula,
          'Read formula string mismatch, cell ' +CellNotation(sheet1, Row, 1));
      end;

    finally
      workbook.Free;
    end;

  finally
    DeleteFile(TempFile);
  end;
end;


procedure TSpreadWriteReadFormulaTests.Test_Write_Read_Calc3DFormula_BIFF5;
begin
  Test_Write_Read_Calc3DFormulas(sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_Calc3DFormula_BIFF8;
begin
  Test_Write_Read_Calc3DFormulas(sfExcel8);
end;

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_Calc3DFormula_OOXML;
begin
  Test_Write_Read_Calc3DFormulas(sfOOXML);
end;
                                       {
procedure TSpreadWriteReadFormulaTests.Test_Write_Read_Calc3DFormula_XML;
begin
  Test_Write_Read_Calc3DFormulas(sfExcelXML);
end;                                    }

procedure TSpreadWriteReadFormulaTests.Test_Write_Read_Calc3DFormula_ODS;
begin
  Test_Write_Read_Calc3DFormulas(sfOpenDocument);
end;


{------------------------------------------------------------------------------}

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormulaTest(ATest: Integer;
  AFormat: TsSpreadsheetFormat);
type
  TSollValues = record
    NumberValue: Integer;
    TextValue: String;
    BoolValue: Boolean;
    ErrorValue: TsErrorValue;
  end;
const
  SollValue: TSollValues = (
    NumberValue: 100;
    TextValue: 'abc';
    BoolValue: false;
    ErrorValue: errIllegalRef
  );
var
  tempfile: String;
  book: TsWorkbook;
  sheet: TsWorksheet;
  x: Float;
  s: String;
  b: Boolean;
  err: TsErrorValue;
  cell: PCell;
begin
  tempFile := GetTempFileName;

  book := TsWorkbook.Create;
  try
    book.Options := book.Options + [boAutoCalc];
    sheet := book.AddWorksheet('Test');
    sheet.WriteFormula(0, 0, '=1+1');
    case ATest of
      0: sheet.WriteNumber(0, 0, sollValue.NumberValue);
      1: sheet.WriteText(0, 0, sollValue.TextValue);
      2: sheet.WriteBoolValue(0, 0, sollValue.BoolValue);
      3: sheet.WriteErrorValue(0, 0, sollValue.ErrorValue);
    end;
    cell := sheet.FindCell(0, 0);
    CheckEquals(true, cell <> nil, 'Cell A1 not found before saving');
    case ATest of
      0: begin
           x := sheet.ReadAsNumber(0, 0);
           CheckEquals(sollValue.NumberValue, x, 'Cell number content mismatch before saving');
         end;
      1: begin
           s := sheet.ReadAsText(0, 0);
           CheckEquals(sollValue.TextValue, s, 'Cell string content mismatch before saving');
         end;
      2: begin
           b := cell^.BoolValue;
           CheckEquals(sollValue.BoolValue, b, 'Cell boolean content mismatch before saving');
         end;
      3: begin
           err := cell^.ErrorValue;
           CheckEquals(ord(sollValue.ErrorValue), ord(err), 'Cell error ontent mismatch before saving');
         end;
    end;
    book.WriteToFile(tempFile, AFormat, true);
  finally
    book.Free;
  end;

  book := TsWorkbook.Create;
  try
    book.Options := book.Options + [boReadFormulas, boAutoCalc];
    book.ReadFromFile(tempFile, AFormat);
    sheet := book.GetWorksheetByIndex(0);
    cell := sheet.FindCell(0, 0);
    CheckEquals(true, cell <> nil, 'Cell A1 not found after reading');
    case ATest of
      0: begin
           x := sheet.ReadAsNumber(Cell);
           CheckEquals(sollValue.NumberValue, x, 'Cell number content mismatch before saving');
         end;
      1: begin
           s := sheet.ReadAsText(cell);
           CheckEquals(sollValue.TextValue, s, 'Cell string content mismatch before saving');
         end;
      2: begin
           b := cell^.BoolValue;
           CheckEquals(sollValue.BoolValue, b, 'Cell boolean content mismatch before saving');
         end;
      3: begin
           err := cell^.ErrorValue;
           CheckEquals(ord(sollValue.ErrorValue), ord(err), 'Cell error ontent mismatch before saving');
         end;
    end;
  finally
    book.Free;
    DeleteFile(tempFile);
  end;
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Number_BIFF2;
begin
  Test_OverwriteFormulaTest(0, sfExcel2);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Number_BIFF5;
begin
  Test_OverwriteFormulaTest(0, sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Number_BIFF8;
begin
  Test_OverwriteFormulaTest(0, sfExcel8);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Number_OOXML;
begin
  Test_OverwriteFormulaTest(0, sfOOXML);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Number_XML;
begin
  Test_OverwriteFormulaTest(0, sfExcelXML);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Number_ODS;
begin
  Test_OverwriteFormulaTest(0, sfOpenDocument);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Text_BIFF2;
begin
  Test_OverwriteFormulaTest(1, sfExcel2);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Text_BIFF5;
begin
  Test_OverwriteFormulaTest(1, sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Text_BIFF8;
begin
  Test_OverwriteFormulaTest(1, sfExcel8);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Text_OOXML;
begin
  Test_OverwriteFormulaTest(1, sfOOXML);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Text_XML;
begin
  Test_OverwriteFormulaTest(1, sfExcelXML);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Text_ODS;
begin
  Test_OverwriteFormulaTest(1, sfOpenDocument);
end;


procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Bool_BIFF2;
begin
  Test_OverwriteFormulaTest(2, sfExcel2);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Bool_BIFF5;
begin
  Test_OverwriteFormulaTest(2, sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Bool_BIFF8;
begin
  Test_OverwriteFormulaTest(2, sfExcel8);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Bool_OOXML;
begin
  Test_OverwriteFormulaTest(2, sfOOXML);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Bool_XML;
begin
  Test_OverwriteFormulaTest(2, sfExcelXML);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Bool_ODS;
begin
  Test_OverwriteFormulaTest(2, sfOpenDocument);
end;


procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Error_BIFF2;
begin
  Test_OverwriteFormulaTest(3, sfExcel2);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Error_BIFF5;
begin
  Test_OverwriteFormulaTest(3, sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Error_BIFF8;
begin
  Test_OverwriteFormulaTest(3, sfExcel8);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Error_OOXML;
begin
  Test_OverwriteFormulaTest(3, sfOOXML);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Error_XML;
begin
  Test_OverwriteFormulaTest(3, sfExcelXML);
end;

procedure TSpreadWriteReadFormulaTests.Test_OverwriteFormula_Error_ODS;
begin
  Test_OverwriteFormulaTest(3, sfOpenDocument);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadWriteReadFormulaTests);


end.

