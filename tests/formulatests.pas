unit formulatests;

{$mode objfpc}{$H+}

{ Deactivate this define in order to bypass tests which will raise an exception
  when the corresponding rpn formula is calculated. }
{.$DEFINE ENABLE_CALC_RPN_EXCEPTIONS}

{ Deactivate this define to include errors in the structure of the rpn formulas.
  Note that Excel report a corrupted file when trying to read this file }
{.DEFINE ENABLE_DEFECTIVE_FORMULAS }


interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpsallformats, fpspreadsheet, fpsfunc,
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
    // Test formula strings
    procedure TestWriteReadFormulaStrings(AFormat: TsSpreadsheetFormat);
    // Test calculation of rpn formulas
    procedure TestCalcRPNFormulas(AFormat: TsSpreadsheetformat);

  published
    // Writes out numbers & reads back.
    // If previous read tests are ok, this effectively tests writing.
    { BIFF2 Tests }
    procedure TestWriteRead_BIFF2_FormulaStrings;
    { BIFF5 Tests }
    procedure TestWriteRead_BIFF5_FormulaStrings;
    { BIFF8 Tests }
    procedure TestWriteRead_BIFF8_FormulaStrings;

    // Writes out and calculates formulas, read back
    { BIFF2 Tests }
    procedure TestWriteRead_BIFF2_CalcRPNFormula;
    { BIFF5 Tests }
    procedure TestWriteRead_BIFF5_CalcRPNFormula;
    { BIFF8 Tests }
    procedure TestWriteRead_BIFF8_CalcRPNFormula;
  end;

implementation

uses
  math, typinfo, lazUTF8, fpsUtils, rpnFormulaUnit;

{ TSpreadWriteReadFormatTests }

procedure TSpreadWriteReadFormulaTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadFormulaTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadFormulaTests.TestWriteReadFormulaStrings(AFormat: TsSpreadsheetFormat);
const
  SHEET = 'Sheet1';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row: Integer;
  TempFile: string; //write xls/xml to this file and read back from it
  expected: String;
  actual: String;
  cell: PCell;
begin
  TempFile := GetTempFileName;

  // Create test workbook
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);

    // Write out all test formulas
    // All formulas are in column B
    WriteRPNFormulaSamples(MyWorksheet, AFormat, true);
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFormulas := true;

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
      if (cell <> nil) and (Length(cell^.RPNFormulaValue) > 0) then begin
        actual := MyWorksheet.ReadRPNFormulaAsString(cell);
        expected := MyWorksheet.ReadAsUTF8Text(Row, 0);
        CheckEquals(expected, actual, 'Test read formula mismatch, cell '+CellNotation(MyWorkSheet,Row,1));
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormulaTests.TestWriteRead_BIFF2_FormulaStrings;
begin
  TestWriteReadFormulaStrings(sfExcel2);
end;

procedure TSpreadWriteReadFormulaTests.TestWriteRead_BIFF5_FormulaStrings;
begin
  TestWriteReadFormulaStrings(sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.TestWriteRead_BIFF8_FormulaStrings;
begin
  TestWriteReadFormulaStrings(sfExcel8);
end;


{ Test calculation of rpn formulas }

procedure TSpreadWriteReadFormulaTests.TestCalcRPNFormulas(AFormat: TsSpreadsheetFormat);
const
  SHEET = 'Sheet1';
  STATS_NUMBERS: Array[0..4] of Double = (1.0, 1.1, 1.2, 0.9, 0.8);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row: Integer;
  TempFile: string;    //write xls/xml to this file and read back from it
  actual: TsArgument;
  expected: TsArgument;
  cell: PCell;
  sollValues: array of TsArgument;
  formula: String;
  s: String;
  t: TTime;
  hr,min,sec,msec: Word;
  ErrorMargin: double;
  k: Integer;
  { When comparing soll and formula values we must make sure that the soll
    values are calculated from double precision numbers, they are used in
    the formula calculation as well. The next variables, along with STATS_NUMBERS
    above, hold the arguments for the direction function calls. }
  number: Double;
  numberArray: array[0..4] of Double;
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
    MyWorkSheet:= MyWorkBook.AddWorksheet(SHEET);
    MyWorkSheet.Options := MyWorkSheet.Options + [soCalcBeforeSaving];
    // Calculation of rpn formulas must be activated explicitly!

    { Write out test formulas.
      This include file creates various rpn formulas and stores the expected
      results in array "sollValues".
      The test file contains the text representation in column A, and the
      formula in column B. }
    Row := 0;
    TempFile:=GetTempFileName;
    {$I testcases_calcrpnformula.inc}
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the workbook
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
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
        cctBool       : actual := CreateBoolArg(cell^.BoolValue);
        cctNumber     : actual := CreateNumberArg(cell^.NumberValue);
        cctError      : actual := CreateErrorArg(cell^.ErrorValue);
        cctUTF8String : actual := CreateStringArg(cell^.UTF8StringValue);
        else            fail('ContentType not supported');
      end;
      expected := SollValues[row];
      CheckEquals(ord(expected.ArgumentType), ord(actual.ArgumentType),
        'Test read calculated formula data type mismatch, formula "' + formula +
        '", cell '+CellNotation(MyWorkSheet,Row,1));

      // The now function result is volatile, i.e. changes continuously. The
      // time for the soll value was created such that we can expect to have
      // the file value in the same second. Therefore we neglect the milliseconds.
      if formula = '=NOW()' then begin
        // Round soll value to seconds
        DecodeTime(expected.NumberValue, hr,min,sec,msec);
        expected.NumberValue := EncodeTime(hr, min, sec, 0);
        // Round formula value to seconds
        DecodeTime(actual.NumberValue, hr,min,sec,msec);
        actual.NumberValue := EncodeTime(hr,min,sec,0);
      end;

      case actual.ArgumentType of
        atBool:
          CheckEquals(BoolToStr(expected.BoolValue), BoolToStr(actual.BoolValue),
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
        atNumber:
          {$if (defined(mswindows)) or (FPC_FULLVERSION>=20701)}
          // FPC 2.6.x and trunk on Windows need this, also FPC trunk on Linux x64
          CheckEquals(expected.NumberValue, actual.NumberValue, ErrorMargin,
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
          {$else}
          // Non-Windows: test without error margin
          CheckEquals(expected.NumberValue, actual.NumberValue,
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
          {$endif}
        atString:
          CheckEquals(expected.StringValue, actual.StringValue,
            'Test read calculated formula result mismatch, formula "' + formula +
            '", cell '+CellNotation(MyWorkSheet,Row,1));
        atError:
          CheckEquals(
            GetEnumName(TypeInfo(TsErrorValue), ord(expected.ErrorValue)),
            GetEnumname(TypeInfo(TsErrorValue), ord(actual.ErrorValue)),
            'Test read calculated formula error value mismatch, formula ' + formula +
            ', cell '+CellNotation(MyWorkSheet,Row,1));
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadFormulaTests.TestWriteRead_BIFF2_CalcRPNFormula;
begin
  TestCalcRPNFormulas(sfExcel2);
end;

procedure TSpreadWriteReadFormulaTests.TestWriteRead_BIFF5_CalcRPNFormula;
begin
  TestCalcRPNFormulas(sfExcel5);
end;

procedure TSpreadWriteReadFormulaTests.TestWriteRead_BIFF8_CalcRPNFormula;
begin
  TestCalcRPNFormulas(sfExcel8);
end;


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadWriteReadFormulaTests);


end.

