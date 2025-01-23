unit calcformulatests;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpstypes, fpspreadsheet, fpsexprparser;

type

  TCalcFormulaTests = class(TTestCase)
  private
    FWorkbook: TsWorkbook;
    FWorksheet: TsWorksheet;
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure Test_ABS;
    procedure Test_CEILING;
    procedure Test_DATE;
    procedure Test_EVEN;
    procedure Test_FLOOR;
    procedure Test_ISERROR;
    procedure Test_MATCH;
    procedure Test_ROUND;
    procedure Test_TIME;
  end;

implementation

procedure TCalcFormulaTests.Setup;
begin
  FWorkbook := TsWorkbook.Create;
  FWorksheet := FWorkbook.AddWorksheet('Sheet1');
end;

procedure TCalcFormulaTests.TearDown;
begin
  FWorkbook.Free;
end;

procedure TCalcFormulaTests.Test_ABS;
begin
  // Positive value
  FWorksheet.WriteNumber(0, 0, +10);
  FWorksheet.WriteFormula(0, 1, 'ABS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula ABS(10) result mismatch');

  // Negative value
  FWorksheet.WriteNumber(0, 0, -10);
  FWorksheet.WriteFormula(0, 1, 'ABS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula ABS(-10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 0, '=1/0');
  FWorksheet.WriteFormula(0, 1, 'ABS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula ABS(1/0) result mismatch');

  // Empty argument
  FWorksheet.WriteBlank(0, 0);
  FWorksheet.WriteFormula(0, 1, 'ABS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula ABS([blank_cell]) result mismatch');
end;

procedure TCalcFormulaTests.Test_CEILING;
begin
  // Examples from https://exceljet.net/functions/ceiling-function
  FWorksheet.WriteFormula(0, 1, '=CEILING(10,3)');
  FWorksheet.CalcFormulas;
  CheckEquals(12, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 CEILING(10,3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=CEILING(36,7)');
  FWorksheet.CalcFormulas;
  CheckEquals(42, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 CEILING(36,7) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=CEILING(610,100)');
  FWorksheet.CalcFormulas;
  CheckEquals(700, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 CEILING(610,100) result mismatch');

  // Negative arguments
  FWorksheet.WriteFormula(0, 1, '=CEILING(-5.4,-1)');
  FWorksheet.CalcFormulas;
  CheckEquals(-5, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 CEILING(-5.4,-1) result mismatch');

  // Zero significance
  FWorksheet.WriteFormula(0, 1, '=CEILING(-5.4,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #5 CEILING(-5.4,0) result mismatch');

  // Different signs of the arguments
  FWorksheet.WriteFormula(0, 1, '=CEILING(-5.4,1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #6 CEILING(-5.4,1) result mismatch');

  // Arguments as string
  FWorksheet.WriteFormula(0, 1, '=CEILING("A",1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 CEILING("A",1) result mismatch');
  FWorksheet.WriteFormula(0, 1, '=CEILING(5.4,"A")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #8 CEILING(5.4,"A") result mismatch');

  // Arguments as boolean
  FWorksheet.WriteFormula(0, 1, '=CEILING(TRUE(),1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #9 CEILING(TRUE(),1) result mismatch');
  FWorksheet.WriteFormula(0, 1, '=CEILING(5.4, TRUE())');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #10 CEILING(5.4, TRUE()) result mismatch');

  // Arguments with errors
  FWorksheet.WriteFormula(0, 1, '=CEILING(1/0,1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #11 CEILING(1/0, 1) result mismatch');
  FWorksheet.WriteFormula(0, 1, '=CEILING(5.4, 1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #12 CEILING(5.4, 1/0) result mismatch');
end;


procedure TCalcFormulaTests.Test_DATE;
var
  actualDate, expectedDate: TDate;
begin
  // Normal date
  FWorksheet.WriteFormula(0, 1, '=DATE(2025,1,22)');
  FWorksheet.CalcFormulas;
  expectedDate := EncodeDate(2025, 1, 22);
  FWorksheet.ReadAsDateTime(0, 1, actualDate);
  CheckEquals(DateToStr(expectedDate), DateToStr(actualDate), '#1 Formula DATE(2025,1,22) result mismatch');

  // Two-digit year
  FWorksheet.WriteFormula(0, 1, '=DATE(90,1,22)');
  FWorksheet.CalcFormulas;
  expectedDate := EncodeDate(1990, 1, 22);
  FWorksheet.ReadAsDateTime(0, 1, actualDate);
  CheckEquals(DateToStr(expectedDate), DateToStr(actualDate), '#2 Formula DATE(90,1,22) result mismatch');

  // Negative year
  FWorksheet.WriteFormula(0, 1, '=DATE(-2000,1,22)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), '#3 Formula DATE(90,1,22) result mismatch');

  // Too-large year
  FWorksheet.WriteFormula(0, 1, '=DATE(10000,1,22)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), '#4 Formula DATE(10000,1,22) result mismatch');

  // Month > 12
  FWorksheet.WriteFormula(0, 1, '=DATE(2008,14,2)');
  FWorksheet.CalcFormulas;
  expectedDate := EncodeDate(2009, 2, 2);
  FWorksheet.ReadAsDateTime(0, 1, actualDate);
  CheckEquals(DateToStr(expectedDate), DateToStr(actualDate), '#5 Formula DATE(2008,14,2) result mismatch');

  // Month < 1
  FWorksheet.WriteFormula(0, 1, '=DATE(2008,-3,2)');
  FWorksheet.CalcFormulas;
  expectedDate := EncodeDate(2007, 9, 2);
  FWorksheet.ReadAsDateTime(0, 1, actualDate);
  CheckEquals(DateToStr(expectedDate), DateToStr(actualDate), '#6 Formula DATE(2008,-3,2) result mismatch');

  // Day > Days in month
  FWorksheet.WriteFormula(0, 1, '=DATE(2008,1,35)');
  FWorksheet.CalcFormulas;
  expectedDate := EncodeDate(2008, 2, 4);
  FWorksheet.ReadAsDateTime(0, 1, actualDate);
  CheckEquals(DateToStr(expectedDate), DateToStr(actualDate), '#7 Formula DATE(2008,1,35) result mismatch');

  // Day < 1
  FWorksheet.WriteFormula(0, 1, '=DATE(2008,1,-15)');
  FWorksheet.CalcFormulas;
  expectedDate := EncodeDate(2007, 12, 16);
  FWorksheet.ReadAsDateTime(0, 1, actualDate);
  CheckEquals(DateToStr(expectedDate), DateToStr(actualDate), '#8 Formula DATE(2008,1,-15) result mismatch');

  // Month > 12 and Day > Days in month
  FWorksheet.WriteFormula(0, 1, '=DATE(2008,14,50)');
  FWorksheet.CalcFormulas;
  expectedDate := EncodeDate(2009, 3, 22);
  FWorksheet.ReadAsDateTime(0, 1, actualDate);
  CheckEquals(DateToStr(expectedDate), DateToStr(actualDate), '#9 Formula DATE(2008,14,50) result mismatch');

  // Month > 12 and Day < 1
  FWorksheet.WriteFormula(0, 1, '=DATE(2008,14,-10)');
  FWorksheet.CalcFormulas;
  expectedDate := EncodeDate(2009, 1, 21);
  FWorksheet.ReadAsDateTime(0, 1, actualDate);
  CheckEquals(DateToStr(expectedDate), DateToStr(actualDate), '#10 Formula DATE(2008,14,-10) result mismatch');

  // Month < 1 and Day > Days in month
  FWorksheet.WriteFormula(0, 1, '=DATE(2008,-3,50)');
  FWorksheet.CalcFormulas;
  expectedDate := EncodeDate(2007,10,20);
  FWorksheet.ReadAsDateTime(0, 1, actualDate);
  CheckEquals(DateToStr(expectedDate), DateToStr(actualDate), '#11 Formula DATE(2008,-3,50) result mismatch');

  // Month < 1 and Day < 1 in month
  FWorksheet.WriteFormula(0, 1, '=DATE(2008,-3,-10)');
  FWorksheet.CalcFormulas;
  expectedDate := EncodeDate(2007,8,21);
  FWorksheet.ReadAsDateTime(0, 1, actualDate);
  CheckEquals(DateToStr(expectedDate), DateToStr(actualDate), '#12 Formula DATE(2008,-3,-10) result mismatch');

  // Error in year
  FWorksheet.WriteFormula(0, 1, '=DATE(1/0,1,22)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), '#13 Formula DATE(1/0,1,22) result mismatch');

  // Error in month
  FWorksheet.WriteFormula(0, 1, '=DATE(2025, 1/0, 22)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), '#14 Formula DATE(2025, 1/0, 22) result mismatch');

  // Error in day
  FWorksheet.WriteFormula(0, 1, '=DATE(2025, 1, 1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), '#15 Formula DATE(2025, 1, 1/0) result mismatch');
end;

procedure TCalcFormulaTests.Test_EVEN;
begin
  FWorksheet.WriteFormula(0, 1, '=EVEN(1.23)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula EVEN(1.23) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=EVEN(2.34)');
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(0, 1), 'Formula EVEN(2.34) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=EVEN(-1.23)');
  FWorksheet.CalcFormulas;
  CheckEquals(-2, FWorksheet.ReadAsNumber(0, 1), 'Formula EVEN(-1.23) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=EVEN(-2.34)');
  FWorksheet.CalcFormulas;
  CheckEquals(-4, FWorksheet.ReadAsNumber(0, 1), 'Formula EVEN(-2.34) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=EVEN(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula EVEN(1/0) result mismatch');
end;

procedure TCalcFormulaTests.Test_FLOOR;
begin
  FWorksheet.WriteFormula(0, 1, '=FLOOR(10,3)');
  FWorksheet.CalcFormulas;
  CheckEquals(9, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 FLOOR(10,3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=FLOOR(36,7)');
  FWorksheet.CalcFormulas;
  CheckEquals(35, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 FLOOR(36,7) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=FLOOR(610,100)');
  FWorksheet.CalcFormulas;
  CheckEquals(600, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 FLOOR(610,100) result mismatch');

  // Negative value, negative significance
  FWorksheet.WriteFormula(0, 1, '=FLOOR(-5.4,-2)');
  FWorksheet.CalcFormulas;
  CheckEquals(-4, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 FLOOR(-5.4,-2) result mismatch');

  // Negative value, positive significance
  FWorksheet.WriteFormula(0, 1, '=FLOOR(-5.4,2)');
  FWorksheet.CalcFormulas;
  CheckEquals(-6, FWorksheet.ReadAsNumber(0, 1), 'Formula #5 FLOOR(-5.4,2) result mismatch');

  // Positive value, negative significance
  FWorksheet.WriteFormula(0, 1, '=FLOOR(5.4,-2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #6 FLOOR(5.4,-2) result mismatch');

  // Zero significance
  FWorksheet.WriteFormula(0, 1, '=FLOOR(-5.4,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 FLOOR(-5.4,0) result mismatch');

  // Arguments as string
  FWorksheet.WriteFormula(0, 1, '=FLOOR("A",1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #8 FLOOR("A",1) result mismatch');
  FWorksheet.WriteFormula(0, 1, '=FLOOR(5.4,"A")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #9 FLOOR(5.4,"A") result mismatch');

  // Arguments as boolean
  FWorksheet.WriteFormula(0, 1, '=FLOOR(TRUE(),1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #10 FLOOR(TRUE(),1) result mismatch');
  FWorksheet.WriteFormula(0, 1, '=FLOOR(5.4, TRUE())');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #11 FLOOR(5.4, TRUE()) result mismatch');

  // Arguments with errors
  FWorksheet.WriteFormula(0, 1, '=FLOOR(1/0,1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #12 FLOOR(1/0, 1) result mismatch');
  FWorksheet.WriteFormula(0, 1, '=FLOOR(5.4, 1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #13 FLOOR(5.4, 1/0) result mismatch');

end;

procedure TCalcFormulaTests.Test_ISERROR;
var
  res: Boolean;
begin
  // Hard coded expression with error
  FWorksheet.WriteFormula(0, 1, '=ISERROR(1/0)');
  FWorksheet.CalcFormulas;
  res := FWorksheet.IsTrueValue(FWorksheet.FindCell(0, 1));
  CheckEquals(true, res, 'Formula #1 ISERROR(1/0) result mismatch');

  // Hard coded expression without error
  FWorksheet.WriteFormula(0, 1, '=ISERROR(0/1)');
  FWorksheet.CalcFormulas;
  res := FWorksheet.IsTrueValue(FWorksheet.FindCell(0, 1));
  CheckEquals(false, res, 'Formula #2 ISERROR(0/1) result mismatch');

  // Reference to cell with error
  FWorksheet.WriteFormula(0, 0, '=1/0');
  FWorksheet.WriteFormula(0, 1, '=ISERROR(A1)');
  FWorksheet.CalcFormulas;
  res := FWorksheet.IsTrueValue(FWorksheet.FindCell(0, 1));
  CheckEquals(true, res, 'Formula #3 ISERROR(A1) result mismatch');

  // Reference to cell without error
  FWorksheet.WriteText(0, 0, 'abc');
  FWorksheet.WriteFormula(0, 1, '=ISERROR(A1)');
  FWorksheet.CalcFormulas;
  res := FWorksheet.IsTrueValue(FWorksheet.FindCell(0, 1));
  CheckEquals(false, res, 'Formula #4 ISERROR(A1) result mismatch');
end;

procedure TCalcFormulaTests.Test_MATCH;
begin
  // *** Match_Type 0, unsorted data in search range

  // Search range to be checked: B1:B4
  FWorksheet.WriteNumber(0, 1, 10);
  FWorksheet.WriteNumber(1, 1, 20);
  FWorksheet.WriteNumber(2, 1, 30);
  FWorksheet.WriteNumber(3, 1, 15);

  // Search for constant, contained in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(10, B1:B4, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #1 MATCH mismatch, match_type 0, in range');

  // Search for constant, below search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(0, B1:B4, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #2 MATCH mismatch, match_type 0, below range');

  // Search for constant, above search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(90, B1:B4, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #3 MATCH mismatch, match_type 0, above range');

  // Search for cell with value in range
  FWorksheet.WriteNumber(0, 0, 20);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B4, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #4 MATCH mismatch, match_type 0, cell in range');
  FWorksheet.WriteBlank(0, 0);

  // Search for cell, but cell is empty
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B4, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #5 MATCH mismatch, match_type 0, empty cell');

  // Search range is empty
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, D1:D3, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #6 MATCH mismatch, match_type -1, empty search range');


  // *** Match_Type 1 (find largest value in range <= value), ascending values in search range

  // Search range to be checked: B1:B3
  FWorksheet.WriteNumber(0, 1, 10);
  FWorksheet.WriteNumber(1, 1, 20);
  FWorksheet.WriteNumber(2, 1, 30);
  FWorksheet.WriteBlank(3, 1);

  // Search for constant, contained in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, B1:B3, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #7 MATCH mismatch, match_type 1, in range');

  // Search for constant,  below search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(8, B1:B3, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #8 MATCH mismatch, match_type 1, below range');

  // Search for constant, above search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(123, B1:B3, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 2), 'Formula MATCH #9 mismatch, match_type 1, above range');

  // Search for cell with value in range
  FWorksheet.WriteNumber(0, 0, 28);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B3, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula MATCH #10 mismatch, match_type 1, cell in range');
  FWorksheet.WriteBlank(0, 0);

  // Search for cell, but cell is empty
  FWorksheet.WriteBlank(0, 0);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B3, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #11 MATCH mismatch, match_type 1, empty cell');

  // Search range is empty
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, D1:D3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula MATCH #12 mismatch, match_type -1, empty search range');


  // *** Match_Type -1 (find smallest value in range >= value), descending values in search range

  // Search range to be checked: B1:B3
  FWorksheet.WriteNumber(0, 1, 30);
  FWorksheet.WriteNumber(1, 1, 20);
  FWorksheet.WriteNumber(2, 1, 10);

  // Search for constant, contained in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, B1:B3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #13 MATCH mismatch, match_type -1, in range');

  // Search for constant,  below search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(8, B1:B3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 2), 'Formula #14 MATCH mismatch, match_type -1, below range');

  // Search for constant, above search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(123, B1:B3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #15 MATCH mismatch, match_type -1, above range');

  // Search for cell with value in range
  FWorksheet.WriteNumber(0, 0, 28);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #16 MATCH mismatch, match_type -1, cell in range');
  FWorksheet.WriteBlank(0, 0);

  // Search for cell, but cell is empty
  FWorksheet.WriteBlank(0, 0);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #17 MATCH mismatch, match_type -1, empty cell');

  // Search range is empty
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, D1:D3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #18 MATCH mismatch, match_type -1, empty search range');


  // **** Error propagation

  // Search for cell, but cell contains error
  FWorksheet.WriteFormula(0, 0, '=1/0');
  FWorksheet.WriteNumber(1, 1, 20);
  FWorksheet.WriteNumber(2, 1, 30);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B4, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #19 MATCH mismatch, match_type 0, error cell');

  // Match_type parameter contains error
  FWorksheet.WriteNumber(0, 1, 10);
  FWorksheet.WriteFormula(0, 5, '=1/0');    // F1
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B3, F1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #20 MATCH mismatch, match_type 0, error in search range');

  // Cell range contains error
  FWorksheet.WriteNumber(0, 1, 10);
  FWorksheet.WriteFormula(1, 1, '=1/0');    // B2 contains a #DIV/0! error now
  FWorksheet.WriteNumber(2, 1, 30);
  // Search for constant, contained in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(20, B1:B3, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #21 MATCH mismatch, match_type 0, error in search range');
    // ArgError because search value is not found
end;

procedure TCalcFormulaTests.Test_ROUND;
begin
  // Round positive value.
  FWorksheet.WriteFormula(0, 1, '=ROUND(123.432, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(123.4, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #1 ROUND(123.432,1) result mismatch');

  // Round positive value. Check that Banker's rounding is not applied
  FWorksheet.WriteFormula(0, 1, '=ROUND(123.456, 2)');
  FWorksheet.CalcFormulas;
  CheckEquals(123.46, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #2 ROUND(123.3456,2) result mismatch');

  // Round negative value.
  FWorksheet.WriteFormula(0, 1, '=ROUND(-123.432, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(-123.4, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #3 ROUND(-123.432,1) result mismatch');

  // Round negative value. Check that Banker's rounding is not applied
  FWorksheet.WriteFormula(0, 1, '=ROUND(-123.456, 2)');
  FWorksheet.CalcFormulas;
  CheckEquals(-123.46, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #4 ROUND(-123.456,2) result mismatch');

  // Negative number of decimals for positive value
  FWorksheet.WriteFormula(0, 1, '=ROUND(123.456, -2)');
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(0, 1), 'Formula #5 ROUND(123.3456,-2) result mismatch');

  // Negative number of decimals for negative value
  FWorksheet.WriteFormula(0, 1, '=ROUND(-123.456, -2)');
  FWorksheet.CalcFormulas;
  CheckEquals(-100, FWorksheet.ReadAsNumber(0, 1), 'Formula #6 ROUND(123.3456,-2) result mismatch');

  // Error in 1st argument
  FWorksheet.WriteFormula(0, 1, '=Round(1/0, 2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #7 ROUND(1/0,2) result mismatch');

  // Error in 2nd argument
  FWorksheet.WriteFormula(0, 1, '=Round(123.456, 1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 ROUND(123.456, 1/0) result mismatch');
end;

procedure TCalcFormulaTests.Test_TIME;
var
  actualTime, expectedTime: TTime;
begin
  // Normal time
  FWorksheet.WriteFormula(0, 1, '=Time(6,32,57)');
  FWorksheet.CalcFormulas;
  expectedTime := EncodeTime(6, 32, 57, 0);
  FWorksheet.ReadAsDateTime(0, 1, actualTime);
  CheckEquals(TimeToStr(expectedTime), TimeToStr(actualTime), 'Formula #1 TIME(6,32,57) result mismatch');

  // Hours < 0
  FWorksheet.WriteFormula(0, 1, '=Time(-6,32,57)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #2 TIME(-6,32,57) result mismatch');

  // Hours > 23
  FWorksheet.WriteFormula(0, 1, '=Time(15,32,57)');
  FWorksheet.CalcFormulas;
  expectedTime := 0.647881944;     // Value read from Excel
  FWorksheet.ReadAsDateTime(0, 1, actualTime);
  CheckEquals(TimeToStr(expectedTime), TimeToStr(actualTime), 'Formula #3 TIME(15,32,57) result mismatch');

  // Minutes > 59
  FWorksheet.WriteFormula(0, 1, '=Time(6,100,57)');
  FWorksheet.CalcFormulas;
  expectedTime := 0.320104167;     // Value read from Excel
  FWorksheet.ReadAsDateTime(0, 1, actualTime);
  CheckEquals(TimeToStr(expectedTime), TimeToStr(actualTime), 'Formula #4 TIME(6,100,57) result mismatch');

  // Minutes < 0
  FWorksheet.WriteFormula(0, 1, '=Time(6,-100,57)');
  FWorksheet.CalcFormulas;
  expectedTime := 0.181215278;     // Value read from Excel
  FWorksheet.ReadAsDateTime(0, 1, actualTime);
  CheckEquals(TimeToStr(expectedTime), TimeToStr(actualTime), 'Formula #5 TIME(6,-100,57) result mismatch');

  // Seconds > 59
  FWorksheet.WriteFormula(0, 1, '=Time(6,32,100)');
  FWorksheet.CalcFormulas;
  expectedTime := 0.27337963;     // Value read from Excel
  FWorksheet.ReadAsDateTime(0, 1, actualTime);
  CheckEquals(TimeToStr(expectedTime), TimeToStr(actualTime), 'Formula #6 TIME(6,32,100) result mismatch');

  // Seconds < 0
  FWorksheet.WriteFormula(0, 1, '=Time(6,32,-100)');
  FWorksheet.CalcFormulas;
  expectedTime := 0.271064815;     // Value read from Excel
  FWorksheet.ReadAsDateTime(0, 1, actualTime);
  CheckEquals(TimeToStr(expectedTime), TimeToStr(actualTime), 'Formula #7 TIME(6,32,-100) result mismatch');

  // Error in hours
  FWorksheet.WriteFormula(0, 1, '=Time(1/0,32,57)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 TIME(1/0,32,57) result mismatch');

  // Error in minutes
  FWorksheet.WriteFormula(0, 1, '=Time(6,1/0,57)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #9 TIME(6,1/0,57) result mismatch');

  // Error in seconds
  FWorksheet.WriteFormula(0, 1, '=Time(6,32,1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #10 TIME(6,32,1/0) result mismatch');
end;

initialization
  RegisterTest(TCalcFormulaTests);
end.

