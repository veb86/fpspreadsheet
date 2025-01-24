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
    procedure Test_AVEDEV;
    procedure Test_AVERAGE;
    procedure Test_CEILING;
    procedure Test_COUNT;
    procedure Test_DATE;
    procedure Test_ERRORTYPE;
    procedure Test_EVEN;
    procedure Test_FLOOR;
    procedure Test_IF;
    procedure Test_IFERROR;
    procedure Test_ISBLANK;
    procedure Test_ISERR;
    procedure Test_ISERROR;
    procedure Test_ISLOGICAL;
    procedure Test_ISNA;
    procedure Test_ISNONTEXT;
    procedure Test_ISNUMBER;
    procedure Test_ISREF;
    procedure Test_ISTEXT;
    procedure Test_MATCH;
    procedure Test_MAX;
    procedure Test_MIN;
    procedure Test_PRODUCT;
    procedure Test_ROUND;
    procedure Test_STDEV;
    procedure Test_STDEVP;
    procedure Test_SUM;
    procedure Test_SUMSQ;
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

procedure TCalcFormulaTests.Test_AVEDEV;
begin
  FWorksheet.WriteFormula(0, 1, '=AVEDEV(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 AVEDEV(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(1,-2,-3,4)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.5, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 AVEDEV(1,-2,-3,4) result mismatch');

  FWorksheet.WriteNumber(0, 0, 1);
  FWorksheet.WriteNumber(1, 0, -2);
  FWorksheet.WriteNumber(2, 0, -3);
  FWorksheet.WriteNumber(3, 0, 4);

  FWorksheet.WriteFormula(4, 0, '=AVEDEV(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(4, 0), 'Formula #3 AVEDEV(A1) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=AVEDEV(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.5, FWorksheet.ReadAsNumber(4, 0), 'Formula #4 AVEDEV(A1,A2) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=AVEDEV(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.5, FWorksheet.ReadAsNumber(4, 0), 'Formula #5 AVEDEV(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=AVEDEV(A1,A2:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.5, FWorksheet.ReadAsNumber(4, 0), 'Formula #6 AVEDEV(A1,A2:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=AVEDEV(A1, 1/0, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(4, 0), 'Formula #7 AVEDEV(A1, 1/0, A2) result mismatch');
end;

procedure TCalcFormulaTests.Test_AVERAGE;
begin
  FWorksheet.WriteFormula(0, 1, '=AVERAGE(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 AVERAGE(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(1,2,3,4)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.5, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 AVERAGE(1,2,3,4) result mismatch');

  FWorksheet.WriteNumber(0, 0, 1);
  FWorksheet.WriteNumber(1, 0, 2);
  FWorksheet.WriteNumber(2, 0, 3);
  FWorksheet.WriteNumber(3, 0, 4);

  FWorksheet.WriteFormula(4, 0, '=AVERAGE(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(4, 0), 'Formula #3 AVERAGE(A1) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=AVERAGE(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.5, FWorksheet.ReadAsNumber(4, 0), 'Formula #4 AVERAGE(A1,A2) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=AVERAGE(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.5, FWorksheet.ReadAsNumber(4, 0), 'Formula #5 AVERAGE(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=AVERAGE(A1,A2:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.5, FWorksheet.ReadAsNumber(4, 0), 'Formula #6 AVERAGE(A1,A2:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=AVERAGE(A1, 1/0, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(4, 0), 'Formula #7 AVERAGE(A1, 1/0, A2) result mismatch');
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

procedure TCalcFormulaTests.Test_COUNT;
begin
{
  FWorksheet.WriteFormula(0, 1, '=COUNT("")');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 COUNT("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 COUNT(10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(20,10,"abc",40)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 COUNT(20,10,"abc",40) result mismatch');
  }

  FWorksheet.WriteNumber(0, 0, 20);
  FWorksheet.WriteNumber(1, 0, 10);
  FWorksheet.WriteText(2, 0, 'abc');
  FWorksheet.WriteNumber(3, 0, 40);

  FWorksheet.WriteFormula(4, 1, '=COUNT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(4, 1), 'Formula #4 COUNT(A1) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNT(A10)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(4, 1), 'Formula #5 COUNT(A10) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNT(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(4, 1), 'Formula #6 COUNT(A1,A2) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNT(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(4, 1), 'Formula #7 COUNT(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNT(A1:A10)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(4, 1), 'Formula #8 COUNT(A1:A10) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNT(A1,A2:A10)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(4, 1), 'Formula #9 COUNT(A1,A2:A10) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNT(A1, 1/0, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(4, 1), 'Formula #10 COUNT(A1, 1/0, A2) result mismatch');
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

procedure TCalcFormulaTests.Test_ERRORTYPE;
begin
  // No error
  FWorksheet.WriteNumber(0, 0, 123);
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 1), 'Formula #1 ERROR.TYPE (no error!) result mismatch');

  // #NULL! error
  FWorksheet.WriteNumber(0, 0, 12);
  FWorksheet.WriteNumber(1, 0, -2);

  // ToDo: Space as argument separator not detected correctly!
{
  This currently is not handled by FPS...
  FWorksheet.WriteFormula(2, 0, '=SUM(A1 A2)');  // missing comma --> #NULL!
  FWorksheet.WriteFormula(2, 0, '=ERROR.TYPE(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errEmptyIntersection), FWorksheet.ReadAsNumber(2, 0), 'Formula #1 ERROR.TYPE (#NULL!) result mismatch');
}

  // #REF! error
  FWorksheet.WriteFormula(2, 0, '=SUM(A1,A2)');
  FWorksheet.DeleteRow(0);     // This creates the #REF! error in the sum cell A2
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errIllegalRef), FWorksheet.ReadAsNumber(0, 1), 'Formula #1 ERROR.TYPE (#REF!) result mismatch');

  // #VALUE! error
  FWorksheet.WriteText(0, 0, 'a');
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(1+A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errWrongType), FWorksheet.ReadAsNumber(0, 1), 'Formula #1 ERROR.TYPE #VALUE! result mismatch');

  // #DIV/0! error
  FWorksheet.WriteFormula(0, 0, '=1/0');
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errDivideByZero), FWorksheet.ReadAsNumber(0, 1), 'Formula #1 ERROR.TYPE #DIV/0! result mismatch');

  // #NUM! error
  FWorksheet.WriteFormula(0, 0, '=SQRT(-1)');
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errOverflow), FWorksheet.ReadAsNumber(0, 1), 'Formula #1 ERROR.TYPE #NUM! result mismatch');

  // ToDo: Create #NAME? error node when identifier is not found. Parser always raises an exception during scanning - maybe there should be a TsErrorExprNode?

  { --- not correctly detected by FPS parser ...
  // #NAME? error
  FWorksheet.WriteFormula(0, 0, '=S_Q_R_T(-1)');
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errWrongName), FWorksheet.ReadAsNumber(0, 1), 'Formula #1 ERROR.TYPE #NAME? result mismatch');
  }

  // #N/A error
  FWorksheet.WriteNumber(0, 0, 10);
  FWorksheet.WriteNumber(1, 0, 20);
  FWorksheet.WriteFormula(2, 0, '=MATCH(-10,A1:A2,0)');
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errArgError), FWorksheet.ReadAsNumber(0, 1), 'Formula #1 ERROR.TYPE #N/A result mismatch');
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

procedure TCalcFormulaTests.Test_IF;
begin
  FWorksheet.WriteNumber(0, 0, 256.0);

  // 3 arguments
  FWorksheet.WriteFormula(0, 1, '=IF(A1>=100,"ok","not ok")');
  FWorksheet.CalcFormulas;
  CheckEquals('ok', FWorksheet.ReadAsText(0, 1), 'Formula #1 IF(A1>=100,"ok","not ok") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=IF(A1<100,"ok","not ok")');
  FWorksheet.CalcFormulas;
  CheckEquals('not ok', FWorksheet.ReadAsText(0, 1), 'Formula #2 IF(A1<100,"ok","not ok") result mismatch');

  // 2 arguments
  FWorksheet.WriteFormula(0, 1, '=IF(A1>=100,"ok")');
  FWorksheet.CalcFormulas;
  CheckEquals('ok', FWorksheet.ReadAsText(0, 1), 'Formula #3 IF(A1>=100,"ok") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=IF(A1<100,"ok")');
  FWorksheet.CalcFormulas;
  CheckEquals('FALSE', FWorksheet.ReadAsText(0, 1), 'Formula #4 IF(A1<100,"ok") result mismatch');

  // Error propagation: error in 3rd argument
  FWorksheet.WriteFormula(0, 1, '=IF(A1>=100, "ok", 1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #5 IF(A1>=100,"ok",1/0) result mismatch');

  // Error propagation: error in 2nd argument
  FWorksheet.WriteFormula(0, 1, '=IF(A1>=100, 1/0,"not ok")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #6 IF(A1>=100,1/0,"not ok") result mismatch');

  // Error propagaton: error in 1st argument
  FWorksheet.WriteFormula(0, 0, '=1/0');
  FWorksheet.WriteFormula(0, 1, '=IF(A1>=100,"ok","not ok")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #7 IF(A1>=100,"ok","not ok") with A1=1/0 result mismatch');
end;

procedure TCalcFormulaTests.Test_IFERROR;
begin
  FWorksheet.WriteFormula(0, 1, '=IFERROR("abc", "ERROR")');
  FWorksheet.CalcFormulas;
  CheckEquals('abc', FWorksheet.ReadAsText(0, 1), 'Formula #1 IFERROR("abc","ERROR") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=IFERROR(#N/A, "ERROR")');
  FWorksheet.CalcFormulas;
  CheckEquals('ERROR', FWorksheet.ReadAsText(0, 1), 'Formula #2 IFERROR(#N/A,"ERROR") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=IFERROR(#N/A, #DIV/0!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #3 IFERROR(#N/A,#DIV/0!) result mismatch');
end;

procedure TCalcFormulaTests.Test_ISBLANK;
var
  cell: PCell;
begin
  cell := FWorksheet.WriteFormula(0, 1, '=ISBLANK(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 ISBLANK(A1) with A1=blank result mismatch');

  FWorksheet.WriteText(0, 0, '');
  cell := FWorksheet.WriteFormula(0, 1, '=ISBLANK(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #2 ISBLANK(A1) with A1='' result mismatch');

  // No argument
  cell := FWorksheet.WriteFormula(0, 1, '=ISBLANK()');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #3 ISBLANK() result mismatch');

  // String
  cell := FWorksheet.WriteFormula(0, 1, '=ISBLANK("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #4 ISBLANK("abc") result mismatch');

  // Some Excel oddity: an empty string is not "blank"...
  cell := FWorksheet.WriteFormula(0, 1, '=ISBLANK("")');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #5 ISBLANK("") result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 0, '=1/0');
  cell := FWorksheet.WriteFormula(0, 1, '=ISBLANK(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #6 ISBLANK(A1) with A1=1/0 result mismatch');
  CheckNotEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #6 ISBLANK(A1) with A1=1/0 result mismatch');
end;

procedure TCalcFormulaTests.Test_ISERR;
var
  cell: PCell;
begin
  // Hard coded expression with error #DIV/0!
  cell := FWorksheet.WriteFormula(0, 1, '=ISERR(#DIV/0!)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 ISERR(1/0) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, '=ISERR(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #2 ISERR(1/0) result mismatch');

  // Hard coded expression without error
  cell := FWorksheet.WriteFormula(0, 1, '=ISERR(0/1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #3 ISERR(0/1) result mismatch');

  // Reference to cell with error
  FWorksheet.WriteFormula(0, 0, '=1/0');
  cell := FWorksheet.WriteFormula(0, 1, '=ISERR(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #4 ISERR(A1) result mismatch');

  // Reference to cell without error
  FWorksheet.WriteText(0, 0, 'abc');
  cell := FWorksheet.WriteFormula(0, 1, '=ISERR(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #5 ISERR(A1) result mismatch (no error in cell)');

  // No error as argument
  cell := FWorksheet.WriteFormula(0, 1, '=ISERR("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #6 ISERR(A1) result mismatch (no error as argument)');

  // #N/A error
  cell := FWorksheet.WriteFormula(0, 1, '=ISERR(#N/A)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #7 ISERR(#N/A) result mismatch (#N/A as argument)');

  FWorksheet.WriteNumber(0, 0, 10);
  FWorksheet.WriteNumber(1, 0, 20);
  FWorksheet.WriteFormula(2, 0, '=MATCH(-10, A1:A2, 0)');  // generates a #N/A error
  cell := FWorksheet.WriteFormula(0, 1, '=ISERR(A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #8 ISERR(#N/A) result mismatch (#N/A as argument)');
end;

procedure TCalcFormulaTests.Test_ISERROR;
var
  cell: PCell;
begin
  // #DIV/0! as argument
  cell := FWorksheet.WriteFormula(0, 1, '=ISERROR(#DIV/0!)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 ISERROR(1/0) result mismatch');

  // Cell with #DIV/0! error
  FWorksheet.WriteFormula(0, 0, '=1/0');
  cell := FWorksheet.WriteFormula(0, 1, '=ISERROR(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #2 ISERROR(A1) (A1 = 1/0) result mismatch');

  // Hard coded expression without error
  cell := FWorksheet.WriteFormula(0, 1, '=ISERROR(0/1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #3 ISERROR(0/1) result mismatch');

  // Reference to cell with error
  FWorksheet.WriteFormula(0, 0, '=1/0');
  cell := FWorksheet.WriteFormula(0, 1, '=ISERROR(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #4 ISERROR(A1) result mismatch');

  // Reference to cell without error
  FWorksheet.WriteText(0, 0, 'abc');
  cell := FWorksheet.WriteFormula(0, 1, '=ISERROR(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #5 ISERROR(A1) result mismatch (no error in cell)');

  // No error as argument
  cell := FWorksheet.WriteFormula(0, 1, '=ISERROR("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #6 ISERROR(A1) result mismatch (no error as argument)');

  // #N/A error
  cell := FWorksheet.WriteFormula(0, 1, '=ISERROR(#N/A)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #7 ISERROR(#N/A) result mismatch (#N/A as argument)');

  FWorksheet.WriteNumber(0, 0, 10);
  FWorksheet.WriteNumber(1, 0, 20);
  FWorksheet.WriteFormula(2, 0, '=MATCH(-10, A1:A2, 0)');  // generates a #N/A error
  cell := FWorksheet.WriteFormula(0, 1, '=ISERROR(A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #8 ISERROR(#N/A) result mismatch (#N/A as argument)');
end;

procedure TCalcFormulaTests.Test_ISLOGICAL;
var
  cell: PCell;
begin
  // Boolean
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 ISLOGICAL result mismatch (true)');
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(FALSE)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #2 ISLOGICAL result mismatch (false)');

  FWorksheet.WriteBoolValue(0, 0, true);
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #3 ISLOGICAL result mismatch (bool cell)');

  // Number
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(1.23)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #4 ISLOGICAL result mismatch (number)');

  FWorksheet.WriteNumber(0, 0, 1.234);
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #5 ISLOGICAL result mismatch (number cell)');

  // Date
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(DATE(2025,1,1))');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #6 ISLOGICAL result mismatch (date)');

  FWorksheet.WriteDateTime(0, 0, EncodeDate(2025,1,1));
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #7 ISLOGICAL result mismatch (date cell)');

  // String (corresponding to 'true')
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL("TRUE")');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #8 ISLOGICAL result mismatch ("true" as string)');

  FWorksheet.WriteText(0, 0, 'TRUE');
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #9 ISLOGICAL result mismatch ("true" as string cell)');

  // Blank
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL()');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #10 ISLOGICAL result mismatch (blank)');

  FWorksheet.WriteBlank(0, 0);
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #11 ISLOGICAL result mismatch (blank cell)');

  // Error
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #12 ISLOGICAL result mismatch (error value)');
  CheckNotEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #12 ISLOGICAL result mismatch (error value)');

  FWorksheet.WriteFormula(0, 0, '=1/0');
  cell := FWorksheet.WriteFormula(0, 1, '=ISLOGICAL(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #13 ISLOGICAL result mismatch (cell with error value)');
  CheckNotEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #13 ISLOGICAL result mismatch (cell with error value)');
end;

procedure TCalcFormulaTests.Test_ISNA;
var
  cell: PCell;
begin
  // Check cell with #N/A error
  FWorksheet.WriteNumber(0, 0, 10);
  FWorksheet.WriteNumber(1, 0, 20);
  FWorksheet.WriteFormula(2, 0, '=MATCH(-10,A1:A2,0)');  // This creates an #N/A error
  cell := FWorksheet.WriteFormula(0, 1, '=ISNA(A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 ISNA result mismatch (#N/A error)');

  // Cell with other error (#DIV/0!)
  FWorksheet.WriteFormula(2, 0, '=1/0');
  cell := FWorksheet.WriteFormula(0, 1, '=ISNA(A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #2 ISNA result mismatch (other error)');

  // Cell with no error
  cell := FWorksheet.WriteFormula(0, 1, '=ISNA(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #3 ISNA result mismatch (no error)');
end;

procedure TCalcFormulaTests.Test_ISNONTEXT;
var
  cell: PCell;
begin
  // Number
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT(1.234)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 ISNONTEXT result mismatch (float)');

  FWorksheet.WriteNumber(0, 0, 1.234);
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #2 ISNONTEXT result mismatch (float cell)');

  // Date
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT(DATE(2025,1,1))');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #3 ISNONTEXT result mismatch (date)');

  FWorksheet.WriteDateTime(0, 0, EncodeDate(2025,1,1));
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #4 ISNONTEXT result mismatch (date cell)');

  // String
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #5 ISNONTEXT result mismatch (float as string)');

  FWorksheet.WriteText(0, 0, 'abc');
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #6 ISNONTEXT result mismatch (float as string cell)');

  // Boolean
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT(TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #7 ISNONTEXT result mismatch (boolean)');

  FWorksheet.WriteFormula(0, 0, '=(1=1)');
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #8 ISNONTEXT result mismatch (boolean cell)');

  // Blank
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT()');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #9 ISNONTEXT result mismatch (blank)');

  FWorksheet.WriteBlank(0, 0);
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #10 ISNONTEXT result mismatch (blank cell)');

  // Error
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #11 ISNONTEXT result mismatch (error value)');
  CheckNotEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #11 ISNONTEXT result mismatch (error value)');

  FWorksheet.WriteFormula(0, 0, '=1/0');
  cell := FWorksheet.WriteFormula(0, 1, '=ISNONTEXT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #12 ISNONTEXT result mismatch (cell with error value)');
  CheckNotEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #12 ISNONTEXT result mismatch (cell with error value)');
end;

procedure TCalcFormulaTests.Test_ISNUMBER;
var
  cell: PCell;
begin
  // Number
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER(1.234)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 IsNumber result mismatch (float)');

  FWorksheet.WriteNumber(0, 0, 1.234);
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #2 IsNumber result mismatch (float cell)');

  // Date
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER(DATE(2025,1,1))');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #3 IsNumber result mismatch (date)');

  FWorksheet.WriteDateTime(0, 0, EncodeDate(2025,1,1));
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #4 IsNumber result mismatch (date cell)');

  // String (corresponds to number)
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER("1.234")');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #5 ISNUMBER result mismatch (float as string)');

  FWorksheet.WriteText(0, 0, '1.234');
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #6 ISNUMBER result mismatch (float as string cell)');

  // Boolean
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER(TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #7 ISNUMBER result mismatch (boolean)');

  FWorksheet.WriteFormula(0, 0, '=(1=1)');
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #8 ISNUMBER result mismatch (boolean cell)');

  // Blank
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER()');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #9 ISNUMBER result mismatch (blank)');

  FWorksheet.WriteBlank(0, 0);
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #10 ISNUMBER result mismatch (blank ceöö)');

  // Error
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #11 ISNUMBER result mismatch (error value)');
  CheckNotEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #11 ISNUMBER result mismatch (error value)');

  FWorksheet.WriteFormula(0, 0, '=1/0');
  cell := FWorksheet.WriteFormula(0, 1, '=ISNUMBER(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #12 ISNUMBER result mismatch (cell with error value)');
  CheckNotEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #12 ISNUMBER result mismatch (cell with error value)');
end;

procedure TCalcFormulaTests.Test_ISREF;
var
  cell: PCell;
begin
  // Cell reference
  cell := FWorksheet.WriteFormula(0, 1, '=ISREF(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 ISREF result mismatch (cell reference)');

  // Cell range
  cell := FWorksheet.WriteFormula(0, 1, '=ISREF(A1:A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #2 ISREF result mismatch (cell range reference)');

  // 3d cell ref
  cell := FWorksheet.WriteFormula(0, 1, '=ISREF(Sheet1!A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #3 ISREF result mismatch (3d cell reference)');

  // 3d cell range ref
  cell := FWorksheet.WriteFormula(0, 1, '=ISREF(Sheet1!A1:A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #3 ISREF result mismatch (3d cell range reference)');

  // no ref
  cell := FWorksheet.WriteFormula(0, 1, '=ISREF("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #3 ISREF result mismatch (string)');
end;

procedure TCalcFormulaTests.Test_ISTEXT;
var
  cell: PCell;
begin
  // Text
  cell := FWorksheet.WriteFormula(0, 1, '=ISTEXT("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 ISTEXT result mismatch (text)');

  FWorksheet.WriteText(0, 0, 'abc');
  cell := FWorksheet.WriteFormula(0, 1, '=ISTEXT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #2 ISTEXT result mismatch (text cell)');

  // Number
  cell := FWorksheet.WriteFormula(0, 1, '=ISTEXT(1.234)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #3 ISTEXT result mismatch (number)');

  FWorksheet.WriteNumber(0, 0, 1.234);
  cell := FWorksheet.WriteFormula(0, 1, '=ISTEXT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #4 ISTEXT result mismatch (number cell)');

  // Blank
  cell := FWorksheet.WriteFormula(0, 1, '=ISTEXT()');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #5 ISTEXT result mismatch (blank)');

  FWorksheet.WriteBlank(0, 0);
  cell := FWorksheet.WriteFormula(0, 1, '=ISTEXT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #6 ISTEXT result mismatch (blank cell)');

  // Error
  cell := FWorksheet.WriteFormula(0, 1, '=ISTEXT(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #7 ISTEXT result mismatch (error value)');
  CheckNotEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #7 ISTEXT result mismatch (error value)');

  FWorksheet.WriteFormula(0, 0, '=1/0');
  cell := FWorksheet.WriteFormula(0, 1, '=ISTEXT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #8 ISTEXT result mismatch (error value)');
  CheckNotEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #8 ISTEXT result mismatch (error value)');
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

procedure TCalcFormulaTests.Test_MIN;
begin
  FWorksheet.WriteFormula(0, 1, '=MIN(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 MIN(10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(20,10,30,40)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 MIN(10,20,30,40) result mismatch');

  FWorksheet.WriteNumber(0, 0, 20);
  FWorksheet.WriteNumber(1, 0, 10);
  FWorksheet.WriteNumber(2, 0, 30);
  FWorksheet.WriteNumber(3, 0, 40);

  FWorksheet.WriteFormula(4, 0, '=MIN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(20, FWorksheet.ReadAsNumber(4, 0), 'Formula #3 MIN(A1) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=MIN(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(4, 0), 'Formula #4 MIN(A1,A2) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=MIN(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(4, 0), 'Formula #5 MIN(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=MIN(A1,A2:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(4, 0), 'Formula #6 MIN(A1,A2:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=MIN(A1, 1/0, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(4, 0), 'Formula #7 MIN(A1, 1/0, A2) result mismatch');
end;

procedure TCalcFormulaTests.Test_MAX;
begin
  FWorksheet.WriteFormula(0, 1, '=MAX(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 MAX(10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(20,10,30,40)');
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 MAX(10,20,30,40) result mismatch');

  FWorksheet.WriteNumber(0, 0, 20);
  FWorksheet.WriteNumber(1, 0, 10);
  FWorksheet.WriteNumber(2, 0, 30);
  FWorksheet.WriteNumber(3, 0, 40);

  FWorksheet.WriteFormula(4, 0, '=MAX(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(20, FWorksheet.ReadAsNumber(4, 0), 'Formula #3 MAX(A1) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=MAX(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(20, FWorksheet.ReadAsNumber(4, 0), 'Formula #4 MAX(A1,A2) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=MAX(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(4, 0), 'Formula #5 MAX(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=MAX(A1,A2:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(4, 0), 'Formula #6 MAX(A1,A2:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=MAX(A1, 1/0, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(4, 0), 'Formula #7 MAX(A1, 1/0, A2) result mismatch');
end;

procedure TCalcFormulaTests.Test_PRODUCT;
begin
  FWorksheet.WriteFormula(0, 1, '=PRODUCT(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 PRODUCT(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(1,2,3,4)');
  FWorksheet.CalcFormulas;
  CheckEquals(24, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 PRODUCT(1,2,3,4) result mismatch');

  FWorksheet.WriteNumber(0, 0, 1);
  FWorksheet.WriteNumber(1, 0, 2);
  FWorksheet.WriteNumber(2, 0, 3);
  FWorksheet.WriteNumber(3, 0, 4);

  FWorksheet.WriteFormula(4, 0, '=PRODUCT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(4, 0), 'Formula #3 PRODUCT(A1) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=PRODUCT(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(4, 0), 'Formula #4 PRODUCT(A1,A2) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=PRODUCT(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(24, FWorksheet.ReadAsNumber(4, 0), 'Formula #5 PRODUCT(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=PRODUCT(A1,A2:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(24, FWorksheet.ReadAsNumber(4, 0), 'Formula #6 PRODUCT(A1,A2:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=PRODUCT(A1, 1/0, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(4, 0), 'Formula #7 PRODUCT(A1, 1/0, A2) result mismatch');
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

procedure TCalcFormulaTests.Test_STDEV;
begin
  FWorksheet.WriteFormula(0, 1, '=STDEV(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #1 STDEV(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(1,-2,-3,4)');
  FWorksheet.CalcFormulas;
  CheckEquals(3.16227766, FWorksheet.ReadAsNumber(0, 1), 1E-8, 'Formula #2 STDEV(1,-2,-3,4) result mismatch');

  FWorksheet.WriteNumber(0, 0, 1);
  FWorksheet.WriteNumber(1, 0, -2);
  FWorksheet.WriteNumber(2, 0, -3);
  FWorksheet.WriteNumber(3, 0, 4);

  FWorksheet.WriteFormula(4, 0, '=STDEV(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(4, 0), 'Formula #3 STDEV(A1) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=STDEV(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.121320344, FWorksheet.ReadAsNumber(4, 0), 1E-8, 'Formula #4 STDEV(A1,A2) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=STDEV(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(3.16227766, FWorksheet.ReadAsNumber(4, 0), 1E-8, 'Formula #5 STDEV(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=STDEV(A1,A2:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(3.16227766, FWorksheet.ReadAsNumber(4, 0), 1E-8, 'Formula #6 STDEV(A1,A2:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=STDEV(A1, 1/0, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(4, 0), 'Formula #7 STDEV(A1, 1/0, A2) result mismatch');
end;

procedure TCalcFormulaTests.Test_STDEVP;
begin
  FWorksheet.WriteFormula(0, 1, '=STDEVP(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 STDEVP(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(1,-2,-3,4)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.738612788, FWorksheet.ReadAsNumber(0, 1), 1E-8, 'Formula #2 STDEVP(1,-2,-3,4) result mismatch');

  FWorksheet.WriteNumber(0, 0, 1);
  FWorksheet.WriteNumber(1, 0, -2);
  FWorksheet.WriteNumber(2, 0, -3);
  FWorksheet.WriteNumber(3, 0, 4);

  FWorksheet.WriteFormula(4, 0, '=STDEVP(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(4, 0), 'Formula #3 STDEVP(A1) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=STDEVP(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.5, FWorksheet.ReadAsNumber(4, 0), 1E-8, 'Formula #4 STDEVP(A1,A2) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=STDEVP(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.738612788, FWorksheet.ReadAsNumber(4, 0), 1E-8, 'Formula #5 STDEVP(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=STDEVP(A1,A2:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.738612788, FWorksheet.ReadAsNumber(4, 0), 1E-8, 'Formula #6 STDEVP(A1,A2:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=STDEVP(A1, 1/0, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(4, 0), 'Formula #7 STDEVP(A1, 1/0, A2) result mismatch');
end;

procedure TCalcFormulaTests.Test_SUM;
begin
  FWorksheet.WriteFormula(0, 1, '=SUM(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 SUM(10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(10,20,30,40)');
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 SUM(10,20,30,40) result mismatch');

  FWorksheet.WriteNumber(0, 0, 10);
  FWorksheet.WriteNumber(1, 0, 20);
  FWorksheet.WriteNumber(2, 0, 30);
  FWorksheet.WriteNumber(3, 0, 40);

  FWorksheet.WriteFormula(4, 0, '=SUM(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(4, 0), 'Formula #3 SUM(A1) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=SUM(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(4, 0), 'Formula #4 SUM(A1,A2) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=SUM(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(4, 0), 'Formula #5 SUM(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=SUM(A1,A2:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(4, 0), 'Formula #6 SUM(A1,A2:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=SUM(A1, 1/0, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(4, 0), 'Formula #7 SUM(A1, 1/0, A2) result mismatch');
end;

procedure TCalcFormulaTests.Test_SUMSQ;
begin
  FWorksheet.WriteFormula(0, 1, '=SUMSQ(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 SUMSQ(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(1,2,3,4)');
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 SUMSQ(1,2,3,4) result mismatch');

  FWorksheet.WriteNumber(0, 0, 1);
  FWorksheet.WriteNumber(1, 0, 2);
  FWorksheet.WriteNumber(2, 0, 3);
  FWorksheet.WriteNumber(3, 0, 4);

  FWorksheet.WriteFormula(4, 0, '=SUMSQ(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(4, 0), 'Formula #3 SUMSQ(A1) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=SUMSQ(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(5, FWorksheet.ReadAsNumber(4, 0), 'Formula #4 SUMSQ(A1,A2) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=SUMSQ(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(4, 0), 'Formula #5 SUMSQ(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=SUMSQ(A1,A2:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(4, 0), 'Formula #6 SUMSQ(A1,A2:A4) result mismatch');

  FWorksheet.WriteFormula(4, 0, '=SUMSQ(A1, 1/0, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(4, 0), 'Formula #7 SUMSQ(A1, 1/0, A2) result mismatch');
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

