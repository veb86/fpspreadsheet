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
    FOtherWorksheet: TsWorksheet;
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure Test_ABS;
    procedure Test_ACOS;
    procedure Test_ADDRESS;
    procedure Test_AND;
    procedure Test_AVEDEV;
    procedure Test_AVERAGE;
    procedure Test_AVERAGEIF;
    procedure Test_CEILING;
    procedure Test_COLUMN;
    procedure Test_CONCATENATE;
    procedure Test_COUNT;
    procedure Test_COUNTA;
    procedure Test_COUNTBLANK;
    procedure Test_COUNTIF;
    procedure Test_DATE;
    procedure Test_ERRORTYPE;
    procedure Test_EVEN;
    procedure Test_EXACT;
    procedure Test_FLOOR;
    procedure Test_IF;
    procedure Test_IFERROR;
    procedure Test_INDEX;
    procedure Test_ISBLANK;
    procedure Test_ISERR;
    procedure Test_ISERROR;
    procedure Test_ISLOGICAL;
    procedure Test_ISNA;
    procedure Test_ISNONTEXT;
    procedure Test_ISNUMBER;
    procedure Test_ISREF;
    procedure Test_ISTEXT;
    procedure Test_LEN;
    procedure Test_LOG;
    procedure Test_LOG10;
    procedure Test_LOWER;
    procedure Test_MATCH;
    procedure Test_MAX;
    procedure Test_MIN;
    procedure Test_NOT;
    procedure Test_ODD;
    procedure Test_OR;
    procedure Test_POWER;
    procedure Test_PRODUCT;
    procedure Test_RADIANS;
    procedure Test_ROUND;
    procedure Test_ROW;
    procedure Test_STDEV;
    procedure Test_STDEVP;
    procedure Test_SUM;
    procedure Test_SUMIF;
    procedure Test_SUMSQ;
    procedure Test_TIME;
    procedure Test_UPPER;
    procedure Test_VAR;
    procedure Test_VARP;
  end;

implementation

procedure TCalcFormulaTests.Setup;
begin
  FWorkbook := TsWorkbook.Create;
  FWorksheet := FWorkbook.AddWorksheet('Sheet1');
  FOtherWorksheet := FWorkbook.AddWorksheet('Sheet2');
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
  FWorksheet.WriteErrorValue(0, 0, errIllegalRef);
  FWorksheet.WriteFormula(0, 1, 'ABS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula ABS(1/0) result mismatch');

  // Empty argument
  FWorksheet.WriteBlank(0, 0);
  FWorksheet.WriteFormula(0, 1, 'ABS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula ABS([blank_cell]) result mismatch');
end;

procedure TCalcFormulaTests.Test_ACOS;
begin
  FWorksheet.WriteFormula(0, 1, '=ACOS(0.5)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/3, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #1 ACOS(0.5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ACOS(0)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #2 ACOS(0) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ACOS(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #3 ACOS(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ACOS(-1)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #4 ACOS(-1) result mismatch');

  // Out-of-domain
  FWorksheet.WriteFormula(0, 1, '=ACOS(2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #5 ACOS(2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ACOS(-2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #6 ACOS(-2) result mismatch');

  // Boolean argument
  FWorksheet.WriteFormula(0, 1, '=ACOS(FALSE)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #7 ACOS(FALSE) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ACOS(TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #8 ACOS(TRUE) result mismatch');

  // Numeric string
  FWorksheet.WriteFormula(0, 1, '=ACOS("1")');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #9 ACOS("1") result mismatch');

  // Non-numeric string
  FWorksheet.WriteFormula(0, 1, '=ACOS("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #10 ACOS("abc") result mismatch');

  // Error argument
  FWorksheet.WriteFormula(0, 1, '=ACOS(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #11 ACOS(1/0) result mismatch');

  // Cell with boolean value
  FWorksheet.WriteFormula(0, 0, '=(1=1)');
  FWorksheet.WriteFormula(0, 1, '=ACOS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #12 ACOS(A1) (A1: (1=1)) result mismatch');

  // Cell with numeric string
  FWorksheet.WriteText(0, 0, '1');
  FWorksheet.WriteFormula(0, 1, '=ACOS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #13 ACOS(A1) (A1: "1") result mismatch');

  // Cell with non-numeric string
  FWorksheet.WriteText(0, 0, 'abc');
  FWorksheet.WriteFormula(0, 1, '=ACOS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #14 ACOS(A1) (A1: "abc") result mismatch');

  // Empty cell
  FWorksheet.WriteBlank(0, 0);                    // Empty A1
  FWorksheet.WriteFormula(0, 1, '=ACOS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 'Formula #15 ACOS(A1) (A1: empty) result mismatch');

  // Cell with error
  FWorksheet.WriteErrorValue(0, 0, errIllegalRef);
  FWorksheet.WriteFormula(0, 1, '=ACOS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #10 ACOS(A1) (A1: #REF!) result mismatch');
end;

procedure TCalcFormulaTests.Test_ADDRESS;
begin
  // Default values only
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,1)');
  FWorksheet.CalcFormulas;
  CheckEquals('$A$1', FWorksheet.ReadAsText(0, 1), 'Formula #1 ADDRESS(1,1) result mismatch');

  // 3rd parameter --> absolute address
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,2,1)');
  FWorksheet.CalcFormulas;
  CheckEquals('$B$1', FWorksheet.ReadAsText(0, 1), 'Formula #2 ADDRESS(1,2,1) result mismatch');

  // 3rd parameter --> absolute row
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,2,2)');
  FWorksheet.CalcFormulas;
  CheckEquals('B$1', FWorksheet.ReadAsText(0, 1), 'Formula #3 ADDRESS(1,2,2) result mismatch');

  // 3rd parameter --> absolute col
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,2,3)');
  FWorksheet.CalcFormulas;
  CheckEquals('$B1', FWorksheet.ReadAsText(0, 1), 'Formula 43 ADDRESS(1,2,3) result mismatch');

  // 3rd parameter --> relative address
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,2,4)');
  FWorksheet.CalcFormulas;
  CheckEquals('B1', FWorksheet.ReadAsText(0, 1), 'Formula #5 ADDRESS(1,2,4) result mismatch');

  // missing 3rd parameter --> absolute address
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,2,)');
  FWorksheet.CalcFormulas;
  CheckEquals('$B$1', FWorksheet.ReadAsText(0, 1), 'Formula #6 ADDRESS(1,2,) result mismatch');

  // Combined with ROW() and COLUMN() formulas
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(ROW(),COLUMN())');
  FWorksheet.CalcFormulas;
  CheckEquals('$B$1', FWorksheet.ReadAsText(0, 1), 'Formula #7 ADDRESS(ROW(), COLUMN()) result mismatch');

  // A1 dialect
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,2,1,TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals('$B$1', FWorksheet.ReadAsText(0, 1), 'Formula #8 ADDRESS(1,2,1,TRUE) result mismatch');

  // R1C1 dialect
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,2,1,FALSE)');
  FWorksheet.CalcFormulas;
  CheckEquals('R1C2', FWorksheet.ReadAsText(0, 1), 'Formula #9 ADDRESS(1,2,1,FALSE) result mismatch');

  // Missing dialect argument (must use A1 then)
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,2,1,)');
  FWorksheet.CalcFormulas;
  CheckEquals('$B$1', FWorksheet.ReadAsText(0, 1), 'Formula #10 ADDRESS(1,2,1,) result mismatch');

  // Sheet name
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,2,1,TRUE,"Sheet1")');
  FWorksheet.CalcFormulas;
  CheckEquals('Sheet1!$B$1', FWorksheet.ReadAsText(0, 1), 'Formula #11 ADDRESS(1,2,1,TRUE,"Sheet1") result mismatch');

  // Quoted sheet name
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(1,2,1,TRUE,"Sheet 1")');
  FWorksheet.CalcFormulas;
  CheckEquals('''Sheet 1''!$B$1', FWorksheet.ReadAsText(0, 1), 'Formula #12 ADDRESS(1,2,1,TRUE,"Sheet 1") result mismatch');

  // Elements of address in cells
  FWorksheet.WriteNumber(0, 0, 1);   // Row (1)
  FWorksheet.WriteNumber(1, 0, 2);   // Column (2)
  FWorksheet.WriteNumber(2, 0, 4);   // Flags (relative)
  FWorksheet.WriteBoolValue(3, 0, FALSE);  // Dialect (R1C1)
  FWorksheet.WriteText(4, 0, 'Sheet 1');   // Worksheet
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(A1,A2,A3,A4,A5)');
  FWorksheet.CalcFormulas;
  CheckEquals('''Sheet 1''!R[1]C[2]', FWorksheet.ReadAsText(0, 1), 'Formula #13 ADDRESS(A1,A2,A3,A4,A5) result mismatch');

  // dto., Sheetname cell empty
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(A1,A2,A3,A4,A10)');
  FWorksheet.CalcFormulas;
  CheckEquals('R[1]C[2]', FWorksheet.ReadAsText(0, 1), 'Formula #13 ADDRESS(A1,A2,A3,A4,A10) (A10 blank) result mismatch');

  // dto., Dialect cell empty (--> is assume to be 0 --> RC dialect)
  FWorksheet.WriteFormula(0, 1, '=ADDRESS(A1,A2,A3,A9,A10)');
  FWorksheet.CalcFormulas;
  CheckEquals('R[1]C[2]', FWorksheet.ReadAsText(0, 1), 'Formula #13 ADDRESS(A1,A2,A3,A9,A11) (A9,A10 blank) result mismatch');
end;

procedure TCalcFormulaTests.Test_AND;
var
  cell: PCell;
begin
  cell := FWorksheet.WriteFormula(0, 1, 'AND(1=1,2=2)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 AND(1=1,2=2) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'AND(1=2,2=2)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #2 AND(1=2,2=2) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'AND(1=1,2=1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #3 AND(1=1,2=1) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'AND(1=2,2=1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #4 AND(1=2,2=1) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'AND(1/0,2=2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #5 AND(1/0,2=2) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'AND(1,TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #6 AND(1,TRUE) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'AND(0,TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #7 AND(0,TRUE) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'AND("0",TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(cell), 'Formula #8 AND("0",TRUE) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'AND("abc",TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(cell), 'Formula #9 AND("abc",TRUE) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'AND(1/0,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #10 AND(1/0,0) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'AND(#REF!,#DIV/0!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(cell), 'Formula #11 AND(#REF!,#DIV/0!) result mismatch');
end;

procedure TCalcFormulaTests.Test_AVEDEV;
const
  EPS = 1E-8;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 1);
  FWorksheet.WriteNumber (1, 0, -2);
  FWorksheet.WriteNumber (2, 0, -3);
  FWorksheet.WriteText   (3, 0, '4');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=AVEDEV(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 AVEDEV(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(1,-2)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #2 AVEDEV(1,-2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV("4")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 AVEDEV("4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(1,-2,-3,"4")');
  FWorksheet.CalcFormulas;
  CheckEquals(2.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #4 AVEDEV(1,2,3,"4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 AVEDEV("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 AVEDEV("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(1,-2,-3,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 AVEDEV(1,-2,-3,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 AVEDEV(1/0) result mismatch');

  // Cell references
  FWorksheet.WriteFormula(0, 1, '=AVEDEV(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 AVEDEV(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #8 AVEDEV(A10)(A10=empty) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #9 AVEDEV(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(1.555555556, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #10 AVEDEV(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(1.555555556, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #11 AVEDEV(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=AVEDEV(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(2.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #12 AVEDEV(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(2.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #13 AVEDEV(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(2.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #14 AVEDEV(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=AVEDEV(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 AVEDEV(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVEDEV(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 AVEDEV(A1:A6) result mismatch');
end;

procedure TCalcFormulaTests.Test_AVERAGE;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 10);
  FWorksheet.WriteNumber (1, 0, 20);
  FWorksheet.WriteNumber (2, 0, 30);
  FWorksheet.WriteText   (3, 0, '40');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=AVERAGE(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 AVERAGE(10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(10,20)');
  FWorksheet.CalcFormulas;
  CheckEquals(15, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 AVERAGE(10,20) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE("40")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 AVERAGE("40") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(10,20,30,"40")');
  FWorksheet.CalcFormulas;
  CheckEquals(25, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 AVERAGE(10,20,30,"40") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 AVERAGE("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 AVERAGE("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(10,20,30,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 AVERAGE(10,20,30,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 AVERAGE(1/0) result mismatch');

  // Cell references
  FWorksheet.WriteFormula(0, 1, '=AVERAGE(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 AVERAGE(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 AVERAGE(A10)(A10=empty) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(15, FWorksheet.ReadAsNumber(0, 1), 'Formula #9 AVERAGE(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(20, FWorksheet.ReadAsNumber(0, 1), 'Formula #10 AVERAGE(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(20, FWorksheet.ReadAsNumber(0, 1), 'Formula #11 AVERAGE(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=AVERAGE(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(25, FWorksheet.ReadAsNumber(0, 1), 'Formula #12 AVERAGE(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(25, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 AVERAGE(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(25, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 AVERAGE(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=AVERAGE(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 AVERAGE(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=AVERAGE(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 AVERAGE(A:A6) result mismatch');
end;

procedure TCalcFormulaTests.Test_AVERAGEIF;
const
  EPS = 1E-8;
begin
  // Test data, value range A1:B9
  // A1:A9 - compare values                             // B1:B9 -- calculation values
  FWorksheet.WriteText(0, 0, 'Abc');                    FWorksheet.WriteNumber (0, 1, 100);
  FWorksheet.WriteText(1, 0, 'Abc');                    FWorksheet.WriteNumber (1, 1, 200);
  FWorksheet.WriteText(2, 0, 'bc');                     FWorksheet.WriteNumber (2, 1, 300);
  FWorksheet.WriteText(3, 0, 'a');                      FWorksheet.WriteNumber (3, 1, 400);
  FWorksheet.WriteText(4, 0, 'bc');                     FWorksheet.WriteText   (4, 1, '500');
  FWorksheet.WriteText(5, 0, 'abc');                    FWorksheet.WriteText   (5, 1, 'no number');
  FWorksheet.WriteText(6, 0, '');                       FWorksheet.WriteNumber (6, 1, 600);
  FWorksheet.WriteDateTime(7, 0, EncodeDate(2025,2,1)); FWorksheet.WriteNumber (7, 1, 700);
  FWorksheet.WriteText(8, 0, 'abc');                    FWorksheet.WriteBoolValue (8, 1, TRUE);
  FWorksheet.WriteText(9, 0, 'abc');                    FWorksheet.WriteErrorValue(9, 1, errIllegalRef);

  // *** two-argument calls

  // Average value of all cells in the second column with are <=200
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(B1:B8,"<=200")');
  FWorksheet.CalcFormulas;
  CheckEquals(150, FWorksheet.ReadAsNumber(0, 2), 'Formula #1 AVERAGEIF(B1:B8,"<=200") result mismatch');

  // Average value of all cells in the second column with are >=400
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(B1:B8,">=400")');
  FWorksheet.CalcFormulas;
  CheckEquals(550, FWorksheet.ReadAsNumber(0, 2), 'Formula #2 AVERAGEIF(B1:B8,">=400") result mismatch');

  // Average value of all cells in the second column with are <0
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(B1:B9,"<0")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #3 AVERAGEIF(B1:B9,"<0") result mismatch');

  // *** three-argument calls

  // Average value of all cells in the second column for which the first column cell is 'abc' (case-insensitive)'
  // ... numeric cells only (incl numeric text cell)
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A5,"abc",B1:B5)');
  FWorksheet.CalcFormulas;
  CheckEquals(150, FWorksheet.ReadAsNumber(0, 2), 'Formula #4 AVERAGEIF(A1:A5,"abc",B1:B5) result mismatch');

  // ... dto, but check case-insensitivity of search phrase
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A5,"ABC",B1:B5)');
  FWorksheet.CalcFormulas;
  CheckEquals(150, FWorksheet.ReadAsNumber(0, 2), 'Formula #5 AVERAGEIF(A1:A5,"ABC",B1:B5) result mismatch');

  // ... including non-numeric text cell
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A7,"abc",B1:B7)');
  FWorksheet.CalcFormulas;
  CheckEquals(150, FWorksheet.ReadAsNumber(0, 2), 'Formula #6 AVERAGEIF(A1:A7,"abc",B1:B7) result mismatch');

  // ... including boolean cell
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A9,"abc",B1:B9)');
  FWorksheet.CalcFormulas;
  CheckEquals(150, FWorksheet.ReadAsNumber(0, 2), 'Formula #7 AVERAGEIF(A1:A9,"abc",B1:B9) result mismatch');

  // ... including error cell
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A10,"abc",B1:B10)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 2), 'Formula #8 AVERAGEIF(A1:A10,"abc",B1:B10) result mismatch');

  // ToDo: CompareStringsWithWildcards does not handle a mask such as "*b" like Excel
  {
  // Search for text cells by wildcards
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A8,"*bc",B1:B8)');
  FWorksheet.CalcFormulas;
  CheckEquals(275, FWorksheet.ReadAsNumber(0, 2), 'Formula #9 AVERAGEIF(A1:A8,"*bc",B1:B8) result mismatch');
  }

  // Search for date cells (matching cell found)
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A8,DATE(2025,2,1),B1:B8)');
  FWorksheet.CalcFormulas;
  CheckEquals(700, FWorksheet.ReadAsNumber(0, 2), 'Formula #9 AVERAGEIF(A1:A8,DATE(2025,2,1),B1:B8) result mismatch');

  // Search for date cell (no matching cell found)
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A8,DATE(2000,2,1),B1:B8)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #9 AVERAGEIF(A1:A8,DATE(2000,2,1),B1:B8) result mismatch');

  // Search for empty cells
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A8,"",B1:B8)');
  FWorksheet.CalcFormulas;
  CheckEquals(600, FWorksheet.ReadAsNumber(0, 2), 'Formula #10 AVERAGEIF(A1:A8,"",B1:B8) result mismatch');

  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A8,"=",B1:B8)');
  FWorksheet.CalcFormulas;
  CheckEquals(600, FWorksheet.ReadAsNumber(0, 2), 'Formula #11 AVERAGEIF(A1:A8,"Abc",B1:B8) result mismatch');

  // Search for non-empty cells
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A8,"<>",B1:B8)');
  FWorksheet.CalcFormulas;
  CheckEquals(2200/6, FWorksheet.ReadAsNumber(0, 2), EPS, 'Formula #12 AVERAGEIF(A1:A8,"<>",B1:B8) result mismatch');

  // Compare with reference cell A20
  FWorksheet.WriteText(19, 0, 'abc');  // A20 = "abc"
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A9,A20,B1:B9)');
  FWorksheet.CalcFormulas;
  CheckEquals(150, FWorksheet.ReadAsNumber(0, 2), 'Formula #13 AVERAGEIF(A1:A9,A20,B1:B9) (A20="abc") result mismatch');

  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A9,"<>"&A20,B1:B9)');
  FWorksheet.CalcFormulas;
  CheckEquals(500, FWorksheet.ReadAsNumber(0, 2), 'Formula #13 AVERAGEIF(A1:A9,"<>"&A20,B1:B9) (A20="abc") result mismatch');

  // Error in first argument
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(#DIV/0!,"<10")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #14 AVERAGEIF(#DIV/0,"<10") result mismatch');

  // Error in second argument
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A9,#DIV/0!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #15 AVERAGEIF(A1:A9,#DIV/0) result mismatch');

  // Error in third argument
  FWorksheet.WriteFormula(0, 2, '=AVERAGEIF(A1:A9,"=",#DIV/0!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #16 AVERAGEIF(A1:A9,"=",#DIV/0) result mismatch');
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

procedure TCalcFormulaTests.Test_COLUMN;
begin
  // Get column number of specified cell
  FWorksheet.WriteFormula(0, 1, '=COLUMN(B2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 COLUMN(B2) result mismatch');

  // Get col of the formula cell --- NOT CORRECTLY IMPLEMENTED IN FPS
  FWorksheet.WriteFormula(0, 1, '=COLUMN()');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 COLUMN() result mismatch');   // This would be the Excel result!
  //CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #2 COLUMN() result mismatch');

  // Error value as argument
  FWorksheet.WriteFormula(0, 1, '=COLUMN(#REF!)');
  FWorksheet.CalcFormulas;
  //CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 COLUMN() result mismatch');
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #3 COLUMN(#REF!) result mismatch');

  // Cell containing an error as argument
  FWorksheet.WriteFormula(1, 1, '=1/0');  // cell B2
  FWorksheet.WriteFormula(0, 2, '=COLUMN(B2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #4 COLUMN(B2) (B2 contains error) result mismatch');

end;

procedure TCalcFormulaTests.Test_CONCATENATE;
begin
  // Test data
  FWorksheet.WriteText   (0, 0, 'abc');                 // A1
  FWorksheet.WriteText   (1, 0, 'def');                 // A2
  FWorksheet.WriteNumber (2, 0, 123);                   // A3
  FWorksheet.WriteBoolValue (3, 0, true);               // A4
  FWorksheet.WriteErrorvalue(4, 0, errIllegalRef);      // A5

  // Concatenate 2 literal strings
  FWorksheet.WriteFormula(0, 1, '=CONCATENATE("abc","def")');
  FWorksheet.CalcFormulas;
  CheckEquals('abcdef', FWorksheet.ReadAsText(0, 1), 'Formula #1 CONCATENATE("abc","def") result mismatch');

  // Concatenate 3 literal strings
  FWorksheet.WriteFormula(0, 1, '=CONCATENATE("abc","def","ghi")');
  FWorksheet.CalcFormulas;
  CheckEquals('abcdefghi', FWorksheet.ReadAsText(0, 1), 'Formula #2 CONCATENATE("abc","def","ghi") result mismatch');

  // Concatenate 2 literal strings and number
  FWorksheet.WriteFormula(0, 1, '=CONCATENATE("abc","def",123)');
  FWorksheet.CalcFormulas;
  CheckEquals('abcdef123', FWorksheet.ReadAsText(0, 1), 'Formula #3 CONCATENATE("abc","def",123) result mismatch');

  // Concatenate two numbers
  FWorksheet.WriteFormula(0, 1, '=CONCATENATE(123,456)');
  FWorksheet.CalcFormulas;
  CheckEquals('123456', FWorksheet.ReadAsText(0, 1), 'Formula #4 CONCATENATE(123,456) result mismatch');

  (*  -- this test will not work in the file because Excel writes a localized "TRUE" to the cell
  // Concatenate string and boolean
  FWorksheet.WriteFormula(0, 1, '=CONCATENATE("abc",TRUE())');
  FWorksheet.CalcFormulas;
  CheckEquals('abcTRUE', FWorksheet.ReadAsText(0, 1), 'Formula #5 CONCATENATE("abc",TRUE()) result mismatch');
  *)

  // Concatenate 2 string cells
  FWorksheet.WriteFormula(0, 1, '=CONCATENATE(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals('abcdef', FWorksheet.ReadAsText(0, 1), 'Formula #5 CONCATENATE(A1,A2) result mismatch');

  // Concatenate string and numeric cells
  FWorksheet.WriteFormula(0, 1, '=CONCATENATE(A1,A3)');
  FWorksheet.CalcFormulas;
  CheckEquals('abc123', FWorksheet.ReadAsText(0, 1), 'Formula #6 CONCATENATE(A1,A3) result mismatch');

  // Concatenate string and error cells
  FWorksheet.WriteFormula(0, 1, '=CONCATENATE(A1,A5)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #6 CONCATENATE(A1,A5) result mismatch');

end;

procedure TCalcFormulaTests.Test_COUNT;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 10);
  FWorksheet.WriteNumber (1, 0, 20);
  FWorksheet.WriteNumber (2, 0, 30);
  FWorksheet.WriteText   (3, 0, '40');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Count literal values
  FWorksheet.WriteFormula(0, 1, '=COUNT(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 COUNT(10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(20,10,"abc",40)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 COUNT(20,10,"abc",40) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT("40")');        // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 COUNT("40") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 COUNT("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT("")');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #5 COUNT("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(1/0)');   // argument error does NOT propagate to formula result
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #6 COUNT(1/0) result mismatch');

  // Count in cell references
  FWorksheet.WriteFormula(0, 1, '=COUNT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 COUNT(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #8 COUNT(A10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #9 COUNT(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 1), 'Formula #10 COUNT(A1:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=COUNT(A1:A4)');   // "real" and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(0, 1), 'Formula #11 COUNT(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(A1:A5)');   // "real" and string  values
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(0, 1), 'Formula #12 COUNT(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(A1:A5,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 COUNT(A1:A5,A8:A10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(A1,A2:A5)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 COUNT(A1,A2:A5) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=COUNT(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #15 COUNT(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNT(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(0, 1), 'Formula #15 COUNT(A1:A6) result mismatch');
end;

procedure TCalcFormulaTests.Test_COUNTA;
begin
  FWorksheet.WriteFormula(0, 1, '=COUNTA("")');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 COUNTA("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNTA(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 COUNTA(10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNTA(20,10,"abc",40)');
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 COUNTA(20,10,"abc",40) result mismatch');

  FWorksheet.WriteNumber(0, 0, 20);
  FWorksheet.WriteNumber(1, 0, 10);
  FWorksheet.WriteText(2, 0, 'abc');
  FWorksheet.WriteNumber(3, 0, 40);

  FWorksheet.WriteFormula(4, 1, '=COUNTA(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(4, 1), 'Formula #4 COUNTA(A1) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNTA(A10)');        // A10 is empty
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(4, 1), 'Formula #5 COUNTA(A10) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNTA(A2,A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(4, 1), 'Formula #6 COUNTA(A2,A3) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNTA(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(4, 1), 'Formula #7 COUNTA(A1:A4) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNTA(A1:A10)');
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(4, 1), 'Formula #8 COUNTA(A1:A10) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNTA(A1,A2:A10)');
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(4, 1), 'Formula #9 COUNTA(A1,A2:A10) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNTA(A1, 1/0, A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(4, 1), 'Formula #10 COUNTA(A1, 1/0, A3) result mismatch');
end;

procedure TCalcFormulaTests.Test_COUNTBLANK;
begin
  // The next 2 tests are successful, but not accepted by Excel which only wants
  // a "range" as argument in COUNTBLANK.

  FWorksheet.WriteFormula(0, 1, '=COUNTBLANK("")');     // empty string is not "blank"
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 COUNTBLANK("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=COUNTBLANK(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 COUNTBLANK(10) result mismatch');

  // The following tests are conformal to Excel.

  FWorksheet.WriteNumber(0, 0, 20);
  FWorksheet.WriteText(1, 0, 'abc');

  FWorksheet.WriteFormula(4, 1, '=COUNTBLANK(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(4, 1), 'Formula #3 COUNTBLANK(A1) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNTBLANK(A10)');        // A10 is empty
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(4, 1), 'Formula #4 COUNTBLANK(A10) result mismatch');

  FWorksheet.WriteFormula(4, 1, '=COUNTBLANK(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(4, 1), 'Formula #5 COUNTBLANK(A1:A4) result mismatch');

  FWorksheet.WriteFormula(2, 0, '=1/0');
  FWorksheet.WriteFormula(4, 1, '=COUNTBLANK(A1:A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(4, 1), 'Formula #6 COUNTBLANK(A1:A4) result mismatch');
end;

procedure TCalcFormulaTests.Test_COUNTIF;
begin
  // Test data, range A1:B5
  FWorksheet.WriteNumber (0, 0, 10);          FWorksheet.WriteFormula(0, 1, '=SQRT(-1)');   // --> #NUM!
  FWorksheet.WriteNumber (1, 0, -20);         FWorksheet.WriteBlank  (1, 1);
  FWorksheet.WriteFormula(2, 0, '=(1=1)');    FWorksheet.WriteNumber (2, 1,  0);
  FWorksheet.WriteText   (3, 0, '');          FWorksheet.WriteText   (3, 1, '5');
  FWorksheet.WriteText   (4, 0, 'abc');       FWorksheet.WriteText   (4, 1, 'ABC');
  FWorksheet.WriteBoolValue(5, 0, false);     FWorksheet.WriteErrorValue(5, 1, errOverflow);   // --> #NUM!

  // Counts the elements in A1:B6 which are equal to "abc" (case-insensitive)
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #1 COUNTIF(A1:B6,"abc") result mismatch');

  // Counts the elements in A1:B6 which are < 0
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,"<0")');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #2 COUNTIF(A1:B6,"<0") result mismatch');

  // Counts empty elements in A1:B6
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,"")');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #3 COUNTIF(A1:B6,"") result mismatch');

  // Counts the elements in A1:B6 which are equal to 0
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #4 COUNTIF(A1:B6,0) result mismatch');

  // Counts the elements in A1:B6 which are TRUE
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #5 COUNTIF(A1:B6,TRUE) result mismatch');

  // Counts the elements in A1:B6 which are FALSE
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,FALSE)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #4 COUNTIF(A1:B6,FALSE) result mismatch');

  // Counts the elements in A1:B5 which are #NUM!
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,#NUM!)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #6 COUNTIF(A1:B6,#NUM!) result mismatch');

  // Error in 1st argument
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(#REF!,1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 2), 'Formula #7 COUNTIF(#REF!,1) result mismatch');

  // Error in both arguments
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(#REF!,#REF!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 2), 'Formula #8 COUNTIF(#REF!,#REF!) result mismatch');

  // Count the elements in A1:B6 which are equal to cell A15 (empty)
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #9 COUNTIF(A1:B6,A15) (A15 empty) result mismatch');

  // Count the elements in A1:B6 which are equal to cell A15 (value 10)
  FWorksheet.WriteNumber(14, 0, 10);
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #10 COUNTIF(A1:B6,A15) (A15 = 10) result mismatch');

  // Count the elements in A1:B6 which are < cell A15 (value 10)
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,"<"&A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 2), 'Formula #11 COUNTIF(A1:B6,"<"&A15) (A15 = 10) result mismatch');

  // Count the elements in A1:B6 which are equal to cell A15 (error value #NUM!)
  FWorksheet.WriteErrorValue(14, 0, errOverflow);
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #12 COUNTIF(A1:B6,A15) (A15 = #NUM!) result mismatch');
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
  // Explicit error type
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(#REF!)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errIllegalRef), FWorksheet.ReadAsNumber(0, 1), 'Formula #1 ERROR.TYPE(#REF!) result mismatch');

  // No error in cell --> #N/A
  FWorksheet.WriteNumber(0, 0, 123);
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 1), 'Formula #2 ERROR.TYPE (no error!) result mismatch');

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
  CheckEquals(ord(errIllegalRef), FWorksheet.ReadAsNumber(0, 1), 'Formula #3 ERROR.TYPE (#REF!) result mismatch');

  // #VALUE! error
  FWorksheet.WriteText(0, 0, 'a');
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(1+A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errWrongType), FWorksheet.ReadAsNumber(0, 1), 'Formula #4 ERROR.TYPE #VALUE! result mismatch');

  // #DIV/0! error
  FWorksheet.WriteFormula(0, 0, '=1/0');
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errDivideByZero), FWorksheet.ReadAsNumber(0, 1), 'Formula #5 ERROR.TYPE #DIV/0! result mismatch');

  // #NUM! error
  FWorksheet.WriteFormula(0, 0, '=SQRT(-1)');
  FWorksheet.WriteFormula(0, 1, '=ERROR.TYPE(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(ord(errOverflow), FWorksheet.ReadAsNumber(0, 1), 'Formula #6 ERROR.TYPE #NUM! result mismatch');

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
  CheckEquals(ord(errArgError), FWorksheet.ReadAsNumber(0, 1), 'Formula #7 ERROR.TYPE #N/A result mismatch');
end;

procedure TCalcFormulaTests.Test_EVEN;
begin
  FWorksheet.WriteFormula(0, 1, '=EVEN(1.23)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 EVEN(1.23) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=EVEN(2.34)');
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 EVEN(2.34) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=EVEN(-1.23)');
  FWorksheet.CalcFormulas;
  CheckEquals(-2, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 EVEN(-1.23) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=EVEN(-2.34)');
  FWorksheet.CalcFormulas;
  CheckEquals(-4, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 EVEN(-2.34) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=EVEN(0.0)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #5 EVEN(0.0) result mismatch');

  // String as argument
  FWorksheet.WriteFormula(0, 1, '=EVEN("1")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 EVEN("1") result mismatch');

  // Empty argument
  FWorksheet.WriteFormula(0, 1, '=EVEN()');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 EVEND() result mismatch');

  // Error in argument
  FWorksheet.WriteFormula(0, 1, '=EVEN(#REF!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #8 EVEN(#REF!) result mismatch');

  // Error in argument cell
  FWorksheet.WriteErrorValue(0, 0, errIllegalRef);
  FWorksheet.WriteFormula(0, 1, '=EVEN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #9 EVEN(A1) result mismatch');
end;

procedure TCalcFormulaTests.Test_EXACT;
var
  cell: PCell;
begin
  // Test cells
  FWorksheet.WriteText(0, 0, 'abc');           // A1
  FWorksheet.WriteText(1, 0, 'ABC');           // A2
  FWorksheet.WriteText(2, 0, 'abcd');          // A3
  FWorksheet.WriteText(3, 0, 'äöü');           // A4
  FWorksheet.WriteText(4, 0, 'Äöü');           // A5
  FWorksheet.WriteErrorValue(5, 0, errDivideByZero);  // A6
  FWorksheet.WriteNumber(6, 0, 123);           // A7
  FWorksheet.WriteText(7, 0, '123');           // A8

  // Same literal strings
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT("abc", "abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #1 EXACT("abc","abc") result type mismatch:');
  CheckEquals(true, cell^.BoolValue, 'Formula #1 EXACT("abc", "abc") result mismatch.');

  // Same strings but different case
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT("abc", "ABC")');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #2 EXACT("abc","ABC") result type mismatch:');
  CheckEquals(false, cell^.BoolValue, 'Formula #2 EXACT("abc", "ABC") result mismatch.');

  // Lengths different
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT("abc", "abcd")');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #3 EXACT("abc","abcd") result type mismatch:');
  CheckEquals(false, cell^.BoolValue, 'Formula #3 EXACT("abc", "abcd") result mismatch.');

  // Same unicode chars
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT("äöü", "äöü")');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #4 EXACT("äöü","äöü") result type mismatch:');
  CheckEquals(true, cell^.BoolValue, 'Formula #4 EXACT("äöü", "äöü") result mismatch.');

  // Different unicode case
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT("äöü", "Äöü")');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #5 EXACT("äöü","Äöü") result type mismatch:');
  CheckEquals(false, cell^.BoolValue, 'Formula #5 EXACT("äöü", "Äöü") result mismatch.');

  // Cells with same ASCII strings
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT(A1, A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #6 EXACT(A1,A1) result type mismatch:');
  CheckEquals(true, cell^.BoolValue, 'Formula #6 EXACT(A1,A1) result mismatch.');

  // Cells with same strings, but different case
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT(A1, A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #7 EXACT(A1,A2) result type mismatch:');
  CheckEquals(false, cell^.BoolValue, 'Formula #7 EXACT(A1,A2) result mismatch.');

  // Cells with almost same strings, but different lengths
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT(A1, A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #8 EXACT(A1,A3) result type mismatch:');
  CheckEquals(false, cell^.BoolValue, 'Formula #8 EXACT(A1,A3) result mismatch.');

  // Cells with ASCII and UTF8
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT(A1, A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #9 EXACT(A1,A4) result type mismatch:');
  CheckEquals(false, cell^.BoolValue, 'Formula #9 EXACT(A1,A4) result mismatch.');

  // Cells with same Unicode strings
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT(A4, A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #10 EXACT(A4,A4) result type mismatch:');
  CheckEquals(true, cell^.BoolValue, 'Formula #10 EXACT(A4,A4) result mismatch.');

  // Cells with almost same Unicode strings, but different case
  cell := FWorksheet.WriteFormula(0, 1, '=EXACT(A4, A5)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #11 EXACT(A4,A5) result type mismatch:');
  CheckEquals(false, cell^.BoolValue, 'Formula #11 EXACT(A4,A5) result mismatch.');

  // Error (#DIV/0!) in one cell
  FWorksheet.WriteFormula(0, 1, '=EXACT(A6, A5)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctError, 'Formula #12 EXACT(A6,A5) result type mismatch:');
  CheckEquals(true, cell^.ErrorValue = errDivideByZero, 'Formula #12 EXACT(A6,A5) result mismatch.');

  // Compare numeric string with same number
  FWorksheet.WriteFormula(0, 1, '=EXACT(12,"12")');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #13 EXACT(12,"12") result type mismatch:');
  CheckEquals(true, cell^.BoolValue, 'Formula #13 EXACT(12,"12") result mismatch.');

  // dto with cells
  FWorksheet.WriteFormula(0, 1, '=EXACT(A7,A8)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, cell^.ContentType = cctBool, 'Formula #14 EXACT(A7,A8) result type mismatch:');
  CheckEquals(true, cell^.BoolValue, 'Formula #14 EXACT(A7,A8) result mismatch.');

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

  FWorksheet.WriteFormula(0, 1, '=IF(1,"ok","not ok")');
  FWorksheet.CalcFormulas;
  CheckEquals('ok', FWorksheet.ReadAsText(0, 1), 'Formula #3 IF(1,"ok","not ok") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=IF(0,"ok","not ok")');
  FWorksheet.CalcFormulas;
  CheckEquals('not ok', FWorksheet.ReadAsText(0, 1), 'Formula #4 IF(0,"ok","not ok") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=IF("1","ok","not ok")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 IF("1","ok","not ok") result mismatch');

  // 2 arguments
  FWorksheet.WriteFormula(0, 1, '=IF(A1>=100,"ok")');
  FWorksheet.CalcFormulas;
  CheckEquals('ok', FWorksheet.ReadAsText(0, 1), 'Formula #6 IF(A1>=100,"ok") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=IF(A1<100,"ok")');
  FWorksheet.CalcFormulas;
  CheckEquals('FALSE', FWorksheet.ReadAsText(0, 1), 'Formula #7 IF(A1<100,"ok") result mismatch');

  // Error propagation: error in 3rd argument - result is non-error argument
  FWorksheet.WriteFormula(0, 1, '=IF(A1>=100, "ok", 1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals('ok', FWorksheet.ReadAsText(0, 1), 'Formula #8 IF(A1>=100,"ok",1/0) result mismatch');

  // Error propagation: error in 3rd argument - result is error argument
  FWorksheet.WriteFormula(0, 1, '=IF(A1<100, "ok", 1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #9 IF(A1<100,"ok",1/0) result mismatch');

  // Error propagation: error in 2nd argument - result is error argument
  FWorksheet.WriteFormula(0, 1, '=IF(A1>=100, 1/0,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #10 IF(A1>=100,1/0,"abc") result mismatch');

  // Error propagation: error in 2nd argument - result is 3rd argument
  FWorksheet.WriteFormula(0, 1, '=IF(A1<100, 1/0,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals('abc', FWorksheet.ReadAsText(0, 1), 'Formula #11 IF(A1<100,1/0,"abc") result mismatch');

  // Error propagaton: error in 1st argument
  FWorksheet.WriteFormula(0, 0, '=1/0');
  FWorksheet.WriteFormula(0, 1, '=IF(A1>=100,"ok","not ok")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #12 IF(A1>=100,"ok","not ok") with A1=1/0 result mismatch');
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

procedure TCalcFormulaTests.Test_INDEX;
var
  sh: TsWorksheet;
begin
  // Sample adapted from https://www.ionos.de/digitalguide/online-marketing/verkaufen-im-internet/excel-index-funktion/
  sh := FWorksheet;
  sh.WriteText(0, 0, 'Name');   sh.WriteText  (0, 1, '1991');  sh.WriteText  (0, 2, '1992');
  sh.WriteText(1, 0, 'Peter');  sh.WriteNumber(1, 1, 78);      sh.WriteNumber(1, 2, 81);
  sh.WriteText(2, 0, 'Frank');  sh.WriteNumber(2, 1, 55);      sh.WriteNumber(2, 2, 66);
  sh.WriteText(3, 0, 'Louise'); sh.WriteNumber(3, 1, 42);      sh.WriteNumber(3, 2, 59);
  sh.WriteText(4, 0, 'Valery'); sh.WriteNumber(4, 1, 12);      sh.WriteNumber(4, 2, 33);
  sh.Writetext(5, 0, 'Eva');    sh.WriteNumber(5, 1, 40);      sh.WriteNumber(5, 2, 66);

  // Cell at third row and first column in range B2:C6
  FWorksheet.WriteFormula(10, 0, '=INDEX(B2:C6,3,1)');
  FWorksheet.CalcFormulas;
  CheckEquals(42, FWorksheet.ReadAsNumber(10,0), 'Formula #1 INDEX(B2:F3,3,1) result mismatch');

  // Sample similar to that in unit formulatests:

  FWorksheet.Clear;
  FWorksheet.WriteText  (0, 0, 'A');     // A1
  FWorksheet.WriteText  (0, 1, 'B');     // B1
  FWorksheet.WriteText  (0, 2, 'C');     // C1
  FWorksheet.WriteNumber(1, 0,  10);     // A2
  FWorksheet.WriteNumber(1, 1,  20);     // B2
  FWorksheet.WriteNumber(1, 2,  30);     // C2
  FWorksheet.WriteNumber(2, 0,  11);     // A3
  FWorksheet.WriteNumber(2, 1,  22);     // B3
  FWorksheet.WriteNumber(2, 2,  33);     // C4

  FWorksheet.WriteFormula(0, 5, 'INDEX(A1:C3,1,1)');
  FWorksheet.CalcFormulas;
  CheckEquals('A', FWorksheet.ReadAsText(0, 5), 'Formula #2 INDEX(A1:C3,1,1) result mismatch');

  FWorksheet.WriteFormula(0, 5, 'INDEX(A1:C1,3)');
  FWorksheet.CalcFormulas;
  CheckEquals('C', FWorksheet.ReadAsText(0, 5), 'Formula #3 INDEX(A1:C1,3) result mismatch');

  FWorksheet.WriteFormula(0, 5, 'INDEX(A1:A3,2)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 5), 'Formula #4 INDEX(A1:A3,3) result mismatch');

  FWorksheet.WriteFormula(0, 5, 'INDEX(A1:C2,1,10)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 5), 'Formula #5 INDEX(A1:C2,1,10) result mismatch');

  FWorksheet.WriteFormula(0, 5, 'SUM(INDEX(A1:C3,0,2))');  // Sum of numbers in 2nd column of A1:C3
  FWorksheet.CalcFormulas;
  CheckEquals(42, FWorksheet.ReadAsNumber(0, 5), 'Formula #6 SUM(INDEX(A1:C3,0,2)) result mismatch');

  FWorksheet.WriteFormula(0, 5, 'SUM(INDEX(A1:C3,2,0))');  // Sum of numbers in 2nd row of A1:C3
  FWorksheet.CalcFormulas;
  CheckEquals(60, FWorksheet.ReadAsNumber(0, 5), 'Formula #7 SUM(INDEX(A1:C3,2,0)) result mismatch');

  // Now the same tests, but across sheets
  FOtherWorksheet.WriteFormula(0, 5, 'INDEX(Sheet1!A1:C3,1,1)');
  FWorkbook.CalcFormulas;
  CheckEquals('A', FOtherWorksheet.ReadAsText(0, 5), 'Formula #8 INDEX(Sheet1!A1:C3,1,1) result mismatch');

  FOtherWorksheet.WriteFormula(0, 5, 'INDEX(Sheet1!A1:C1,3)');
  FWorkbook.CalcFormulas;
  CheckEquals('C', FOtherWorksheet.ReadAsText(0, 5), 'Formula #9 INDEX(Sheet1!A1:C1,3) result mismatch');

  FOtherWorksheet.WriteFormula(0, 5, 'INDEX(Sheet1!A1:A3,2)');
  FWorkbook.CalcFormulas;
  CheckEquals(10, FOtherWorksheet.ReadAsNumber(0, 5), 'Formula #10 INDEX(Sheet1!A1:A3,3) result mismatch');

  FOtherWorksheet.WriteFormula(0, 5, 'INDEX(Sheet1!A1:C2,1,10)');
  FWorkbook.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FOtherWorksheet.ReadAsText(0, 5), 'Formula #11 INDEX(Sheet1!A1:C2,1,10) result mismatch');

  FOtherWorksheet.WriteFormula(0, 5, 'SUM(INDEX(Sheet1!A1:C3,0,2))');  // Sum of numbers in 2nd column of A1:C3
  FWorkbook.CalcFormulas;
  CheckEquals(42, FOtherWorksheet.ReadAsNumber(0, 5), 'Formula #6 SUM(Sheet1!INDEX(A1:C3,0,2)) result mismatch');

  FOtherWorksheet.WriteFormula(0, 5, 'SUM(INDEX(Sheet1!A1:C3,2,0))');  // Sum of numbers in 2nd row of A1:C3
  FWorkbook.CalcFormulas;
  CheckEquals(60, FOtherWorksheet.ReadAsNumber(0, 5), 'Formula #7 SUM(Sheet1!INDEX(A1:C3,2,0)) result mismatch');
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
  {  In Excel the "no argument" case is not accepted (runtime error, no error code)
  cell := FWorksheet.WriteFormula(0, 1, '=ISBLANK()');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #3 ISBLANK() result mismatch');
  }

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

procedure TCalcFormulaTests.Test_LEN;
begin
  FWorksheet.WriteText(0, 0, 'abc');                   // A1
  FWorksheet.WriteText(1, 0, 'Äbγ');                   // A2
  FWorksheet.WriteErrorValue(2, 0, errIllegalRef);     // A3

  // Literal ASCII string
  FWorksheet.WriteFormula(0, 1, '=LEN("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 LEN("abc") result mismatch');

  // Literal Unicode string
  FWorksheet.WriteFormula(0, 1, '=LEN("Äbγ")');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 LEN("Äbc") result mismatch');

  // Another Unicode string
  FWorksheet.WriteFormula(0, 1, '=LEN("αβγδε")');
  FWorksheet.CalcFormulas;
  CheckEquals(5, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 LEN("αβγδε") result mismatch');

  // Cell with ASCII string
  FWorksheet.WriteFormula(0, 1, '=LEN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 LEN(A1) result mismatch');

  // Cell with unicode string
  FWorksheet.WriteFormula(0, 1, '=LEN(A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 1), 'Formula #5 LEN(A2) result mismatch');

  // Cell with error
  FWorksheet.WriteFormula(0, 1, '=LEN(A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #6 LEN(A3) (#REF!) result mismatch');

  // Empty cell
  FWorksheet.WriteFormula(0, 1, '=LEN(A99)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 LEN(A99) (empty cell) result mismatch');
end;

procedure TCalcFormulaTests.Test_LOG;
begin
  // Correct formula
  FWorksheet.WriteFormula(0, 1, '=LOG(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 LOG(10) result mismatch');

  // Argument is zero
  FWorksheet.WriteFormula(0, 1, '=LOG(0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #2 LOG(0) result mismatch');

  // Negative argument
  FWorksheet.WriteFormula(0, 1, '=LOG(-10)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #3 LOG(-10) result mismatch');

  // Non-numeric argument
  FWorksheet.WriteFormula(0, 1, '=LOG("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #4 LOG("abc") result mismatch');

  // Error argument
  FWorksheet.WriteFormula(0, 1, '=LOG(#REF!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #5 LOG(#REF!) result mismatch');

  // Two argument cases

  // Correct formula
  FWorksheet.WriteFormula(0, 1, '=LOG(8,2)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 1), 'Formula #6 LOG(8, 2) result mismatch');

  //  2nd argument negative
  FWorksheet.WriteFormula(0, 1, '=LOG(8,-2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #7 LOG(8,-2) result mismatch');

  // Non-numeric 2nd argument
  FWorksheet.WriteFormula(0, 1, '=LOG(8,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #8 LOG(8,"abc") result mismatch');

  // Missing 1st argument
  FWorksheet.WriteFormula(0, 1, '=LOG(,2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #9 LOG(,2) result mismatch');

  // Missing 2nd argument
  FWorksheet.WriteFormula(0, 1, '=LOG(10,)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #10 LOG(10,) result mismatch');

  // Error in 2nd argument
  FWorksheet.WriteFormula(0, 1, '=LOG(8,#REF!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #11 LOG(8,#REF!) result mismatch');
end;

procedure TCalcFormulaTests.Test_LOG10;
begin
  // Correct formula
  FWorksheet.WriteFormula(0, 1, '=LOG10(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 LOG10(10) result mismatch');

  // Argument is zero
  FWorksheet.WriteFormula(0, 1, '=LOG10(0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #2 LOG10(0) result mismatch');

  // Negative argument
  FWorksheet.WriteFormula(0, 1, '=LOG10(-10)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #3 LOG10(-10) result mismatch');

  // Non-numeric argument
  FWorksheet.WriteFormula(0, 1, '=LOG10("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #4 LOG10("abc") result mismatch');

  // Error argument
  FWorksheet.WriteFormula(0, 1, '=LOG10(#REF!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #5 LOG10(#REF!) result mismatch');
end;

procedure TCalcFormulaTests.Test_LOWER;
begin
  FWorksheet.WriteText(0, 0, 'abc');                   // A1
  FWorksheet.WriteText(1, 0, 'Äbγδ');                  // A2
  FWorksheet.WriteErrorValue(2, 0, errIllegalRef);     // A3

  // ASCII, already lower case
  FWorksheet.WriteFormula(0, 1, '=LOWER("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals('abc', FWorksheet.ReadAsText(0, 1), 'Formula #1 LOWER("abc") result mismatch');

  // ASCII, mixed upper/lower case
  FWorksheet.WriteFormula(0, 1, '=LOWER("Abc")');
  FWorksheet.CalcFormulas;
  CheckEquals('abc', FWorksheet.ReadAsText(0, 1), 'Formula #2 LOWER("Abc") result mismatch');

  // Unicode, already lower case
  FWorksheet.WriteFormula(0, 1, '=LOWER("äöü αβγ")');
  FWorksheet.CalcFormulas;
  CheckEquals('äöü αβγ', FWorksheet.ReadAsText(0, 1), 'Formula #3 LOWER("äöü αβγ") result mismatch');

  // Unicode, mixed upper/lower case
  FWorksheet.WriteFormula(0, 1, '=LOWER("Äöü αβΓ")');
  FWorksheet.CalcFormulas;
  CheckEquals('äöü αβγ', FWorksheet.ReadAsText(0, 1), 'Formula #4 LOWER("Äöü αβΓ") result mismatch');

  // Mixed unicode, ASCII, number
  FWorksheet.WriteFormula(0, 1, '=LOWER("Äöü αβΓ 123")');
  FWorksheet.CalcFormulas;
  CheckEquals('äöü αβγ 123', FWorksheet.ReadAsText(0, 1), 'Formula #5 LOWER("Äöü αβΓ 123") result mismatch');

  // Error in argument
  FWorksheet.WriteFormula(0, 1, '=LOWER(#REF!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #6 LOWER(#REF!) result mismatch');

  // Cell with error
  FWorksheet.WriteFormula(0, 1, '=LOWER(A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #7 LOWER(A3) (#REF!) result mismatch');

  // Empty cell
  FWorksheet.WriteFormula(0, 1, '=LOWER(A99)');
  FWorksheet.CalcFormulas;
  CheckEquals('', FWorksheet.ReadAsText(0, 1), 'Formula #8 LOWER(A9) (empty) result mismatch');
end;

procedure TCalcFormulaTests.Test_MATCH;
begin
  // *** Match_Type 0, unsorted data in search range, find first value

  // Search range to be checked: B1:B4
  FWorksheet.WriteNumber(0, 1, 10);
  FWorksheet.WriteNumber(1, 1, 20);
  FWorksheet.WriteNumber(2, 1, 30);
  FWorksheet.WriteNumber(3, 1, 15);
  FWorksheet.WriteNumber(4, 1, 20);

  // Search for constant, contained in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(10, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #1 MATCH(10,B1_B5,0) mismatch, value in range');

  // Search for constant, contained several times in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(20, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #2 MATCH(20,B1:B5,0) mismatch, value above range');

  // Search for constant, below search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(0, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #3 MATCH(0,B1:B5,0) mismatch, value below range');

  // Search for constant, above search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(90, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #4 MATCH(90,B1:B5,0) mismatch, value above range');

  // Search for cell with value in range
  FWorksheet.WriteNumber(0, 0, 20);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #5 MATCH(A1,B1:B5,0) mismatch, cell value in range');

  // Search for cell, but cell is empty
  FWorksheet.WriteFormula(0, 2, '=MATCH(A99, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #6 MATCH(A99,B1:B5,0) mismatch, empty cell');

  // Search range is empty
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, D1:D3, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #7 MATCH mismatch, match_type -1, empty search range');


  // *** Match_Type 1 (find largest value in range <= value), ascending values in search range

  // Search range to be checked: B1:B3
  FWorksheet.WriteNumber(0, 1, 10);
  FWorksheet.WriteNumber(1, 1, 20);
  FWorksheet.WriteNumber(2, 1, 30);
  FWorksheet.WriteBlank(3, 1);

  // Search for constant, contained in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, B1:B3, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #8 MATCH mismatch, match_type 1, in range');

  // Search for constant,  below search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(8, B1:B3, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #9 MATCH mismatch, match_type 1, below range');

  // Search for constant, above search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(123, B1:B3, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 2), 'Formula MATCH #10 mismatch, match_type 1, above range');

  // Search for cell with value in range
  FWorksheet.WriteNumber(0, 0, 28);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B3, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula MATCH #11 mismatch, match_type 1, cell in range');
  FWorksheet.WriteBlank(0, 0);

  // Search for cell, but cell is empty
  FWorksheet.WriteBlank(0, 0);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B3, 1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #12 MATCH mismatch, match_type 1, empty cell');

  // Search range is empty
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, D1:D3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula MATCH #13 mismatch, match_type -1, empty search range');


  // *** Match_Type -1 (find smallest value in range >= value), descending values in search range

  // Search range to be checked: B1:B3
  FWorksheet.WriteNumber(0, 1, 30);
  FWorksheet.WriteNumber(1, 1, 20);
  FWorksheet.WriteNumber(2, 1, 10);

  // Search for constant, contained in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, B1:B3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #14 MATCH mismatch, match_type -1, in range');

  // Search for constant,  below search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(8, B1:B3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 2), 'Formula #15 MATCH mismatch, match_type -1, below range');

  // Search for constant, above search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(123, B1:B3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #16 MATCH mismatch, match_type -1, above range');

  // Search for cell with value in range
  FWorksheet.WriteNumber(0, 0, 28);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #17 MATCH mismatch, match_type -1, cell in range');
  FWorksheet.WriteBlank(0, 0);

  // Search for cell, but cell is empty
  FWorksheet.WriteBlank(0, 0);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #18 MATCH mismatch, match_type -1, empty cell');

  // Search range is empty
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, D1:D3, -1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #19 MATCH mismatch, match_type -1, empty search range');


  // **** Error propagation

  // Search for cell, but cell contains error
  FWorksheet.WriteFormula(0, 0, '=1/0');
  FWorksheet.WriteNumber(1, 1, 20);
  FWorksheet.WriteNumber(2, 1, 30);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B4, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #20 MATCH mismatch, match_type 0, error cell');

  // Match_type parameter contains error
  FWorksheet.WriteNumber(0, 1, 10);
  FWorksheet.WriteFormula(0, 5, '=1/0');    // F1
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B3, F1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #21 MATCH mismatch, match_type 0, error in search range');

  // Cell range contains error
  FWorksheet.WriteNumber(0, 1, 10);
  FWorksheet.WriteFormula(1, 1, '=1/0');    // B2 contains a #DIV/0! error now
  FWorksheet.WriteNumber(2, 1, 30);
  // Search for constant, contained in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(20, B1:B3, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #22 MATCH mismatch, match_type 0, error in search range');
    // ArgError because search value is not found


  // **** Partial text

  FWorksheet.WriteText(0, 0, 'abc');
  FWorksheet.WriteText(1, 0, 'axy');
  FWorksheet.WriteText(2, 0, 'xxy');
  FWorksheet.WriteText(3, 0, 'ayc');
  FWorksheet.WriteText(4, 0, 'äbc');

  FWorksheet.WriteFormula(0, 2, '=MATCH("a*",A1:A4,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #23 MATCH mismatch, partial text "a*"');

  FWorksheet.WriteFormula(0, 2, '=MATCH("z*",A1:A4,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #24 MATCH mismatch, partial text "z*"');

  FWorksheet.WriteFormula(0, 2, '=MATCH("*y",A1:A4,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #25 MATCH mismatch, partial text "*y"');

  FWorksheet.WriteFormula(0, 2, '=MATCH("*z",A1:A4,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #26 MATCH mismatch, partial text "*z"');

  FWorksheet.WriteFormula(0, 2, '=MATCH("*z*",A1:A4,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #27 MATCH mismatch, partial text "*z*"');

  FWorksheet.WriteFormula(0, 2, '=MATCH("ay?",A1:A4,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(0, 2), 'Formula #28 MATCH mismatch, partial text "ay?');

  FWorksheet.WriteFormula(0, 2, '=MATCH("a?",A1:A4,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #29 MATCH mismatch, partial text "a?');

  FWorksheet.WriteFormula(0, 2, '=MATCH("Ä*",A1:A5,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(5, FWorksheet.ReadAsNumber(0, 2), 'Formula #30 MATCH mismatch, partial text "Ä*');

end;

procedure TCalcFormulaTests.Test_MAX;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 10);
  FWorksheet.WriteNumber (1, 0, 20);
  FWorksheet.WriteNumber (2, 0, 30);
  FWorksheet.WriteText   (3, 0, '40');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=MAX(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 MAX(10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(10,20)');
  FWorksheet.CalcFormulas;
  CheckEquals(20, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 MAX(10,20) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX("40")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 MAX("40") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(10,20,30,"40")');
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 MAX(10,20,30,"40") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 MAX("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 MAX("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(10,20,30,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 MAX(10,20,30,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 MAX(1/0) result mismatch');

  // Cell references
  FWorksheet.WriteFormula(0, 1, '=MAX(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 MAX(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #8 MAX(A10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(20, FWorksheet.ReadAsNumber(0, 1), 'Formula #9 MAX(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #10 MAX(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #11 MAX(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=MAX(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(0, 1), 'Formula #12 MAX(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 MAX(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 MAX(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=MAX(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 MAX(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 MAX(A:A6) result mismatch');
end;

procedure TCalcFormulaTests.Test_MIN;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 10);
  FWorksheet.WriteNumber (1, 0, 20);
  FWorksheet.WriteNumber (2, 0, 30);
  FWorksheet.WriteText   (3, 0, '-40');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=MIN(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 MIN(10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(10,20)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 MIN(10,20) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN("40")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 MIN("40") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(10,20,30,"40")');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 MIN(10,20,30,"40") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 MIN("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 MIN("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(10,20,30,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 MIN(10,20,30,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MAX(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 MAX(1/0) result mismatch');

  // Cell references
  FWorksheet.WriteFormula(0, 1, '=MIN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 MIN(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #8 MIN(A10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #9 MIN(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #10 MIN(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #11 MIN(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=MIN(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(-40, FWorksheet.ReadAsNumber(0, 1), 'Formula #12 MIN(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(-40, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 MIN(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(-40, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 MIN(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=MIN(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 MIN(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=MIN(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 MIN(A:A6) result mismatch');
end;

procedure TCalcFormulaTests.Test_NOT;
var
  cell: PCell;
begin
  cell := FWorksheet.WriteFormula(0, 1, 'NOT(1=1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #1 NOT(1=1) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'NOT(1=2)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #2 NOT(1=2) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'NOT(0)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #3 NOT(0) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'NOT(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #4 NOT(1) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'NOT(12)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #5 NOT(12) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'NOT("0")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(cell), 'Formula #6 NOT("0") result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'NOT("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(cell), 'Formula #7 NOT("abc") result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'NOT(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #8 NOT(1/0) result mismatch');
end;

procedure TCalcFormulaTests.Test_ODD;
begin
  FWorksheet.WriteFormula(0, 1, '=ODD(0.5)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 ODD(0.5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ODD(1.5)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 ODD(1.5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ODD(-0.5)');
  FWorksheet.CalcFormulas;
  CheckEquals(-1, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 ODD(-0.5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ODD(0.0)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 ODD(0.0) result mismatch');

  // String as argument
  FWorksheet.WriteFormula(0, 1, '=ODD("1")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 ODD("1") result mismatch');

  // Empty as argument
  FWorksheet.WriteFormula(0, 1, '=ODD()');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 ODD() result mismatch');

  // Error in argument
  FWorksheet.WriteFormula(0, 1, '=ODD(#REF!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #7 ODD(#REF!) result mismatch');

  // Error in argument cell
  FWorksheet.WriteErrorValue(0, 0, errIllegalRef);
  FWorksheet.WriteFormula(0, 1, '=ODD(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #8 ODD(A1) result mismatch');
end;

procedure TCalcFormulaTests.Test_OR;
var
  cell: PCell;
begin
  cell := FWorksheet.WriteFormula(0, 1, 'OR(1=1,2=2)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #1 OR(1=1,2=2) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'OR(1=2,2=2)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #2 OR(1=2,2=2) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'OR(1=1,2=1)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #3 OR(1=1,2=1) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'OR(1=2,2=1)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #4 OR(1=2,2=1) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'OR(1/0,2=2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #5 OR(1/0,2=2) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'OR(1,TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(true, FWorksheet.IsTrueValue(cell), 'Formula #6 OR(1,TRUE) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'OR(0,FALSE)');
  FWorksheet.CalcFormulas;
  CheckEquals(false, FWorksheet.IsTrueValue(cell), 'Formula #7 OR(0,TRUE) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'OR("0",TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(cell), 'Formula #8 OR("0",TRUE) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'OR("abc",TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(cell), 'Formula #9 OR("abc",TRUE) result mismatch');

  cell := FWorksheet.WriteFormula(0, 1, 'OR(1/0,1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(cell), 'Formula #10 OR(1/0,1/0) result mismatch');
end;

procedure TCalcFormulaTests.Test_POWER;
begin
  FWorksheet.WriteFormula(0, 1, '=POWER(3,2)');
  FWorksheet.CalcFormulas;
  CheckEquals(9, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 POWER(3,2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=POWER(2,-3)');
  FWorksheet.CalcFormulas;
  CheckEquals(1/8, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 POWER(2,-3) result mismatch');

  // x^0
  FWorksheet.WriteFormula(0, 1, '=POWER(2,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 POWER(2,0) result mismatch');

  // 0^x
  FWorksheet.WriteFormula(0, 1, '=POWER(0,2)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 POWER(0,2) result mismatch');

  // 0^0
  FWorksheet.WriteFormula(0, 1, '=POWER(0,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #5 POWER(0,0) result mismatch');

  // Boolean argument
  FWorksheet.WriteFormula(0, 1, '=POWER(TRUE,"2")');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #6 POWER(TRUE,"2") result mismatch');

  // Numeric string arguments
  FWorksheet.WriteFormula(0, 1, '=POWER("3",2)');
  FWorksheet.CalcFormulas;
  CheckEquals(9, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 POWER("3",2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=POWER(3,"2")');
  FWorksheet.CalcFormulas;
  CheckEquals(9, FWorksheet.ReadAsNumber(0, 1), 'Formula #8 POWER(3,"2") result mismatch');

  // Non-numeric string arguments
  FWorksheet.WriteFormula(0, 1, '=POWER("abc",2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #9 POWER("abc",2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=POWER(3,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #10 POWER(3,"abc") result mismatch');
end;

procedure TCalcFormulaTests.Test_PRODUCT;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 1);
  FWorksheet.WriteNumber (1, 0, 2);
  FWorksheet.WriteNumber (2, 0, 3);
  FWorksheet.WriteText   (3, 0, '4');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=PRODUCT(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 PRODUCT(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(1,2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 PRODUCT(1,2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT("4")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(4, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 PRODUCT("4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(1,2,3,"4")');
  FWorksheet.CalcFormulas;
  CheckEquals(24, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 PRODUCT(1,2,3,"4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 PRODUCT("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 PRODUCT("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(1,2,3,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 PRODUCT(1,2,3,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 PRODUCT(1/0) result mismatch');

  // Count in cell references
  FWorksheet.WriteFormula(0, 1, '=PRODUCT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 PRODUCT(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #8 PRODUCT(A10) (A10 = empty) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #9 PRODUCT(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(6, FWorksheet.ReadAsNumber(0, 1), 'Formula #10 PRODUCT(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(6, FWorksheet.ReadAsNumber(0, 1), 'Formula #11 PRODUCT(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=PRODUCT(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(24, FWorksheet.ReadAsNumber(0, 1), 'Formula #12 PRODUCT(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(24, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 PRODUCT(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(24, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 PRODUCT(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=PRODUCT(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 PRODUCT(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=PRODUCT(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 PRODUCT(A:A6) result mismatch');
end;

procedure TCalcFormulaTests.Test_RADIANS;
begin
  FWorksheet.WriteFormula(0, 1, '=RADIANS(90)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #1 RADIANS(90) result mismatch');

  // Boolean argument
  FWorksheet.WriteFormula(0, 1, '=RADIANS(TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi*1/180, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #2 RADIANS(TRUE) result mismatch');

  // Numeric string
  FWorksheet.WriteFormula(0, 1, '=RADIANS("90")');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #3 RADIANS("90") result mismatch');

  // Non-numeric string
  FWorksheet.WriteFormula(0, 1, '=RADIANS("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #4 RADIANS("abc") result mismatch');

  // Error argument
  FWorksheet.WriteFormula(0, 1, '=RADIANS(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #5 RADIANS(1/0) result mismatch');

  // Cell with boolean value
  FWorksheet.WriteFormula(0, 0, '=(1=1)');
  FWorksheet.WriteFormula(0, 1, '=RADIANS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi*1/180, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #6 RADIANS(A1) (A1: (1=1)) result mismatch');

  // Cell with numeric string
  FWorksheet.WriteText(0, 0, '90');
  FWorksheet.WriteFormula(0, 1, '=RADIANS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #7 RADIANS(A1) (A1: "90") result mismatch');

  // Cell with non-numeric string
  FWorksheet.WriteText(0, 0, 'abc');
  FWorksheet.WriteFormula(0, 1, '=RADIANS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #8 RADIANS(A1) (A1: "abc") result mismatch');

  // Empty cell
  FWorksheet.WriteBlank(0, 0);                    // Empty A1
  FWorksheet.WriteFormula(0, 1, '=RADIANS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #9 RADIANS(A1) (A1: empty) result mismatch');

  // Cell with error
  FWorksheet.WriteErrorValue(0, 0, errIllegalRef);
  FWorksheet.WriteFormula(0, 1, '=RADIANS(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #10 RADIANS(A1) (A1: #REF!) result mismatch');
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

procedure TCalcFormulaTests.Test_ROW;
begin
  // Get row of specified cell
  FWorksheet.WriteFormula(0, 1, '=ROW(B2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 ROW(B2) result mismatch');

  // Get row of the formula cell --- NOT CORRECTLY IMPLEMENTED IN FPS
  FWorksheet.WriteFormula(0, 1, '=ROW()');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 ROW() result mismatch');   // This would be the Excel result!
  //CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #2 ROW() result mismatch');

  // Error value as argument
  FWorksheet.WriteFormula(0, 1, '=ROW(#REF!)');
  FWorksheet.CalcFormulas;
  //CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 ROW() result mismatch');
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #3 ROW(#REF!) result mismatch');

  // Cell containing an error as argument
  FWorksheet.WriteFormula(1, 1, '=1/0');
  FWorksheet.WriteFormula(0, 1, '=ROW(B2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 ROW(B2) (B2 contains error) result mismatch');
end;

procedure TCalcFormulaTests.Test_STDEV;
const
  EPS = 1E-8;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 1);
  FWorksheet.WriteNumber (1, 0, -2);
  FWorksheet.WriteNumber (2, 0, -3);
  FWorksheet.WriteText   (3, 0, '4');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=STDEV(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #1 STDEV(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(1,-2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.121320344, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #2 STDEV(1,-2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV("4")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #3 STDEV("4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(1,-2,-3,"4")');
  FWorksheet.CalcFormulas;
  CheckEquals(3.16227766, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #4 STDEV(1,2,3,"4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 STDEV("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 STDEV("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(1,-2,-3,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 STDEV(1,-2,-3,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 STDEV(1/0) result mismatch');

  // Cell references
  FWorksheet.WriteFormula(0, 1, '=STDEV(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #7 STDEV(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 STDEV(A10)(A10=empty) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.121320344, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #9 STDEV(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(2.081665999, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #10 STDEV(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(2.081665999, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #11 STDEV(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=STDEV(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(3.16227766, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #12 STDEV(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(3.16227766, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #13 STDEV(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(3.16227766, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #14 STDEV(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=STDEV(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 STDEV(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEV(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 STDEV(A:A6) result mismatch');
end;

procedure TCalcFormulaTests.Test_STDEVP;
const
  EPS = 1E-8;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 1);
  FWorksheet.WriteNumber (1, 0, -2);
  FWorksheet.WriteNumber (2, 0, -3);
  FWorksheet.WriteText   (3, 0, '4');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=STDEVP(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 STDEVP(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(1,-2)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #2 STDEVP(1,-2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP("4")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 STDEVP("4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(1,-2,-3,"4")');
  FWorksheet.CalcFormulas;
  CheckEquals(2.738612788, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #4 STDEVP(1,2,3,"4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 STDEVP("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 STDEVP("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(1,-2,-3,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 STDEVP(1,-2,-3,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 STDEVP(1/0) result mismatch');

  // Cell references
  FWorksheet.WriteFormula(0, 1, '=STDEVP(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 STDEVP(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 STDEVP(A10)(A10=empty) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #9 STDEVP(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(1.699673171, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #10 STDEVP(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(1.699673171, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #11 STDEVP(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=STDEVP(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(2.738612788, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #12 STDEVP(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(2.738612788, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #13 STDEVP(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(2.738612788, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #14 STDEVP(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=STDEVP(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 STDEVP(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=STDEVP(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 STDEVP(A:A6) result mismatch');
end;

procedure TCalcFormulaTests.Test_SUM;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 10);
  FWorksheet.WriteNumber (1, 0, 20);
  FWorksheet.WriteNumber (2, 0, 30);
  FWorksheet.WriteText   (3, 0, '40');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=SUM(10)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 SUM(10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(10,20)');
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 SUM(10,20) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM("40")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(40, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 SUM("40") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(10,20,30,"40")');
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 SUM(10,20,30,"40") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 SUM("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 SUM("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(10,20,30,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 SUM(10,20,30,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 SUM(1/0) result mismatch');

  // Count in cell references
  FWorksheet.WriteFormula(0, 1, '=SUM(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 SUM(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #8 SUM(A10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #9 SUM(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(60, FWorksheet.ReadAsNumber(0, 1), 'Formula #10 SUM(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(60, FWorksheet.ReadAsNumber(0, 1), 'Formula #11 SUM(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(0, 1), 'Formula #12 SUM(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 SUM(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 SUM(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=SUM(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 SUM(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 SUM(A:A6) result mismatch');
end;

procedure TCalcFormulaTests.Test_SUMIF;
begin
  // Test data, range A1:B5
  FWorksheet.WriteNumber (0, 0, 10);          FWorksheet.WriteNumber (0, 1, -1);
  FWorksheet.WriteNumber (1, 0, 20);          FWorksheet.WriteNumber (1, 1, -2);
  FWorksheet.WriteNumber (2, 0, 40);          FWorksheet.WriteNumber (2, 1,  6);
  FWorksheet.WriteText   (3, 0, '-40');       FWorksheet.WriteText   (3, 1, '5');
  FWorksheet.WriteText   (4, 0, 'abc');       FWorksheet.WriteText   (4, 1, 'ABC');

  // Work data, range A8:B12
  FWorksheet.WriteNumber ( 7, 0, 100);         FWorksheet.WriteNumber ( 7, 1, -100);
  FWorksheet.WriteNumber ( 8, 0, 200);         FWorksheet.WriteNumber ( 8, 1, -200);
  FWorksheet.WriteNumber ( 9, 0, 300);         FWorksheet.WriteNumber ( 9, 1, -300);
  FWorksheet.WriteText   (10, 0, '400');       FWorksheet.WriteText   (10, 1, '-500');
  FWorksheet.WriteText   (11, 0, 'xyz');       FWorksheet.WriteText   (11, 1, 'XYZ');

  // *** Range contains numbers only ***

  // Calculate sum of the elements in A1:B3 which are equal to 0
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B3,0)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 2), 'Formula #1 SUMIF(A1:B3,0) result mismatch');

  // Calculate sum of the elements in A1:B3 which are < 0
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B3,"<0")');
  FWorksheet.CalcFormulas;
  CheckEquals(-3, FWorksheet.ReadAsNumber(0, 2), 'Formula #2 SUMIF(A1:B3,"<0") result mismatch');

  // Calculate sum of the elements in A8:B10 for which the elements in A1:B3 are equal to 10
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B3,10,A8:B10)');
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(0, 2), 'Formula #3 SUMIF(A1:B3,10,A8:B10) result mismatch');

  // Compare cell A15
  FWorksheet.WriteNumber( 14, 0, 10);

  // Calculate sum of the elements in A1:B3 which are equal to cell A15 (value 10)
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B3,A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 2), 'Formula #4 SUMIF(A1:B3,A15) result mismatch');

  // Calculate sum of the elements in A1:B3 which are < cell A15
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B3,"<"&A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 2), 'Formula #5 SUMIF(A1:B3,"<"&A15) result mismatch');

  // Calculate sum of the elements in A8:B10 for which the elements in A1:B3 are equal to 10
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B3,"<"&A15,A8:B10)');
  FWorksheet.CalcFormulas;
  CheckEquals(-600, FWorksheet.ReadAsNumber(0, 2), 'Formula #6 SUMIF(A1:B3,"<"&A15,A8:B10) result mismatch');


  // *** Range contains also numeric strings ***

  // Calculate sum of the elements in A1:B4 which are equal to -40
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B4,-40)');
  FWorksheet.CalcFormulas;
  CheckEquals(-40, FWorksheet.ReadAsNumber(0, 2), 'Formula #7 SUMIF(A1:B4,-40) result mismatch');

  // Calculate sum of the elements in A1:B4 which are < 0
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B4,"<0")');
  FWorksheet.CalcFormulas;
  CheckEquals(-43, FWorksheet.ReadAsNumber(0, 2), 'Formula #8 SUMIF(A1:B4,"<0") result mismatch');

  // Calculate sum of the elements in A8:B11 for which the elements in A1:B4 are equal to -40
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B4,-40,A8:B11)');
  FWorksheet.CalcFormulas;
  CheckEquals(400, FWorksheet.ReadAsNumber(0, 2), 'Formula #9 SUMIF(A1:B4,-40,A8:B11) result mismatch');

  // Compare cell A15
  FWorksheet.WriteNumber( 14, 0, -40);

  // Calculate sum of the elements in A1:B4 which are equal to cell A15
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B4,A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(-40, FWorksheet.ReadAsNumber(0, 2), 'Formula #10 SUMIF(A1:B4,A15) result mismatch');

  // Calculate sum of the elements in A1:B4 which are equal <= cell A15
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B4,"<="&A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(-40, FWorksheet.ReadAsNumber(0, 2), 'Formula #11 SUMIF(A1:B4,"<="&A15) result mismatch');

  // Calculate sum of the elements in A8:B11 for which the elements in A1:B4 are equal to cell A15
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B4,A15,A8:B11)');
  FWorksheet.CalcFormulas;
  CheckEquals(400, FWorksheet.ReadAsNumber(0, 2), 'Formula #12 SUMIF(A1:B4,A15,A8:B11) result mismatch');


  // *** Range contains also non-numeric strings ***

  // Calculate sum of the elements in A1:B5 which are equal to -40
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B5,-40)');
  FWorksheet.CalcFormulas;
  CheckEquals(-40, FWorksheet.ReadAsNumber(0, 2), 'Formula #13 SUMIF(A1:B5,-40) result mismatch');

  // Calculate sum of the elements in A1:B5 which are < 0
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B5,"<0")');
  FWorksheet.CalcFormulas;
  CheckEquals(-43, FWorksheet.ReadAsNumber(0, 2), 'Formula #14 SUMIF(A1:B5,"<0") result mismatch');

  // Calculate sum of the elements in A8:B12 for which the elements in A1:B5 are equal to -40
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B5,-40,A8:B12)');
  FWorksheet.CalcFormulas;
  CheckEquals(400, FWorksheet.ReadAsNumber(0, 2), 'Formula #15 SUMIF(A1:B5,-40,A8:B12) result mismatch');

  // *** Range contains also error cells ***

  // Calculate sum of the elements in A1:B5 which are equal to -40  --> error cell must be ignored
  FWorksheet.WriteErrorValue(0, 0, errIllegalRef);  // add error to A1
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B5,-40)');
  FWorksheet.CalcFormulas;
  CheckEquals(-40, FWorksheet.ReadAsNumber(0, 2), 'Formula #16 SUMIF(A1:B5,-40) result mismatch');

  // Calculate sum of the elements in A8:B13 for which the elements in A1:B6 are equal to 40
  FWorksheet.WriteErrorValue(9, 0, errIllegalRef);   // The value in A10 corresponding to 40 is an error now
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B5,40,A8:B12)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 2), 'Formula #17 SUMIF(A1:B5,40,A8:B12) result mismatch');

  // *** Error in arguments

  // Error in first argument
  FWorksheet.WriteFormula(0, 2, '=SUMIF(#DIV/0!,"<10")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #18 SUMIF(#DIV/0,"<10") result mismatch');

  // Error in second argument
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:A9,#DIV/0!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #19 SUMIF(A1:A9,#DIV/0) result mismatch');

  // Error in third argument
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:A9,"=",#DIV/0!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #20 SUMIF(A1:A9,"=",#DIV/0) result mismatch');

  // Write compare cell to contain an error (A15)
  FWorksheet.WriteFormula( 14, 0, '=1/0');   // A15

  // Calculate sum of the elements in A1:B5 which are equal to cell A15 (containing #DIV/0!)
  // but this error does not exist in any cell of A1:B5
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B5,A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 2), 'Formula #21 SUMIF(A1:B5,A15) (A15=#DIV/0!) result mismatch');

  // Calculate sum of the elements in A1:B5 which are equal to cell A15 (containing #DIV/0!)
  // Cell A1 does have this error, so we expect the result to be 1 (1 cell with #DIV/0! error)
  FWorksheet.WriteFormula(0, 0, '=1/0');
  FWorksheet.WriteFormula(0, 2, '=SUMIF(A1:B5,A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #22 SUMIF(A1:B5,A15) (A1=A15=#DIV/0!) result mismatch');
  //CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 2), 'Formula #21 SUMIF(A1:B5,A15) (A15=#DIV/0!) result mismatch');

end;

procedure TCalcFormulaTests.Test_SUMSQ;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 1);
  FWorksheet.WriteNumber (1, 0, 2);
  FWorksheet.WriteNumber (2, 0, 3);
  FWorksheet.WriteText   (3, 0, '4');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=SUMSQ(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 SUMSQ(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(1,2)');
  FWorksheet.CalcFormulas;
  CheckEquals(5, FWorksheet.ReadAsNumber(0, 1), 'Formula #2 SUMSQ(1,2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ("4")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(16, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 SUMSQ("4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(1,2,3,"4")');
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 SUMSQ(1,2,3,"4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 SUMSQ("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 SUMSQ("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(1,2,3,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 SUMSQ(1,2,3,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 SUMSQ(1/0) result mismatch');

  // Count in cell references
  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 SUMSQ(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #8 SUMSQ(A10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(5, FWorksheet.ReadAsNumber(0, 1), 'Formula #9 SUMSQ(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(14, FWorksheet.ReadAsNumber(0, 1), 'Formula #10 SUMSQ(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(14, FWorksheet.ReadAsNumber(0, 1), 'Formula #11 SUMSQ(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A4)');   // "real" and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #12 SUMSQ(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A5)');   // "real" and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 SUMSQ(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 SUMSQ(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A5,1/0)');     // error in argument --> error in result
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 SUMSQ(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A6)');        // error in cell --> error in result
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 SUMSQ(A1:A6) result mismatch');
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

procedure TCalcFormulaTests.Test_UPPER;
begin
  FWorksheet.WriteText(0, 0, 'abc');                   // A1
  FWorksheet.WriteText(1, 0, 'Äbγδ');                  // A2
  FWorksheet.WriteErrorValue(2, 0, errIllegalRef);     // A3

  // ASCII, already upper case
  FWorksheet.WriteFormula(0, 1, '=UPPER("ABC")');
  FWorksheet.CalcFormulas;
  CheckEquals('ABC', FWorksheet.ReadAsText(0, 1), 'Formula #1 UPPER("ABC") result mismatch');

  // ASCII, mixed upper/lower case
  FWorksheet.WriteFormula(0, 1, '=UPPER("Abc")');
  FWorksheet.CalcFormulas;
  CheckEquals('ABC', FWorksheet.ReadAsText(0, 1), 'Formula #2 UPPER("Abc") result mismatch');

  // Unicode, already upper case
  FWorksheet.WriteFormula(0, 1, '=UPPER("ÄÖÜ ΓΔΣ")');
  FWorksheet.CalcFormulas;
  CheckEquals('ÄÖÜ ΓΔΣ', FWorksheet.ReadAsText(0, 1), 'Formula #3 UPPER("ÄÖÜ ΓΔΣ") result mismatch');

  // Unicode, mixed upper/lower case
  FWorksheet.WriteFormula(0, 1, '=UPPER("Äöü Γδσ")');
  FWorksheet.CalcFormulas;
  CheckEquals('ÄÖÜ ΓΔΣ', FWorksheet.ReadAsText(0, 1), 'Formula #4 UPPER("Äöü Γδσ") result mismatch');

  // Mixed unicode, ASCII, number
  FWorksheet.WriteFormula(0, 1, '=UPPER("Äöü Γδσ 123")');
  FWorksheet.CalcFormulas;
  CheckEquals('ÄÖÜ ΓΔΣ 123', FWorksheet.ReadAsText(0, 1), 'Formula #5 UPPER("Äöü Γδσ 123") result mismatch');

  // Error in argument
  FWorksheet.WriteFormula(0, 1, '=UPPER(#REF!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #6 UPPER(#REF!) result mismatch');

  // Cell with error
  FWorksheet.WriteFormula(0, 1, '=UPPER(A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #7 UPPER(A3) (#REF!) result mismatch');

  // Empty cell
  FWorksheet.WriteFormula(0, 1, '=UPPER(A99)');
  FWorksheet.CalcFormulas;
  CheckEquals('', FWorksheet.ReadAsText(0, 1), 'Formula #8 UPPER(A9) (empty) result mismatch');
end;

procedure TCalcFormulaTests.Test_VAR;
const
  EPS = 1E-8;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 1);
  FWorksheet.WriteNumber (1, 0, -2);
  FWorksheet.WriteNumber (2, 0, -3);
  FWorksheet.WriteText   (3, 0, '4');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=VAR(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #1 VAR(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(1,-2)');
  FWorksheet.CalcFormulas;
  CheckEquals(4.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #2 VAR(1,-2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR("4")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #3 VAR("4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(1,-2,-3,"4")');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #4 VAR(1,-2,-3,"4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 VAR("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 VAR("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(1,-2,-3,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 VAR(1,-2,-3,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 VAR(1/0) result mismatch');

  // Cell references
  FWorksheet.WriteFormula(0, 1, '=VAR(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #7 VAR(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 VAR(A10)(A10=empty) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(4.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #9 VAR(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(4.333333333, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #10 VAR(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(4.333333333, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #11 VAR(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=VAR(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #12 VAR(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #13 VAR(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #14 VAR(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=VAR(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 VAR(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VAR(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 VAR(A:A6) result mismatch');
end;

procedure TCalcFormulaTests.Test_VARP;
const
  EPS = 1E-8;
begin
  // Test data
  FWorksheet.WriteNumber (0, 0, 1);
  FWorksheet.WriteNumber (1, 0, -2);
  FWorksheet.WriteNumber (2, 0, -3);
  FWorksheet.WriteText   (3, 0, '4');
  FWorksheet.WriteText   (4, 0, 'abc');
  FWorksheet.WriteFormula(5, 0, '=1/0');

  // Literal values
  FWorksheet.WriteFormula(0, 1, '=VARP(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #1 VAR(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(1,-2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.25, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #2 VARP(1,-2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP("4")');     // although string considered to be numeric
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #3 VARP("4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(1,-2,-3,"4")');
  FWorksheet.CalcFormulas;
  CheckEquals(7.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #4 VARP(1,-2,-3,"4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #5 VARP("abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP("")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #6 VARP("") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(1,-2,-3,"abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #7 VARP(1,-2,-3,"abc") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 VARP(1/0) result mismatch');

  // Cell references
  FWorksheet.WriteFormula(0, 1, '=VARP(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #7 VARP(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #8 VARP(A10)(A10=empty) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(2.25, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #9 VARP(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(2.888888889, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #10 VARP(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(2.888888889, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #11 VARP(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=VARP(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(7.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #12 VARP(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(7.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #13 VARP(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(7.5, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #14 VARP(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=VARP(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #15 VARP(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=VARP(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #16 VARP(A:A6) result mismatch');
end;

initialization
  RegisterTest(TCalcFormulaTests);
end.

