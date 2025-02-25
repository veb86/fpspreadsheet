{ included by CalcFormulaTests.pas }

procedure TCalcMathFormulaTests.Test_ABS;
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

procedure TCalcMathFormulaTests.Test_ACOS;
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
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #16 ACOS(A1) (A1: #REF!) result mismatch');
end;

procedure TCalcMathFormulaTests.Test_ASIN;
begin
  FWorksheet.WriteFormula(0, 1, '=ASIN(0.5)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/6, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #1 ASIN(0.5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ASIN(0)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #2 ASIN(0) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ASIN(1)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #3 ASIN(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ASIN(-1)');
  FWorksheet.CalcFormulas;
  CheckEquals(-pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #4 ASIN(-1) result mismatch');

  // Out-of-domain
  FWorksheet.WriteFormula(0, 1, '=ASIN(2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #5 ASIN(2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ASIN(-2)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #6 ASIN(-2) result mismatch');

  // Boolean argument
  FWorksheet.WriteFormula(0, 1, '=ASIN(FALSE)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #7 ASIN(FALSE) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ASIN(TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #8 ASIN(TRUE) result mismatch');

  // Numeric string
  FWorksheet.WriteFormula(0, 1, '=ASIN("1")');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #9 ASIN("1") result mismatch');

  // Non-numeric string
  FWorksheet.WriteFormula(0, 1, '=ASIN("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #10 ASIN("abc") result mismatch');

  // Error argument
  FWorksheet.WriteFormula(0, 1, '=ASIN(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #11 ASIN(1/0) result mismatch');

  // Cell with boolean value
  FWorksheet.WriteFormula(0, 0, '=(1=1)');
  FWorksheet.WriteFormula(0, 1, '=ASIN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #12 ASIN(A1) (A1: (1=1)) result mismatch');

  // Cell with numeric string
  FWorksheet.WriteText(0, 0, '1');
  FWorksheet.WriteFormula(0, 1, '=ASIN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #13 ASIN(A1) (A1: "1") result mismatch');

  // Cell with non-numeric string
  FWorksheet.WriteText(0, 0, 'abc');
  FWorksheet.WriteFormula(0, 1, '=ASIN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #14 ASIN(A1) (A1: "abc") result mismatch');

  // Empty cell
  FWorksheet.WriteBlank(0, 0);                    // Empty A1
  FWorksheet.WriteFormula(0, 1, '=ASIN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #15 ASIN(A1) (A1: empty) result mismatch');

  // Cell with error
  FWorksheet.WriteErrorValue(0, 0, errIllegalRef);
  FWorksheet.WriteFormula(0, 1, '=ASIN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #16 ASIN(A1) (A1: #REF!) result mismatch');
end;

procedure TCalcMathFormulaTests.Test_ATAN;
begin
  FWorksheet.WriteFormula(0, 1, '=ATAN(0)');
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #1 ATAN(0) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ATAN(1)');
  FWorksheet.CalcFormulas;
  // Soll result from Wolfram Alpha
  CheckEquals(0.78539816339744830961566084, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #2 ATAN(1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ATAN(-1)');
  FWorksheet.CalcFormulas;
  // Soll result from Wolfram Alpha
  CheckEquals(-0.78539816339744830961566084, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #3 ATAN(-1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ATAN(1E300)');
  FWorksheet.CalcFormulas;
  CheckEquals(pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #4 ATAN(1E300) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ATAN(-1E300)');
  FWorksheet.CalcFormulas;
  CheckEquals(-pi/2, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #5 ATAN(-1E300) result mismatch');

  // Boolean argument
  FWorksheet.WriteFormula(0, 1, '=ATAN(FALSE)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #6 ATAN(FALSE) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=ATAN(TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.78539816339744830961566084, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #7 ATAN(TRUE) result mismatch');

  // Numeric string
  FWorksheet.WriteFormula(0, 1, '=ATAN("1")');
  FWorksheet.CalcFormulas;
  CheckEquals(0.78539816339744830961566084, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #8 ATAN("1") result mismatch');

  // Non-numeric string
  FWorksheet.WriteFormula(0, 1, '=ATAN("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #9 ATAN("abc") result mismatch');

  // Error argument
  FWorksheet.WriteFormula(0, 1, '=ATAN(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #10 ATAN(1/0) result mismatch');

  // Cell with boolean value
  FWorksheet.WriteFormula(0, 0, '=(1=1)');
  FWorksheet.WriteFormula(0, 1, '=ATAN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.78539816339744830961566084, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #11 ATAN(A1) (A1: (1=1)) result mismatch');

  // Cell with numeric string
  FWorksheet.WriteText(0, 0, '1');
  FWorksheet.WriteFormula(0, 1, '=ATAN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.78539816339744830961566084, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #12 ATAN(A1) (A1: "1") result mismatch');

  // Cell with non-numeric string
  FWorksheet.WriteText(0, 0, 'abc');
  FWorksheet.WriteFormula(0, 1, '=ATAN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #13 ATAN(A1) (A1: "abc") result mismatch');

  // Empty cell
  FWorksheet.WriteBlank(0, 0);                    // Empty A1
  FWorksheet.WriteFormula(0, 1, '=ATAN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 ATAN(A1) (A1: empty) result mismatch');

  // Cell with error
  FWorksheet.WriteErrorValue(0, 0, errIllegalRef);
  FWorksheet.WriteFormula(0, 1, '=ATAN(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #15 ATAN(A1) (A1: #REF!) result mismatch');
end;

procedure TCalcMathFormulaTests.Test_CEILING;
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

procedure TCalcMathFormulaTests.Test_EVEN;
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

procedure TCalcMathFormulaTests.Test_FLOOR;
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

procedure TCalcMathFormulaTests.Test_LOG;
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

procedure TCalcMathFormulaTests.Test_LOG10;
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

procedure TCalcMathFormulaTests.Test_ODD;
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

procedure TCalcMathFormulaTests.Test_POWER;
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

procedure TCalcMathFormulaTests.Test_RADIANS;
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

procedure TCalcMathFormulaTests.Test_ROUND;
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

procedure TCalcMathFormulaTests.Test_SQRT;
begin
  FWorksheet.WriteFormula(0, 1, '=SQRT(0.0)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #1 SQRT(0.0) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SQRT(1.0)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #2 SQRT(1.0) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SQRT(16.0)');
  FWorksheet.CalcFormulas;
  CheckEquals(4.0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #3 SQRT(16.0) result mismatch');

  // Out-of-domain
  FWorksheet.WriteFormula(0, 1, '=SQRT(-1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #4 SQRT(-1) result mismatch');

  // Boolean argument
  FWorksheet.WriteFormula(0, 1, '=SQRT(FALSE)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #5 SQRT(FALSE) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SQRT(TRUE)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #6 SQRTS(TRUE) result mismatch');

  // Numeric string
  FWorksheet.WriteFormula(0, 1, '=SQRT("1")');
  FWorksheet.CalcFormulas;
  CheckEquals(1.0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #7 SQRT("1") result mismatch');

  // Non-numeric string
  FWorksheet.WriteFormula(0, 1, '=SQRT("abc")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #8 SQRT("abc") result mismatch');

  // Error argument
  FWorksheet.WriteFormula(0, 1, '=SQRT(1/0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #9 SQRT(1/0) result mismatch');

  // Cell with boolean value
  FWorksheet.WriteFormula(0, 0, '=(1=1)');
  FWorksheet.WriteFormula(0, 1, '=SQRT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #10 SQRT(A1) (A1: (1=1)) result mismatch');

  // Cell with numeric string
  FWorksheet.WriteText(0, 0, '1');
  FWorksheet.WriteFormula(0, 1, '=SQRT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(1.0, FWorksheet.ReadAsNumber(0, 1), 1e-8, 'Formula #11 SQRTS(A1) (A1: "1") result mismatch');

  // Cell with non-numeric string
  FWorksheet.WriteText(0, 0, 'abc');
  FWorksheet.WriteFormula(0, 1, '=SQRT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_WRONG_TYPE, FWorksheet.ReadAsText(0, 1), 'Formula #12 SQRT(A1) (A1: "abc") result mismatch');

  // Empty cell
  FWorksheet.WriteBlank(0, 0);                    // Empty A1
  FWorksheet.WriteFormula(0, 1, '=SQRT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(0.0, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 SQRT(A1) (A1: empty) result mismatch');

  // Cell with error
  FWorksheet.WriteErrorValue(0, 0, errIllegalRef);
  FWorksheet.WriteFormula(0, 1, '=SQRT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #14 SQRT(A1) (A1: #REF!) result mismatch');
end;


