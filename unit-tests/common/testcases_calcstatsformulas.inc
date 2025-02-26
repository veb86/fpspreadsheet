{ included by CalcFormulaTests.pas }

procedure TCalcStatsFormulaTests.Test_AVEDEV;
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

procedure TCalcStatsFormulaTests.Test_AVERAGE;
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

procedure TCalcStatsFormulaTests.Test_AVERAGEIF;
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

  // ToDo: CompareStringsWithWildcards does not handle a mask such as "*b" like Excel            // wp: really?  I thought this was fixed!
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


procedure TCalcStatsFormulaTests.Test_COUNT;
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

procedure TCalcStatsFormulaTests.Test_COUNTA;
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

procedure TCalcStatsFormulaTests.Test_COUNTBLANK;
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

procedure TCalcStatsFormulaTests.Test_COUNTIF;
var
  t: TDateTime;
begin
  // Test data, range A1:B5
  FWorksheet.WriteNumber   (0, 0, 10);        FWorksheet.WriteFormula   (0, 1, '=SQRT(-1)');   // --> #NUM!
  FWorksheet.WriteNumber   (1, 0, -20);       FWorksheet.WriteBlank     (1, 1);
  FWorksheet.WriteFormula  (2, 0, '=(1=1)');  FWorksheet.WriteNumber    (2, 1,  0);
  FWorksheet.WriteText     (3, 0, '');        FWorksheet.WriteText      (3, 1, '5');
  FWorksheet.WriteText     (4, 0, 'abc');     FWorksheet.WriteText      (4, 1, 'ABC');
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
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #6 COUNTIF(A1:B6,FALSE) result mismatch');

  // Counts the elements in A1:B5 which are #NUM!
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,#NUM!)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #7 COUNTIF(A1:B6,#NUM!) result mismatch');

  // Error in 1st argument
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(#REF!,1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 2), 'Formula #8 COUNTIF(#REF!,1) result mismatch');

  // Error in both arguments
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(#REF!,#REF!)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 2), 'Formula #9 COUNTIF(#REF!,#REF!) result mismatch');

  // Count the elements in A1:B6 which are equal to cell A15 (empty)
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #10 COUNTIF(A1:B6,A15) (A15 empty) result mismatch');

  // Count the elements in A1:B6 which are equal to cell A15 (value 10)
  FWorksheet.WriteNumber(14, 0, 10);
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #11 COUNTIF(A1:B6,A15) (A15 = 10) result mismatch');

  // Count the elements in A1:B6 which are < cell A15 (value 10)
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,"<"&A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(3, FWorksheet.ReadAsNumber(0, 2), 'Formula #12 COUNTIF(A1:B6,"<"&A15) (A15 = 10) result mismatch');

  // Count the elements in A1:B6 which are equal to cell A15 (error value #NUM!)
  FWorksheet.WriteErrorValue(14, 0, errOverflow);
  FWorksheet.WriteFormula(0, 2, '=COUNTIF(A1:B6,A15)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #13 COUNTIF(A1:B6,A15) (A15 = #NUM!) result mismatch');

  // Huge range
  FWorksheet.WriteFormula(0, 1, '=COUNTIF(A2:AZ999999,"abc")');   // A2: do not include formula cell in range
  t := Now;
  FWorksheet.CalcFormulas;
  t := Now-t;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 COUNTIF(A2:AZ999999,"abc") result mismatch');
  CheckEquals(true, t < 0.1/(24*60*60), 'Formula #14 COUNTIF(A2:AZ999999,"abc") takes too long');
end;

procedure TCalcStatsFormulaTests.Test_MAX;
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

procedure TCalcStatsFormulaTests.Test_MIN;
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

procedure TCalcStatsFormulaTests.Test_PRODUCT;
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

procedure TCalcStatsFormulaTests.Test_STDEV;
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

procedure TCalcStatsFormulaTests.Test_STDEVP;
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

procedure TCalcStatsFormulaTests.Test_SUM;
var
  t: TDateTime;
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
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #9 SUM(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #10 SUM(A10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #11 SUM(A1,A2) result mismatch');

  // "real" numeric values
  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(60, FWorksheet.ReadAsNumber(0, 1), 'Formula #12 SUM(A1:A3) result mismatch');

  // Several ranges
  FWorksheet.WriteFormula(0, 1, '=SUM(A1,A2:A3)');
  FWorksheet.CalcFormulas;
  CheckEquals(60, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 SUM(A1,A2:A3) result mismatch');

  // Huge range
  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A3,A10:A999999)');
  t := Now;
  FWorksheet.CalcFormulas;
  t := Now-t;
  CheckEquals(60, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 SUM(A1:A3,A10:A999999) result mismatch');
  CheckEquals(true, t < 0.1/(24*60*60), 'Formula #14 SUM(A1:A3,A10:A999999) takes too long');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A4)');   // real and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(0, 1), 'Formula #15 SUM(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A5)');   // real and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(0, 1), 'Formula #16 SUM(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(100, FWorksheet.ReadAsNumber(0, 1), 'Formula #17 SUM(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=SUM(A1, 1/0, A2)');     // error in argument
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #18 SUM(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(A1:A6)');     // error in cell
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #19 SUM(A:A6) result mismatch');
end;

procedure TCalcStatsFormulaTests.Test_SUMIF;
var
  t: TDateTime;
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

  // Huge range
  FWorksheet.WriteFormula(0, 1, '=SUMIF(A2:Z999999,0)');   // A2: do not include formula cell in range
  t := Now;
  FWorksheet.CalcFormulas;
  t := Now-t;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #22 SUMIF(A2:Z999999,0) result mismatch');
  CheckEquals(true, t < 0.1/(24*60*60), 'Formula #22 SUMIF(A2:Z999999,0) takes too long');

end;

procedure TCalcStatsFormulaTests.Test_SUMSQ;
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
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 1), 'Formula #9 SUMSQ(A1) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A10)');     // empty cell
  FWorksheet.CalcFormulas;
  CheckEquals(0, FWorksheet.ReadAsNumber(0, 1), 'Formula #10 SUMSQ(A10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1,A2)');
  FWorksheet.CalcFormulas;
  CheckEquals(5, FWorksheet.ReadAsNumber(0, 1), 'Formula #11 SUMSQ(A1,A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A3)');   // "real" numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(14, FWorksheet.ReadAsNumber(0, 1), 'Formula #12 SUMSQ(A1:A3) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1,A2:A3)');    // several ranges
  FWorksheet.CalcFormulas;
  CheckEquals(14, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 SUMSQ(A1,A2:A3) result mismatch');

  // Cell references pointing to string cells
  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A4)');   // "real" and string numeric values
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 SUMSQ(A1:A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A5)');   // "real" and string values --> ignore string
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #15 SUMSQ(A1:A5) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A4,A8:A10)');   // real and string values and blanks
  FWorksheet.CalcFormulas;
  CheckEquals(30, FWorksheet.ReadAsNumber(0, 1), 'Formula #16 SUMSQ(A1:A4,A8:A10) result mismatch');

  // Error propagation
  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A5,1/0)');     // error in argument --> error in result
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #17 SUMSQ(A1, 1/0, A2) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUMSQ(A1:A6)');        // error in cell --> error in result
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #18 SUMSQ(A1:A6) result mismatch');
end;

procedure TCalcStatsFormulaTests.Test_VAR;
const
  EPS = 1E-8;
var
  t: TDateTime;
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

  // Huge range
  FWorksheet.WriteFormula(0, 1, '=VAR(A1:A3,A10:A9999999)');
  t := Now();
  FWorksheet.CalcFormulas;
  t := Now() - t;
  CheckEquals(4.333333333, FWorksheet.ReadAsNumber(0, 1), EPS, 'Formula #17 VAR(A1:A3,A10:9999999) result mismatch');
  CheckEquals(true, t < 0.1/(24*60*60), 'Formula #17 VAR(A1:A3,A10:A9999999) takes too long');
end;

procedure TCalcStatsFormulaTests.Test_VARP;
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


