{ included by CalcFormulaTests.pas }

procedure TCalcLookupFormulaTests.Test_ADDRESS;
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

procedure TCalcLookupFormulaTests.Test_COLUMN;
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

procedure TCalcLookupFormulaTests.Test_INDEX;
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

procedure TCalcLookupFormulaTests.Test_INDIRECT;
begin
  // *** Test data ***
  FWorksheet.WriteNumber (0, 0, 10);                    // A1
  FWorksheet.WriteNumber (1, 0, 20);                    // A2
  FWorksheet.WriteNumber (2, 0, 30);                    // A3
  FWorksheet.WriteText   (3, 0, 'A1');                  // A4
  FWorksheet.WriteErrorValue(4, 0, errDivideByZero);    // A5
  FWorksheet.WriteText   (5, 0, 'A');                   // A6
  FWorksheet.WriteNumber (6, 0, 1);                     // A7
  FWorksheet.WriteText   (7, 0, 'Sheet2');              // A8
  FWorksheet.WriteText   (8, 0, 'Sheet2!A1:A3');        // A9
  FWorksheet.WriteFormula(9, 0, '=A6&A7');              // A10
  FOtherWorksheet.WriteNumber(0, 0, 1000);              // Sheet2!A1
  FOtherWorksheet.WriteNumber(1, 0, 2000);              // Sheet2!A2
  FOtherWorksheet.WriteNumber(2, 0, 3000);              // Sheet2!A3

  // *** Single cell references ***
  FWorksheet.WriteFormula(0, 1, '=INDIRECT("A1")');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #1 INDIRECT("A1") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=INDIRECT("A4")');
  FWorksheet.CalcFormulas;
  CheckEquals('A1', FWorksheet.ReadAsText(0, 1), 'Formula #2 INDIRECT("A4") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=INDIRECT(A4)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #3 INDIRECT(A4) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=INDIRECT("Sheet2!A1")');
  FWorksheet.CalcFormulas;
  CheckEquals(1000, FWorksheet.ReadAsNumber(0, 1), 'Formula #4 INDIRECT("Sheet2!A1") result mismatch');

  // Constructing cell address from other cells
  FWorksheet.WriteFormula(0, 1, '=INDIRECT(A6&A7)');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #5 INDIRECT(A6&A7) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=INDIRECT(A10)');     // A10 = A6&A7
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #6 INDIRECT(A10) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=INDIRECT(A8&"!"&A6&A7)');     // --> "Sheet2!A1"
  FWorksheet.CalcFormulas;
  CheckEquals(1000, FWorksheet.ReadAsNumber(0, 1), 'Formula #7 INDIRECT(A8&"!"&A6&A7) result mismatch');

  // Constructing cell address from other cells and constant
  FWorksheet.WriteFormula(0, 1, '=INDIRECT(A6&"1")');
  FWorksheet.CalcFormulas;
  CheckEquals(10, FWorksheet.ReadAsNumber(0, 1), 'Formula #8 INDIRECT(A6&"1") result mismatch');

  // Error in indirectly addressed cell
  FWorksheet.WriteFormula(0, 1, '=INDIRECT(A5)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_DIVIDE_BY_ZERO, FWorksheet.ReadAsText(0, 1), 'Formula #9 INDIRECT(A5) result mismatch');

  // Circular reference
  FWorksheet.WriteFormula(0, 1, '=INDIRECT(A1)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ILLEGAL_REF, FWorksheet.ReadAsText(0, 1), 'Formula #10 INDIRECT(A1) result mismatch');


  // *** Cell range references ***

  FWorksheet.WriteFormula(0, 1, '=SUM(INDIRECT("A1:A3"))');
  FWorksheet.CalcFormulas;
  CheckEquals(60, FWorksheet.ReadAsNumber(0, 1), 'Formula #11 SUM(INDIRECT("A1:A3")) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(INDIRECT("Sheet2!A1:A3"))');
  FWorksheet.CalcFormulas;
  CheckEquals(6000, FWorksheet.ReadAsNumber(0, 1), 'Formula #12 SUM(INDIRECT("Sheet2!A1:A3")) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(INDIRECT(A9))');
  FWorksheet.CalcFormulas;
  CheckEquals(6000, FWorksheet.ReadAsNumber(0, 1), 'Formula #13 SUM(INDIRECT(A9)) result mismatch');

  FWorksheet.WriteFormula(0, 1, '=SUM(INDIRECT(A8&"!"&A4&":A3"))');
  FWorksheet.CalcFormulas;
  CheckEquals(6000, FWorksheet.ReadAsNumber(0, 1), 'Formula #14 SUM(INDIRECT(A8&"!"&A4&":A3")) result mismatch');
end;

procedure TCalcLookupFormulaTests.Test_MATCH;
begin
  // *** Match_Type 0, unsorted data in search range, find first value

  // Search range to be checked: B1:B4
  FWorksheet.WriteNumber(0, 1, 10);
  FWorksheet.WriteNumber(1, 1, 20);
  FWorksheet.WriteNumber(2, 1, 30);
  FWorksheet.WriteNumber(3, 1, 15);
  FWorksheet.WriteNumber(4, 1, 20);

  // Search range in other sheet
  FOtherWorksheet.WriteNumber(0, 1, 100);
  FOtherWorksheet.WriteNumber(1, 1, 200);
  FOtherWorksheet.WriteNumber(2, 1, 300);
  FOtherWorksheet.WriteNumber(3, 1, 150);
  FOtherWorksheet.WriteNumber(4, 1, 200);

  // Search for constant, contained in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(10, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #1 MATCH(10,B1:B5,0) mismatch, value in range');

  // dto., but in other sheet
  FWorksheet.WriteFormula(0, 2, '=MATCH(100, Sheet2!B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(1, FWorksheet.ReadAsNumber(0, 2), 'Formula #2 MATCH(100,Sheet2!B1:B5,0) mismatch, value in range');

  // Search for constant, contained several times in search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(20, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #3 MATCH(20,B1:B5,0) mismatch, value above range');

  // dto., but in other sheet
  FWorksheet.WriteFormula(0, 2, '=MATCH(200, Sheet2!B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #4 MATCH(200,Sheet2!B1:B5,0) mismatch, value above range');

  // Search for constant, below search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(0, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #5 MATCH(0,B1:B5,0) mismatch, value below range');

  // dto., but in other sheet
  FWorksheet.WriteFormula(0, 2, '=MATCH(0, Sheet2!B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #6 MATCH(0,Sheet2!B1:B5,0) mismatch, value below range');

  // Search for constant, above search range
  FWorksheet.WriteFormula(0, 2, '=MATCH(90, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #7 MATCH(90,B1:B5,0) mismatch, value above range');

  // dto., but in other sheet
  FWorksheet.WriteFormula(0, 2, '=MATCH(900, Sheet2!B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #8 MATCH(900,Sheet2!B1:B5,0) mismatch, value above range');

  // Search for cell with value in range
  FWorksheet.WriteNumber(0, 0, 20);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #9 MATCH(A1,B1:B5,0) mismatch, cell value in range');

  // dto, but in other sheet
  FWorksheet.WriteNumber(0, 0, 200);
  FWorksheet.WriteFormula(0, 2, '=MATCH(A1, Sheet2!B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(2, FWorksheet.ReadAsNumber(0, 2), 'Formula #10 MATCH(A1,Sheet2!B1:B5,0) mismatch, cell value in range');

  // Search for cell, but cell is empty
  FWorksheet.WriteFormula(0, 2, '=MATCH(A99, B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #11 MATCH(A99,B1:B5,0) mismatch, empty cell');

  // dto., but in other sheet
  FWorksheet.WriteFormula(0, 2, '=MATCH(A99, Sheet2!B1:B5, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #12 MATCH(A99,Sheet2!B1:B5,0) mismatch, empty cell');

  // Search range is empty
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, D1:D3, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #13 MATCH mismatch, match_type -1, empty search range');

  // dto., but in other sheet
  FWorksheet.WriteFormula(0, 2, '=MATCH(28, Sheet2!D1:D3, 0)');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_ARG_ERROR, FWorksheet.ReadAsText(0, 2), 'Formula #14 MATCH mismatch, match_type -1, empty search range');

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

procedure TCalcLookupFormulaTests.Test_ROW;
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

