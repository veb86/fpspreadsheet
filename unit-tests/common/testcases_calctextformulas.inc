{ included by CalcFormulaTests.pas }

// *** String formula tests

procedure TCalcTextFormulaTests.Test_CONCATENATE;
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

  {  -- this test will not work in the file because Excel writes a localized "TRUE" to the cell
  // Concatenate string and boolean
  FWorksheet.WriteFormula(0, 1, '=CONCATENATE("abc",TRUE())');
  FWorksheet.CalcFormulas;
  CheckEquals('abcTRUE', FWorksheet.ReadAsText(0, 1), 'Formula #5 CONCATENATE("abc",TRUE()) result mismatch');
  }

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

procedure TCalcTextFormulaTests.Test_EXACT;
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

procedure TCalcTextFormulaTests.Test_LEN;
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

procedure TCalcTextFormulaTests.Test_LOWER;
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

procedure TCalcTextFormulaTests.Test_UPPER;
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


