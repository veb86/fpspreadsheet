{ included by CalcFormulaTests.pas }

procedure TCalcInfoFormulaTests.Test_ERRORTYPE;
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

procedure TCalcInfoFormulaTests.Test_IFERROR;
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


procedure TCalcInfoFormulaTests.Test_ISBLANK;
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

procedure TCalcInfoFormulaTests.Test_ISERR;
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

procedure TCalcInfoFormulaTests.Test_ISERROR;
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

procedure TCalcInfoFormulaTests.Test_ISLOGICAL;
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

procedure TCalcInfoFormulaTests.Test_ISNA;
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

procedure TCalcInfoFormulaTests.Test_ISNONTEXT;
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

procedure TCalcInfoFormulaTests.Test_ISNUMBER;
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

procedure TCalcInfoFormulaTests.Test_ISREF;
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

procedure TCalcInfoFormulaTests.Test_ISTEXT;
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

