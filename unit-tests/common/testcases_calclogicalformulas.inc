{ included by CalcFormulaTests.pas }

procedure TCalcLogicalFormulaTests.Test_AND;
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

procedure TCalcLogicalFormulaTests.Test_IF;
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

procedure TCalcLogicalFormulaTests.Test_NOT;
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

procedure TCalcLogicalFormulaTests.Test_OR;
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


