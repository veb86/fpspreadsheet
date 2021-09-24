{ include file for "formulatests.pas", containing the test cases for the
  calc3dformula test. }

  // Setting up some test numbers
  sheet1.WriteText(0, 4, 'abc');       // E1 = 'abc'
  sheet1.WriteNumber(1, 5, 12.0);      // F2 = 12.0

  sheet2.WriteText(2, 1, 'A');         // B3 = 'A'
  sheet2.WriteNumber(1, 4, 1.0);       // E2 = 1.0
  sheet2.WriteNumber(2, 4, -1.0);      // E3 = -1.0
  sheet2.WriteNumber(3, 4, 10.0);      // E4 = 10.0

  sheet3.WriteText(1, 2, 'B');         // C2 = 'B'
  sheet3.WriteNumber(1, 1, 2.0);       // B2 = 2.0
  sheet3.WriteNumber(3, 4, 100);       // E4 = 100.0

//------------------------------------------------------------------------------

  Row := 0;
  formula := 'Sheet2!B3';                                   { A1 }
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, formula);
  SetLength(SollValues, Row+1);
  SollValues[Row] := StringResult('A');

  inc(Row);
  formula := 'Sheet2!B3&Sheet3!C2';                         { A2 }
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, formula);
  SetLength(SollValues, Row+1);
  SollValues[Row] := StringResult('AB');

  inc(Row);
  formula := 'Sheet2!E2';                                   { A3 }
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, formula);
  SetLength(SollValues, Row+1);
  SollValues[Row] := FloatResult(1.0);

  inc(Row);
  formula := 'Sheet2!E2+Sheet3!B2';                         { A4 }
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, formula);
  SetLength(SollValues, Row+1);
  SollValues[Row] := FloatResult(3.0);

  inc(Row);
  formula := 'E1&Sheet2!B3';                                { A5 }
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, 'E1&Sheet2!B3');
  SetLength(SollValues, Row+1);
  SollValues[Row] := StringResult('abcA');

  inc(Row);                                                 { A6 }
  formula := 'F2-Sheet2!E2-11';
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, formula);
  SetLength(SollValues, Row+1);
  SollValues[Row] := FloatResult(0.0);

  inc(Row);
  formula := 'Sheet2!$B$3';                                 { A7 }
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, formula);
  SetLength(SollValues, Row+1);
  SollValues[Row] := StringResult('A');

  inc(Row);
  formula := 'Sheet2!B$3&Sheet3!$C2';                       { A8 }
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, formula);
  SetLength(SollValues, Row+1);
  SollValues[Row] := StringResult('AB');

  inc(Row);
  formula := 'SUM(Sheet2!E2:E4)';                           { A9 }
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, formula);
  SetLength(SollValues, Row+1);
  SollValues[Row] := FloatResult(10.0);

  inc(Row);                                                { A10 }
  formula := 'SUM(Sheet2:Sheet3!E4)';
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, formula);
  SetLength(SollValues, Row+1);
  SollValues[Row] := FloatResult(110.0);

  {
  inc(Row);
  formula := 'D1&Sheet2!B3%"BC"';
  sheet1.WriteText(Row, 0, formula);
  sheet1.WriteFormula(Row, 1, 'D1&Sheet2!B3%"BC"');
  SetLength(SollValues, Row+1);
  SollValues[Row] := StringResult('abcABC');
}

