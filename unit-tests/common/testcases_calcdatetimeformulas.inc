{ included by CalcFormulaTests.pas }

procedure TCalcDateTimeFormulaTests.Test_DATE;
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

procedure TCalcDateTimeFormulaTests.Test_TIME;
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


