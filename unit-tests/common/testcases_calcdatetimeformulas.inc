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

procedure TCalcDateTimeFormulaTests.Test_DateAdd;
var
  date1, date2, delta: Double;
  cell: PCell;
begin
  date1 := EncodeDate(2024,10,1);
  date2 := EncodeDate(2024,10,11);
  delta := 10;
  FWorksheet.WriteDateTime(0, 0, date1, nfShortDate);
  FWorksheet.WriteDateTime(1, 0, date2, nfShortDate);
  FWorksheet.WriteNumber(2, 0, delta);

  // Date cell + number --> Date cell
  cell := FWorksheet.WriteFormula(0, 1, 'A1+10');
  FWorksheet.CalcFormulas;
  CheckEquals(date1+delta, FWorksheet.ReadAsNumber(cell), 'Formula #1 A1+10 result mismatch');
  CheckEquals(ord(cctDateTime), ord(cell^.ContentType), 'Formula #1 A1+10 content type mismatch');

  // Number + Date cell --> Date cell
  cell := FWorksheet.WriteFormula(0, 1, '10+A1');
  FWorksheet.CalcFormulas;
  CheckEquals(delta+date1, FWorksheet.ReadAsNumber(cell), 'Formula #2 10+A1 result mismatch');
  CheckEquals(ord(cctDateTime), ord(cell^.ContentType), 'Formula #2 10+A1 content type mismatch');

  // Date cell + Date cell --> Number cell
  cell := FWorksheet.WriteFormula(0, 1, 'A1+A2');
  FWorksheet.CalcFormulas;
  CheckEquals(date1+date2, FWorksheet.ReadAsNumber(cell), 'Formula #3 A1+A2 result mismatch');
  CheckEquals(ord(cctNumber), ord(cell^.ContentType), 'Formula #3 A1+A2 content type mismatch');

  (*     --- This crashes in FPS, it works in Excel It is kind of nonsense, anyway...
  // Date string + number = Number cell
  cell := FWorksheet.WriteFormula(0, 1, Format('="%s"+10', [DateToStr(date1)]));
  FWorksheet.CalcFormulas;
  CheckEquals(date1+10, FWorksheet.ReadAsNumber(cell), 'Formula #4 datestring+10 result mismatch');
  CheckEquals(ord(cctNumber), ord(cell^.ContentType), 'Formula #4 datestring+10 content type mismatch');
  *)
end;

procedure TCalcDateTimeFormulaTests.Test_DATEDIF;
var
  date1, date2: TDateTime;
  expectedDiff: Integer;
  actualDiff: Double;
begin
  date1 := EncodeDate(2024,10,1);
  date2 := EncodeDate(2025,11,10);

  FWorksheet.WriteDateTime(0, 0, date1, nfShortDate);
  FWorksheet.WriteDateTime(1, 0, date2, nfShortDate);

  // DATEDIF as days
  // ... between cell references
  expectedDiff := DaysBetween(date1, date2);
  FWorksheet.WriteFormula(0, 1, '=DATEDIF(A1,A2,"D")');
  FWorksheet.CalcFormulas;
  actualDiff := FWorksheet.ReadAsNumber(0, 1);
  CheckEquals(expectedDiff, actualDiff, 'Formula #1 DATEDIFF(A1,A2,"D") result mismatch.');
  // ... between number constants
  FWorksheet.WriteFormula(0, 1, Format('=DATEDIF(%.0f,%.0f,"D")',[date1,date2]));
  FWorksheet.CalcFormulas;
  actualDiff := FWorksheet.ReadAsNumber(0, 1);
  CheckEquals(expectedDiff, actualDiff, 'Formula #2 DATEDIFF(date1,date2,"D") result mismatch.');
  // ... between strings
  FWorksheet.WriteFormula(0, 1, Format('=DATEDIF("%s","%s","D")',[DateToStr(date1),DateToStr(date2)]));
  FWorksheet.CalcFormulas;
  actualDiff := FWorksheet.ReadAsNumber(0, 1);
  CheckEquals(expectedDiff, actualDiff, 'Formula #3 DATEDIFF("date1","date2","D") result mismatch.');
  // ... mixed: between cell and string
  FWorksheet.WriteFormula(0, 1, Format('=DATEDIF(A1,"%s","D")',[DateToStr(date2)]));
  FWorksheet.CalcFormulas;
  actualDiff := FWorksheet.ReadAsNumber(0, 1);
  CheckEquals(expectedDiff, actualDiff, 'Formula #4 DATEDIFF(A1,"date2","D") result mismatch.');
  // ... mixed: between integer const and string
  FWorksheet.WriteFormula(0, 1, Format('=DATEDIF(%.0f,"%s","D")',[date1,DateToStr(date2)]));
  FWorksheet.CalcFormulas;
  actualDiff := FWorksheet.ReadAsNumber(0, 1);
  CheckEquals(expectedDiff, actualDiff, 'Formula #5 DATEDIFF(date1,"date2","D") result mismatch.');

  // DATEDIF as monthsbetween cell references (months)
  // ... between cell references
  expectedDiff := MonthsBetween(date1, date2);
  FWorksheet.WriteFormula(0, 1, '=DATEDIF(A1,A2,"M")');
  FWorksheet.CalcFormulas;
  actualDiff := FWorksheet.ReadAsNumber(0, 1);
  CheckEquals(expectedDiff, actualDiff, 'Formula #6 DATEDIFF(A1,A2,"M") result mismatch.');

  // DATEDIF between cell references (years)
  expectedDiff := YearsBetween(date1, date2);
  FWorksheet.WriteFormula(0, 1, '=DATEDIF(A1,A2,"Y")');
  FWorksheet.CalcFormulas;
  actualDiff := FWorksheet.ReadAsNumber(0, 1);
  CheckEquals(expectedDiff, actualDiff, 'Formula #7 DATEDIFF(A1,A2,"Y") result mismatch.');

  // Reverse order --> error #NUM!
  FWorksheet.WriteFormula(0, 1, '=DATEDIF(A2,A1,"Y")');
  FWorksheet.CalcFormulas;
  CheckEquals(STR_ERR_OVERFLOW, FWorksheet.ReadAsText(0, 1), 'Formula #8 DATEDIFF(A2,A1,"M") error not detected.');
end;


procedure TCalcDateTimeFormulaTests.Test_DATEVALUE;
var
  expectedDateValue, actualDateValue: Double;
  s: String;
begin
  expectedDateValue := EncodeDate(2000,1,12);

  // String formatted in system's formatsettings;
  FWorksheet.WriteFormula(0, 1, '=DATEVALUE("' + DateToStr(expectedDateValue) + '")');
  FWorksheet.CalcFormulas;
  FWorksheet.ReadAsDateTime(0, 1, actualDateValue);
  CheckEquals(expectedDateValue, actualDateValue, 'Formula #1 DATEVALUE("<2000-01-12 in system format>") result mismatch');

  // String formatted according to arbitrary FormatSettings
  FWorksheet.WriteFormula(0, 1, '=DATEVALUE("2000-01-12")');
  FWorksheet.CalcFormulas;
  FWorksheet.ReadAsDateTime(0, 1, actualDateValue);
  CheckEquals(expectedDateValue, actualDateValue, 'Formula #2 DATEVALUE("2000-01-12") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=DATEVALUE("2000/01/12")');
  FWorksheet.CalcFormulas;
  FWorksheet.ReadAsDateTime(0, 1, actualDateValue);
  CheckEquals(expectedDateValue, actualDateValue, 'Formula #3 DATEVALUE("2000/01/12") result mismatch');

  FWorksheet.WriteFormula(0, 1, '=DATEVALUE("2000.01.12")');
  FWorksheet.CalcFormulas;
  FWorksheet.ReadAsDateTime(0, 1, actualDateValue);
  CheckEquals(expectedDateValue, actualDateValue, 'Formula #4 DATEVALUE("2000.01.12") result mismatch');

  { ABBREVIATED OR FULL MONTH NAMES SEEMINGLY ARE NOT SUPPORTED BY FPS HERE....
  // With abbreviated month name
  s := FormatDateTime('yyyy-mmmm-d', expectedDateValue);
  FWorksheet.WriteFormula(0, 1, '=DATEVALUE("' + s + '")');
  FWorksheet.CalcFormulas;
  FWorksheet.ReadAsDateTime(0, 1, actualDateValue);
  CheckEquals(expectedDateValue, actualDateValue, 'Formula #5 DATEVALUE("' + s + '") result mismatch');
  }

  // Reference to string cell
  FWorksheet.WriteText(0,0, DateToStr(expectedDateValue));
  FWorksheet.WriteFormula(0, 1, '=DATEVALUE(A1)');
  FWorksheet.CalcFormulas;
  FWorksheet.ReadAsDateTime(0, 1, actualDateValue);
  CheckEquals(expectedDateValue, actualDateValue, 'Formula #5 DATEVALUE(A1) result mismatch');

  FWorksheet.WriteText(0,0, '2000-01-12');
  FWorksheet.WriteFormula(0, 1, '=DATEVALUE(A1)');
  FWorksheet.CalcFormulas;
  FWorksheet.ReadAsDateTime(0, 1, actualDateValue);
  CheckEquals(expectedDateValue, actualDateValue, 'Formula #6 DATEVALUE(A1) (A1 = "2000-01-12") result mismatch');

  FWorksheet.WriteText(0,0, '2000/01/12');
  FWorksheet.WriteFormula(0, 1, '=DATEVALUE(A1)');
  FWorksheet.CalcFormulas;
  FWorksheet.ReadAsDateTime(0, 1, actualDateValue);
  CheckEquals(expectedDateValue, actualDateValue, 'Formula #7 DATEVALUE(A1) (A1 = "2000/01/12") result mismatch');

  FWorksheet.WriteText(0,0, '2000.01.12');
  FWorksheet.WriteFormula(0, 1, '=DATEVALUE(A1)');
  FWorksheet.CalcFormulas;
  FWorksheet.ReadAsDateTime(0, 1, actualDateValue);
  CheckEquals(expectedDateValue, actualDateValue, 'Formula #8 DATEVALUE(A1) (A1 = "2000.01.12") result mismatch');

  // Year missing (will be replaced by current year)
  expectedDateValue := EncodeDate(YearOf(Date), 1, 12);
  FWorksheet.WriteFormula(0, 1, '=DATEVALUE("01-12")');
  FWorksheet.CalcFormulas;
  FWorksheet.ReadAsDateTime(0, 1, actualDateValue);
  CheckEquals(expectedDateValue, actualDateValue, 'Formula #9 DATEVALUE("01-12") result mismatch');

end;

procedure TCalcDateTimeFormulaTests.Test_DateSubtract;
var
  date1, date2, delta: Double;
  cell: PCell;
begin
  date1 := EncodeDate(2024,10,1);
  date2 := EncodeDate(2024,10,11);
  delta := 10;
  FWorksheet.WriteDateTime(0, 0, date1, nfShortDate);
  FWorksheet.WriteDateTime(1, 0, date2, nfShortDate);
  FWorksheet.WriteNumber(2, 0, delta);

  // Date cell - number --> Date cell
  cell := FWorksheet.WriteFormula(0, 1, 'A1-10');
  FWorksheet.CalcFormulas;
  CheckEquals(date1-delta, FWorksheet.ReadAsNumber(cell), 'Formula #1 A1-10 result mismatch');
  CheckEquals(ord(cctDateTime), ord(cell^.ContentType), 'Formula #1 A1-10 content type mismatch');

  // Number - Date cell = Date cell (displayed in grid as #######... when negative
  cell := FWorksheet.WriteFormula(0, 1, '50000-A1');
  FWorksheet.CalcFormulas;
  CheckEquals(50000-date1, FWorksheet.ReadAsNumber(cell), 'Formula #2 50000-A1 result mismatch');
  CheckEquals(ord(cctDateTime), ord(cell^.ContentType), 'Formula #2 50000-A1 content type mismatch');

  // Date cell - Date cell = Number cell
  cell := FWorksheet.WriteFormula(0, 1, 'A2-A1');
  FWorksheet.CalcFormulas;
  CheckEquals(date2-date1, FWorksheet.ReadAsNumber(cell), 'Formula #3 A2-A1 result mismatch');
  CheckEquals(ord(cctNumber), ord(cell^.ContentType), 'Formula #3 A2-A1 content type mismatch');

  (*     --- This crashes in FPS, it works in Excel. It is kind of nonsense, anyway...
  // Date string - number = Date cell
  cell := FWorksheet.WriteFormula(0, 1, Format('="%s"-10', [DateToStr(date1)]));
  FWorksheet.CalcFormulas;
  CheckEquals(date1-10, FWorksheet.ReadAsNumber(cell), 'Formula #4 datestring-10 result mismatch');
  CheckEquals(ord(cctNumber), ord(cell^.ContentType), 'Formula #4 datestring-10 content type mismatch');
  *)
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


