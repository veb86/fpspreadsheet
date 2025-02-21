program demo_conditional_formatting;

uses
  sysUtils,
  fpsTypes, fpsUtils, fpspreadsheet, fpsConditionalFormat,
  xlsxooxml, xlsxml, fpsOpenDocument;

var
  wb: TsWorkbook;
  sh: TsWorksheet;
  fmt: TsCellFormat;
  fmtIdx: Integer;
  row: Integer;
  i: Integer;
  lastCol: Integer;
  dir: String;

begin
  wb := TsWorkbook.Create;
  try
    sh := wb.AddWorksheet('test');
    sh.WriteDefaultColWidth(15, suMillimeters);

    sh.WriteText(0, 0, 'Condition');
    sh.WriteColWidth(0, 70, suMillimeters);
    sh.WriteText(0, 1, 'Format');
    sh.WriteColWidth(1, 90, suMillimeters);
    sh.WriteText(0, 2, 'Test values');

    row := 2;
    for i := row to row+42 do
    begin
      sh.WriteNumber(i, 2, 1.0);
      sh.WriteNumber(i, 3, 2.0);
      sh.WriteNumber(i, 4, 3.0);
      sh.WriteNumber(i, 5, 4.0);
      sh.WriteNumber(i, 6, 5.0);
      sh.WriteNumber(i, 7, 6.0);
      sh.WriteNumber(i, 8, 7.0);
      sh.WriteNumber(i, 9, 8.0);
      sh.WriteNumber(i, 10, 9.0);
      //sh.WriteNumber(i, 11, 10.0);
      sh.WriteText(i, 12, 'abc');
      sh.WriteText(i, 13, 'abc');
      sh.WriteBlank(i, 14);
//      sh.WriteText(i, 14, '');
      sh.WriteText(i, 15, 'def');
      sh.WriteText(i, 16, 'defg');
      sh.WriteFormula(i, 17, '=1.0/0.0');
      sh.WriteFormula(i, 18, '=1.0/1.0');

      sh.WriteDateTime(i, 19, Now()- 30);
      sh.WriteDateTime(i, 20, Now() - 7);
      sh.WritedateTime(i, 21, Now() - 1);
      sh.WriteDatetime(i, 22, Now());
      sh.WriteDateTime(i, 23, Now() + 1);
      sh.WriteDateTime(i, 24, Now() + 7);
      sh.WriteDateTime(i, 25, Now() + 30);
    end;
    lastCol := 25;

    for i := 19 to 25 do
      sh.WriteColWidth(i, 30, suMillimeters);

    // conditional format #1: equal to number constant
    sh.WriteText(row, 0, 'equal to constant 5');
    sh.WriteText(row, 1, 'background yellow');

    // prepare cell format tempate
    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    // Write conditional format
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcEqual, 5, fmtIdx);

    // conditional format #2: equal to text constant
    inc(row);
    sh.WriteText(row, 0, 'equal to text "abc"');
    sh.WriteText(row, 1, 'background green, bold text');
    fmt.SetBackgroundColor(scGreen);
    fmt.SetFont(2);  // Font #2 in fps is bold, by default.
    fmtIdx := wb.AddCellFormat(fmt);
    // Write conditional format
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcEqual, 'abc', fmtIdx);

    // conditional format #3: greater than cell reference
    inc(row);
    sh.WriteText(row, 0, 'greater than cell F4');
    sh.WriteText(row, 1, 'all borders, red, thick line');
    InitFormatRecord(fmt);
    fmt.SetBorders([cbEast, cbWest, cbNorth, cbSouth], scRed, lsThick);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcGreaterThan, 'F4', fmtIdx);  // Absolute ref needed but generated automatically

    // conditional format #4: less than formula
    inc(row);
    sh.WriteText(row, 0, 'less than formula "=$C$4+$D$4"');
    sh.WriteText(row, 1, 'background red');
    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcLessThan, '=$C$4+$D$4', fmtIdx);    // Absolute ref required

    // conditional format #5: greater equal constant
    inc(row);
    sh.WriteText(row, 0, 'greater equal constant 5');
    sh.WriteText(row, 1, 'background gray');
    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scGray);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcGreaterEqual, 5, fmtIdx);

    // conditional format #6: less equal constant
    inc(row);
    sh.WriteText(row, 0, 'less equal constant 5');
    sh.WriteText(row, 1, 'background gray');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcLessEqual, 5, fmtIdx);

    // conditional format #7: between
    inc(row);
    sh.WriteText(row, 0, 'between 2 and 7');
    sh.WriteText(row, 1, 'background light gray');
    fmt.SetBackgroundColor($EEEEEE);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcBetween, 2, 7, fmtIdx);

    // conditional format #8: not between
    inc(row);
    sh.WriteText(row, 0, 'not between 2 and 7');
    sh.WriteText(row, 1, 'background light gray');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcNotBetween, 2, 7, fmtIdx);

    // conditional format #9: above average
    inc(row);
    sh.WriteText(row, 0, '> average');
    sh.WriteText(row, 1, 'hatched background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsThinStripeDiagUp, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, 10), cfcAboveAverage, fmtIdx);   // only 1..9 -> ave = 5

    // conditional format #10: below average
    inc(row);
    sh.WriteText(row, 0, '< average');
    sh.WriteText(row, 1, 'dotted background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsGray25, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, 10), cfcBelowAverage, fmtIdx);   // only 1..9 -> ave = 5

    // conditional format #11: above or equal to average
    inc(row);
    sh.WriteText(row, 0, '>= average');
    sh.WriteText(row, 1, 'hor striped background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsThinStripeHor, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, 10), cfcAboveEqualAverage, fmtIdx);  // only 1..9 -> ave = 5

    // conditional format #12: below or equal to average
    inc(row);
    sh.WriteText(row, 0, '<= average');
    sh.WriteText(row, 1, 'vert striped background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsThinStripeVert, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, 10), cfcBelowEqualAverage, fmtIdx);  // only 1..9 -> ave = 5

    // conditional format #13: top 3 values
    inc(row);
    sh.WriteText(row, 0, 'top 3 values');
    sh.WriteText(row, 1, 'background green');
    fmt.SetBackgroundColor(scGreen);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcTop, 3, fmtIdx);

    // conditional format #14: smallest 3 values
    inc(row);
    sh.WriteText(row, 0, 'smallest 3 values');
    sh.WriteText(row, 1, 'background bright blue');
    fmt.SetBackgroundColor($FFC0C0);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcBottom, 3, fmtIdx);

    // conditional format #15: top 10 percent
    inc(row);
    sh.WriteText(row, 0, 'top 10 percent');
    sh.WriteText(row, 1, 'background green');
    fmt.SetBackgroundColor(scGreen);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcTopPercent, 10, fmtIdx);

    // conditional format #16: smallest 10 percent
    inc(row);
    sh.WriteText(row, 0, 'smallest 10 percent');
    sh.WriteText(row, 1, 'background bright blue');
    fmt.SetBackgroundColor($FFC0C0);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcBottomPercent, 10, fmtIdx);

    // conditional format #17: duplicates
    inc(row);
    sh.WriteText(row, 0, 'duplicate values');
    sh.WriteText(row, 1, 'background bright red');
    fmt.SetBackgroundColor($D0D0FF);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcDuplicate, fmtIdx);

    // conditional format #18: unique
    inc(row);
    sh.WriteText(row, 0, 'unique values');
    sh.WriteText(row, 1, 'borders all sides');
    InitFormatRecord(fmt);
    fmt.SetBorders(ALL_BORDERS);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcUnique, fmtIdx);

    // conditional format #19: contains any text
    inc(row);
    sh.WriteText(row, 0, 'contains any text');
    sh.WriteText(row, 1, 'background red');
    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcContainsText, '', fmtIdx);

    // conditional format #20: empty
    inc(row);
    sh.WriteText(row, 0, 'empty');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcNotContainsText, '', fmtIdx);

    // conditional format #21: text begins with 'ab'
    inc(row);
    sh.WriteText(row, 0, 'text begins with "ab"');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcBeginsWith, 'ab', fmtIdx);

    // conditional format #22: text ends with 'g'
    inc(row);
    sh.WriteText(row, 0, 'text ends with "g"');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcEndsWith, 'g', fmtIdx);

    // conditional format #23: text contains 'ef'
    inc(row);
    sh.WriteText(row, 0, 'text contains "ef"');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcContainsText, 'ef', fmtIdx);

    // conditional format #24: text does NOT contain 'ef'
    inc(row);
    sh.WriteText(row, 0, 'text does not contain "ef"');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcNotContainsText, 'ef', fmtIdx);

    // conditional format #25: contains error
    inc(row);
    sh.WriteText(row, 0, 'contains error');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcContainsErrors, fmtIdx);

    // conditional format #26: no errors
    inc(row);
    sh.WriteText(row, 0, 'no errors');
    sh.WriteText(row, 1, 'background yellow, font "Courier New"/red/bold/14');
    fmt.SetBackgroundColor(scYellow);
    fmt.SetFont(wb.AddFont('Courier New', 14, [fssBold], scRed));
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcNotContainsErrors, fmtIdx);

    // conditional date formats
    inc(row);
    sh.WriteText(row, 0, 'yesterday');
    sh.WriteText(row, 1, 'background yellow');
    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scYellow);
    fmt.SetNumberFormat(wb.AddNumberFormat('yyyy\-mm\-dd'));
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcYesterday, fmtIdx);

    inc(row);
    sh.WriteText(row, 0, 'today');
    sh.WriteText(row, 1, 'background yellow');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcToday, fmtIdx);

    inc(row);
    sh.WriteText(row, 0, 'tomorrow');
    sh.WriteText(row, 1, 'background yellow');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcTomorrow, fmtIdx);

    inc(row);
    sh.WriteText(row, 0, 'last 7 days');
    sh.WriteText(row, 1, 'background yellow');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcLast7Days, fmtIdx);

    inc(row);
    sh.WriteText(row, 0, 'last week');
    sh.WriteText(row, 1, 'background yellow');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcLastWeek, fmtIdx);

    inc(row);
    sh.WriteText(row, 0, 'this week');
    sh.WriteText(row, 1, 'background yellow');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcThisWeek, fmtIdx);

    inc(row);
    sh.WriteText(row, 0, 'next week');
    sh.WriteText(row, 1, 'background yellow');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcNextWeek, fmtIdx);

    inc(row);
    sh.WriteText(row, 0, 'last month');
    sh.WriteText(row, 1, 'background yellow');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcLastMonth, fmtIdx);
                            (*
    inc(row);
    sh.WriteText(row, 0, 'tomorrow');
    sh.WriteText(row, 1, 'background yellow');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcThisMonth, fmtIdx);

    inc(row);
    sh.WriteText(row, 0, 'tomorrow');
    sh.WriteText(row, 1, 'background yellow');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcNextMonth, fmtIdx);
                       *)
    // conditional format: expression
    inc(row);
    sh.WriteText(row, 0, 'expression: ISNUMBER($E$5)');
    sh.WriteText(row, 1, 'background blue');
    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scBlue);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, 2), cfcExpression, '=ISNUMBER($E$5)', fmtIdx);

    // conditional format: expression
    inc(row);
    sh.WriteText(row, 0, 'expression: ISNUMBER(E5)');
    sh.WriteText(row, 1, 'background blue');
    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scBlue);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, 5), cfcExpression, '=ISNUMBER(E5)', fmtIdx);

    // Two rules in the same conditional format
    inc(row);
    sh.WriteText(row, 0, 'Two rules: #1: equal to 5, #2: equal to 3');
    sh.WriteText(row, 1, '#1: background yellow, #2: background green');
    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcEqual, 5, fmtIdx);
    fmt.SetBackgroundColor(scGreen);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcEqual, 3, fmtIdx);  // use the same cell range

    // Rotated text
    inc(row);
    sh.WriteText(row, 0, 'Equal to "abc"');
    sh.WriteText(row, 1, 'Rotated text (90 CCW), hor center, vert top');
    InitFormatRecord(fmt);
    fmt.SetTextRotation(rt90DegreeCounterClockwiseRotation);
    fmt.SetHorAlignment(haCenter);
    fmt.SetVertAlignment(vaTop);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcEqual, 'abc', wb.AddCellFormat(fmt));

    // Databar
    inc(row);
    sh.WriteText(row, 0, 'Data bar');
    sh.WriteDatabars(Range(Row, 2, row, 12), scRed);

    // ColorRange
    inc(row);
    sh.WriteText(row, 0, 'Color Range');
    sh.WriteText(row, 1, 'yellow -> blue -> red');
    sh.WriteColorRange(Range(Row, 2, row, 12), scYellow, scBlue, scRed);

    // ColorRange
    inc(row);
    sh.WriteText(row, 0, 'Color Range');
    sh.WriteText(row, 1, 'yellow -> red');
    sh.WriteColorRange(Range(Row, 2, row, 12), scYellow, scRed);

    // Icon sets
    inc(row);
    sh.WriteText(row, 0, 'IconSet');
    sh.WriteText(row, 1, '3 flags');
    sh.WriteIconSet(Range(Row, 2, row, 12), is3Flags);

    inc(row);
    sh.WriteText(row, 0, 'IconSet');
    sh.WriteText(row, 1, '5 quarters');
    sh.WriteIconSet(Range(Row, 2, row, 12), is5Quarters);

    { ------ Save workbook to file-------------------------------------------- }
    dir := ExtractFilePath(ParamStr(0));
    wb.WriteToFile(dir + 'test.xlsx', true);
    wb.WriteToFile(dir + 'test.ods', true);
    wb.WriteToFile(dir + 'test.xml', true);

    if wb.ErrorMsg <> '' then
      WriteLn(wb.ErrorMsg);

  finally
    wb.Free;
  end;

  if ParamCount = 0 then
  begin
    {$IFDEF MSWINDOWS}
    WriteLn('Press [ENTER] to quit...');
    ReadLn;
    {$ENDIF}
  end;
end.

