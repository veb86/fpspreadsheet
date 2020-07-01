program demo_conditional_formatting;

uses
  sysUtils,
  fpsTypes, fpsUtils, fpspreadsheet, fpsConditionalFormat,
  xlsxooxml, fpsOpenDocument;

var
  wb: TsWorkbook;
  sh: TsWorksheet;
  fmt: TsCellFormat;
  fmtIdx: Integer;
  font: TsFont;
  row: Integer;
  i: Integer;
  lastCol: Integer;
begin
  wb := TsWorkbook.Create;
  try
    sh := wb.AddWorksheet('test');
    sh.WriteDefaultColWidth(20, suMillimeters);

    sh.WriteText(0, 0, 'Condition');
    sh.WriteColWidth(0, 50, suMillimeters);
    sh.WriteText(0, 1, 'Format');
    sh.WriteColWidth(1, 70, suMillimeters);
    sh.WriteText(0, 2, 'Test values');

    row := 2;
    for i := row to row+30 do
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
    end;
    lastCol := 18;

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
    sh.WriteText(row, 1, 'background green');
    fmt.SetBackgroundColor(scGreen);
    fmtIdx := wb.AddCellFormat(fmt);
    // Write conditional format
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcEqual, 'abc', fmtIdx);

    // conditional format #3: greater than cell reference
    inc(row);
    sh.WriteText(row, 0, 'greater than cell C3');
    sh.WriteText(row, 1, 'all borders, red, thick line');
    InitFormatRecord(fmt);
    fmt.SetBorders([cbEast, cbWest, cbNorth, cbSouth], scRed, lsThick);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcGreaterThan, 'C3', fmtIdx);

    // conditional format #4: less than formula
    inc(row);
    sh.WriteText(row, 0, 'less than formula "=1+3"');
    sh.WriteText(row, 1, 'background red');
    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcLessThan, '=1+3', fmtIdx);

    // conditional format #5: greater equal constant
    inc(row);
    sh.WriteText(row, 0, 'greater equal constant 5');
    sh.WriteText(row, 1, 'background gray');
    fmt.SetBackgroundColor(scGray);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcGreaterEqual, 5, fmtIdx);

    // conditional format #6: less equal constant
    inc(row);
    sh.WriteText(row, 0, 'less equal constant 5');
    sh.WriteText(row, 1, 'background gray');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcLessEqual, 5, fmtIdx);

    // conditional format #6: between
    inc(row);
    sh.WriteText(row, 0, 'between 3 and 7');
    sh.WriteText(row, 1, 'background light gray');
    fmt.SetBackgroundColor($EEEEEE);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcBetween, 2, 7, fmtIdx);

    // conditional format #6: not between
    inc(row);
    sh.WriteText(row, 0, 'not between 3 and 7');
    sh.WriteText(row, 1, 'background light gray');
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcNotBetween, 2, 7, fmtIdx);

    // conditional format #6: above average
    inc(row);
    sh.WriteText(row, 0, '> average');
    sh.WriteText(row, 1, 'hatched background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsThinStripeDiagUp, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, 10), cfcAboveAverage, fmtIdx);   // only 1..9 -> ave = 5

    // conditional format #6: below average
    inc(row);
    sh.WriteText(row, 0, '< average');
    sh.WriteText(row, 1, 'dotted background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsGray25, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, 10), cfcBelowAverage, fmtIdx);   // only 1..9 -> ave = 5

    // conditional format #6: above or equal to average
    inc(row);
    sh.WriteText(row, 0, '>= average');
    sh.WriteText(row, 1, 'hor striped background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsThinStripeHor, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, 10), cfcAboveEqualAverage, fmtIdx);  // only 1..9 -> ave = 5

    // conditional format #6: below or equal to average
    inc(row);
    sh.WriteText(row, 0, '<= average');
    sh.WriteText(row, 1, 'vert striped background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsThinStripeVert, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, 10), cfcBelowEqualAverage, fmtIdx);  // only 1..9 -> ave = 5

    // conditional format #6: top 3 values
    inc(row);
    sh.WriteText(row, 0, 'top 3 values');
    sh.WriteText(row, 1, 'background green');
    fmt.SetBackgroundColor(scGreen);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcTop, 3, fmtIdx);

    // conditional format #6: smallest 3 values
    inc(row);
    sh.WriteText(row, 0, 'smallest 3 values');
    sh.WriteText(row, 1, 'background bright blue');
    fmt.SetBackgroundColor($FFC0C0);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcBottom, 3, fmtIdx);

    // conditional format #6: top 30 percent
    inc(row);
    sh.WriteText(row, 0, 'top 10 percent');
    sh.WriteText(row, 1, 'background green');
    fmt.SetBackgroundColor(scGreen);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcTopPercent, 10, fmtIdx);

    // conditional format #6: smallest 3 values
    inc(row);
    sh.WriteText(row, 0, 'smallest 10 percent');
    sh.WriteText(row, 1, 'background bright blue');
    fmt.SetBackgroundColor($FFC0C0);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcBottomPercent, 10, fmtIdx);

    // conditional format #6: duplicates
    inc(row);
    sh.WriteText(row, 0, 'duplicate values');
    sh.WriteText(row, 1, 'background bright red');
    fmt.SetBackgroundColor($D0D0FF);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcDuplicate, fmtIdx);

    // conditional format #6: unique
    inc(row);
    sh.WriteText(row, 0, 'unique values');
    sh.WriteText(row, 1, 'background bright red');
    fmt.SetBackgroundColor($D0D0FF);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcUnique, fmtIdx);

    // conditional format #6: contains any text
    inc(row);
    sh.WriteText(row, 0, 'contains any text');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcContainsText, '', fmtIdx);

    // conditional format #6: empty
    inc(row);
    sh.WriteText(row, 0, 'empty');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcNotContainsText, '', fmtIdx);

    // conditional format #6: text begins with 'ab'
    inc(row);
    sh.WriteText(row, 0, 'text begins with "ab"');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcBeginsWith, 'ab', fmtIdx);

    // conditional format #6: text ends with 'g'
    inc(row);
    sh.WriteText(row, 0, 'text ends with "g"');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcEndsWith, 'g', fmtIdx);

    // conditional format #6: text contains 'ef'
    inc(row);
    sh.WriteText(row, 0, 'text contains "ef"');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcContainsText, 'ef', fmtIdx);

    // conditional format #6: text does NOT contain 'ef'
    inc(row);
    sh.WriteText(row, 0, 'text does not contain "ef"');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcNotContainsText, 'ef', fmtIdx);

    // conditional format #6: contains error
    inc(row);
    sh.WriteText(row, 0, 'contains error');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcContainsErrors, fmtIdx);

    // conditional format #6: no errors
    inc(row);
    sh.WriteText(row, 0, 'no errors');
    sh.WriteText(row, 1, 'background red');
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcNotContainsErrors, fmtIdx);

    WriteLn('row = ', row);
    WriteLn('wb.GetNumcellFormats = ', wb.GetNumCellFormats);
    WriteLn('wb.GetNumConditionalFormats = ',  wb.GetNumConditionalFormats);

    { ------ Save workbook to file-------------------------------------------- }
    wb.WriteToFile('test.xlsx', true);
    wb.WriteToFile('test.ods', true);
  finally
    wb.Free;
  end;

  ReadLn;
end.

