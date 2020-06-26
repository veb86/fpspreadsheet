program demo_conditional_formatting;

uses
  sysUtils,
  fpsTypes, fpsUtils, fpspreadsheet, xlsxooxml, fpsconditionalformat;

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
      sh.WriteNumber(i, 11, 10.0);
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
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcAboveAverage, fmtIdx);

    // conditional format #6: below average
    inc(row);
    sh.WriteText(row, 0, '< average');
    sh.WriteText(row, 1, 'dotted background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsGray25, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcBelowAverage, fmtIdx);

    // conditional format #6: above or equal to average
    inc(row);
    sh.WriteText(row, 0, '>= average');
    sh.WriteText(row, 1, 'hor striped background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsThinStripeHor, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcAboveEqualAverage, fmtIdx);

    // conditional format #6: below or equal to average
    inc(row);
    sh.WriteText(row, 0, '<= average');
    sh.WriteText(row, 1, 'vert striped background yellow on red');
    InitFormatRecord(fmt);
    fmt.SetBackground(fsThinStripeVert, scRed, scYellow);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(row, 2, row, lastCol), cfcBelowEqualAverage, fmtIdx);

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

    (*

    sh.Wri

'

    { ------ 1st conditional format : cfcEqual ------------------------------- }
    sh.WriteNumber(0, 0, 1.0);
    sh.WriteNumber(1, 0, 2.0);
    sh.WriteNumber(2, 0, 3.0);
    sh.WriteNumber(3, 0, 4.0);
    sh.WriteNumber(4, 0, 5.0);

    // Prepare the format record
    InitFormatRecord(fmt);
    // ... set the background color
    fmt.SetBackgroundColor(scYellow);
    // ... set the borders
    fmt.SetBorders([cbNorth, cbEast, cbWest], scBlack, lsThin);
    fmt.SetBorders([cbSouth], scRed, lsThick);
    // ... set the font (bold)   ---- NOT SUPPORTED AT THE MOMENT FOR WRITING TO XLSX...
    font := wb.CloneFont(0);
    font.Style := [fssBold, fssItalic];
    font.Color := scRed;
    fmt.SetFont(wb.AddFont(font));
    // Add format record to format list
    fmtIdx := wb.AddCellFormat(fmt);
    // Use the format as conditional format of A1:A6 when cells are equal to 3.
    sh.WriteConditionalCellFormat(Range(0, 0, 5, 0), cfcEqual, 3.0, fmtIdx);

    { ------- 2nd conditional format : cfcBelowEqualAverage ------------------ }
    sh.WriteNumber(0, 2, 10.0);
    sh.WriteNumber(1, 2, 20.0);
    sh.WriteNumber(2, 2, 15.0);
    sh.WriteNumber(3, 2, 11.0);
    sh.WriteNumber(4, 2, 19.0);

    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(0, 2, 4, 2), cfcBelowEqualAverage, fmtIdx);

    { ------- 3rd and 4th conditional formats : beginWith, containsText ------ }
    sh.WriteText(0, 4, 'abc');
    sh.WriteText(1, 4, 'def');
    sh.WriteText(2, 4, 'bac');
    sh.WriteText(3, 4, 'dbc');
    sh.WriteText(4, 4, 'acb');
    sh.WriteText(5, 4, 'aca');

    InitFormatRecord(fmt);
    fmt.SetBackgroundColor($DEF1F4);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(0, 4, 5, 4), cfcBeginsWith, 'a', fmtIdx);

    fmt.SetBackgroundColor($D08330);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(0, 4, 5, 4), cfcContainsText, 'bc', fmtIdx);

    { ------ 5th conditional format: containsErrors -------------------------- }
    sh.WriteFormula(0, 6, '=1.0/0.0');
    sh.WriteFormula(1, 6, '=1.0/1.0');
    sh.WriteFormula(2, 6, '=1.0/2.0');

    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scGreen);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(0, 6, 5, 6), cfcNotContainsErrors, fmtIdx);

    // Condition for ContainsErrors after NoContainsErrors to get higher priority
    fmt.SetBackgroundColor(scRed);
    fmtIdx := wb.AddCellFormat(fmt);
    sh.WriteConditionalCellFormat(Range(0, 0, 100, 100), cfcContainsErrors, fmtIdx);

    { ------ 6th conditional format: unique/duplicate values ----------------- }
    sh.WriteNumber(0, 1, 1.0);
    sh.WriteNumber(1, 1, 99.0);
    InitFormatRecord(fmt);
    fmt.SetBackgroundColor(scSilver);
    sh.WriteConditionalCellFormat(Range(0, 0, 1, 1), cfcUnique, wb.AddCellFormat(fmt));
    fmt.SetBackgroundColor(scGreen);
    sh.WriteConditionalCellFormat(Range(0, 0, 1, 1), cfcDuplicate, wb.AddCellFormat(fmt));
                                                                                   *)

    { ------ Save workbook to file-------------------------------------------- }
    wb.WriteToFile('test.xlsx', true);
  finally
    wb.Free;
  end;

end.

