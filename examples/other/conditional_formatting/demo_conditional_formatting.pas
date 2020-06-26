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

begin
  wb := TsWorkbook.Create;
  try
    sh := wb.AddWorksheet('test');

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

    { ------ Save workbook to file-------------------------------------------- }
    wb.WriteToFile('test.xlsx', true);
  finally
    wb.Free;
  end;

end.

