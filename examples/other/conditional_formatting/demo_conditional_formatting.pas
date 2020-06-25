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
    // ... set the font (bold)   ---- NOT SUPPORTED AT THE MOMENT...
    font := wb.CloneFont(0);
    font.Style := [fssBold, fssItalic];
    font.Color := scRed;
    fmt.SetFont(wb.AddFont(font));
    // Add format record to format list
    fmtIdx := wb.AddCellFormat(fmt);

    // Use the format as conditional format of A1:A6 when cells are equal to 3.
    sh.WriteConditionalCellFormat(Range(0, 0, 5, 0), cfcEqual, 3.0, fmtIdx);

    wb.WriteToFile('test.xlsx', true);
  finally
    wb.Free;
  end;

end.

