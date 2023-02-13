program demo_frozen_rows_cols;
uses
  SysUtils,
  FPSpreadsheet, fpsTypes, xlsxOOXML;
var
  wb: TsWorkbook;
  ws: TsWorksheet;
  r, c: Integer;
begin
  wb := TsWorkbook.Create;
  try
    ws := wb.AddWorksheet('Sheet1');

    // Fill worksheet with some data
    for r := 0 to 100 do
      for c := 0 to 10 do
        ws.WriteText(r, c, Format('R%d C%d', [r, c]));

    // Prepare frozen columns and frozen rows
    ws.LeftPaneWidth := 1;  // There should be 1 frozen column
    ws.TopPaneHeight := 2;  // There should be 2 frozen rows
    ws.Options := ws.Options + [soHasFrozenPanes];  // Activate this feature.

    // Save to file
    wb.WriteToFile('test.xlsx', true);
  finally
    wb.Free;
  end;
end.

