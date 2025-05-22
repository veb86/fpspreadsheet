program zoom_demo;
uses
  fpspreadsheet, fpsTypes, xlsxOOXML, fpsOpenDocument, xlsbiff8;
var
  book: TsWorkbook;
  sheet: TsWorksheet;
begin
  book := TsWorkbook.Create;
  try
    sheet := book.AddWorksheet('Sheet');
    sheet.WriteNumber(0, 0, 1.0);
    sheet.WriteNumber(0, 1, 2.0);
    sheet.WriteNumber(1, 0, 5.12);
    sheet.ZoomFactor := 2.0;
    book.Options := book.Options + [boWriteZoomFactor];
    book.WriteToFile('zoom.xlsx', true);
    book.WriteToFile('zoom.ods', true);
    book.WriteToFile('zoom.xls', true);
  finally
    book.Free;
  end;
end.

