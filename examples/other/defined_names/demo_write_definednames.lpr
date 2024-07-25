program demo_write_definednames;
uses
  fpspreadsheet, fpsTypes, fpsUtils, fpsClasses, fpsAllFormats;
var
  wb: TsWorkbook;
  ws: TsWorksheet;
  wsIdx: Integer;
  i: Integer;
begin
  wb := TsWorkbook.Create;
  try
    wb.Options := [boAutoCalc];
    ws := wb.AddWorksheet('Test');
    wsIdx := wb.GetWorksheetIndex(ws);

    wb.DefinedNames.Add('distance', wsIdx, 1, 2);
    ws.WriteText(1, 1, 'distance');     ws.WriteNumber(1, 2, 120);     ws.WriteFormula(1, 3, '=distance');

    ws.WriteText(2, 1, 'time');         ws.WriteNumber(2, 2, 60);
    wb.DefinedNames.Add('time', wsIdx, 2, 2);

    wb.DefinedNames.Add('speed', wsIdx, 3, 2);
    ws.WriteText(3, 1, 'speed');        ws.WriteFormula(3, 2, '=distance/time');

    wb.WriteToFile('test_defnames.xlsx', true);
    wb.WriteToFile('test_defnames.ods', true);
  finally
    wb.Free;
  end;
end.

