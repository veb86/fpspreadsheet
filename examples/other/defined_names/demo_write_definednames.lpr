program demo_write_definednames;
uses
  fpspreadsheet, fpsTypes, fpsUtils, fpsClasses, fpsAllFormats;
var
  wb: TsWorkbook;
  ws: TsWorksheet;
  wsIdx0, wsIdx1: Integer;
  i: Integer;
begin
  wb := TsWorkbook.Create;
  try
    wb.Options := [boAutoCalc];

    {----------}

    // Single cell defined names
    ws := wb.AddWorksheet('Simple');
    wsIdx0 := wb.GetWorksheetIndex(ws);

    wb.DefinedNames.Add('distance', wsIdx0, 1, 2);
    ws.WriteText(1, 1, 'distance');     ws.WriteNumber(1, 2, 120);     ws.WriteFormula(1, 3, '=distance');

    ws.WriteText(2, 1, 'time');         ws.WriteNumber(2, 2, 60);
    wb.DefinedNames.Add('time', wsIdx0, 2, 2);

    wb.DefinedNames.Add('speed', wsIdx0, 3, 2);
    ws.WriteText(3, 1, 'speed');        ws.WriteFormula(3, 2, '=distance/time');

    {----------}

    // Cell range as defined name
    ws := wb.AddWorksheet('Range');
    wsIdx1 := wb.GetWorksheetIndex(ws);

    ws.WriteText(0, 0, 'Data');
    ws.WriteNumber(1, 0, 1.0);
    ws.WriteNumber(2, 0, 2.0);
    ws.WriteNumber(3, 0, 3.0);
    wb.DefinedNames.Add('data', wsIdx1, wsIdx1, 1, 0, 3, 0);
    ws.WriteFormula(4, 0, '=SUM(data)');

    {----------}

    // Defined name in other sheet
    ws := wb.AddWorksheet('Range from other sheet');

    wb.DefinedNames.Add('data', wsIdx1, wsIdx1, 1, 0, 3, 0);  // wsIdx refers to sheet "Range"
    ws.WriteFormula(4, 0, '=SUM(data)');

    {----------}

    ws := wb.AddWorksheet('3D range');
    wb.DefinedNames.Add('data_1_2', wsIdx0, wsIdx1, 0, 0, 5, 5);
    ws.WriteText(0, 0, 'Count of cells in 1st and 2nd sheet');
    ws.WriteFormula(1, 0, '=COUNTA(data_1_2)');

    ws.WriteText(3, 0, 'Count of numeric cells in 1st and 2nd sheet');
    ws.WriteFormula(4, 0, '=COUNT(data_1_2)');

    ws.WriteText(6, 0, 'Sum of numeric cells in 1st and 2nd sheet');
    ws.WriteFormula(7, 0, '=SUM(data_1_2)');

    {----------}

    wb.WriteToFile('test_defnames.xlsx', true);
    wb.WriteToFile('test_defnames.ods', true);
  finally
    wb.Free;
  end;
end.

