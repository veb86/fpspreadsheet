program demo_read_definednames;
{$DEFINE ODS}
uses
  fpspreadsheet, fpsTypes, fpsClasses, fpsUtils, fpsAllFormats;
var
  wb: TsWorkbook;
  ws: TsWorksheet;
  cell: PCell;
  i, j: Integer;
  fn: String;
begin
  {$IFDEF ODS}
  fn := 'test_defnames.ods';
  {$ELSE}
  fn := 'test_defnames.xlsx';
  {$ENDIF}

  //fn := 'test3D.xlsx';

  wb := TsWorkbook.Create;
  try
    wb.Options := [boAutoCalc, boReadFormulas];

    WriteLn('FILE: ', fn, LineEnding);
    wb.ReadFromFile(fn);

    WriteLn('DEFINED NAMES (GLOBAL)');
    for i := 0 to wb.DefinedNames.Count-1 do
      WriteLn('  "', wb.DefinedNames[i].Name, '" --> ', wb.DefinedNames[i].RangeAsString(wb));

    WriteLn('--------------------------------------------------------');

    for i := 0 to wb.GetWorksheetCount - 1 do
    begin
      ws := wb.GetWorksheetByIndex(i);
      WriteLn('WORKSHEET "', ws.Name, '"');

      WriteLn('  DEFINED NAMES (LOCAL)');
      if ws.DefinedNames.Count = 0 then
        WriteLn('    (none)')
      else
        for j := 0 to ws.DefinedNames.Count-1 do
          WriteLn('    "', ws.DefinedNames[i].Name, '" --> ', ws.DefinedNames[i].RangeAsString(wb));

      WriteLn('  CELLS');
      for cell in ws.Cells do
      begin
        Write('    ', GetCellString(cell^.Row, cell^.Col), ' --> ', ws.ReadAsText(cell));
        if HasFormula(cell) then
          Write(' (formula: "=', ws.GetFormula(cell)^.Text, '")');
        WriteLn;
      end;
      WriteLn('--------------------------------------------------------');
    end;

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

