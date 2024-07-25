program demo_read_definednames;
{.$DEFINE ODS}
uses
  fpspreadsheet, fpsTypes, fpsClasses, fpsUtils, fpsAllFormats;
var
  wb: TsWorkbook;
  ws: TsWorksheet;
  cell: PCell;
  i: Integer;
  fn: String;
begin
  {$IFDEF ODS}
  fn := 'test_defnames.ods';
  {$ELSE}
  fn := 'test_defnames.xlsx';
  {$ENDIF}

  wb := TsWorkbook.Create;
  try
    wb.Options := [boAutoCalc, boReadFormulas];

    wb.ReadFromFile(fn);
    ws := wb.GetFirstWorksheet;

    WriteLn('FILE: ', fn, LineEnding);

    WriteLn('DEFINED NAMES');
    for i := 0 to wb.DefinedNames.Count-1 do
      WriteLn(wb.DefinedNames[i].Name, ' --> ', wb.DefinedNames[i].RangeAsString(wb));

    WriteLn;
    WriteLn('CELLS');
    for cell in ws.Cells do
    begin
      Write(GetCellString(cell^.Row, cell^.Col), ' --> ', ws.ReadAsText(cell));
      if HasFormula(cell) then
        Write(' (formula: ', ws.GetFormula(cell)^.Text, ')');
      WriteLn;
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

