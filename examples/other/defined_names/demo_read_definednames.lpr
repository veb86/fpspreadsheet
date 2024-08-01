program demo_read_definednames;
uses
  SysUtils, fpspreadsheet, fpsTypes, fpsClasses, fpsUtils, fpsAllFormats;
var
  wb: TsWorkbook;
  ws: TsWorksheet;
  cell: PCell;
  i, j: Integer;
  fn: String;
  fmt: TsSpreadsheetFormat = sfOpendocument;
//  fmt: TsSpreadsheetFormat = xlsxOOXML;
begin
  fn := 'test_defnames';
  fn := 'Mappe_illegalRef';
  fn := 'Mappe3';
  case fmt of
    sfOpenDocument: fn := fn + '.ods';
    sfOOXML: fn := fn + '.xlsx';
    else raise Exception.Create('Format not supported:');
  end;

  wb := TsWorkbook.Create;
  try
    wb.Options := [boAutoCalc, boReadFormulas];

    WriteLn('FILE: ', fn, LineEnding);
    wb.ReadFromFile(fn);

    WriteLn('DEFINED NAMES (GLOBAL)');
    for i := 0 to wb.DefinedNames.Count-1 do
    begin
      Write('  "', wb.DefinedNames[i].Name, '" --> ');
      case fmt of
        sfOOXML: WriteLn(wb.DefinedNames[i].RangeAsString(wb));
        sfOpenDocument:  WriteLn(wb.DefinedNames[i].RangeAsString_ODS(wb));
      end;
    end;

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
        begin
          Write('  "', ws.DefinedNames[i].Name, '" --> ');
          case ExtractFileExt(fn) of
            '.xlsx': WriteLn(ws.DefinedNames[i].RangeAsString(wb));
            '.ods':  WriteLn(ws.DefinedNames[i].RangeAsString_ODS(wb));
          end;
        end;

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

