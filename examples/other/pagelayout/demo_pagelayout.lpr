program demo_pagelayout;

uses
  SysUtils, FPSpreadSheet, FPSTypes, FPSAllFormats;
var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
begin
  WriteLn('Creating a worksheet with pagelayout for printing and print preview...');
  workbook := TsWorkbook.Create;
  try
    worksheet := workbook.AddWorksheet('test');
    worksheet.WriteText(0, 0, 'left/top');
    worksheet.WriteText(1, 1, 'center');
    worksheet.WriteText(2, 2, 'right/bottom');
    worksheet.PageLayout.Orientation := spoLandscape;
    worksheet.PageLayout.Options := worksheet.Pagelayout.Options + [poHorCentered, poVertCentered, poPrintGridLines, poPrintHeaders];
    workbook.WriteToFile('test5.xls', sfExcel5, true);
    workbook.WriteToFile('test8.xls', sfExcel8, true);
    workbook.WriteToFile('test.xlsx', true);
    workbook.WriteToFile('test.ods', true);
  finally
    workbook.Free;
  end;

  if ParamCount = 0 then
  begin
    {$IFDEF MSWINDOWS}
    WriteLn('Press [ENTER] to quit...');
    ReadLn;
    {$ENDIF}
  end;
end.

