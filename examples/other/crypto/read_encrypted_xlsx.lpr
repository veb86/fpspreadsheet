program read_encrypted_xlsx;

uses
  SysUtils,
  fpSpreadsheet, fpsTypes, fpsUtils, xlsxOOXML_Crypto;

const
  FILENAME = 'pwd 123.xlsx';
  PASSWORD = '123';

var
  wb: TsWorkbook;
  ws: TsWorksheet;
  cell: PCell;
  t: TDateTime;
begin
  t := Now;
  wb := TsWorkbook.Create;
  try
    wb.ReadFromFile(FILENAME, sfidOOXML_Crypto, PASSWORD, []);
    ws := wb.GetFirstWorksheet;
    if ws <> nil then
      for cell in ws.Cells do
        WriteLn('cell ', GetCellString(cell^.Row, cell^.Col), ' = "', ws.ReadAsText(cell), '"');
  finally
    wb.Free;
  end;
  t := Now - t;
  WriteLn('Time to decrypt and load: ', FormatDateTime('nn:ss.zzz', t), ' seconds');
  WriteLn;
  Write('Press ENTER to close...');
  ReadLn;
end.

