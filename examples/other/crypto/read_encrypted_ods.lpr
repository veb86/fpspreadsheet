program read_encrypted_ods;

uses
  SysUtils,
  fpSpreadsheet, fpsTypes, fpsUtils, fpsOpenDocument_Crypto;

const
  FILENAME = 'pwd 123.ods';
  PASSWORD = '123';

var
  wb: TsWorkbook;
  ws: TsWorksheet;
  cell: PCell;
  fn: String;
  fmtID: TsSpreadFormatID;
  t: TDateTime;
begin
  t := Now;
  wb := TsWorkbook.Create;
  try
    wb.ReadFromFile(FILENAME, sfidOpenDocument_Crypto, PASSWORD, []);
    ws := wb.GetFirstWorksheet;
    if ws <> nil then
      for cell in ws.Cells do
        WriteLn('cell ', GetCellString(cell^.Row, cell^.Col), ' = "', ws.ReadAsText(cell), '"');
  finally
    wb.Free;
  end;
  t := Now - t;
  WriteLn('Time to decrypt and load: ', FormatDateTime('nn:ss.zzz', t), ' seconds');

  Write('Press ENTER to close...');
  ReadLn;
end.

