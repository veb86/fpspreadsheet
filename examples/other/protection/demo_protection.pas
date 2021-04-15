program demo_protection;

{$mode objfpc}{$H+}

uses
  Classes, SysUtils,
  fpstypes, fpspreadsheet, fpsallformats, fpsutils, fpscrypto;

const
  PASSWORD = 'lazarus';

var
  book: TsWorkbook;
  sheet: TsWorksheet;
  cell: PCell;
  c: TsCryptoInfo;
  dir: String;

begin
  book := TsWorkbook.Create;
  try
    sheet := book.AddWorksheet('Sheet1');

    // Add an unprotected cell
    cell := sheet.WriteText(0, 0, 'Unprotected cell');
    sheet.WriteCellProtection(cell, []);

    // Add a protected cell
    sheet.WriteText(1, 0, 'Protected cell');

    // Activate worksheet protection such that a password is required to
    // change the protection state
    InitCryptoInfo(c);
    c.Algorithm := caExcel;
    c.PasswordHash := Format('%.4x', [ExcelPasswordHash(PASSWORD)]);
    sheet.CryptoInfo := c;
    sheet.Protection := [spDeleteRows, spDeleteColumns, spInsertRows, spInsertColumns];
    sheet.Protect(true);

    dir := ExtractFilePath(ParamStr(0));
    book.WriteToFile(dir + 'protected.xls', sfExcel8, true);
    book.WriteToFile(dir + 'protected.xlsx', sfOOXML, true);
    // Note ODS does not write the excel password correctly, yet. --> protection cannot be removed.
    book.WriteToFile(dir + 'protected.ods', sfOpenDocument, true);

  finally
    book.Free;
  end;

  WriteLn('Open the files "protected.*" in your spreadsheet application.');
  WriteLn('Only cell A1 can be modifed.');

  if ParamCount = 0 then
  begin
    {$IFDEF MSWINDOWS}
    WriteLn('Press [ENTER] to quit...');
    ReadLn;
    {$ENDIF}
  end;

end.

