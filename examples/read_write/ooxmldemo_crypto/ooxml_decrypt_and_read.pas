{
ooxml_decrypt_and_read.lpr

Demonstrates how to read an Excel 2007 xlsx file which is workbook-protected
and thus encrypted by an internal password.

Basic operating procedure
- Add package laz_fpspreadsheet_crypto
- Use xlsxooxml_crypto (instead of xlsxooxml)
- In Workbook.ReadFromFile specify the file format id spfidOOXML_crypto instead
  of the the file format sfOOXML.
}

program ooxml_decrypt_and_read;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, LazUTF8, fpstypes, fpspreadsheet, //laz_fpspreadsheet,
  xlsxooxml_crypto;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  InputFilename: String;
  MyDir: string;
  cell: PCell;
  i: Integer;
  password: String;
  Prot_enc: Integer = 0;  // 0 - protected, 1 - encrypted workbook

begin
  MyDir := ExtractFilePath(ParamStr(0));

  // Open the input file
  MyDir := ExtractFilePath(ParamStr(0));

  case Prot_enc of
    0: begin
         InputFileName := MyDir + 'protected_workbook.xlsx';
         password := '';
       end;
    1: begin
         InputFileName := MyDir + 'encrypted_workbook.xlsx';
         password := 'test';
       end;
  end;

  if not FileExists(InputFileName) then begin
    WriteLn('Input file ', InputFileName, ' does not exist. Please run opendocwrite first.');
    Halt;
  end;
  WriteLn('Opening input file ', InputFilename);

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;

  MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];
  MyWorkbook.ReadFromFile(InputFilename, sfidOOXML_crypto, password);

  MyWorksheet := MyWorkbook.GetFirstWorksheet;

  // Write all cells with contents to the console
  WriteLn('');
  WriteLn('Contents of the first worksheet of the file:');
  WriteLn('');

  for cell in MyWorksheet.Cells do
    WriteLn(
      'Row: ', cell^.Row,
      ' Col: ', cell^.Col,
      ' Value: ', UTF8ToConsole(MyWorkSheet.ReadAsText(cell^.Row, cell^.Col))
    );

  // Finalization
  MyWorkbook.Free;

  {$ifdef WINDOWS}
  WriteLn;
  WriteLn('Press ENTER to quit...');
  ReadLn;
  {$ENDIF}
end.

