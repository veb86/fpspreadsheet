{
ooxmlwrite.lpr

Demonstrates how to write an OOXML file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program ooxmlwrite;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpstypes, fpspreadsheet, fpsallformats, fpscell;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
  i: Integer;
  MyCell: PCell;

begin
  // Open the output file
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorksheet := MyWorkbook.AddWorksheet('My Worksheet');

  MyWorksheet.WriteNumber(0, 0, 1.0);

  MyWorksheet.WriteNumberFormat(0, 0, nfFixed, 2);

  // Save the spreadsheet to a file
  MyWorkbook.WriteToFile(MyDir + 'test.xlsx', sfOOXML, true);
  MyWorkbook.Free;

  WriteLn('Workbook written to "' + Mydir + 'test.xlsx' + '".');

  {$IFDEF MSWINDOWS}
  WriteLn('Press ENTER to quit...');
  ReadLn;
  {$ENDIF}
end.

