{
excel2read.lpr

Demonstrates how to read an Excel 2.x file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel2read;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, LazUTF8, fpsTypes, fpsUtils, fpspreadsheet, xlsbiff2;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  InputFilename: string;
  MyDir: string;
  i: Integer;
  CurCell: PCell;

begin
  // Open the input file
  MyDir := ExtractFilePath(ParamStr(0));
  InputFileName := MyDir + 'test' + STR_EXCEL_EXTENSION;
  if not FileExists(InputFileName) then begin
    WriteLn('Input file ', InputFileName, ' does not exist. Please run excel2write first.');
    Halt;
  end;

  WriteLn('Opening input file ', InputFilename);

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas, boAutoCalc];
    MyWorkbook.ReadFromFile(InputFilename, sfExcel2);

    MyWorksheet := MyWorkbook.GetFirstWorksheet;

    // Write all cells with contents to the console
    WriteLn('');
    WriteLn('Contents of the first worksheet of the file:');
    WriteLn('');

    for CurCell in MyWorksheet.Cells do
    begin
      Write('Row: ', CurCell^.Row, ' Col: ', CurCell^.Col, ' Value: ',
        UTF8ToConsole(MyWorkSheet.ReadAsText(CurCell^.Row, CurCell^.Col))
      );
      if HasFormula(CurCell) then
        Write(' (Formula ', MyWorksheet.ReadFormula(CurCell), ')');
      WriteLn;
    end;

  finally
    // Finalization
    MyWorkbook.Free;
  end;

  if ParamCount = 0 then
  begin
    {$ifdef Windows}
    WriteLn;
    WriteLn('Press ENTER to quit...');
    ReadLn;
    {$endif}
  end;
end.

