{
excel8read.lpr

Demonstrates how to read an Excel 8.x file using the fpspreadsheet library

AUTHORS: Felipe Monteiro de Carvalho
}
program excel8read;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, LazUTF8, fpsTypes, fpspreadsheet, xlsbiff8,
  fpsutils;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  InputFilename: string;
  MyDir: string;
  i: Integer;
  CurCell: PCell;

{$R *.res}

begin
  // Open the input file
  MyDir := ExtractFilePath(ParamStr(0));
  InputFileName := MyDir + 'test.xls';

  if not FileExists(InputFileName) then begin
    WriteLn('Input file ', InputFileName, ' does not exist. Please run excel8write first.');
    Halt;
  end;
  WriteLn('Opening input file ', InputFilename);

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];
    MyWorkbook.ReadFromFile(InputFilename, sfExcel8);
    MyWorksheet := MyWorkbook.GetFirstWorksheet;

    // Write all cells with contents to the console
    WriteLn('');
    WriteLn('Contents of the first worksheet of the file:');
    WriteLn('');

    for CurCell in MyWorksheet.Cells do
    begin
      Write('Row: ', CurCell^.Row,
       ' Col: ', CurCell^.Col, ' Value: ',
       UTF8ToConsole(MyWorkSheet.ReadAsText(CurCell^.Row,
         CurCell^.Col))
       );
      if HasFormula(CurCell) then
        WriteLn(' Formula: ', MyWorkSheet.ReadFormulaAsString(CurCell))
      else
        WriteLn;
    end;

  finally
    // Finalization
    MyWorkbook.Free;
  end;

  if ParamCount = 0 then
  begin
    {$IFDEF WINDOWS}
    WriteLn;
    WriteLn('Press ENTER to quit...');
    ReadLn;
    {$ENDIF}
  end;
end.

