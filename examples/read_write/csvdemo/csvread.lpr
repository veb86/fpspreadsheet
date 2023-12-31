{
csvread.dpr

Demonstrates how to read a CSV file using the fpspreadsheet library
}

program csvread;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, LazUTF8, fpstypes, fpsutils, fpspreadsheet, fpscsv;

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
  InputFileName := MyDir + 'test' + STR_COMMA_SEPARATED_EXTENSION;
  if not FileExists(InputFileName) then begin
    WriteLn('Input file ', InputFileName, ' does not exist. Please run csvwrite first.');
    Halt;
  end;

  WriteLn('Opening input file ', InputFilename);

  // Tab-delimited
  CSVParams.Delimiter := #9;
  CSVParams.QuoteChar := '''';

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];
    MyWorkbook.ReadFromFile(InputFilename, sfCSV);

    MyWorksheet := MyWorkbook.GetFirstWorksheet;

    // Write all cells with contents to the console
    WriteLn('');
    WriteLn('Contents of the first worksheet of the file:');
    WriteLn('');

    for CurCell in MyWorksheet.Cells do
    begin
      if HasFormula(CurCell) then
        WriteLn('Row: ', CurCell^.Row, ' Col: ', CurCell^.Col, ' Formula: ', MyWorksheet.ReadFormulaAsString(CurCell))
      else
      WriteLn(
        'Row: ', CurCell^.Row,
        ' Col: ', CurCell^.Col,
        ' Value: ', UTF8ToConsole(MyWorkSheet.ReadAsText(CurCell^.Row, CurCell^.Col))
       );
    end;

  finally
    // Finalization
    MyWorkbook.Free;
  end;

  {$ifdef WINDOWS}
  if ParamCount = 0 then
  begin
    WriteLn;
    WriteLn('Press ENTER to quit...');
    ReadLn;
  end;
  {$ENDIF}
end.

