{
htmlread.dpr

Demonstrates how to read a html file using the fpspreadsheet library.
IMPORTANT: Requires the output file of the htmlwrite demo.
}

program htmlread;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, LazUTF8, fpstypes, fpsutils, fpspreadsheet, fpshtml;

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
  InputFileName := MyDir + 'test' + STR_HTML_EXTENSION;
  if not FileExists(InputFileName) then begin
    WriteLn('Input file ', InputFileName, ' does not exist. Please run htmlwrite first.');
    Halt;
  end;

  WriteLn('Opening input file ', InputFilename);

  // Parameters
  HTMLParams.TableIndex := 0;

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];
    MyWorkbook.ReadFromFile(InputFilename, sfHTML);

    MyWorksheet := MyWorkbook.GetFirstWorksheet;

    // Write all cells with contents to the console
    WriteLn('');
    WriteLn('Contents of the first worksheet of the file:');
    WriteLn('');

    for CurCell in MyWorksheet.Cells do
    begin
      WriteLn(
        'Row: ', CurCell^.Row,
        ' Col: ', CurCell^.Col,
        ' Value: ', UTF8ToConsole(MyWorkSheet.ReadAsUTF8Text(CurCell^.Row, CurCell^.Col))
       );
    end;

  finally
    // Finalization
    MyWorkbook.Free;
  end;

 {$IFDEF WINDOWS}
  WriteLn;
  WriteLn('Press ENTER to quit...');
  ReadLn;
 {$ENDIF}
end.

