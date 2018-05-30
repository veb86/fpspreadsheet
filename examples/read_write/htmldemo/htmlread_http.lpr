{
htmlread_http.dpr

Demonstrates how to read a html file from the internet using the fpspreadsheet library.
}

program htmlread_http;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, LazUTF8, fphttpclient, fpstypes, fpsutils, fpspreadsheet, fpshtml;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  i: Integer;
  CurCell: PCell;
  stream: TMemoryStream;

const
//  url = 'http://unicode.e-workers.de/entities.php';
  url = 'https://www.freepascal.org/docs.var';

begin
  stream := TMemoryStream.Create;
  try
    // Get file from the internet
    with TFPHttpClient.Create(nil) do
      try
        Get(url, stream);
      finally
        Free;
      end;

    // Parameters
    HTMLParams.TableIndex := 0;

    // Create the spreadsheet
    MyWorkbook := TsWorkbook.Create;
    try
      MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas];
      MyWorkbook.ReadFromStream(stream, sfHTML);

      MyWorksheet := MyWorkbook.GetFirstWorksheet;

      // Write all cells with contents to the console
      WriteLn('');
      WriteLn('Contents of the first worksheet of the file:');
      WriteLn('');

      for CurCell in MyWorksheet.Cells do
      begin
        Write(
          'Row: ', CurCell^.Row,
          ' Col: ', CurCell^.Col,
          ' Value: ', UTF8ToConsole(MyWorkSheet.ReadAsText(CurCell^.Row, CurCell^.Col))
         );
        if MyWorksheet.HasHyperlink(CurCell) then
          Write(' Hyperlink: ', MyWorksheet.ReadHyperlink(CurCell).Target);
        WriteLn;
      end;

    finally
      // Finalization
      MyWorkbook.Free;
    end;
  finally
    stream.Free;
  end;

 {$IFDEF WINDOWS}
  WriteLn;
  WriteLn('Press ENTER to quit...');
  ReadLn;
 {$ENDIF}
end.

