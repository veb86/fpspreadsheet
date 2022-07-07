program comments_demo_read;

uses
  SysUtils, fpspreadsheet, fpstypes, fpsutils, fpsallformats;
const
  FILE_NAME = 'test';
  
function RemoveLinebreaks(s: String): String;
var
  i: Integer;
begin
  SetLength(Result, Length(s));
  for i := 1 to Length(s) do
    if s[i] in [#10, #13] then
      Result[i] := ' '
    else
      Result[i] := s[i];
end;
  
var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  cmnt: PsComment;
  txt: string;
  i: Integer;
begin
  workbook := TsWorkbook.Create;
  try
    workbook.ReadFromFile(FILE_NAME + '.xlsx', sfOOXML);
    for i := 0 to workbook.GetWorksheetCount-1 do
    begin
      worksheet := workbook.GetWorksheetByIndex(i);
      WriteLn('Worksheet "', worksheet.Name, '":');
      for cmnt in worksheet.Comments do
      begin
        txt := RemoveLinebreaks(cmnt^.Text);
        WriteLn('  Comment in cell ', GetCellString(cmnt^.Row, cmnt^.Col), ': "', txt, '"');
      end;
      WriteLn;
    end;
  finally
    workbook.Free;
  end;
  
  WriteLn;
  WriteLn('Press ENTER to close...');
  ReadLn;
end.

