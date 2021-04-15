program demo_metadata;

uses
  {$IFDEF MSWINDOWS}
  windows,
  {$ENDIF}
  SysUtils,
  fpspreadsheet, fpstypes, xlsxooxml, fpsopendocument, xlsxml;

function GetUserName: String;
// http://forum.lazarus.freepascal.org/index.php/topic,23171.msg138057.html#msg138057
{$IFDEF WINDOWS}
const
  MaxLen = 256;
var
  Len: DWORD;
  WS: WideString = '';
  Res: windows.BOOL;
{$ENDIF}
begin
  Result := '';
  {$IFDEF UNIX}
  {$IF (DEFINED(LINUX)) OR (DEFINED(FREEBSD))}
  Result := SysToUtf8(users.GetUserName(fpgetuid));   //GetUsername in unit Users, fpgetuid in unit BaseUnix
  {$ELSE Linux/BSD}
  Result := GetEnvironmentVariableUtf8('USER');
  {$ENDIF UNIX}
  {$ELSE}
  {$IFDEF WINDOWS}
  Len := MaxLen;
  {$IFnDEF WINCE}
  if Win32MajorVersion <= 4 then begin
    SetLength(Result,MaxLen);
    Res := Windows.GetuserName(@Result[1], Len);
    //writeln('GetUserNameA = ',Res);
    if Res then begin
      SetLength(Result,Len-1);
//      Result := SysToUtf8(Result);
    end else
      SetLength(Result,0);
  end
  else
  {$ENDIF NOT WINCE}
  begin
    SetLength(WS, MaxLen-1);
    Res := Windows.GetUserNameW(@WS[1], Len);
    //writeln('GetUserNameW = ',Res);
    if Res then begin
      SetLength(WS, Len - 1);
      Result := ws;
    end else
      SetLength(Result,0);
  end;
  {$ENDIF WINDOWS}
  {$ENDIF UNIX}
end;

var
  book: TsWorkbook;
  sheet: TsWorksheet;
  i: Integer;
begin
  book := TsWorkbook.Create;
  try
    book.MetaData.CreatedBy := 'Donald Duck & Dagobert Duck';
    book.MetaData.Authors.Add('Donald Duck II');
    book.MetaData.DateCreated := EncodeDate(2020, 1, 1) + EncodeTime(12, 30, 40, 20);
    book.MetaData.DateLastModified := Now();
    book.MetaData.LastModifiedBy := 'Dagobert Duck';
    book.MetaData.Title := 'Test of metadata äöü';
    book.Metadata.Subject := 'FPSpreadsheet demos & tests';
    book.MetaData.Comments.Add('This is a test of spreadsheet metadata.');
    book.MetaData.Comments.Add('Assign the author to the field CreatedBy.');
    book.MetaData.Comments.Add('Assign the creation date to the field CreatedAt.');
    book.MetaData.Keywords.Add('Test1,Test2,Test3&4');
    book.MetaData.Keywords.Add('FPSpreadsheet');
    book.MetaData.AddCustom('Comparny', 'Disney');
    book.MetaData.AddCustom('Status', 'finished');

    sheet := book.AddWorksheet('Test');
    sheet.WriteText(2, 3, 'abc');
    sheet.WriteBackgroundColor(2, 3, scYellow);
    book.WriteToFile('test.xlsx', true);
    book.WriteToFile('test.ods', true);
    book.WriteToFile('test.xml', true)
  finally
    book.Free;
  end;

  book := TsWorkbook.Create;
  try
    // Select one of these
//    book.ReadFromFile('test.ods');
//    book.ReadFromFile('test.xlsx');
    book.ReadFromFile('test.xml');
    WriteLn('Created by         : ', book.MetaData.CreatedBy);
    WriteLn('Date created       : ', DateTimeToStr(book.MetaData.DateCreated));
    WriteLn('Modified by        : ', book.MetaData.LastModifiedBy);
    WriteLn('Date last modified : ', DateTimeToStr(book.MetaData.DateLastModified));
    WriteLn('Title              : ', book.MetaData.Title);
    WriteLn('Subject            : ', book.MetaData.Subject);
    WriteLn('Keywords           : ', book.MetaData.Keywords.CommaText);
    WriteLn('Custom             : ', 'Name':20, 'Value':20);
    for i := 0 to book.MetaData.Custom.Count-1 do
      WriteLn('                     ', book.MetaData.Custom.Names[i]:20, book.MetaData.Custom.ValueFromIndex[i]:20);
    WriteLn('Comments: ');
    WriteLn(book.MetaData.Comments.Text);
  finally
    book.Free;
  end;

  if ParamCount = 0 then
  begin
    {$IFDEF MSWINDOWS}
    WriteLn('Press [ENTER] to quit...');
    ReadLn;
    {$ENDIF}
  end;
end.

