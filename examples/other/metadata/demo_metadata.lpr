program demo_metadata;

uses
  {$IFDEF MSWINDOWS}
  windows,
  {$ENDIF}
  SysUtils,
  fpspreadsheet, fpstypes, xlsxooxml, fpsopendocument;

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
begin
  book := TsWorkbook.Create;
  try
    book.MetaData.CreatedBy := 'Donald Duck';
    book.MetaData.CreatedAt := EncodeDate(2020, 1, 1) + EncodeTime(12, 30, 40, 20);
    book.MetaData.Title := 'Test of metadata äöü';
    book.MetaData.Comments.Add('This is a test of spreadsheet metadata.');
    book.MetaData.Comments.Add('Assign the author to the field CreatedBy.');
    book.MetaData.Comments.Add('Assign the creation date to the field CreatedAt.');
    book.MetaData.Keywords.Add('Test');
    book.MetaData.Keywords.Add('FPSpreadsheet');

    sheet := book.AddWorksheet('Test');
    sheet.WriteText(2, 3, 'abc');
    sheet.WriteBackgroundColor(2, 3, scYellow);
    book.WriteToFile('test.xlsx', true);
    book.WritetoFile('test.ods', true);
  finally
    book.Free;
  end;

  book := TsWorkbook.Create;
  try
    book.ReadFromFile('test.xlsx');
    book.MetaData.ModifiedAt := Now();
    book.MetaData.ModifiedBy := GetUserName;
    WriteLn('CreatedBy  : ', book.MetaData.CreatedBy);
    WriteLn('CreatedAt  : ', DateTimeToStr(book.MetaData.CreatedAt));
    WriteLn('ModifiedBy : ', book.MetaData.ModifiedBy);
    WriteLn('ModifiedAt : ', DateTimeToStr(book.MetaData.ModifiedAt));
    WriteLn('Title      : ', book.MetaData.Title);
    WriteLn('Comments   : ');
    WriteLn(book.MetaData.Comments.Text);
    WriteLn('Keywords   : ', book.MetaData.Keywords.CommaText);
  finally
    book.Free;
  end;

  ReadLn;
end.

