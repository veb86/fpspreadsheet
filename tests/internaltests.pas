unit internaltests;

{ Other units test file read/write capability.
This unit tests functions, procedures and properties that fpspreadsheet provides.
}
{$mode objfpc}{$H+}

interface

{
Adding tests/test data:
- just add your new test procedure
}

uses
  // Not using lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpstypes, fpsallformats, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  fpsutils, fpsstreams, testsutility, md5;

type
  { TSpreadReadInternalTests }
  // Tests fpspreadsheet functionality, especially internal functions
  // Excel/LibreOffice/OpenOffice import/export compatibility should *NOT* be tested here

  { TSpreadInternalTests }

  TSpreadInternalTests= class(TTestCase)
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;

    procedure FractionTest(AMaxDigits: Integer);
    procedure WriteToStreamTest(AFormat: TsSpreadsheetFormat);

  published
    // Tests getting Excel style A1 cell locations from row/column based locations.
    // Bug 26447
    procedure TestCellString;
    //todo: add more calls, rename sheets, try to get sheets with invalid indexes etc
    //(see strings tests for how to deal with expected exceptions)
    procedure GetSheetByIndex;
    // Verify GetSheetByName returns the correct sheet number
    // GetSheetByName was implemented in SVN revision 2857
    procedure GetSheetByName;
    // Test for invalid sheet names
    procedure InvalidSheetName;
    // Tests whether overwriting existing file works
    procedure OverwriteExistingFile;
    // Write out date cell and try to read as UTF8; verify if contents the same
    procedure ReadDateAsUTF8;
    // Test buffered stream
    procedure TestReadBufStream;
    procedure TestWriteBufStream;
    // Test write to stream
    procedure TestWriteToStream_Biff8;
    procedure TestWriteToStream_Biff5;
    // Test fractions
//    procedure FractionTest_0;
    procedure FractionTest_1;
    procedure FractionTest_2;
    procedure FractionTest_3;
  end;

implementation

uses
  Math;

const
  InternalSheet = 'Internal'; //worksheet name

procedure TSpreadInternalTests.GetSheetByIndex;
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
begin
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:=MyWorkBook.AddWorksheet(InternalSheet);
    MyWorkSheet:=nil;
    MyWorkSheet:=MyWorkBook.GetWorksheetByIndex(0);
    CheckFalse((MyWorksheet=nil),'GetWorksheetByIndex should return a valid index');
  finally
    MyWorkbook.Free;
  end;
end;

procedure TSpreadInternalTests.GetSheetByName;
const
  AnotherSheet='AnotherSheet';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
begin
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:=MyWorkBook.AddWorksheet(InternalSheet);
    MyWorkSheet:=MyWorkBook.AddWorksheet(AnotherSheet);
    MyWorkSheet:=nil;
    MyWorkSheet:=MyWorkBook.GetWorksheetByName(InternalSheet);
    CheckFalse((MyWorksheet=nil),'GetWorksheetByName should return a valid index');
    CheckEquals(MyWorksheet.Name,InternalSheet,'GetWorksheetByName should return correct name.');
  finally
    MyWorkbook.Free;
  end;
end;

procedure TSpreadInternalTests.InvalidSheetName;
type
  TSheetNameCheck = record
    Valid: Boolean;
    SheetName: String;
  end;
const
  TestCases: array[0..9] of TSheetNameCheck = (
    (Valid: true;  SheetName:'Sheet'),
    (Valid: true;  SheetName:'äöü'),    // UFt8-characters are ok
    (Valid: false; SheetName:'Test'),   // duplicate
    (Valid: false; SheetName:'TEST'),   // duplicate since case is ignored
    (Valid: false; SheetName:''),       // empty string
    (Valid: false; SheetName:'Very very very very very very very very long'),  // too long
    (Valid: false; SheetName:'[sheet]'), // forbidden characters in following cases
    (Valid: false; SheetName:'/sheet/'),
    (Valid: false; SheetName:'\sheet\'),
    (Valid: false; SheetName:'***sheet***')
  );
var
  i: Integer;
  MyWorkbook: TsWorkbook;
  ok: Boolean;
begin
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.AddWorksheet('Test');
    for i:=0 to High(TestCases) do
    begin
      ok := MyWorkbook.ValidWorksheetName(TestCases[i].SheetName);
      CheckEquals(TestCases[i].Valid, ok, 'Sheet name validity check mismatch: ' + TestCases[i].SheetName);
    end;
  finally
    MyWorkbook.Free;
  end;
end;

procedure TSpreadInternalTests.OverwriteExistingFile;
const
  FirstFileCellText='Old version';
  SecondFileCellText='New version';
var
  FirstFileHash: string;
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  TempFile: string;
begin
  // Write out first file
  MyWorkbook:=TsWorkbook.Create;
  try
    MyWorkSheet:=MyWorkBook.AddWorksheet(InternalSheet);
    MyWorkSheet.WriteUTF8Text(0,0,FirstFileCellText);
    TempFile:=NewTempFile;
    MyWorkBook.WriteToFile(TempFile, sfExcel8, false);
  finally
    MyWorkbook.Free;
  end;

  if not(FileExists(TempFile)) then
    fail('Trying to write first file did not work.');
  FirstFileHash:=MD5Print(MD5File(TempFile));

  // Now overwrite with second file
  MyWorkbook:=TsWorkbook.Create;
  try
    MyWorkSheet:=MyWorkBook.AddWorksheet(InternalSheet);
    MyWorkSheet.WriteUTF8Text(0,0,SecondFileCellText);
    MyWorkBook.WriteToFile(TempFile,sfExcel8,true);
  finally
    MyWorkbook.Free;
  end;
  if FirstFileHash=MD5Print(MD5File(TempFile)) then
    fail('File contents are still those of the first file.');
end;

procedure TSpreadInternalTests.ReadDateAsUTF8;
var
  ActualDT: TDateTime;
  ActualDTString: string; //Result from ReadAsUTF8Text
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  Row,Column: Cardinal;
  TestDT: TDateTime;
begin
  Row:=0;
  Column:=0;
  TestDT:=EncodeDate(1969,7,21)+EncodeTime(2,56,0,0);

  MyWorkbook:=TsWorkbook.Create;
  try
    MyWorkSheet:=MyWorkBook.AddWorksheet(InternalSheet);
    MyWorkSheet.WriteDateTime(Row,Column,TestDT); //write date

    // Reading as date/time should just work
    if not(MyWorksheet.ReadAsDateTime(Row,Column,ActualDT)) then
      Fail('Could not read date time for cell '+CellNotation(MyWorkSheet,Row,Column));
    CheckEquals(TestDT,ActualDT,'Test date/time value mismatch '
      +'cell '+CellNotation(MyWorkSheet,Row,Column));

    //Check reading as string, convert to date & compare
    ActualDTString:=MyWorkSheet.ReadAsUTF8Text(Row,Column);
    ActualDT:=StrToDateTimeDef(ActualDTString,EncodeDate(1906,1,1));
    CheckEquals(TestDT,ActualDT,'Date/time mismatch using ReadAsUTF8Text');

  finally
    MyWorkbook.Free;
  end;
end;

procedure TSpreadInternalTests.TestReadBufStream;
const
  BUF_SIZE = 1024;
  FILE_SIZE = 2000;
var
  tempFileName: String;
  stream: TStream;
  writedata: array of Byte;
  readdata: array of Byte;
  i, n, nread: Integer;
begin
  RandSeed := 0;

  // Create a test file
  tempFileName := GetTempFileName;
  stream := TFileStream.Create(tempFileName, fmCreate);
  try
    SetLength(writedata, FILE_SIZE);
    for i:=0 to High(writedata) do
      writedata[i] := random(256);
    stream.WriteBuffer(writedata[0], Length(writedata));
  finally
    stream.Free;
  end;

  // Use a TBufStream to read parts of the file back
  stream := TBufStream.Create(tempFilename, fmOpenRead, BUF_SIZE);
  try
    // Check stream size
    CheckEquals(FILE_SIZE, stream.Size, 'Size mismatch');

    // Read first 100 bytes and compare with data
    nread := 100;
    SetLength(readdata, nread);
    n := stream.Read(readdata[0], nread);
    CheckEquals(nread, n, 'Bytes count mismatch');
    for i:=0 to nread-1 do
      CheckEquals(writedata[i], readdata[i], Format('Read mismatch at position %d', [i]));

    // Check stream size
    CheckEquals(FILE_SIZE, stream.Size, 'Size mismatch');

    // Read next 100 bytes and compare
    stream.ReadBuffer(readdata[0], nread);
    for i:=0 to nread-1 do
      CheckEquals(writedata[i+nread], readdata[i], Format('Read mismatch at position %d', [i+nread]));

    // Go to position 1000, this is 24 bytes to the end of the buffer, and read
    // 100 bytes again - this process will require to refresh the buffer
    stream.Position := 1000;
    stream.ReadBuffer(readdata[0], nread);
    for i:=0 to nread-1 do
      CheckEquals(writedata[i+1000], readdata[i], Format('Read mismatch at position %d', [i+1000]));

    // Check stream size
    CheckEquals(FILE_SIZE, stream.Size, 'Size mismatch');

    // Read next 100 bytes
    stream.ReadBuffer(readdata[0], nread);
    for i:=0 to nread-1 do
      CheckEquals(writedata[i+1000+nread], readdata[i], Format('Read mismatch at position %d', [i+1000+nread]));

    // Go back to start and fill the memory stream again with bytes 0..1023
    stream.Position := 0;
    stream.ReadBuffer(readdata[0], nread);

    // Now read 100 bytes which are not in the buffer
    stream.Position := 1500;  // this is past the buffered range
    stream.ReadBuffer(readdata[0], 100);
    for i:=0 to nread-1 do
      CheckEquals(writedata[i+1500], readdata[i], Format('Read mismatch at position %d', [i+1500]));

    // Go back to start and fill the memory stream again with bytes 0..1023
    stream.Position := 0;
    stream.ReadBuffer(readdata[0], 100);

    // Read last 100 bytes
    stream.Seek(nread, soFromEnd);
    stream.ReadBuffer(readdata[0], nread);
    for i:=0 to nread-1 do
      CheckEquals(writedata[i+FILE_SIZE-nread], readdata[i],
        Format('Read mismatch at position %d', [i+FILE_SIZE-nread]));

  finally
    stream.Free;
    DeleteFile(tempFileName);
  end;
end;

procedure TSpreadInternalTests.TestWriteBufStream;
const
  BUFSIZE = 1024;
var
  stream: TBufStream;
  readBuf, writeBuf1, writeBuf2: array of byte;
  nRead, nWrite1, nWrite2: Integer;
  i: Integer;
begin
  stream := TBufStream.Create(BUFSIZE);
  try
    // Write 100 random bytes. They fit into the BUFSIZE of the memory buffer
    nWrite1 := 100;
    SetLength(writeBuf1, nWrite1);
    for i:=0 to nWrite1-1 do writeBuf1[i] := Random(255);
    stream.WriteBuffer(writeBuf1[0], nWrite1);

    // Check stream size - must be equal to nWrite
    CheckEquals(nWrite1, stream.Size, 'Stream size mismatch (#1)');

    // Check stream position must be equal to nWrite
    CheckEquals(nWrite1, stream.Position, 'Stream position mismatch (#2)');

    // Bring stream pointer back to start
    stream.Position := 0;
    CheckEquals(0, stream.Position, 'Stream position mismatch (#3)');

    // Read the first 10 bytes just written and compare
    nRead := 10;
    SetLength(readBuf, nRead);
    nRead := stream.Read(readBuf[0], nRead);
    CheckEquals(10, nRead, 'Read/write size mismatch (#4)');
    for i:=0 to 9 do
      CheckEquals(writeBuf1[i], readBuf[i], Format('Read/write mismatch at position %d (#5)', [i]));

    // Back to start, and read the entire stream
    stream.Position := 0;
    nRead := stream.Size;
    Setlength(readBuf, nRead);
    nRead := stream.Read(readBuf[0], stream.Size);
    CheckEquals(nWrite1, nRead, 'Stream read size mismatch (#6)');
    for i:=0 to nWrite1-1 do
      CheckEquals(writeBuf1[i], readBuf[i], Format('Read/write mismatch at position %d (#7)', [i]));

    // Now put stream pointer to end and write another 2000 bytes. This crosses
    // the size of the memory buffer, and the stream must swap to file.
    stream.Seek(0, soFromEnd);
    CheckEquals(stream.Size, stream.Position, 'Stream position not at end (#8)');

    nWrite2 := 2000;
    SetLength(writeBuf2, nWrite2);
    for i:=0 to nWrite2-1 do writeBuf2[i] := Random(255);
    stream.WriteBuffer(writeBuf2[0], nWrite2);

    // The stream pointer must be at 100+2000, same for the size
    CheckEquals(nWrite1+nWrite2, stream.Position, 'Stream position mismatch (#9)');
    CheckEquals(nWrite1+nWrite2, stream.Size, 'Stream size mismatch (#10)');

    // Read the last 10 bytes and compare
    Stream.Seek(10, soFromEnd);
    SetLength(readBuf, 10);
    Stream.ReadBuffer(readBuf[0], 10);
    for i:=0 to 9 do
      CheckEquals(writeBuf2[nWrite2-10+i], readBuf[i], Format('Read/write mismatch at position %d from end (#11)', [i]));

    // Now read all from beginning
    Stream.Position := 0;
    SetLength(readBuf, stream.Size);
    nRead := Stream.Read(readBuf[0], stream.Size);
    CheckEquals(nWrite1+nWrite2, nRead, 'Read/write size mismatch (#4)');
    for i:=0 to nRead-1 do
      if i < nWrite1 then
        CheckEquals(writeBuf1[i], readBuf[i], Format('Read/write mismatch at position %d (#11)', [i]))
      else
        CheckEquals(writeBuf2[i-nWrite1], readBuf[i], Format('Read/write mismatch at position %d (#11)', [i]));

  finally
    stream.Free;
  end;
end;

procedure TSpreadInternalTests.WriteToStreamTest(AFormat: TsSpreadsheetFormat);
var
  myworkbook: TsWorkbook;
  myworksheet: TsWorksheet;
  memstream: TMemoryStream;
  filestream: TMemoryStream;
  tempFile: String;
  pf, pm: Pointer;
  i, p: Integer;
begin
  tempFile := GetTempFileName;

  myworkbook := TsWorkbook.Create;
  myworksheet := myworkbook.AddWorksheet('Test');
  memstream := TMemoryStream.Create;
  filestream := TMemoryStream.Create;
  try
    myworksheet.WriteText(0, 0, 'Text');
    myworksheet.WriteNumber(0, 1, 12.345);
    myworksheet.WriteDateTime(0, 2, now() );

    // Write to file
    myworkbook.WriteToFile(tempfile, AFormat);

    // Write to memory stream
    myworkbook.WriteToStream(memstream, AFormat);

    // Determine length of "used" data, there seems to be scap at the end
    memstream.Position := 0;
    myworkbook.ReadFromStream(memstream, AFormat);
    p := memstream.Position;

    // Read file back into memory stream
    filestream.LoadFromFile(tempfile);

    // Compare both streams
    CheckEquals(filestream.Size, memstream.Size, 'Stream size mismatch');

    pf := filestream.Memory;
    pm := memStream.Memory;
    for i:=0 to p-1 do
    begin
      CheckEquals(PByte(pf)^, PByte(pm)^, 'Stream mismatch at position ' + IntToStr(i));
      inc(pf);
      inc(pm);
    end;

  finally
    filestream.Free;
    memstream.Free;
    myworkbook.Free;
  end;

  DeleteFile(tempFile);
end;

procedure TSpreadInternalTests.TestWriteToStream_Biff5;
begin
  WriteToStreamTest(sfExcel5);
end;

procedure TSpreadInternalTests.TestWriteToStream_Biff8;
begin
  WriteToStreamTest(sfExcel8);
end;


procedure TSpreadInternalTests.TestCellString;
var
  r,c: Cardinal;
  s: String;
  flags: TsRelFlags;
begin
  CheckEquals('$A$1',GetCellString(0,0,[]));
  CheckEquals('$Z$1',GetCellString(0,25,[])); //bug 26447
  CheckEquals('$AA$2',GetCellString(1,26,[])); //just past the last letter
  CheckEquals('$GW$5',GetCellString(4,204,[])); //some big value
  CheckEquals('$IV$1',GetCellString(0,255,[])); //the last column of xls
  CheckEquals('$IW$1',GetCellString(0,256,[])); //the first column beyond xls
  CheckEquals('$XFD$1',GetCellString(0,16383,[])); // the last column of xlsx
  CheckEquals('$XFE$1',GetCellString(0,16384,[])); // the first column beyond xlsx

  // Something VERY big, beyond xlsx
//  s := 'ZZZZ1';   // this is case is no longer possible because max column count has been cut down to 65536
  s := 'CRAA1';
  ParseCellString(s, r, c, flags);
  CheckEquals(s, GetCellString(r, c, flags));
end;

procedure TSpreadInternalTests.FractionTest(AMaxDigits: Integer);
const
  N = 300;
var
  j: Integer;
  sollNum, sollDenom: Integer;
  sollValue: Double;
  actualNum, actualDenom: Int64;
  max: Integer;
  prec: Double;
begin
  max := Round(IntPower(10, AMaxDigits));
  prec := 0.001/max;
  for sollDenom := 1 to max-1 do
    for sollNum := 1 to sollDenom-1 do begin
      sollValue := StrToFloat(FormatFloat('0.000000000', sollNum/sollDenom));
      FloatToFraction(sollValue, max, actualNum, actualDenom);
      //FloatToFraction(sollValue, prec, max, max, actualNum, actualDenom);
      if (actualnum*solldenom div actualdenom <> sollnum) then
        fail(Format('Conversion error: %g = %d/%d turns to %d/%d (=%g)', [sollValue, sollNum, sollDenom, actualNum, actualDenom, actualNum/actualdenom]));
    end;
end;
                             (*
procedure TSpreadInternalTests.FractionTest_0;
const
  N = 300;
var
  j: Integer;
  sollNum, sollDenom: Integer;
  sollvalue: Double;
  actualNum, actualDenom: Int64;
begin
  sollNum := 1;
  for j := 1 to N do
  begin
    sollDenom := j;
//    sollValue := StrToFloat(FormatFloat('0.00000', sollNum/sollDenom));
    sollValue := 1.0/sollDenom;
//    FloatToFraction(sollvalue, 0.1/DIGITS, DIGITS, DIGITS, actualNum, actualDenom);
    FloatToFraction(sollvalue, 1000, actualNum, actualDenom);
    if actualDenom > sollDenom then
      fail(Format('Conversion error: approximated %d/%d turns to %d/%d', [sollNum, sollDenom, actualNum, actualDenom]));
  end;
end;
*)

procedure TSpreadInternalTests.FractionTest_1;
begin
  FractionTest(1);
end;

procedure TSpreadInternalTests.FractionTest_2;
begin
  FractionTest(2);
end;

procedure TSpreadInternalTests.FractionTest_3;
begin
  FractionTest(3);
end;

procedure TSpreadInternalTests.SetUp;
begin
end;

procedure TSpreadInternalTests.TearDown;
begin

end;

initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadInternalTests);

end.


