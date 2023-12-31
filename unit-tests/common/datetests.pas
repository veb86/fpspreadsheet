unit datetests;

{$mode objfpc}{$H+}

{
Adding tests/test data:
1. Add a new value to column A in the relevant worksheet, and save the spreadsheet read-only
   (for dates, there are 2 files, with different datemodes. Use them both...)
   Repeat this for all supported spreadsheet formats (Excel XLS, ODF, etc)
2. Increase SollDates array size
3. Add value from 1) to InitNormVariables so you can test against it
4. Add your read test(s), read and check read value against SollDates[<added number>]
}

interface

uses
  // Not using lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, {%H-}fpsallformats, fpspreadsheet, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

var
  // Norm to test against - list of dates/times that should occur in spreadsheet
  SollDates: array[0..37] of TDateTime; //"Soll" is a German word in Dutch accountancy jargon meaning "normative value to check against". There ;)
  // Initializes Soll*/normative variables.
  // Useful in test setup procedures to make sure the norm is correct.
  procedure InitSollDates;

type
  { TSpreadReadDateTests }
  // Read from xls/xml file with known values to test interoperability with Excel/LibreOffice/OpenOffice
  TSpreadReadDateTests= class(TTestCase)
  private
    // Tries to read date from the external file in column A, specified (0-based) row
    procedure TestReadDate(FileName: string; Row: integer);
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
  published
    // Reads dates, date/time and time values from spreadsheet and checks against list
    // One cell per test so some tests can fail and those further below may still work
    procedure TestReadDate1904_0; //date tests
    procedure TestReadDate1904_1; //date and time
    procedure TestReadDate1904_2;
    procedure TestReadDate1904_3;
    procedure TestReadDate1904_4; //time only tests start here
    procedure TestReadDate1904_5;
    procedure TestReadDate1904_6;
    procedure TestReadDate1904_7;
    procedure TestReadDate1904_8;
    procedure TestReadDate1904_9;
    procedure TestReadDate1904_10;
    procedure TestReadDate1904_11;
    procedure TestReadDate1904_12;
    procedure TestReadDate1904_13;
    procedure TestReadDate1904_14;
    procedure TestReadDate1904_15;
    procedure TestReadDate1904_16;
    procedure TestReadDate1904_17;
    procedure TestReadDate1904_18;
    procedure TestReadDate1904_19;
    procedure TestReadDate1904_20;
    procedure TestReadDate1904_21;
    procedure TestReadDate1904_22;
    procedure TestReadDate1904_23;
    procedure TestReadDate1904_24;
    procedure TestReadDate1904_25;
    procedure TestReadDate1904_26;
    procedure TestReadDate1904_27;
    procedure TestReadDate1904_28;
    procedure TestReadDate1904_29;
    procedure TestReadDate1904_30;
    procedure TestReadDate1904_31;
    procedure TestReadDate1904_32;
    procedure TestReadDate1904_33;
    procedure TestReadDate1904_34;
    procedure TestReadDate1904_35;
    procedure TestReadDate1904_36;
    procedure TestReadDate1904_37;

    procedure TestReadDate1899_0; //same as above except with the 1899/1900 date system set
    procedure TestReadDate1899_1;
    procedure TestReadDate1899_2;
    procedure TestReadDate1899_3;
    procedure TestReadDate1899_4;
    procedure TestReadDate1899_5;
    procedure TestReadDate1899_6;
    procedure TestReadDate1899_7;
    procedure TestReadDate1899_8;
    procedure TestReadDate1899_9;
    procedure TestReadDate1899_10;
    procedure TestReadDate1899_11;
    procedure TestReadDate1899_12;
    procedure TestReadDate1899_13;
    procedure TestReadDate1899_14;
    procedure TestReadDate1899_15;
    procedure TestReadDate1899_16;
    procedure TestReadDate1899_17;
    procedure TestReadDate1899_18;
    procedure TestReadDate1899_19;
    procedure TestReadDate1899_20;
    procedure TestReadDate1899_21;
    procedure TestReadDate1899_22;
    procedure TestReadDate1899_23;
    procedure TestReadDate1899_24;
    procedure TestReadDate1899_25;
    procedure TestReadDate1899_26;
    procedure TestReadDate1899_27;
    procedure TestReadDate1899_28;
    procedure TestReadDate1899_29;
    procedure TestReadDate1899_30;
    procedure TestReadDate1899_31;
    procedure TestReadDate1899_32;
    procedure TestReadDate1899_33;
    procedure TestReadDate1899_34;
    procedure TestReadDate1899_35;
    procedure TestReadDate1899_36;
    procedure TestReadDate1899_37;

    procedure TestReadODFDate1904_0; // same as above except OpenDocument/ODF format
    procedure TestReadODFDate1904_1; //date and time
    procedure TestReadODFDate1904_2;
    procedure TestReadODFDate1904_3;
    procedure TestReadODFDate1904_4; //time only tests start here
    procedure TestReadODFDate1904_5;
    procedure TestReadODFDate1904_6;
    procedure TestReadODFDate1904_7;
    procedure TestReadODFDate1904_8;
    procedure TestReadODFDate1904_9;
    procedure TestReadODFDate1904_10;
    procedure TestReadODFDate1904_11;
    procedure TestReadODFDate1904_12;
    procedure TestReadODFDate1904_13;
    procedure TestReadODFDate1904_14;
    procedure TestReadODFDate1904_15;
    procedure TestReadODFDate1904_16;
    procedure TestReadODFDate1904_17;
    procedure TestReadODFDate1904_18;
    procedure TestReadODFDate1904_19;
    procedure TestReadODFDate1904_20;
    procedure TestReadODFDate1904_21;
    procedure TestReadODFDate1904_22;
    procedure TestReadODFDate1904_23;
    procedure TestReadODFDate1904_24;
    procedure TestReadODFDate1904_25;
    procedure TestReadODFDate1904_26;
    procedure TestReadODFDate1904_27;
    procedure TestReadODFDate1904_28;
    procedure TestReadODFDate1904_29;
    procedure TestReadODFDate1904_30;
    procedure TestReadODFDate1904_31;
    procedure TestReadODFDate1904_32;
    procedure TestReadODFDate1904_33;
    procedure TestReadODFDate1904_34;
    procedure TestReadODFDate1904_35;
    procedure TestReadODFDate1904_36;
    procedure TestReadODFDate1904_37;

    procedure TestReadODFDate1899_0; //same as above except with the 1899/1900 date system set
    procedure TestReadODFDate1899_1;
    procedure TestReadODFDate1899_2;
    procedure TestReadODFDate1899_3;
    procedure TestReadODFDate1899_4;
    procedure TestReadODFDate1899_5;
    procedure TestReadODFDate1899_6;
    procedure TestReadODFDate1899_7;
    procedure TestReadODFDate1899_8;
    procedure TestReadODFDate1899_9;
    procedure TestReadODFDate1899_10;
    procedure TestReadODFDate1899_11;
    procedure TestReadODFDate1899_12;
    procedure TestReadODFDate1899_13;
    procedure TestReadODFDate1899_14;
    procedure TestReadODFDate1899_15;
    procedure TestReadODFDate1899_16;
    procedure TestReadODFDate1899_17;
    procedure TestReadODFDate1899_18;
    procedure TestReadODFDate1899_19;
    procedure TestReadODFDate1899_20;
    procedure TestReadODFDate1899_21;
    procedure TestReadODFDate1899_22;
    procedure TestReadODFDate1899_23;
    procedure TestReadODFDate1899_24;
    procedure TestReadODFDate1899_25;
    procedure TestReadODFDate1899_26;
    procedure TestReadODFDate1899_27;
    procedure TestReadODFDate1899_28;
    procedure TestReadODFDate1899_29;
    procedure TestReadODFDate1899_30;
    procedure TestReadODFDate1899_31;
    procedure TestReadODFDate1899_32;
    procedure TestReadODFDate1899_33;
    procedure TestReadODFDate1899_34;
    procedure TestReadODFDate1899_35;
    procedure TestReadODFDate1899_36;
    procedure TestReadODFDate1899_37;

    procedure TestReadOOXMLDate1904_0; // same as above except Excel xlsx format
    procedure TestReadOOXMLDate1904_1; //date and time
    procedure TestReadOOXMLDate1904_2;
    procedure TestReadOOXMLDate1904_3;
    procedure TestReadOOXMLDate1904_4; //time only tests start here
    procedure TestReadOOXMLDate1904_5;
    procedure TestReadOOXMLDate1904_6;
    procedure TestReadOOXMLDate1904_7;
    procedure TestReadOOXMLDate1904_8;
    procedure TestReadOOXMLDate1904_9;
    procedure TestReadOOXMLDate1904_10;
    procedure TestReadOOXMLDate1904_11;
    procedure TestReadOOXMLDate1904_12;
    procedure TestReadOOXMLDate1904_13;
    procedure TestReadOOXMLDate1904_14;
    procedure TestReadOOXMLDate1904_15;
    procedure TestReadOOXMLDate1904_16;
    procedure TestReadOOXMLDate1904_17;
    procedure TestReadOOXMLDate1904_18;
    procedure TestReadOOXMLDate1904_19;
    procedure TestReadOOXMLDate1904_20;
    procedure TestReadOOXMLDate1904_21;
    procedure TestReadOOXMLDate1904_22;
    procedure TestReadOOXMLDate1904_23;
    procedure TestReadOOXMLDate1904_24;
    procedure TestReadOOXMLDate1904_25;
    procedure TestReadOOXMLDate1904_26;
    procedure TestReadOOXMLDate1904_27;
    procedure TestReadOOXMLDate1904_28;
    procedure TestReadOOXMLDate1904_29;
    procedure TestReadOOXMLDate1904_30;
    procedure TestReadOOXMLDate1904_31;
    procedure TestReadOOXMLDate1904_32;
    procedure TestReadOOXMLDate1904_33;
    procedure TestReadOOXMLDate1904_34;
    procedure TestReadOOXMLDate1904_35;
    procedure TestReadOOXMLDate1904_36;
    procedure TestReadOOXMLDate1904_37;

    procedure TestReadOOXMLDate1899_0; //same as above except with the 1899/1900 date system set
    procedure TestReadOOXMLDate1899_1;
    procedure TestReadOOXMLDate1899_2;
    procedure TestReadOOXMLDate1899_3;
    procedure TestReadOOXMLDate1899_4;
    procedure TestReadOOXMLDate1899_5;
    procedure TestReadOOXMLDate1899_6;
    procedure TestReadOOXMLDate1899_7;
    procedure TestReadOOXMLDate1899_8;
    procedure TestReadOOXMLDate1899_9;
    procedure TestReadOOXMLDate1899_10;
    procedure TestReadOOXMLDate1899_11;
    procedure TestReadOOXMLDate1899_12;
    procedure TestReadOOXMLDate1899_13;
    procedure TestReadOOXMLDate1899_14;
    procedure TestReadOOXMLDate1899_15;
    procedure TestReadOOXMLDate1899_16;
    procedure TestReadOOXMLDate1899_17;
    procedure TestReadOOXMLDate1899_18;
    procedure TestReadOOXMLDate1899_19;
    procedure TestReadOOXMLDate1899_20;
    procedure TestReadOOXMLDate1899_21;
    procedure TestReadOOXMLDate1899_22;
    procedure TestReadOOXMLDate1899_23;
    procedure TestReadOOXMLDate1899_24;
    procedure TestReadOOXMLDate1899_25;
    procedure TestReadOOXMLDate1899_26;
    procedure TestReadOOXMLDate1899_27;
    procedure TestReadOOXMLDate1899_28;
    procedure TestReadOOXMLDate1899_29;
    procedure TestReadOOXMLDate1899_30;
    procedure TestReadOOXMLDate1899_31;
    procedure TestReadOOXMLDate1899_32;
    procedure TestReadOOXMLDate1899_33;
    procedure TestReadOOXMLDate1899_34;
    procedure TestReadOOXMLDate1899_35;
    procedure TestReadOOXMLDate1899_36;
    procedure TestReadOOXMLDate1899_37;

  end;

  { TSpreadWriteReadDateTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadDateTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    // Reads dates, date/time and time values from spreadsheet and checks against list
    // One cell per test so some tests can fail and those further below may still work
    procedure TestWriteReadDates(AFormat: TsSpreadsheetFormat);
    procedure TestWriteReadMilliseconds(AFormat: TsSpreadsheetFormat);

  published
    procedure TestWriteReadDates_BIFF2;
    procedure TestWriteReadDates_BIFF5;
    procedure TestWriteReadDates_BIFF8;
    procedure TestWriteReadDates_ODS;
    procedure TestWriteReadDates_OOXML;
    procedure TestWriteReadDates_XML;

    procedure TestWriteReadMilliseconds_BIFF2;
    procedure TestWriteReadMilliseconds_BIFF5;
    procedure TestWriteReadMilliseconds_BIFF8;
    procedure TestWriteReadMilliseconds_ODS;
    procedure TestWriteReadMilliseconds_OOXML;
    procedure TestWriteReadMilliseconds_XML;
  end;

  { TSpreadWriteReadYear1900Tests }
  { Tests to check whether the year-1900 bug in Excel is handled correctly }
  TSpreadWriteReadYear1900Tests = class(TTestCase)
  private
  protected
    procedure Setup; override;
    procedure TearDown; override;
    procedure TestWriteReadYear1900Dates(AFormat: TsSpreadsheetFormat);
  published
    procedure TestWriteReadYear1900Dates_BIFF2;
    procedure TestWriteReadYear1900Dates_BIFF5;
    procedure TestWriteReadYear1900Dates_BIFF8;
    procedure TestWriteReadYear1900Dates_ODS;
    procedure TestWriteReadYear1900Dates_OOXML;
    procedure TestWriteReadYear1900Dates_XML;
  end;


implementation

var
  TestWorksheet: TsWorksheet = nil;
  TestWorkbook: TsWorkbook = nil;
  TestFileName: String = '';

const
  DatesSheet = 'Dates'; //worksheet name

// Initialize array with variables that represent the values
// we expect to be in the test spreadsheet files.
//
// When adding tests, add values to this array
// and increase array size in variable declaration
procedure InitSollDates;
begin
  // Set up norm - MUST match spreadsheet cells exactly
  SollDates[0]:=EncodeDate(1905,09,12); //FPC number 2082
  SollDates[1]:=EncodeDate(1908,09,12)+EncodeTime(12,0,0,0); //noon
  SollDates[2]:=EncodeDate(2013,11,24);
  SollDates[3]:=EncodeDate(2030,12,31);
  SollDates[4]:=EncodeTime(0,0,0,0);
  SollDates[5]:=EncodeTime(0,0,1,0);
  SollDates[6]:=EncodeTime(1,0,0,0);
  SollDates[7]:=EncodeTime(3,0,0,0);
  SollDates[8]:=EncodeTime(12,0,0,0);
  SollDates[9]:=EncodeTime(18,0,0,0);
  SollDates[10]:=EncodeTime(23,59,0,0);
  SollDates[11]:=EncodeTime(23,59,59,0);

  SollDates[12]:=SollDates[1];  // #1 formatted as nfShortDateTime
  SollDates[13]:=SollDates[1];  // #1 formatted as nfShortTime
  SollDates[14]:=SollDates[1];  // #1 formatted as nfLongTime
  SollDates[15]:=SollDates[1];  // #1 formatted as nfShortTimeAM
  SollDates[16]:=SollDates[1];  // #1 formatted as nfLongTimeAM
  SollDates[17]:=SollDates[1];  // #1 formatted as nfCustom dd/mmm
  SollDates[18]:=SollDates[1];  // #1 formatted as nfCustom mmm/yy
  SollDates[19]:=SollDates[1];  // #1 formatted as nfCustom mm:ss

  SollDates[20]:=SollDates[5];  // #5 formatted as nfShortDateTime
  SollDates[21]:=SollDates[5];  // #5 formatted as nfShortTime
  SollDates[22]:=SollDates[5];  // #5 formatted as nfLongTime
  SollDates[23]:=SollDates[5];  // #5 formatted as nfShortTimeAM
  SollDates[24]:=SollDates[5];  // #5 formatted as nfLongTimeAM
  SollDates[25]:=SollDates[5];  // #5 formatted as nfCustom dd:mmm
  SollDates[26]:=SollDates[5];  // #5 formatted as nfCustom mmm:yy
  SollDates[27]:=SollDates[5];  // #5 formatted as nfCustom mm:ss

  SollDates[28]:=SollDates[11];  // #11 formatted as nfShortDateTime
  SollDates[29]:=SollDates[11];  // #11 formatted as nfShortTime
  SollDates[30]:=SollDates[11];  // #11 formatted as nfLongTime
  SollDates[31]:=SollDates[11];  // #11 formatted as nfShortTimeAM
  SollDates[32]:=SollDates[11];  // #11 formatted as nfLongTimeAM
  SollDates[33]:=SollDates[11];  // #11 formatted as nfCustom dd/mmm
  SollDates[34]:=SollDates[11];  // #11 formatted as nfCustom mmm/yy
  SollDates[35]:=SollDates[11];  // #11 formatted as nfCustom mmm:ss

  SollDates[36]:=EncodeTime(3,45,12,0);     // formatted as nfTimeDuration
  SollDates[37]:=EncodeTime(3,45,12,0) + 1  // formatted as nfTimeDuration
end;


{ TSpreadWriteReadDateTests }

procedure TSpreadWriteReadDateTests.SetUp;
begin
  inherited SetUp;
  InitSollDates;
end;

procedure TSpreadWriteReadDateTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadDateTests.TestWriteReadDates(AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  ActualDateTime: TDateTime;
  Row: Cardinal;
  TempFile: string; //write xls/xml to this file and read back from it
  ErrorMargin: TDateTime;
begin
  ErrorMargin := 1.0/(24*60*60*1000*100); // 0.01 ms
  TempFile:=NewTempFile;
  {// Not needed: use workbook.writetofile with overwrite=true
  if fileexists(TempFile) then
    DeleteFile(TempFile);
  }
  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:=MyWorkBook.AddWorksheet(DatesSheet);
    for Row := Low(SollDates) to High(SollDates) do
    begin
      // The last two test dates are assumed to be formatted as time-interval
      if Row >= High(SollDates) then
        MyWorksheet.WriteDateTime(Row, 0, SollDates[Row], nfCustom, '[h]:nn:ss')
      else
        MyWorkSheet.WriteDateTime(Row, 0, SollDates[Row], nfShortDateTime);
      // Some checks inside worksheet itself
      if not(MyWorkSheet.ReadAsDateTime(Row,0,ActualDateTime)) then
        Fail('Failed writing date time for cell '+CellNotation(MyWorkSheet,Row));
      CheckEquals(SollDates[Row], ActualDateTime,
        'Test date/time value mismatch cell '+CellNotation(MyWorksheet,Row));
    end;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook,DatesSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    // Read test data from A column & compare if written=original
    for Row := Low(SollDates) to High(SollDates) do
    begin
      if not(MyWorkSheet.ReadAsDateTime(Row,0,ActualDateTime)) then
        Fail('Could not read date time for cell '+CellNotation(MyWorkSheet,Row));
      CheckEquals(SollDates[Row], ActualDateTime, ErrorMargin,
        'Test date/time value mismatch cell '+CellNotation(MyWorkSheet,Row));
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadDateTests.TestWriteReadMilliseconds(
  AFormat: TsSpreadsheetFormat);
type
  TMillisecondTestParam = record
    h, m, s, ms: word;
    str1, str2, str3: String;
  end;
const
  SOLL_TIMES: array[0..2] of TMillisecondTestParam = (
    (h:12; m: 0; s: 0; ms:  0; str1:'12:00:00.0'; str2:'12:00:00.00'; str3:'12:00:00.000'),
    (h:23; m:59; s:59; ms: 10; str1:'23:59:59.0'; str2:'23:59:59.01'; str3:'23:59:59.010'),
    (h:23; m:59; s:59; ms:191; str1:'23:59:59.2'; str2:'23:59:59.19'; str3:'23:59:59.191')
  );
  FORMAT_STRINGS: array[1..3] of string = (
    'hh:nn:ss.z', 'hh:nn:ss.zz', 'hh:nn:ss.zzz');
  EPS = 0.0005*60*60*24;  // 0.5 ms
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  actualDateTime: TDateTime;
  actualStr: String;
  r, c: Cardinal;
  t: TTime;
  tempFile: String;
begin
  tempFile := NewTempFile;

  // Write out all test values
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.FormatSettings.DecimalSeparator := '.';
    MyWorkSheet := MyWorkBook.AddWorksheet(DatesSheet);
    for r := Low(SOLL_TIMES) to High(SOLL_TIMES) do
    begin
      with SOLL_TIMES[r] do t := EncodeTime(h, m, s, ms);
      for c := Low(FORMAT_STRINGS) to High(FORMAT_STRINGS) do
      begin
        MyWorkSheet.WriteDateTime(r, c, t, FORMAT_STRINGS[c]);

        // Some checks inside worksheet itself, before writing
        if not(MyWorkSheet.ReadAsDateTime(r, c, actualDateTime)) then
          Fail('Failed writing date time for cell '+CellNotation(MyWorkSheet, r, c));
        CheckEquals(t, actualDateTime, EPS,
          'Test date/time value mismatch cell '+CellNotation(MyWorksheet, r, c));
        actualStr := MyWorksheet.ReadAsText(r, c);
        case c of
          1: CheckEquals(SOLL_TIMES[r].str1, actualstr,
               'Cell string mismatch, cell '+CellNotation(Myworksheet, r, c));
          2: CheckEquals(SOLL_TIMES[r].str2, actualstr,
               'Cell string mismatch, cell '+CellNotation(Myworksheet, r, c));
          3: CheckEquals(SOLL_TIMES[r].str3, actualstr,
               'Cell string mismatch, cell '+CellNotation(Myworksheet, r, c));
        end;
      end;
    end;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.FormatSettings.DecimalSeparator := '.';
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    if AFormat = sfExcel2 then
      MyWorksheet := MyWorkbook.GetFirstWorksheet
    else
      MyWorksheet := GetWorksheetByName(MyWorkBook,DatesSheet);
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    // Read test data from A column & compare if written=original
    for r := Low(SOLL_TIMES) to High(SOLL_TIMES) do
    begin
      with SOLL_TIMES[r] do t := EncodeTime(h, m, s, ms);
      for c := Low(FORMAT_STRINGS) to High(FORMAT_STRINGS) do begin
        if not(MyWorkSheet.ReadAsDateTime(r, c, actualDateTime)) then
          Fail('Could not read date time for cell '+CellNotation(MyWorkSheet, r, c));
        CheckEquals(r, actualDateTime, EPS,
          'Test date/time value mismatch cell '+CellNotation(MyWorkSheet, r, c));
        actualStr := MyWorksheet.ReadAsText(r, c);
        case c of
          1: CheckEquals(SOLL_TIMES[r].str1, actualstr,
               'Cell string mismatch, cell '+CellNotation(Myworksheet, r, c));
          2: CheckEquals(SOLL_TIMES[r].str2, actualstr,
               'Cell string mismatch, cell '+CellNotation(Myworksheet, r, c));
          3: CheckEquals(SOLL_TIMES[r].str3, actualstr,
               'Cell string mismatch, cell '+CellNotation(Myworksheet, r, c));
        end;
      end;
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;

end;

procedure TSpreadWriteReadDateTests.TestWriteReadDates_BIFF2;
begin
  TestWriteReadDates(sfExcel2);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadDates_BIFF5;
begin
  TestWriteReadDates(sfExcel5);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadDates_BIFF8;
begin
  TestWriteReadDates(sfExcel8);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadDates_ODS;
begin
  TestWriteReadDates(sfOpenDocument);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadDates_OOXML;
begin
  TestWriteReadDates(sfOOXML);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadDates_XML;
begin
  TestWriteReadDates(sfExcelXML);
end;


{ TSpreadReadDateTests }

procedure TSpreadReadDateTests.TestReadDate(FileName: string; Row: integer);
var
  ActualDateTime: TDateTime;
  ErrorMargin: TDateTime; //margin for error in comparison test
begin
  ErrorMargin := 1E-5/(24*60*60*1000);  // = 10 nsec = 1E-8 sec (1 ns fails)

  if Row > High(SollDates) then
    fail('Error in test code: array bounds overflow. Check array size is correct.');

  // Load the file only if is the file name changes.
  if TestFileName <> FileName then
  begin
    if TestWorkbook <> nil then
      TestWorkbook.Free;

    // Open the spreadsheet
    TestWorkbook := TsWorkbook.Create;
    case UpperCase(ExtractFileExt(FileName)) of
      '.XLSX': TestWorkbook.ReadFromFile(FileName, sfOOXML);
      '.ODS' : TestWorkbook.ReadFromFile(FileName, sfOpenDocument);
      // Excel XLS/BIFF
      else TestWorkbook.ReadFromFile(FileName, sfExcel8);
    end;
    TestWorksheet := GetWorksheetByName(TestWorkBook, DatesSheet);
    if TestWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet '+DatesSheet);

    TestFileName := FileName;
  end;

  // We know these are valid time/date/datetime values....
  // Just test for empty string; we'll probably end up in a maze of localized date/time stuff
  // if we don't.
  CheckNotEquals(TestWorkSheet.ReadAsText(Row, 0), '',
    'Could not read date time as string for cell '+CellNotation(TestWorkSheet,Row));

  if not(TestWorkSheet.ReadAsDateTime(Row, 0, ActualDateTime)) then
    Fail('Could not read date time value for cell '+CellNotation(TestWorkSheet,Row));
  {$if (defined(mswindows)) or (FPC_FULLVERSION>=20701)}
  // FPC 2.6.x and trunk on Windows need this, also FPC trunk on Linux x64
  CheckEquals(SollDates[Row],ActualDateTime,ErrorMargin,'Test date/time value mismatch, '
    +'cell '+CellNotation(TestWorksheet,Row));
  {$else}
  // Non-windows: test without error margin
  CheckEquals(SollDates[Row],ActualDateTime,'Test date/time value mismatch, '
    +'cell '+CellNotation(TestWorksheet,Row));
  {$endif}

  // Don't free the workbook here - it will be reused. It is destroyed at finalization.
end;

procedure TSpreadReadDateTests.SetUp;
begin
  InitSollDates;
end;

procedure TSpreadReadDateTests.TearDown;
begin

end;


{ BIFF8 1904 datemode tests }

procedure TSpreadReadDateTests.TestReadDate1904_0;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,0);
end;

procedure TSpreadReadDateTests.TestReadDate1904_1;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,1);
end;

procedure TSpreadReadDateTests.TestReadDate1904_2;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,2);
end;

procedure TSpreadReadDateTests.TestReadDate1904_3;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,3);
end;

procedure TSpreadReadDateTests.TestReadDate1904_4;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,4);
end;

procedure TSpreadReadDateTests.TestReadDate1904_5;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,5);
end;

procedure TSpreadReadDateTests.TestReadDate1904_6;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,6);
end;

procedure TSpreadReadDateTests.TestReadDate1904_7;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,7);
end;

procedure TSpreadReadDateTests.TestReadDate1904_8;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,8);
end;

procedure TSpreadReadDateTests.TestReadDate1904_9;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,9);
end;

procedure TSpreadReadDateTests.TestReadDate1904_10;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,10);
end;

procedure TSpreadReadDateTests.TestReadDate1904_11;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,11);
end;

procedure TSpreadReadDateTests.TestReadDate1904_12;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,12);
end;

procedure TSpreadReadDateTests.TestReadDate1904_13;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,13);
end;

procedure TSpreadReadDateTests.TestReadDate1904_14;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,14);
end;

procedure TSpreadReadDateTests.TestReadDate1904_15;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,15);
end;

procedure TSpreadReadDateTests.TestReadDate1904_16;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,16);
end;

procedure TSpreadReadDateTests.TestReadDate1904_17;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,17);
end;

procedure TSpreadReadDateTests.TestReadDate1904_18;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,18);
end;

procedure TSpreadReadDateTests.TestReadDate1904_19;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,19);
end;

procedure TSpreadReadDateTests.TestReadDate1904_20;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,20);
end;

procedure TSpreadReadDateTests.TestReadDate1904_21;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,21);
end;

procedure TSpreadReadDateTests.TestReadDate1904_22;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,22);
end;

procedure TSpreadReadDateTests.TestReadDate1904_23;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,23);
end;

procedure TSpreadReadDateTests.TestReadDate1904_24;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,24);
end;

procedure TSpreadReadDateTests.TestReadDate1904_25;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,25);
end;

procedure TSpreadReadDateTests.TestReadDate1904_26;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,26);
end;

procedure TSpreadReadDateTests.TestReadDate1904_27;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,27);
end;

procedure TSpreadReadDateTests.TestReadDate1904_28;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,28);
end;

procedure TSpreadReadDateTests.TestReadDate1904_29;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,29);
end;

procedure TSpreadReadDateTests.TestReadDate1904_30;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,30);
end;

procedure TSpreadReadDateTests.TestReadDate1904_31;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,31);
end;

procedure TSpreadReadDateTests.TestReadDate1904_32;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,32);
end;

procedure TSpreadReadDateTests.TestReadDate1904_33;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,33);
end;

procedure TSpreadReadDateTests.TestReadDate1904_34;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,34);
end;

procedure TSpreadReadDateTests.TestReadDate1904_35;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,35);
end;

procedure TSpreadReadDateTests.TestReadDate1904_36;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,36);
end;

procedure TSpreadReadDateTests.TestReadDate1904_37;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1904,37);
end;


{ BIFF8 1899 datemode tests }

procedure TSpreadReadDateTests.TestReadDate1899_0;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,0);
end;

procedure TSpreadReadDateTests.TestReadDate1899_1;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,1);
end;

procedure TSpreadReadDateTests.TestReadDate1899_2;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,2);
end;

procedure TSpreadReadDateTests.TestReadDate1899_3;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,3);
end;

procedure TSpreadReadDateTests.TestReadDate1899_4;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,4);
end;

procedure TSpreadReadDateTests.TestReadDate1899_5;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,5);
end;

procedure TSpreadReadDateTests.TestReadDate1899_6;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,6);
end;

procedure TSpreadReadDateTests.TestReadDate1899_7;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,7);
end;

procedure TSpreadReadDateTests.TestReadDate1899_8;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,8);
end;

procedure TSpreadReadDateTests.TestReadDate1899_9;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,9);
end;

procedure TSpreadReadDateTests.TestReadDate1899_10;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,10);
end;

procedure TSpreadReadDateTests.TestReadDate1899_11;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,11);
end;

procedure TSpreadReadDateTests.TestReadDate1899_12;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,12);
end;

procedure TSpreadReadDateTests.TestReadDate1899_13;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,13);
end;

procedure TSpreadReadDateTests.TestReadDate1899_14;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,14);
end;

procedure TSpreadReadDateTests.TestReadDate1899_15;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,15);
end;

procedure TSpreadReadDateTests.TestReadDate1899_16;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,16);
end;

procedure TSpreadReadDateTests.TestReadDate1899_17;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,17);
end;

procedure TSpreadReadDateTests.TestReadDate1899_18;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,18);
end;

procedure TSpreadReadDateTests.TestReadDate1899_19;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,19);
end;

procedure TSpreadReadDateTests.TestReadDate1899_20;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,20);
end;

procedure TSpreadReadDateTests.TestReadDate1899_21;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,21);
end;

procedure TSpreadReadDateTests.TestReadDate1899_22;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,22);
end;

procedure TSpreadReadDateTests.TestReadDate1899_23;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,23);
end;

procedure TSpreadReadDateTests.TestReadDate1899_24;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,24);
end;

procedure TSpreadReadDateTests.TestReadDate1899_25;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,25);
end;

procedure TSpreadReadDateTests.TestReadDate1899_26;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,26);
end;

procedure TSpreadReadDateTests.TestReadDate1899_27;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,27);
end;

procedure TSpreadReadDateTests.TestReadDate1899_28;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,28);
end;

procedure TSpreadReadDateTests.TestReadDate1899_29;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,29);
end;

procedure TSpreadReadDateTests.TestReadDate1899_30;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,30);
end;

procedure TSpreadReadDateTests.TestReadDate1899_31;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,31);
end;

procedure TSpreadReadDateTests.TestReadDate1899_32;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,32);
end;

procedure TSpreadReadDateTests.TestReadDate1899_33;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,33);
end;

procedure TSpreadReadDateTests.TestReadDate1899_34;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,34);
end;

procedure TSpreadReadDateTests.TestReadDate1899_35;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,35);
end;

procedure TSpreadReadDateTests.TestReadDate1899_36;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,36);
end;

procedure TSpreadReadDateTests.TestReadDate1899_37;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileBIFF8_1899,37);
end;


{ ODS 1904 datemode tests }

procedure TSpreadReadDateTests.TestReadODFDate1904_0;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,0);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_1;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,1);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_2;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,2);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_3;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,3);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_4;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,4);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_5;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,5);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_6;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,6);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_7;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,7);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_8;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,8);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_9;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,9);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_10;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,10);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_11;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,11);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_12;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,12);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_13;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,13);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_14;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,14);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_15;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,15);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_16;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,16);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_17;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,17);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_18;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,18);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_19;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,19);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_20;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,20);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_21;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,21);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_22;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,22);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_23;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,23);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_24;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,24);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_25;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,25);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_26;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,26);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_27;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,27);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_28;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,28);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_29;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,29);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_30;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,30);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_31;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,31);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_32;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,32);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_33;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,33);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_34;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,34);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_35;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,35);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_36;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,36);
end;

procedure TSpreadReadDateTests.TestReadODFDate1904_37;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1904,37);
end;


{ ODS 1899 datemode tests }

procedure TSpreadReadDateTests.TestReadODFDate1899_0;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,0);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_1;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,1);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_2;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,2);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_3;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,3);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_4;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,4);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_5;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,5);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_6;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,6);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_7;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,7);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_8;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,8);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_9;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,9);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_10;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,10);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_11;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,11);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_12;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,12);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_13;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,13);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_14;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,14);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_15;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,15);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_16;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,16);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_17;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,17);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_18;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,18);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_19;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,19);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_20;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,20);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_21;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,21);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_22;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,22);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_23;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,23);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_24;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,24);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_25;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,25);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_26;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,26);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_27;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,27);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_28;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,28);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_29;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,29);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_30;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,30);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_31;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,31);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_32;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,32);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_33;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,33);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_34;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,34);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_35;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,35);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_36;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,36);
end;

procedure TSpreadReadDateTests.TestReadODFDate1899_37;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileODF_1899,37);
end;


{ Excel xlsx 1904 datemode tests }

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_0;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,0);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_1;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,1);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_2;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,2);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_3;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,3);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_4;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,4);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_5;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,5);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_6;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,6);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_7;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,7);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_8;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,8);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_9;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,9);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_10;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,10);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_11;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,11);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_12;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,12);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_13;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,13);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_14;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,14);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_15;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,15);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_16;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,16);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_17;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,17);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_18;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,18);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_19;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,19);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_20;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,20);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_21;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,21);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_22;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,22);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_23;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,23);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_24;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,24);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_25;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,25);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_26;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,26);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_27;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,27);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_28;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,28);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_29;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,29);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_30;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,30);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_31;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,31);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_32;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,32);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_33;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,33);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_34;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,34);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_35;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,35);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_36;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,36);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1904_37;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1904,37);
end;


{ ODS 1899 datemode tests }

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_0;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,0);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_1;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,1);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_2;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,2);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_3;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,3);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_4;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,4);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_5;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,5);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_6;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,6);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_7;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,7);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_8;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,8);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_9;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,9);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_10;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,10);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_11;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,11);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_12;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,12);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_13;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,13);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_14;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,14);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_15;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,15);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_16;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,16);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_17;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,17);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_18;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,18);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_19;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,19);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_20;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,20);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_21;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,21);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_22;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,22);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_23;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,23);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_24;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,24);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_25;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,25);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_26;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,26);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_27;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,27);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_28;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,28);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_29;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,29);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_30;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,30);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_31;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,31);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_32;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,32);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_33;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,33);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_34;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,34);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_35;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,35);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_36;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,36);
end;

procedure TSpreadReadDateTests.TestReadOOXMLDate1899_37;
begin
  TestReadDate(ExtractFilePath(ParamStr(0)) + TestFileOOXML_1899,37);
end;

//------------------------------------------------------------------------------

procedure TSpreadWriteReadDateTests.TestWriteReadMilliseconds_BIFF2;
begin
  TestWriteReadMilliseconds(sfExcel2);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadMilliseconds_BIFF5;
begin
  TestWriteReadMilliseconds(sfExcel5);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadMilliseconds_BIFF8;
begin
  TestWriteReadMilliseconds(sfExcel8);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadMilliseconds_ODS;
begin
  TestWriteReadMilliseconds(sfOpenDocument);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadMilliseconds_OOXML;
begin
  TestWriteReadMilliseconds(sfOOXML);
end;

procedure TSpreadWriteReadDateTests.TestWriteReadMilliseconds_XML;
begin
  TestWriteReadMilliseconds(sfExcelXML);
end;


{ =============================================================================}

var
  Y1900_SollDates: array of TDate;
  Y1900_SollTimes: array of TTime;
  Y1900_SollIntervals: array of TDateTime;

procedure InitY1900_SollDates;
begin
  SetLength(Y1900_SollDates, 5);
  Y1900_SollDates[0] := EncodeDate(1900, 1, 1);
  Y1900_SollDates[1] := EncodeDate(1900, 1, 2);
  Y1900_SollDates[2] := EncodeDate(1900, 2, 28);
  Y1900_SollDates[3] := Encodedate(1900, 3, 1);
  Y1900_SollDates[4] := Encodedate(1900, 3, 2);

  SetLength(Y1900_SollTimes, 3);
  Y1900_SollTimes[0] := EncodeTime(0, 0, 0, 0);
  Y1900_SollTimes[1] := EncodeTime(12, 0, 0, 0);
  Y1900_SollTimes[2] := EncodeTime(23, 59, 59, 0);

  SetlengtH(Y1900_SollIntervals, 7);
  Y1900_SollIntervals[0] := EncodeTime(0, 0, 0, 0);
  Y1900_SollIntervals[1] := EncodeTime(23, 59, 59, 0);
  Y1900_SollIntervals[2] := EncodeTime(23, 59, 59, 0) + 1.0;
  Y1900_SollIntervals[3] := EncodeTime(23, 59, 59, 0) + 59.0;
  Y1900_SollIntervals[4] := EncodeTime(23, 59, 59, 0) + 60.0;
  Y1900_SollIntervals[5] := EncodeTime(23, 59, 59, 0) + 61.0;
  Y1900_SollIntervals[6] := EncodeTime(23, 59, 59, 0) + 62.0;
end;

procedure TSpreadWriteReadYear1900Tests.Setup;
begin
  inherited;
  InitY1900_SollDates;
end;

procedure TSpreadWriteReadYear1900Tests.TearDown;
begin
  inherited;
end;

procedure TSpreadWriteReadYear1900Tests.TestWriteReadYear1900Dates(AFormat: TsSpreadsheetFormat);
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  r: Cardinal;
  i: Integer;
  actualDateTime: TDateTime;
  ok: boolean;
  tempFile: String;
  ErrorMargin: TDateTime; //margin for error in comparison test
begin
  ErrorMargin := 1E-5/(24*60*60*1000);  // = 10 nsec = 1E-8 sec (1 ns fails)
  tempFile := NewTempFile;

  book := TsWorkbook.Create;
  try
    sheet := book.Addworksheet('Year1900');
    r := 0;
    for i := Low(Y1900_SollDates) to High(Y1900_SollDates) do
    begin
      sheet.WriteDateTime(r, 0, Y1900_SollDates[i], nfShortDateTime);
      ok := sheet.ReadAsDateTime(r, 0, actualDateTime);
      CheckEquals(true, ok,
        'Test date detection error, cell '+CellNotation(sheet, r));
      CheckEquals(Y1900_SollDates[i], actualDateTime,
        'Test date value memory reading mismatch, cell '+CellNotation(sheet, r));
      inc(r);
    end;

    for i := Low(Y1900_SollTimes) to High(Y1900_SollTimes) do
    begin
      sheet.WriteDateTime(r, 0, Y1900_SollTimes[i], nfLongTime);
      ok := sheet.ReadAsDateTime(r, 0, actualDateTime);
      CheckEquals(true, ok,
        'Test time detection error, cell '+CellNotation(sheet, r));
      CheckEquals(Y1900_SollTimes[i], actualDateTime, ErrorMargin,
        'Test time value memory reading mismatch, cell '+CellNotation(sheet, r));
      inc(r);
    end;

    for i := Low(Y1900_SollIntervals) to High(Y1900_SollIntervals) do
    begin
      sheet.WriteDateTime(r, 0, Y1900_SollIntervals[i], nfCustom, '[h]:nn:ss');
      ok := sheet.ReadAsDateTime(r, 0, actualDateTime);
      CheckEquals(true, ok,
        'Test time detection error, cell '+CellNotation(sheet, r));
      CheckEquals(Y1900_SollIntervals[i], actualDateTime, ErrorMargin,
        'Test interval value memory reading mismatch, cell '+CellNotation(sheet, r));
      inc(r);
    end;

    book.WriteToFile(tempFile, AFormat, true);
  finally
    book.Free;
  end;

  book := TsWorkbook.Create;
  try
    book.ReadFromFile(tempFile, AFormat);
    sheet := book.GetFirstWorksheet;
    r := 0;
    for i := Low(Y1900_SollDates) to High(Y1900_SollDates) do
    begin
      ok := sheet.ReadAsDateTime(r, 0, actualDateTime);
      CheckEquals(true, ok,
        'Test date detection error, cell '+CellNotation(sheet, r));
      CheckEquals(Y1900_SollDates[i], actualDateTime,
        'Test date value file reading mismatch, cell '+CellNotation(sheet, r));
      inc(r);
    end;
    for i := Low(Y1900_SollTimes) to High(Y1900_SollTimes) do
    begin
      ok := sheet.ReadAsDateTime(r, 0, actualDateTime);
      CheckEquals(true, ok,
        'Test time detection error, cell '+CellNotation(sheet, r));
      CheckEquals(Y1900_SollTimes[i], actualDateTime, ErrorMargin,
        'Test time value file reading mismatch, cell '+CellNotation(sheet, r));
      inc(r);
    end;
    for i := Low(Y1900_SollIntervals) to High(Y1900_SollIntervals) do
    begin
      ok := sheet.ReadAsDateTime(r, 0, actualDateTime);
      CheckEquals(true, ok,
        'Test interval detection error, cell '+CellNotation(sheet, r));
      CheckEquals(Y1900_SollIntervals[i], actualDateTime, ErrorMargin,
        'Test interval value file reading mismatch, cell '+CellNotation(sheet, r));
      inc(r);
    end;
  finally
    book.Free;
  end;
  DeleteFile(TempFile);
end;

procedure TSpreadWriteReadYear1900Tests.TestWriteReadYear1900Dates_BIFF2;
begin
  TestWriteReadYear1900Dates(sfExcel2);
end;

procedure TSpreadWriteReadYear1900Tests.TestWriteReadYear1900Dates_BIFF5;
begin
  TestWriteReadYear1900Dates(sfExcel5);
end;

procedure TSpreadWriteReadYear1900Tests.TestWriteReadYear1900Dates_BIFF8;
begin
  TestWriteReadYear1900Dates(sfExcel8);
end;

procedure TSpreadWriteReadYear1900Tests.TestWriteReadYear1900Dates_ODS;
begin
  TestWriteReadYear1900Dates(sfOpenDocument);
end;

procedure TSpreadWriteReadYear1900Tests.TestWriteReadYear1900Dates_OOXML;
begin
  TestWriteReadYear1900Dates(sfOOXML);
end;

procedure TSpreadWriteReadYear1900Tests.TestWriteReadYear1900Dates_XML;
begin
  TestWriteReadYear1900Dates(sfExcelXML);
end;


{
begin
end;
published
  procedure TestWriteReadYear1900Dates_BIFF2;
  procedure TestWriteReadYear1900Dates_BIFF5;
  procedure TestWriteReadYear1900Dates_BIFF8;
  procedure TestWriteReadYear1900Dates_ODS;
  procedure TestWriteReadYear1900Dates_OOXML;
  procedure TestWriteReadYear1900Dates_XML;
end;
 }


initialization
  // Register so these tests are included in a full run
  RegisterTest(TSpreadReadDateTests);
  RegisterTest(TSpreadWriteReadDateTests);
  RegisterTest(TSpreadWriteReadYear1900Tests);

  InitSollDates; //useful to have norm data if other code want to use this unit

finalization
  FreeAndNil(TestWorkbook);

end.


