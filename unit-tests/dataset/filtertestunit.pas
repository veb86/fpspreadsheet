unit FilterTestUnit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpcunit, testregistry,
  DB,
  fpspreadsheet, fpstypes, fpsutils, fpsdataset;

type
  TFilterTest= class(TTestCase)
  private
    function CreateAndOpenDataset: TsWorksheetDataset;
    procedure Filter_01(Dataset: TDataset; var Accept: Boolean);    // 'IntCol < 2'
    procedure Filter_10(Dataset: TDataset; var Accept: Boolean);    // 'StringCol = 'abc
    procedure Filter_11(Dataset: TDataset; var Accept: Boolean);    // 'UPPER(StringCol) = 'ABC'
    procedure Filter_12(Dataset: TDataset; var Accept: Boolean);    // 'StringCol = 'ä'
    procedure Filter_13(Dataset: TDataset; var Accept: Boolean);    // 'StringCol > 'α'
    procedure Filter_20(Dataset: TDataset; var Accept: Boolean);    // 'WideStringCol = 'wABC'
    procedure Filter_21(Dataset: TDataset; var Accept: Boolean);    // 'UPPER(WideStringCol) = 'WABC'
    procedure Filter_22(Dataset: TDataset; var Accept: Boolean);    // 'WideStringCol = 'wä'
  protected
    procedure FilterTest(TestIndex: Integer);
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure FilterTest_01_Int;
    procedure FilterTest_10_String;
    procedure FilterTest_11_UpperString;
    procedure FilterTest_12_StringUTF8;
    procedure FilterTest_13_StringUTF8;

    procedure FilterTest_ByEvent_101_Int;
    procedure FilterTest_ByEvent_110_String;
    procedure FilterTest_ByEvent_111_UpperString;
    procedure FilterTest_ByEvent_112_String_UTF8;
    procedure FilterTest_ByEvent_113_String_UTF8;
    procedure FilterTest_ByEvent_120_WideString;
    procedure FilterTest_ByEvent_121_UpperWideString;
    procedure FilterTest_ByEvent_122_WideString_UTF8;
  end;

implementation

const
  FILE_NAME = 'testfile.xlsx';
  SHEET_NAME = 'Sheet';
  INT_COL = 0;
  STRING_COL = 1;
  WIDESTRING_COL = 2;
  INT_FIELD = 'IntCol';
  STRING_FIELD = 'StringCol';
  WIDESTRING_FIELD = 'WideStringCol';

var
  DataFileName: String;

type
  TTestRow = record
    IntValue: Integer;
    StringValue: String;
    WideStringValue: Widestring;
  end;

const
  // Unfiltered test values
  UNFILTERED: array[0..7] of TTestRow = (       // Index
    (IntValue: 10; StringValue: 'abc'; WideStringValue: 'wabc'),         // 0
    (IntValue:  1; StringValue: 'ABC'; WideStringvalue: 'wABC'),         // 1
    (IntValue:  1; StringValue: 'a';   WideStringValue: 'wa'),           // 2
    (IntValue:  2; StringValue: 'A';   WideStringValue: 'wA'),           // 3
    (IntValue: -1; StringValue: 'xyz'; WideStringValue: 'wxyz'),         // 4
    (IntValue: 25; StringValue: 'ä';   WideStringValue: 'wä'),           // 5
    (IntValue: 30; StringValue: 'Äöü'; WideStringValue: 'wÄöü'),         // 6
    (IntValue:  5; StringValue: 'αβγä';WideStringValue: 'wαβγä')         // 7
  );

  // These are the indexes into the UNFILTERED array after filtering
  FILTERED_01: array[0..2] of Integer = (1, 2, 4);  // 'IntCol < 2'
  FILTERED_10: array[0..0] of Integer = (0);        // 'StringCol = 'abc'
  FILTERED_11: array[0..1] of Integer = (0, 1);     // 'UPPER(StringCol) = 'ABC'
  FILTERED_12: array[0..0] of Integer = (5);        // StringCol = 'ä'
  FILTERED_13: array[0..0] of Integer = (7);        // StringCol >= 'α'
  FILTERED_20: array[0..0] of Integer = (1);        // 'WideStringCol = 'wABC'
  FILTERED_21: array[0..1] of Integer = (0, 1);     // 'UPPER(WideStringCol) = 'WABC'
  FILTERED_22: array[0..0] of Integer = (5);        // WideStringCol = 'wä'

  EXPRESSION_01 = 'IntCol < 2';
  EXPRESSION_10 = 'StringCol = "abc"';
  EXPRESSION_11 = 'UPPER(StringCol) = "ABC"';
  EXPRESSION_12 = 'StringCol = "ä"';
  EXPRESSION_13 = 'StringCol >= "α"';
  EXPRESSION_20 = 'WideStringCol = "wABC"';
  EXPRESSION_21 = 'UPPER(WideStringCol) = "WABC"';
  EXPRESSION_22 = 'WideStringCol = "wä"';


procedure TFilterTest.Filter_01(Dataset: TDataset; var Accept: Boolean);
begin
  Accept := Dataset.FieldByName(INT_FIELD).AsInteger < 2;
end;

procedure TFilterTest.Filter_10(Dataset: TDataset; var Accept: Boolean);
begin
  Accept := Dataset.FieldByName(STRING_FIELD).AsString = 'abc';
end;

procedure TFilterTest.Filter_11(Dataset: TDataset; var Accept: Boolean);
begin
  Accept := UpperCase(Dataset.FieldByName(STRING_FIELD).AsString) = 'ABC';
end;

procedure TFilterTest.Filter_12(Dataset: TDataset; var Accept: Boolean);
begin
  Accept := Dataset.FieldByName(STRING_FIELD).AsString = 'ä';
end;

procedure TFilterTest.Filter_13(Dataset: TDataset; var Accept: Boolean);
begin
  Accept := Dataset.FieldByName(STRING_FIELD).AsString >= 'α';
end;

procedure TFilterTest.Filter_20(Dataset: TDataset; var Accept: Boolean);
begin
  Accept := Dataset.FieldByName(WIDESTRING_FIELD).AsWideString = WideString('wABC');
end;

procedure TFilterTest.Filter_21(Dataset: TDataset; var Accept: Boolean);
begin
  Accept := Uppercase(Dataset.FieldByName(WIDESTRING_FIELD).AsWideString) = WideString('WABC');
end;

procedure TFilterTest.Filter_22(Dataset: TDataset; var Accept: Boolean);
begin
  Accept := Dataset.FieldByName(WIDESTRING_FIELD).AsWideString = WideString('wä');
end;

function TFilterTest.CreateAndOpenDataset: TsWorksheetDataset;
begin
  Result := TsWorksheetDataset.Create(nil);
  Result.FileName := DataFileName;
  Result.SheetName := SHEET_NAME;
  Result.AutoFieldDefs := false;
  Result.AddFieldDef(INT_FIELD, ftInteger);
  Result.AddFieldDef(STRING_FIELD, ftString, 20);
  Result.AddFieldDef(WIDESTRING_FIELD, ftWideString, 20);
  Result.CreateTable;
  Result.Open;
end;

procedure TFilterTest.FilterTest(TestIndex: Integer);
var
  dataset: TsWorksheetDataset;
  intField: TField;
  stringField: TField;
  widestringField: TField;
  actualInt: Integer;
  actualString: String;
  actualWideString: WideString;
  expectedInt: Integer;
  expectedString: String;
  expectedWideString: WideString;
  expectedRecordCount: Integer;
  i, idx: Integer;
begin
  dataset := CreateAndOpenDataset;
  try
    dataset.Filter := '';
    dataset.OnFilterRecord := nil;
    case TestIndex of
      // Tests using the Filter property
       1: dataset.Filter := EXPRESSION_01;    // Integer test
      10: dataset.Filter := EXPRESSION_10;    // String tests
      11: dataset.Filter := EXPRESSION_11;
      12: dataset.Filter := EXPRESSION_12;
      13: dataset.Filter := EXPRESSION_13;
      20: dataset.Filter := EXPRESSION_20;     // widestring tests
      21: dataset.Filter := EXPRESSION_21;
      22: dataset.Filter := EXPRESSION_22;
      // Tests using the OnFilterRecord event
      101: dataset.OnFilterRecord := @Filter_01;
      110: dataset.OnFilterRecord := @Filter_10;
      111: dataset.OnFilterRecord := @Filter_11;
      112: dataset.OnFilterRecord := @Filter_12;
      113: dataset.OnFilterRecord := @Filter_13;
      120: dataset.OnFilterRecord := @Filter_20;
      121: dataset.OnFilterRecord := @Filter_21;
      122: dataset.OnFilterRecord := @Filter_22;
    end;
    dataset.Filtered := true;

    case (TestIndex mod 100) of
       1: expectedRecordCount := Length(FILTERED_01);
      10: expectedRecordCount := Length(FILTERED_10);
      11: expectedRecordCount := Length(FILTERED_11);
      12: expectedRecordCount := Length(FILTERED_12);
      13: expectedRecordCount := Length(FILTERED_13);
      20: expectedRecordCount := Length(FILTERED_20);
      21: expectedRecordCount := Length(FILTERED_21);
      22: expectedRecordCount := Length(FILTERED_22);
    end;

    intField := dataset.FieldByName(INT_FIELD);
    stringField := dataset.FieldByName(STRING_FIELD);
    wideStringField := dataset.FieldByName(WIDESTRING_FIELD);

    dataset.First;
    i := 0;
    while not dataset.EOF do
    begin
      CheckEquals(true, i < expectedRecordCount, 'Record count mismatch.');

      case TestIndex mod 100 of
         1: idx := FILTERED_01[i];
        10: idx := FILTERED_10[i];
        11: idx := FILTERED_11[i];
        12: idx := FILTERED_12[i];
        13: idx := FILTERED_13[i];
        20: idx := FILTERED_20[i];
        21: idx := FILTERED_21[i];
        22: idx := FILTERED_22[i];
      end;

      actualInt := intField.AsInteger;
      actualString := stringField.AsString;
      actualWideString := wideStringField.AsWideString;

      expectedInt := UNFILTERED[idx].IntValue;
      expectedString := UNFILTERED[idx].StringValue;
      expectedWideString := UNFILTERED[idx].WideStringValue;

      CheckEquals(
        expectedInt,
        actualInt,
        'Integer field value mismatch in row ' + IntToStr(i)
      );
      CheckEquals(
        expectedString,
        actualString,
        'String field value mismatch in row ' + IntToStr(i)
      );
      CheckEquals(
        expectedWideString,
        actualWideString,
        'Widestring field value mismatch in row ' + IntToStr(i)
      );

      inc(i);
      dataset.Next;
    end;

    CheckEquals(true, i = expectedRecordCount, 'Record count mismatch.');

  finally
    dataset.Free;
  end;
end;

procedure TFilterTest.FilterTest_01_Int;
begin
  FilterTest(1);
end;

procedure TFilterTest.FilterTest_10_String;
begin
  FilterTest(10);
end;

procedure TFilterTest.FilterTest_11_UpperString;
begin
  FilterTest(11);
end;

procedure TFilterTest.FilterTest_12_StringUTF8;
begin
  FilterTest(12);
end;

procedure TFilterTest.FilterTest_13_StringUTF8;
begin
  FilterTest(13);
end;

procedure TFilterTest.FilterTest_ByEvent_101_Int;
begin
  FilterTest(101);
end;

procedure TFilterTest.FilterTest_ByEvent_110_String;
begin
  FilterTest(110);
end;

procedure TFilterTest.FilterTest_ByEvent_111_UpperString;
begin
  FilterTest(111);
end;

procedure TFilterTest.FilterTest_ByEvent_112_String_UTF8;
begin
  FilterTest(112);
end;

procedure TFilterTest.FilterTest_ByEvent_113_String_UTF8;
begin
  FilterTest(113);
end;

procedure TFilterTest.FilterTest_ByEvent_120_WideString;
begin
  FilterTest(120);
end;

procedure TFilterTest.FilterTest_ByEvent_121_UpperWideString;
begin
  FilterTest(121);
end;

procedure TFilterTest.FilterTest_ByEvent_122_WideString_UTF8;
begin
  FilterTest(122);
end;

procedure TFilterTest.SetUp;
var
  i, r: Integer;
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
begin
  // Create test spreadsheet file
  workbook := TsWorkbook.Create;
  try
    // Create worksheet
    worksheet := workbook.AddWorkSheet(SHEET_NAME);

    // Write headers (= field names)
    worksheet.WriteText(0, INT_COL, INT_FIELD);
    worksheet.WriteText(0, STRING_COL, STRING_FIELD);
    worksheet.WriteText(0, WIDESTRING_COL, WIDESTRING_FIELD);

    // Write values
    for i := Low(UNFILTERED) to High(UNFILTERED) do
    begin
      r := 1 + (i - Low(UNFILTERED));
      worksheet.WriteNumber(r, INT_COL, UNFILTERED[i].IntValue, nfFixed, 0);
      worksheet.WriteText(r, STRING_COL, UNFILTERED[i].StringValue);
      worksheet.WriteText(r, WIDESTRING_COL, UNFILTERED[i].WideStringValue);
    end;

    // Save
    DataFileName := GetTempDir + FILE_NAME;
    workbook.WriteToFile(DataFileName, true);
  finally
    workbook.Free;
  end;
end;

procedure TFilterTest.TearDown;
begin
  if FileExists(DataFileName) then DeleteFile(DataFileName);
end;


initialization
  RegisterTest(TFilterTest);

end.

