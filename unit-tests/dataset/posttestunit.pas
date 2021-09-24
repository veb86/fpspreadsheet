{ - Creates a new WorksheetDataset with a variety of fields
  - Appends a record and posts the dataset
  - Opens the created spreadsheet file and compares its cells with the
    posted data.
}

unit PostTestUnit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpcunit, testregistry,
  DB,
  fpsdataset, fpspreadsheet, fpstypes, fpsutils;

type

  TPostTest= class(TTestCase)
  protected
    procedure RunPostTest(ADataType: TFieldType; ASize: Integer = 0);
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure PostTest_Int;
    procedure PostTest_String_20;
    procedure PostTest_String_10;
    procedure PostTest_Widestring_20;
    procedure PostTest_Widestring_10;
  end;

implementation

uses
  LazUTF8, LazUTF16;

const
  FILE_NAME = 'testfile.xls';
  SHEET_NAME = 'Sheet';
  COL_NAME = 'TestCol';

var
  DataFileName: String;

type
  TTestRecord = record
    IntValue: Integer;
    StringValue: String;
    WideStringValue: WideString;
  end;

const
  TestData: Array[0..5] of TTestRecord = (
    (IntValue:  10; StringValue: 'abc';           WideStringValue: 'abc'),            // 0
    (IntValue: -20; StringValue: 'äöüαβγ';        WideStringvalue: 'äöüαβγ'),         // 1
    (IntValue: 100; StringValue: 'a234567890';    WideStringvalue: 'a234567890'),     // 2
    (IntValue:   0; StringValue: 'a234567890123'; WideStringvalue: 'a234567890123'),  // 3
    (IntValue: 501; StringValue: 'äα34567890';    WideStringValue: 'äα34567890'),     // 4
    (IntValue: 502; StringValue: 'äα34567890123'; WideStringValue: 'äα34567890123')   // 5
  );

procedure TPostTest.RunPostTest(ADataType: TFieldType; ASize: Integer = 0);
var
  dataset: TsWorksheetDataset;
  field: TField;
  i: Integer;
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  row, lastRow: Integer;
  actualIntValue: Integer;
  actualStringValue: String;
  actualWideStringValue: WideString;
  expectedIntValue: Integer;
  expectedStringValue: String;
  expectedWideStringValue: WideString;
begin
  dataset := TsWorksheetDataset.Create(nil);
  try
    dataset.FileName := DataFileName;
    dataset.SheetName := SHEET_NAME;
    dataset.AddFieldDef(COL_NAME, ADataType, ASize);
    dataset.CreateTable;
    dataset.Open;

    field := dataset.FieldByName(COL_NAME);
    for i := 0 to High(TestData) do
    begin
      dataset.Append;
      case ADataType of
        ftInteger    : field.AsInteger := TestData[i].IntValue;
        ftString     : field.AsString := TestData[i].StringValue;
        ftWideString : field.AsString := UTF8Decode(TestData[i].WideStringValue);
      end;
      dataset.Post;
    end;
    dataset.Close;
  finally
    dataset.Free;
  end;

  CheckEquals(
    true,
    FileExists(DatafileName),
    'Spreadsheet data file not found'
  );

  workbook := TsWorkbook.Create;
  try
    workbook.ReadFromFile(DataFileName);
    worksheet := workbook.GetWorksheetByName(SHEET_NAME);
    CheckEquals(
      true,
      worksheet <> nil,
      'Worksheet not found'
    );

    lastRow := worksheet.GetLastRowIndex(true);
    CheckEquals(
      Length(TestData),
      lastRow,
      'Row count mismatch in worksheet'
    );

    actualStringValue := worksheet.ReadAsText(0, 0);
    CheckEquals(
      COL_NAME,
      actualStringValue,
      'Column name mismatch'
    );

    i := 0;
    for row := 1 to lastRow do
    begin
      case ADataType of
        ftInteger:
          begin
            expectedIntValue := TestData[i].IntValue;
            actualIntValue := Round(worksheet.ReadAsNumber(row, 0));
            CheckEquals(
              expectedIntValue,
              actualIntValue,
              'Integer field mismatch, row ' + IntToStr(row)
            );
          end;
        ftString:
          begin
            expectedStringValue := UTF8Copy(TestData[i].StringValue, 1, ASize);
            actualStringValue := worksheet.ReadAsText(row, 0);
            CheckEquals(
              expectedStringValue,
              actualStringValue,
              'String field mismatch, Row ' + IntToStr(row)
            );
          end;
        ftWideString:
          begin
            expectedWideStringValue := UTF16Copy(TestData[i].WideStringValue, 1, ASize);
            actualWideStringValue  := UTF8Decode(worksheet.ReadAsText(row, 0));
            CheckEquals(
              expectedWidestringValue,
              actualWideStringValue,
              'Widestring field mismatch, row ' + IntToStr(row)
            );
          end;
        else
          raise Exception.Create('Field type not tested here.');
      end;
      inc(i);
    end;
  finally
    workbook.Free;
  end;
end;

procedure TPostTest.PostTest_Int;
begin
  RunPostTest(ftInteger);
end;

procedure TPostTest.PostTest_String_20;
begin
  RunPostTest(ftString, 20);
end;

procedure TPostTest.PostTest_String_10;
begin
  RunPostTest(ftString, 10);
end;

procedure TPostTest.PostTest_WideString_20;
begin
  RunPostTest(ftWideString, 20);
end;

procedure TPostTest.PostTest_WideString_10;
begin
  RunPostTest(ftWideString, 10);
end;

procedure TPostTest.SetUp;
begin
  DataFileName := GetTempDir + FILE_NAME;
end;

procedure TPostTest.TearDown;
begin
  if FileExists(DataFileName) then DeleteFile(DataFileName);
end;


initialization
  RegisterTest(TPostTest);

end.

