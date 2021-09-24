{ These tests check whether empty columns in the worksheet are ignored when
  FieldDefs are determined. }

unit EmptyColumnsTestUnit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpcunit, testregistry,
  DB,
  fpSpreadsheet, fpsTypes, fpsDataset;

type

  TEmptyColumnsTest = class(TTestCase)
  private
    function CreateAndOpenDataset(
      ATestIndex: Integer; AutoFieldDefs: Boolean): TsWorksheetDataset;
    procedure CreateWorksheet(ATestIndex: Integer);
  protected
    procedure TestFieldDefs(ATestIndex: Integer; AutoFieldDefs: Boolean);
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure Test_0;
    procedure Test_1;
    procedure Test_2;
    procedure Test_3;
    procedure Test_4;
    procedure Test_5;
    procedure Test_6;
    procedure Test_0_AutoFieldDefs;
    procedure Test_1_AutoFieldDefs;
    procedure Test_2_AutoFieldDefs;
    procedure Test_3_AutoFieldDefs;
    procedure Test_4_AutoFieldDefs;
    procedure Test_5_AutoFieldDefs;
    procedure Test_6_AutoFieldDefs;
  end;


implementation

const
  FILE_NAME = 'testfile.xlsx';
  SHEET_NAME = 'Sheet';

var
  DataFileName: String;

type
  TDataRec = record
    ColumnType: TFieldType;
    FieldDefIndex: Integer;
  end;
  TTestData = array [0..3] of TDataRec;  // colums 0..3 in worksheet

const
  TestCases: array[0..6] of TTestData = (
    (  //0
      (ColumnType:ftInteger; FieldDefIndex: 0),
      (ColumnType:ftFloat;   FieldDefIndex: 1),
      (ColumnType:ftString;  FieldDefIndex: 2),
      (ColumnType:ftDate;    FieldDefIndex: 3)
    ),
    ( // 1
      (ColumnType:ftInteger; FieldDefIndex: 0),
      (ColumnType:ftUnknown; FieldDefIndex:-1),
      (ColumnType:ftFloat;   FieldDefIndex: 1),
      (ColumnType:ftDate;    FieldDefIndex: 2)
    ),
    ( // 2
      (ColumnType:ftInteger; FieldDefIndex: 0),
      (ColumnType:ftFloat;   FieldDefIndex: 1),
      (ColumnType:ftUnknown; FieldDefIndex:-1),
      (ColumnType:ftString;  FieldDefIndex: 2)
    ),
    ( // 3
      (ColumnType:ftUnknown; FieldDefIndex:-1),
      (ColumnType:ftInteger; FieldDefIndex: 0),
      (ColumnType:ftFloat;   FieldDefIndex: 1),
      (ColumnType:ftString;  FieldDefIndex: 2)
    ),
    ( // 4
      (ColumnType:ftInteger; FieldDefIndex: 0),
      (ColumnType:ftString;  FieldDefIndex: 1),
      (ColumnType:ftDate;    FieldDefIndex: 2),
      (ColumnType:ftUnknown; FieldDefIndex:-1)
    ),
    ( // 5
      (ColumnType:ftInteger; FieldDefIndex: 0),
      (ColumnType:ftUnknown; FieldDefIndex:-1),
      (ColumnType:ftUnknown; FieldDefIndex:-1),
      (ColumnType:ftFloat;   FieldDefIndex: 1)
    ),
    ( // 6
      (ColumnType:ftUnknown; FieldDefIndex:-1),
      (ColumnType:ftUnknown; FieldDefIndex:-1),
      (ColumnType:ftUnknown; FieldDefIndex:-1),
      (ColumnType:ftInteger; FieldDefIndex: 0)
    )
  );

function TEmptyColumnsTest.CreateAndOpenDataset(
  ATestIndex: Integer; AutoFieldDefs: Boolean): TsWorksheetDataset;
var
  i: Integer;
begin
  Result := TsWorksheetDataset.Create(nil);
  Result.FileName := DataFileName;
  Result.SheetName := SHEET_NAME;
  Result.AutoFieldDefs := AutoFieldDefs;
  if not AutoFieldDefs then
  begin
    for i := 0 to Length(TTestData)-1 do
    begin
      case TestCases[ATestIndex][i].ColumnType of
        ftUnknown: ;
        ftInteger: Result.AddFieldDef('IntCol', ftInteger, 0, i);
        ftFloat: Result.AddFieldDef('FloatCol', ftFloat, 0, i);
        ftString: Result.AddFieldDef('StringCol', ftString, 20, i);
        ftDate: Result.AddFieldDef('DateCol', ftDate, 0, i);
        else raise Exception.Create('Field type not expected in this test.');
      end;
    end;
    Result.CreateTable;
  end;
  Result.Open;
end;

{ Creates a worksheet with columns as defined by the TestColumns.
  ftUnknown will become an empty column. }
procedure TEmptyColumnsTest.CreateWorksheet(ATestIndex: Integer);
const
  NumRows = 10;
var
  r, c: Integer;
  s: String;
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
begin
  // Create test spreadsheet file
  workbook := TsWorkbook.Create;
  try
    // Create worksheet
    worksheet := workbook.AddWorkSheet(SHEET_NAME);
    // Write headers (= field names) and record values
    for c := 0 to Length(TTestData)-1 do
    begin
      case TestCases[ATestIndex][c].ColumnType of
        ftUnknown: ;
        ftInteger:
          begin
            worksheet.WriteText(0, c, 'IntCol');
            for r := 1 to NumRows do
              worksheet.WriteNumber(r, c, Random(100));
          end;
        ftFloat:
          begin
            worksheet.WriteText(0, c, 'FloatCol');
            for r := 1 to NumRows do
              worksheet.WriteNumber(r, c, Random*100);
          end;
        ftString:
          begin
            worksheet.WriteText(0, c, 'StringCol');
            for r := 1 to NumRows do
              worksheet.WriteText(r, c, char(ord('a') + random(26)));
          end;
        ftDate:
          begin
            worksheet.WriteText(0, c, 'DateCol');
            for r := 1 to NumRows do
              worksheet.WriteDateTime(r, c, EncodeDate(2000,1,1) + Random(1000), nfShortDate);
          end;
      end;
    end;

    // Save
    workbook.WriteToFile(DataFileName, true);
  finally
    workbook.Free;
  end;
end;

procedure TEmptyColumnsTest.TestFieldDefs(ATestIndex: Integer; AutoFieldDefs: Boolean);
var
  dataset: TsWorksheetDataset;
  c, i: Integer;
  expectedFieldDefIndex, actualFieldDefIndex: Integer;
begin
  CreateWorksheet(ATestIndex);
  dataset := CreateAndOpenDataset(ATestIndex, AutoFieldDefs);
  try
    for i := 0 to dataset.FieldDefs.Count-1 do
    begin
      c := TsFieldDef(dataset.FieldDefs[i]).ColIndex;
      expectedFieldDefIndex := TestCases[ATestIndex][c].FieldDefIndex;
      actualFieldDefIndex := i;
      CheckEquals(
        expectedFieldDefIndex,
        actualFieldDefIndex,
        'FieldDef index mismatch, fieldDef #' + IntToStr(i)
      );
    end;
  finally
    dataset.Free;
  end;
end;

procedure TEmptyColumnsTest.Test_0;
begin
  TestFieldDefs(0, false);
end;

procedure TEmptyColumnsTest.Test_1;
begin
  TestFieldDefs(1, false);
end;

procedure TEmptyColumnsTest.Test_2;
begin
  TestFieldDefs(2, false);
end;

procedure TEmptyColumnsTest.Test_3;
begin
  TestFieldDefs(3, false);
end;

procedure TEmptyColumnsTest.Test_4;
begin
  TestFieldDefs(4, false);
end;

procedure TEmptyColumnsTest.Test_5;
begin
  TestFieldDefs(5, false);
end;

procedure TEmptyColumnsTest.Test_6;
begin
  TestFieldDefs(6, false);
end;

procedure TEmptyColumnsTest.Test_0_AutoFieldDefs;
begin
  TestFieldDefs(0, true);
end;

procedure TEmptyColumnsTest.Test_1_AutoFieldDefs;
begin
  TestFieldDefs(1, true);
end;

procedure TEmptyColumnsTest.Test_2_AutoFieldDefs;
begin
  TestFieldDefs(2, true);
end;

procedure TEmptyColumnsTest.Test_3_AutoFieldDefs;
begin
  TestFieldDefs(3, true);
end;

procedure TEmptyColumnsTest.Test_4_AutoFieldDefs;
begin
  TestFieldDefs(4, true);
end;

procedure TEmptyColumnsTest.Test_5_AutoFieldDefs;
begin
  TestFieldDefs(5, true);
end;

procedure TEmptyColumnsTest.Test_6_AutoFieldDefs;
begin
  TestFieldDefs(6, true);
end;

procedure TEmptyColumnsTest.SetUp;
begin
  DataFileName := GetTempDir + FILE_NAME;
end;

procedure TEmptyColumnsTest.TearDown;
begin
  if FileExists(DataFileName) then DeleteFile(DataFileName);
end;

initialization

  RegisterTest(TEmptyColumnsTest);
end.

