unit CopyFromDatasetUnit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpcunit, testutils, testregistry,
  DB, dbf,
  fpspreadsheet, fpsDataset;

type

  { TCopyFromDatasetTest }

  TCopyFromDatasetTest= class(TTestCase)
  private
    function CreateDbf: TDbf;
    procedure CopyDatasetTest(ATestIndex: Integer);
  protected
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure CopyDatasetTest_FieldDefs;
    procedure CopyDatasetTest_Fields;
    procedure CopyDatasetTest_Records;
  end;

implementation

uses
  TypInfo;

const
  DBF_FILE_NAME = 'testdata.dbf';
  FILE_NAME = 'testfile.xlsx';
  SHEET_NAME = 'Sheet';

  STRING_FIELD = 'StringCol';
  INT_FIELD = 'IntegerCol';
  FLOAT_FIELD = 'FloatCol';

  NUM_RECORDS = 10;

var
  DataFileName: String;
  DbfPath: String;


function TCopyFromDatasetTest.CreateDbf: TDbf;
var
  i: Integer;
begin
  Result := TDbf.Create(nil);
  Result.FilePathFull := DbfPath;
  Result.TableName := DBF_FILE_NAME;
  Result.FieldDefs.Add(STRING_FIELD, ftString, 20);
  Result.FieldDefs.Add(INT_FIELD, ftInteger);
  Result.FieldDefs.Add(FLOAT_FIELD, ftFloat);
  Result.CreateTable;
  Result.Open;
  for i := 1 to NUM_RECORDS do
  begin
    Result.Append;
    Result.FieldByName(STRING_FIELD).AsString := 'abc' + IntToStr(i);
    Result.FieldByName(INT_FIELD).AsInteger := -5 + i;
    Result.FieldByName(FLOAT_FIELD).AsFloat := -5.1 * (i + 5.1);
    Result.Post;
  end;
end;

procedure TCopyFromDatasetTest.CopyDatasetTest(ATestIndex: Integer);
const
  DEBUG = false;
var
  dbf: TDbf;
  dataset: TsWorksheetDataset;
  i: Integer;
begin
  dbf := CreateDbf;

  if DEBUG then
  begin
    dbf.Close;
    dbf.Open;
  end;

  dataset := TsWorksheetDataset.Create(nil);
  try
    dataset.CopyFromDataset(dbf, DataFileName, dbf.TableName);

    // Save for debugging
    if DEBUG then
    begin
      dataset.Close;
      dataset.Open;
    end;

    case ATestIndex of
      // FIELD DEFS
      0: begin
           CheckEquals(     // Compare FieldDef count
             dbf.FieldDefs.Count,
             dataset.FieldDefs.Count,
             'Mismatch in number of FieldDefs'
           );

           // Compare FieldDefs
           for i := 0 to dbf.FieldDefs.Count-1 do
           begin
             CheckEquals(
               dbf.FieldDefs[i].Name,
               dataset.FieldDefs[i].Name,
               'Mismatch in FieldDefs[' + IntToStr(i) + '].Name'
             );
             CheckEquals(
               GetEnumName(TypeInfo(TFieldType), integer(dbf.FieldDefs[i].DataType)),
               GetEnumName(TypeInfo(TFieldType), integer(dataset.FieldDefs[i].DataType)),
               'Mismatch in FieldDefs[' + IntToStr(i) + '].DataType'
             );
             CheckEquals(
               dbf.FieldDefs[i].Size,
               dataset.FieldDefs[i].Size,
               'Mismatch in FieldDefs[' + IntToStr(i) + '].Size'
             );
           end;
         end;

      // FIELDS
      1: begin
           // Compare field count
           CheckEquals(
             dbf.FieldCount,
             dataset.FieldCount,
             'Mismatch in FieldCount'
           );

           // Compare fields
           for i := 0 to dbf.FieldCount-1 do
           begin
             CheckEquals(
               dbf.Fields[i].FieldName,
               dataset.Fields[i].FieldName,
               'Mismatch in Fields[' + IntToStr(i) + '].FieldName'
             );
             CheckEquals(
               GetEnumName(TypeInfo(TFieldType), integer(dbf.Fields[i].DataType)),
               GetEnumName(TypeInfo(TFieldType), integer(dataset.Fields[i].DataType)),
               'Mismatch in Fields[' + IntToStr(i) + '].DataType'
             );
           end;
         end;

      // RECORDS
      2: begin
           // Compare record count
           CheckEquals(
             dbf.RecordCount,
             dataset.RecordCount,
             'Mismatch in RecordCount'
           );

           dbf.First;
           dataset.First;
           while not dbf.EoF do
           begin
             for i := 0 to dbf.FieldCount-1 do
             begin
               CheckEquals(
                 dbf.Fields[i].AsString,
                 dataset.Fields[i].AsString,
                 'Record value mismatch, Field #[' + IntToStr(i) + '], RecNo ' + IntToStr(dbf.RecNo)
               );
             end;
             dbf.Next;
             dataset.Next;
           end;
         end;
    end;

  finally
    dataset.Free;
    dbf.Free;
  end;
end;

procedure TCopyFromDatasetTest.CopyDatasetTest_FieldDefs;
begin
  CopyDatasetTest(0);
end;

procedure TCopyFromDatasetTest.CopyDatasetTest_Fields;
begin
  CopyDatasetTest(1);
end;

procedure TCopyFromDatasetTest.CopyDatasetTest_Records;
begin
  CopyDatasetTest(2);
end;

procedure TCopyFromDatasetTest.SetUp;
begin
  DataFileName := GetTempDir + FILE_NAME;
  DbfPath := GetTempDir;
end;

procedure TCopyFromDatasetTest.TearDown;
begin
  if FileExists(DataFileName) then DeleteFile(DataFileName);
  if FileExists(DbfPath + DBF_FILE_NAME) then DeleteFile(DbfPath + DBF_FILE_NAME);
end;


initialization
  RegisterTest(TCopyFromDatasetTest);

end.

