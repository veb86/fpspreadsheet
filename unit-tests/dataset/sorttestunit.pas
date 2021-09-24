unit SortTestUnit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpcunit, testregistry,
  DB,
  fpspreadsheet, fpstypes, fpsutils, fpsdataset;

type

  TSortTest= class(TTestCase)
  private
    function CreateAndOpenDataset: TsWorksheetDataset;
  protected
    procedure SetUp; override;
    procedure TearDown; override;
    procedure SortTest(SortField: String; Descending, CaseInsensitive: Boolean);
  published
    procedure SortTest_IntField_Ascending;
    procedure SortTest_IntField_Descending;
    procedure SortTest_TextField_Ascending_CaseSensitive;
    procedure SortTest_TextField_Descending_CaseSensitive;
    procedure SortTest_TextField_Ascending_CaseInsensitive;
    procedure SortTest_TextField_Descending_CaseInsensitive;
  end;

implementation

const
  FILE_NAME = 'testfile.xlsx';
  SHEET_NAME = 'Sheet';
  INT_COL = 0;
  TEXT_COL = 1;
  INT_FIELD = 'IntCol';
  TEXT_FIELD = 'TextCol';

var
  DataFileName: String;

type
  TTestRow = record
    IntValue: Integer;
    TextValue: String;
  end;

const
  // Unsorted test values
  UNSORTED: array[0..4] of TTestRow = (       // Index
    (IntValue: 10; TextValue: 'abc'),         // 0
    (IntValue:  1; TextValue: 'ABC'),         // 1
    (IntValue:  1; TextValue: 'a'),           // 2
    (IntValue:  2; TextValue: 'A'),           // 3
    (IntValue: -1; TextValue: 'xyz')          // 4
  );

  // These are the indexes into the UNSORTED array after sorting
  SORTED_BY_INT_ASCENDING: array[0..4] of Integer = (4, 1, 2, 3, 0);
  SORTED_BY_INT_DESCENDING: array[0..4] of Integer = (0, 3, 2, 1, 4);
  SORTED_BY_TEXT_ASCENDING_CASESENS: array[0..4] of Integer = (2, 0, 3, 1, 4);
  SORTED_BY_TEXT_DESCENDING_CASESENS: array[0..4] of Integer = (4, 1, 3, 0, 2);
  SORTED_BY_TEXT_ASCENDING_CASEINSENS: array[0..4] of Integer = (3, 2, 1, 0, 4);
  SORTED_BY_TEXT_DESCENDING_CASEINSENS: array[0..4] of Integer = (4, 1, 0, 3, 2);
  // Note on case-insensitive sorting: Depending on implementation of the
  // sorting algorithms different results can be obtained for which the
  // uppercased texts are the same. Therefore, Excel yields different result
  // than FPSpreadsheet. Above indices are for FPSpreadsheet.


function TSortTest.CreateAndOpenDataset: TsWorksheetDataset;
begin
  Result := TsWorksheetDataset.Create(nil);
  Result.FileName := DataFileName;
  Result.SheetName := SHEET_NAME;
  Result.Open;
end;

procedure TSortTest.SortTest(SortField: String; Descending, CaseInsensitive: Boolean);
var
  dataset: TsWorksheetDataset;
  options: TsSortOptions;
  intField: TField;
  textField: TField;
  actualInt: Integer;
  actualText: String;
  expectedInt: Integer;
  expectedText: String;
  i, sortedIdx: Integer;
begin
  options := [];
  if Descending then Include(options, ssoDescending);
  if CaseInsensitive then Include(options, ssoCaseInsensitive);

  dataset := CreateAndOpenDataset;
  try
    dataset.SortOnField(SortField, options);

    // For debugging
    dataset.Close;  // to write the worksheet to file
    dataset.Open;

    intField := dataset.FieldByName(INT_FIELD);
    textField := dataset.FieldByName(TEXT_FIELD);

    dataset.First;
    i := 0;
    while not dataset.EOF do
    begin
      if SortField = INT_FIELD then
      begin
        if Descending then
          sortedIdx := SORTED_BY_INT_DESCENDING[i]
        else
          sortedIdx := SORTED_BY_INT_ASCENDING[i];
      end else
      if SortField = TEXT_FIELD then
      begin
        if Descending then
        begin
          if CaseInsensitive then
            sortedIdx := SORTED_BY_TEXT_DESCENDING_CASEINSENS[i]
          else
            sortedIdx := SORTED_BY_TEXT_DESCENDING_CASESENS[i];
        end else
        begin
          if CaseInsensitive then
            sortedIdx := SORTED_BY_TEXT_ASCENDING_CASEINSENS[i]
          else
            sortedIdx := SORTED_BY_TEXT_ASCENDING_CASESENS[i];
        end;
      end;

      expectedInt := UNSORTED[sortedIdx].IntValue;
      expectedText := UNSORTED[sortedIdx].TextValue;
      actualInt := intField.AsInteger;
      actualText := textField.AsString;

      CheckEquals(
        expectedInt,
        actualInt,
        'Integer field value mismatch in row ' + IntToStr(i)
      );
      CheckEquals(
        expectedText,
        actualText,
        'Text field value mismatch in row ' + IntToStr(i)
      );

      inc(i);
      dataset.Next;
    end;

  finally
    dataset.Free;
  end;
end;

procedure TSortTest.SortTest_IntField_Ascending;
begin
  SortTest(INT_FIELD, false, false);
end;

procedure TSortTest.SortTest_IntField_Descending;
begin
  SortTest(INT_FIELD, true, false);
end;

procedure TSortTest.SortTest_TextField_Ascending_CaseSensitive;
begin
  SortTest(TEXT_FIELD, false, false);
end;

procedure TSortTest.SortTest_TextField_Descending_CaseSensitive;
begin
  SortTest(TEXT_FIELD, true, false);
end;

procedure TSortTest.SortTest_TextField_Ascending_CaseInsensitive;
begin
  SortTest(TEXT_FIELD, false, true);
end;
procedure TSortTest.SortTest_TextField_Descending_CaseInsensitive;
begin
  SortTest(TEXT_FIELD, true, true);
end;

procedure TSortTest.SetUp;
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
    worksheet.WriteText(0, TEXT_COL, TEXT_FIELD);

    // Write values
    for i := Low(UNSORTED) to High(UNSORTED) do
    begin
      r := 1 + (i - Low(UNSORTED));
      worksheet.WriteNumber(r, INT_COL, UNSORTED[i].IntValue, nfFixed, 0);
      worksheet.WriteText(r, TEXT_COL, UNSORTED[i].TextValue);
    end;

    // Save
    DataFileName := GetTempDir + FILE_NAME;
    workbook.WriteToFile(DataFileName, true);
  finally
    workbook.Free;
  end;
end;

procedure TSortTest.TearDown;
begin
  if FileExists(DataFileName) then DeleteFile(DataFileName);
end;


initialization
  RegisterTest(TSortTest);

end.

