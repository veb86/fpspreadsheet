unit SearchTestUnit;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpcunit, testutils, testregistry,
  DB,
  fpspreadsheet, fpsTypes, fpsDataset;

type

  TSearchTest = class(TTestCase)
  private
    function CreateAndOpenDataset: TsWorksheetDataset;
    procedure LocateTest(SearchInField: String; SearchValue: Variant;
      ExpectedRecNo: Integer; Options: TLocateOptions = []);
    procedure LookupTest(SearchInField: String; SearchValue: Variant;
      ResultFields: String; ExpectedValues: Variant);

  protected
    procedure SetUp; override;
    procedure TearDown; override;

  published
    procedure LocateTest_Int_Found;
    procedure LocateTest_Int_NotFound;
    procedure LocateTest_String_Found;
    procedure LocateTest_String_Found_CaseInsensitive;
    procedure LocateTest_String_NotFound;
    procedure LocateTest_NonASCIIString_Found;
    procedure LocateTest_NonASCIIString_Found_CaseInsensitive;
    procedure LocateTest_NonASCIIString_NotFound;
    procedure LocateTest_WideString_Found;
    procedure LocateTest_WideString_Found_CaseInsensitive;
    procedure LocateTest_WideString_NotFound;
    procedure LocateTest_NonASCIIWideString_Found;
    procedure LocateTest_NonASCIIWideString_Found_CaseInsensitive;
    procedure LocateTest_NonASCIIWideString_NotFound;

    procedure LookupTest_Int_Found;
    procedure LookupTest_Int_NotFound;
    procedure LookupTest_String_Found;
    procedure LookupTest_String_NotFound;
    procedure LookupTest_NonASCIIString_Found;
    procedure LookupTest_NonASCIIString_NotFound;
    procedure LookupTest_WideString_Found;
    procedure LookupTest_WideString_NotFound;
    procedure LookupTest_NonASCIIWideString_Found;
    procedure LookupTest_NonASCIIWideString_NotFound;

  end;

implementation

uses
  Variants, LazUTF8;

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

const
  NUM_ROWS = 5;
var
  INT_VALUES: array[1..NUM_ROWS] of Integer = (
    12, 20, -10, 83, 3
  );
  STRING_VALUES: array[1..NUM_ROWS] of String = (
    'abc', 'a', 'Hallo', 'ijk', 'äöüαβγ'
  );
  WIDESTRING_VALUES: array[1..NUM_ROWS] of String = (  // Strings are converted to wide at runtime
    'ABC', 'A', 'Test', 'ÄöüΓ', 'xyz'
  );

function TSearchTest.CreateAndOpenDataset: TsWorksheetDataset;
begin
  Result := TsWorksheetDataset.Create(nil);
  Result.FileName := DataFileName;
  Result.SheetName := SHEET_NAME;
  Result.AutoFieldDefs := false;
  Result.AddFieldDef(INT_FIELD, ftInteger);
  Result.AddFieldDef(STRING_FIELD, ftString, 20);
  Result.AddFieldDef(WIDESTRING_FIELD, ftWideString, 20);
  Result.Open;
end;

procedure TSearchTest.LocateTest(SearchInField: String; SearchValue: Variant;
  ExpectedRecNo: Integer; Options: TLocateOptions = []);
var
  dataset: TsWorksheetDataset;
  actualRecNo: Integer;
  found: Boolean;
  f: TField;
begin
  dataset := CreateAndOpenDataset;
  try
    found := dataset.Locate(SearchInField, SearchValue, options);

    if ExpectedRecNo = -1 then
      CheckEquals(
        false,
        found,
        'Record found unexpectedly.'
      )
    else
      CheckEquals(
        true,
        found,
        'Existing record not found.'
      );

    if found then
    begin
      actualRecNo := dataset.RecNo;
      CheckEquals(
        ExpectedRecNo,
        actualRecNo,
        'Mismatch of found RecNo.'
      );

      for f in dataset.Fields do
        case f.FieldName of
          INT_FIELD:
            CheckEquals(
              INT_VALUES[actualRecNo],
              f.AsInteger,
              'Value mismatch in integer field'
            );
          STRING_FIELD:
            CheckEquals(
              STRING_VALUES[actualRecNo],
              f.AsString,
              'Value mismatch in string field'
            );
          WIDESTRING_FIELD:
            CheckEquals(
              UTF8ToUTF16(WIDESTRING_VALUES[actualRecNo]),
              f.AsWideString,
              'Value mismatch in widestring field'
            );
        end;
    end;
  finally
    dataset.Free;
  end;
end;

procedure TSearchTest.LocateTest_Int_Found;
begin
  LocateTest(INT_FIELD, -10, 3);
end;

procedure TSearchTest.LocateTest_Int_NotFound;
begin
  LocateTest(INT_FIELD, 1000, -1);
end;

procedure TSearchTest.LocateTest_String_Found;
begin
  LocateTest(STRING_FIELD, 'a', 2);
end;

procedure TSearchTest.LocateTest_String_Found_CaseInsensitive;
begin
  LocateTest(STRING_FIELD, 'ABC', 1, [loCaseInsensitive]);
end;

procedure TSearchTest.LocateTest_String_NotFound;
begin
  LocateTest(STRING_FIELD, 'ttt', -1);
end;

procedure TSearchTest.LocateTest_NonASCIIString_Found;
begin
  LocateTest(STRING_FIELD, 'äöüαβγ', 5);
end;

procedure TSearchTest.LocateTest_NonASCIIString_Found_CaseInsensitive;
begin
  LocateTest(STRING_FIELD, 'ÄöÜαβΓ', 5, [loCaseInsensitive]);
end;

procedure TSearchTest.LocateTest_NonASCIIString_NotFound;
begin
  LocateTest(STRING_FIELD, 'ä', -1);
end;

procedure TSearchTest.LocateTest_WideString_Found;
begin
  LocateTest(WIDESTRING_FIELD, WideString('ABC'), 1);
end;

procedure TSearchTest.LocateTest_WideString_Found_CaseInsensitive;
begin
  LocateTest(WIDESTRING_FIELD, WideString('Abc'), 1, [loCaseInsensitive]);
end;

procedure TSearchTest.LocateTest_WideString_NotFound;
begin
  LocateTest(WIDESTRING_FIELD, WideString('abc'), -1);
end;

procedure TSearchTest.LocateTest_NonASCIIWideString_Found;
var
  ws: WideString;
begin
  ws := UTF8ToUTF16('ÄöüΓ');
  LocateTest(WIDESTRING_FIELD, ws, 4);
end;

procedure TSearchTest.LocateTest_NonASCIIWideString_Found_CaseInsensitive;
var
  ws: Widestring;
begin
  ws := UTF8ToUTF16('Äöüγ');
  LocateTest(WIDESTRING_FIELD, ws, 4, [loCaseInsensitive]);
end;

procedure TSearchTest.LocateTest_NonASCIIWideString_NotFound;
var
  ws: WideString;
begin
  ws := UTF8ToUTF16('ä-α');
  LocateTest(WIDESTRING_FIELD, ws, -1);
end;

// -----------------------------------------------------------------------------

procedure TSearchTest.LookupTest(SearchInField: String; SearchValue: Variant;
  ResultFields: String; ExpectedValues: Variant);
var
  dataset: TsWorksheetDataset;
  savedRecNo: Integer;
  i, j: Integer;
  actualValues: Variant;
  expectedInt, actualInt: Integer;
  expectedStr, actualStr: String;
  expectedWideStr, actualWideStr: WideString;
  L: TStringList;
begin
  dataset := CreateAndOpenDataset;
  try
    savedRecNo := dataset.RecNo;
    actualValues := dataset.Lookup(SearchInField, SearchValue, ResultFields);

    // The active record position must not be changed
    CheckEquals(
      savedRecNo,
      dataset.RecNo,
      'Lookup must not move the active record.'
    );

    // Compare count of elements in value arrays
    CheckEquals(
      VarArrayDimCount(ExpectedValues),
      VarArrayDimCount(actualValues),
      'Mismatch in found field values.'
    );

    if VarIsNull(ExpectedValues) then
    begin
      CheckEquals(
        true,
        varIsNull(actualValues),
        'Record found but not expected.'
      );
      exit;
    end;

    if not VarIsNull(ExpectedValues) then
      CheckEquals(
        false,
        varIsNull(actualValues),
        'Record expected but not found.'
      );

    L := TStringList.Create;
    L.StrictDelimiter := true;
    L.Delimiter := ';';
    L.DelimitedText := ResultFields;

    // Compare lookup values with expected values
    for i := 0 to dataset.Fields.Count-1 do
    begin
      j := L.IndexOf(dataset.Fields[i].FieldName);
      if j = -1 then
        continue;

      case dataset.Fields[i].DataType of
        ftInteger:
          begin
            expectedInt := ExpectedValues[j];
            actualInt := actualvalues[j];
            CheckEquals(
              expectedInt,
              actualInt,
              'Integer field lookup value mismatch'
            );
          end;
        ftString:
          begin
            expectedStr := VarToStr(ExpectedValues[j]);
            actualStr := VarToStr(actualValues[j]);
            CheckEquals(
              expectedStr,
              actualStr,
              'String field lookup value mismatch'
            );
          end;
        ftWideString:
          begin
            expectedWideStr := VarToWideStr(ExpectedValues[j]);
            actualWideStr := VarToWideStr(actualValues[j]);
            CheckEquals(
              ExpectedWideStr,
              actualWideStr,
              'Widestring field lookup value mismatch'
            );
          end;
        else
          raise Exception.Create('Unsupported field type in LookupTest');
      end;
    end;
    L.Free;
  finally
    dataset.Free;
  end;
end;

procedure TSearchTest.LookupTest_Int_Found;
var
  ws: wideString;
begin
  ws := UTF8ToUTF16(WIDESTRING_VALUES[2]);
  LookupTest(INT_FIELD, 20, STRING_FIELD+';'+WIDESTRING_FIELD, VarArrayOf(['a', ws]));
end;

procedure TSearchTest.LookupTest_Int_NotFound;
begin
  LookupTest(INT_FIELD, 200, STRING_FIELD+';'+WIDESTRING_FIELD, Null);
end;

procedure TSearchTest.LookupTest_String_Found;
var
  ws: wideString;
begin
  ws := UTF8ToUTF16(WIDESTRING_VALUES[3]);
  LookupTest(STRING_FIELD, 'Hallo', INT_FIELD+';'+WIDESTRING_FIELD, VarArrayOf([-10, ws]));
end;

procedure TSearchTest.LookupTest_String_NotFound;
begin
  LookupTest(STRING_FIELD, 'Halloooo', INT_FIELD+';'+WIDESTRING_FIELD, Null);
end;

procedure TSearchTest.LookupTest_NonASCIIString_Found;
var
  ws: wideString;
begin
  ws := UTF8ToUTF16('xyz');
  LookupTest(STRING_FIELD, 'äöüαβγ', INT_FIELD+';'+WIDESTRING_FIELD, VarArrayOf([3, ws]));
end;

procedure TSearchTest.LookupTest_NonASCIIString_NotFound;
begin
  LookupTest(STRING_FIELD, 'ÄÄÄÄα', INT_FIELD+';'+WIDESTRING_FIELD, Null);
end;

procedure TSearchTest.LookupTest_WideString_Found;
var
  ws: wideString;
begin
  ws := UTF8ToUTF16('ABC');
  LookupTest(WIDESTRING_FIELD, ws, INT_FIELD+';'+STRING_FIELD, VarArrayOf([12, 'abc']));
end;

procedure TSearchTest.LookupTest_WideString_NotFound;
var
  ws: wideString;
begin
  ws := UTF8ToUTF16('ABCD');
  LookupTest(WIDESTRING_FIELD, ws, INT_FIELD+';'+STRING_FIELD, null);
end;

procedure TSearchTest.LookupTest_NonASCIIWideString_Found;
var
  ws: wideString;
begin
  ws := UTF8ToUTF16('ÄöüΓ');
  LookupTest(WIDESTRING_FIELD, ws, INT_FIELD+';'+STRING_FIELD, VarArrayOf([83, 'ijk']));
end;

procedure TSearchTest.LookupTest_NonASCIIWideString_NotFound;
var
  ws: wideString;
begin
  ws := UTF8ToUTF16('Äöαβ');
  LookupTest(WIDESTRING_FIELD, ws, INT_FIELD+';'+STRING_FIELD, null);
end;

// -----------------------------------------------------------------------------

procedure TSearchTest.SetUp;
var
  r: Integer;
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
    for r := 1 to NUM_ROWS do
    begin
      worksheet.WriteNumber(r, INT_COL, INT_VALUES[r], nfFixed, 0);
      worksheet.WriteText(r, STRING_COL, STRING_VALUES[r]);
      worksheet.WriteText(r, WIDESTRING_COL, WIDESTRING_VALUES[r]);
    end;

    // Save
    DataFileName := GetTempDir + FILE_NAME;
    workbook.WriteToFile(DataFileName, true);
  finally
    workbook.Free;
  end;
end;

procedure TSearchTest.TearDown;
begin
  if FileExists(DataFileName) then DeleteFile(DataFileName);
end;

initialization
  RegisterTest(TSearchTest);

end.

