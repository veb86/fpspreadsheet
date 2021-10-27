unit ReadFieldsTestUnit;

{$mode objfpc}{$H+}

{$IF FPC_FullVersion >= 30300}
  {$DEFINE TEST_BYTE_FIELD}
{$IFEND}

interface

uses
  Classes, SysUtils, fpcunit, {%H-}testutils, testregistry,
  DB,
  fpspreadsheet, fpstypes, fpsdataset;

type

  TReadFieldsTest= class(TTestCase)
  private
    function CreateAndOpenDataset(AutoFieldDefs: Boolean): TsWorksheetDataset;
    procedure ReadFieldTest(Col: Integer; FieldName: String; AutoFieldDefs: Boolean);
  protected
    procedure SetUp; override;
    procedure TearDown; override;
  published
    procedure ReadIntegerField;
    {$IFDEF TEST_BYTE_FIELD}
    procedure ReadByteField;
    {$IFEND}
    procedure ReadWordField;
    procedure ReadFloatField;
    procedure ReadCurrencyField;
    procedure ReadBCDField;
    procedure ReadFmtBCDField;
    procedure ReadStringField;
    procedure ReadMemoField;
    procedure ReadBoolField;
    procedure ReadDateField;
    procedure ReadTimeField;
    procedure ReadDateTimeField;

    procedure ReadIntegerField_AutoFieldDefs;
    procedure ReadByteField_AutoFieldDefs;
    procedure ReadWordField_AutoFieldDefs;
    procedure ReadFloatField_AutoFieldDefs;
    procedure ReadCurrencyField_AutoFieldDefs;
    procedure ReadStringField_AutoFieldDefs;
    procedure ReadMemoField_AutoFieldDefs;
    procedure ReadBoolField_AutoFieldDefs;
    procedure ReadDateField_AutoFieldDefs;
    procedure ReadTimeField_AutoFieldDefs;
    procedure ReadDateTimeField_AutoFieldDefs;

  end;

implementation

const
  FILE_NAME = 'testfile.xlsx';
  SHEET_NAME = 'Sheet';

  INT_COL = 0;
  BYTE_COL = 1;
  WORD_COL = 2;
  FLOAT_COL = 3;
  CURRENCY_COL = 4;
  BCD_COL = 5;
  FMTBCD_COL = 6;
  STRING_COL = 7;
  BOOL_COL = 8;
  DATE_COL = 9;
  TIME_COL = 10;
  DATETIME_COL = 11;
  MEMO_COL = 12;

  TestText: array[0..3] of string = (
    'Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua',
    'At vero eos et accusam et justo duo dolores et ea rebum. Stet clita kasd gubergren, no sea takimata sanctus est Lorem ipsum dolor sit amet',
    'Статья 1 Все люди рождаются свободными и равными в своем достоинстве и правах.',
    'ϰαὶ τότ'' ἐγὼ Κύϰλωπα προσηύδων ἄγχι παραστάς, '
  );


var
  DataFileName: String;

function TReadFieldsTest.CreateAndOpenDataset(AutoFieldDefs: Boolean): TsWorksheetDataset;
begin
  Result := TsWorksheetDataset.Create(nil);
  Result.AutoFieldDefs:= true;
  Result.FileName := DataFileName;
  Result.SheetName := SHEET_NAME;
  Result.AutoFieldDefs := AutoFieldDefs;
  if not AutoFieldDefs then
  begin
    Result.AddFieldDef('IntCol', ftInteger);
    {$IFDEF TEST_BYTE_FIELD}
    Result.AddFieldDef('ByteCol', ftByte);
    {$ELSE}
    Result.AddFieldDef('ByteCol', ftInteger);  // No ftByte in too old FPC
    {$ENDIF}
    Result.AddFieldDef('WordCol', ftWord);
    Result.AddFieldDef('FloatCol', ftFloat);
    Result.AddFieldDef('CurrencyCol', ftCurrency);
    Result.AddFieldDef('BCDCol', ftBCD);
    Result.AddFieldDef('FmtBCDCol', ftFmtBCD);
    Result.AddFieldDef('StringCol', ftString, 30);
    Result.AddFieldDef('BoolCol', ftBoolean);
    Result.AddFieldDef('DateCol', ftDate);
    Result.AddFieldDef('TimeCol', ftTime);
    Result.AddFieldDef('DateTimeCol', ftDateTime);
    Result.AddFieldDef('MemoCol', ftMemo);
    Result.CreateTable;
  end;
  Result.Open;
end;

procedure TReadFieldsTest.SetUp;
const
  NumRows = 10;
var
  r: Integer;
  s: String;
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
begin
  // Create test spreadsheet file
  workbook := TsWorkbook.Create;
  try
    // Create worksheet
    worksheet := workbook.AddWorkSheet(SHEET_NAME);
    // Write headers (= field names)
    worksheet.WriteText(0, INT_COL, 'IntCol');
    worksheet.WriteText(0, BYTE_COL, 'ByteCol');
    worksheet.WriteText(0, WORD_COL, 'WordCol');
    worksheet.WriteText(0, FLOAT_COL, 'FloatCol');
    worksheet.WriteText(0, CURRENCY_COL, 'CurrencyCol');
    worksheet.WriteText(0, BCD_COL, 'BCDCol');
    worksheet.WriteText(0, FMTBCD_Col, 'FmtBCDCol');
    worksheet.WriteText(0, STRING_COL, 'StringCol');
    worksheet.WriteText(0, BOOL_COL, 'BoolCol');
    worksheet.WriteText(0, DATE_COL, 'DateCol');
    worksheet.WriteText(0, TIME_COL, 'TimeCol');
    worksheet.WriteText(0, DATETIME_COL, 'DateTimeCol');
    worksheet.Writetext(0, MEMO_COL, 'MemoCol');
    for r := 1 to NumRows do begin
      // Write values to IntCol
      worksheet.WriteNumber(r, INT_COL, r*120- 50, nfFixed, 0);
      // Write values to ByteCol
      worksheet.WriteNumber(r, BYTE_COL, r*2, nfFixed, 0);
      //Write values to WordCol
      worksheet.WriteNumber(r, WORD_COL, r*3, nfFixed, 0);
      // Write values to FloatCol
      worksheet.WriteNumber(r, FLOAT_COL, r*1.1-5.1, nfFixed, 2);
      // Write values to CurrencyCol
      worksheet.WriteCurrency(r, CURRENCY_COL, r*1000, nfCurrency);
      // Write values to BCDcol
      worksheet.WriteNumber(r, BCD_COL, r*1.2-3);
      // Write values to FmtBCDCol
      worksheet.WriteNumber(r, FMTBCD_COL, r*12.3-60);
      // Write values to StringCol
      case r of
        1: s := 'Статья';
        2: s := 'Λορεμ ιπσθμ δολορ σιτ αμετ';
        else s := char(ord('A') + r-1) + char(ord('b') + r-1) + char(ord('c') + r-1);
      end;
      worksheet.WriteText(r, STRING_COL, s);
      // Write values to BoolCol
      worksheet.WriteBoolValue(r, BOOL_COL, odd(r));
      // Write values to DateCol
      worksheet.WriteDateTime(r, DATE_COL, EncodeDate(2021, 8, 1) + r-1, nfShortDate);
      // Write values to TimeCol
      worksheet.WriteDateTime(r, TIME_COL, EncodeTime(8, 0, 0, 0) + (r-1) / (24*60), nfShortTime);
      // Write value to DateTimeCol
      worksheet.WriteDateTime(r, DATETIME_COL, EncodeDate(2021, 8, 1) + EncodeTime(8, 0, 0, 0) + (r-1) + (r-1)/24, nfShortDateTime);
      // Write value to MemoCol
      worksheet.WriteText(r, MEMO_COL, TestText[r mod Length(TestText)]);
    end;

    // Save
    DataFileName := GetTempDir + FILE_NAME;
    workbook.WriteToFile(DataFileName, true);
  finally
    workbook.Free;
  end;
end;

procedure TReadFieldsTest.TearDown;
begin
  if FileExists(DataFileName) then DeleteFile(DataFileName);
end;

procedure TReadFieldsTest.ReadFieldTest(Col: Integer; FieldName: String;
  AutoFieldDefs: Boolean);
const
  FLOAT_EPS = 1E-9;
var
  dataset: TDataset;
  row: Integer;
  f: TField;
  dt: TDateTime;
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
begin
  dataset := CreateAndOpenDataset(AutoFieldDefs);
  try
    workbook := TsWorkbook.Create;
    try
      workbook.ReadFromFile(DataFileName);
      worksheet := workbook.GetFirstWorksheet;

      f := dataset.FieldByName(FieldName);

      CheckEquals(
        worksheet.ReadAsText(0, col),
        f.FieldName,
        'Column header / FieldName mismatch'
      );

      CheckEquals(
        col,
        f.FieldNo-1,
        'Field number mismatch'
      );

      CheckEquals(
        worksheet.GetLastRowIndex(true),
        dataset.RecordCount,
        'Row count / record count mismatch'
      );

      dataset.First;
      row := 1;
      while not dataset.EoF do
      begin
        if (f.DataType in [ftString, ftWideString, ftMemo]) then
          CheckEquals(
            worksheet.ReadAsText(row, col),
            f.AsString,
            'Text mismatch in row ' + IntToStr(row)
          )
        else
        if (f.DataType in [
          ftInteger, {$IFDEF TEST_BYTE_FIELD}ftByte, {$ENDIF}
          ftWord, ftSmallInt, ftLargeInt])
        then
          CheckEquals(
            round(worksheet.ReadAsNumber(row, col)),
            f.AsInteger,
            'Integer value mismatch in row ' + IntToStr(row)
          )
        else if (f.DataType in [ftFloat, ftCurrency, ftBCD, ftFmtBCD]) then
          CheckEquals(
            worksheet.ReadAsNumber(row, col),
            f.AsFloat,
            FLOAT_EPS,
            'Float value mismatch in row ' + IntToStr(row)
          )
        else if (f.DataType = ftDate) then
        begin
          CheckEquals(
            true,
            worksheet.ReadAsDateTime(row, col, dt),
            'Invalid date in row ' + IntToStr(row)
          );
          CheckEquals(
            dt,
            f.AsDateTime,
            FLOAT_EPS,
            'Date value mismatch in row ' + IntToStr(row)
          )
        end
        else if (f.DataType = ftTime) then
        begin
          CheckEquals(
            true,
            worksheet.ReadAsDateTime(row, col, dt),
            'Invalid time in row ' + IntToStr(row)
          );
          CheckEquals(
            dt,
            f.AsDateTime,
            FLOAT_EPS,
            'Time value mismatch in row ' + IntToStr(row)
          )
        end
        else if (f.DataType = ftDateTime) then
        begin
          CheckEquals(
            true,
            worksheet.ReadAsDateTime(row, col, dt),
            'Invalid date/time in row ' + IntToStr(row)
          );
          CheckEquals(
            dt,
            f.AsDateTime,
            FLOAT_EPS,
            'Date/time value mismatch in row ' + IntToStr(row)
          );
        end;
        inc(row);
        dataset.Next;
      end;
    finally
      workbook.Free;
    end;
  finally
    dataset.Free;
  end;
end;


procedure TReadFieldsTest.ReadIntegerField_AutoFieldDefs;
begin
  ReadFieldTest(INT_COL, 'IntCol', true);
end;

procedure TReadFieldsTest.ReadByteField_AutoFieldDefs;
begin
  ReadFieldTest(BYTE_COL, 'ByteCol', true);
end;

procedure TReadFieldsTest.ReadWordField_AutoFieldDefs;
begin
  ReadFieldTest(WORD_COL, 'WordCol', true);
end;

procedure TReadFieldsTest.ReadFloatField_AutoFieldDefs;
begin
  ReadFieldTest(FLOAT_COL, 'FloatCol', true);
end;

procedure TReadFieldsTest.ReadCurrencyField_AutoFieldDefs;
begin
  ReadFieldTest(CURRENCY_COL, 'CurrencyCol', true);
end;

procedure TReadFieldsTest.ReadStringField_AutoFieldDefs;
begin
  ReadFieldTest(STRING_COL, 'StringCol', true);
end;

procedure TReadFieldsTest.ReadMemoField_AutoFieldDefs;
begin
  ReadFieldTest(MEMO_COL, 'MemoCol', true);
end;

procedure TReadFieldsTest.ReadBoolField_AutoFieldDefs;
begin
  ReadFieldTest(BOOL_COL, 'BoolCol', true);
end;

procedure TReadFieldsTest.ReadDateField_AutoFieldDefs;
begin
  ReadFieldTest(DATE_COL, 'DateCol', true);
end;

procedure TReadFieldsTest.ReadTimeField_AutoFieldDefs;
begin
  ReadFieldTest(TIME_COL, 'TimeCol', true);
end;

procedure TReadFieldsTest.ReadDateTimeField_AutoFieldDefs;
begin
  ReadFieldTest(DATETIME_COL, 'DateTimeCol', true);
end;


procedure TReadFieldsTest.ReadIntegerField;
begin
  ReadFieldTest(INT_COL, 'IntCol', false);
end;

{$IFDEF TEST_BYTE_FIELD}
procedure TReadFieldsTest.ReadByteField;
begin
  ReadFieldTest(BYTE_COL, 'ByteCol', false);
end;
{$ENDIF}

procedure TReadFieldsTest.ReadWordField;
begin
  ReadFieldTest(WORD_COL, 'WordCol', false);
end;

procedure TReadFieldsTest.ReadFloatField;
begin
  ReadFieldTest(FLOAT_COL, 'FloatCol', false);
end;

procedure TReadFieldsTest.ReadCurrencyField;
begin
  ReadFieldTest(CURRENCY_COL, 'CurrencyCol', false);
end;

procedure TReadFieldsTest.ReadBCDField;
begin
  ReadFieldTest(BCD_COL, 'BCDCol', false);
end;

procedure TReadFieldsTest.ReadFmtBCDField;
begin
  ReadFieldTest(FMTBCD_COL, 'FmtBCDCol', false);
end;

procedure TReadFieldsTest.ReadStringField;
begin
  ReadFieldTest(STRING_COL, 'StringCol', false);
end;

procedure TReadFieldsTest.ReadMemoField;
begin
  ReadFieldTest(MEMO_COL, 'MemoCol', false);
end;

procedure TReadFieldsTest.ReadBoolField;
begin
  ReadFieldTest(BOOL_COL, 'BoolCol', false);
end;

procedure TReadFieldsTest.ReadDateField;
begin
  ReadFieldTest(DATE_COL, 'DateCol', false);
end;

procedure TReadFieldsTest.ReadTimeField;
begin
  ReadFieldTest(TIME_COL, 'TimeCol', false);
end;

procedure TReadFieldsTest.ReadDateTimeField;
begin
  ReadFieldTest(DATETIME_COL, 'DateTimeCol', false);
end;


initialization
  RegisterTest(TReadFieldsTest);

end.

