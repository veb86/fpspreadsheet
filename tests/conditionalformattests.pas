{ Tests for conditional formatting
  These unit tests write out to and read back from files.
}

unit conditionalformattests;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry, testsutility,
  fpsTypes, fpsUtils, fpsAllFormats, fpSpreadsheet, fpsConditionalFormat;

type
  { TSpreadWriteReadCFTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadCFTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;

    // Test conditional cell format
    procedure TestWriteRead_CF_Number(AFormat: TsSpreadsheetFormat;
      ACondition: TsCFCondition; AValue1: Integer = MaxInt; AValue2: Integer = MaxInt);

  published
    procedure TestWriteRead_CF_Number_XLSX_Equal_Const;
    procedure TestWriteRead_CF_Number_XLSX_NotEqual_Const;
    procedure TestWriteRead_CF_Number_XLSX_GreaterThan_Const;
    procedure TestWriteRead_CF_Number_XLSX_LessThan_Const;
    procedure TestWriteRead_CF_Number_XLSX_GreaterEqual_Const;
    procedure TestWriteRead_CF_Number_XLSX_LessEqual_Const;
    procedure TestWriteRead_CF_Number_XLSX_Between_Const;
    procedure TestWriteRead_CF_Number_XLSX_NotBetween_Const;
  end;

implementation

uses
  Math, TypInfo;


{ TSpreadWriteReadCFTests }

procedure TSpreadWriteReadCFTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadCFTests.TearDown;
begin
  inherited TearDown;
end;


{ CFCellFormat tests. Detected cells get a red background. }

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number(
  AFormat: TsSpreadsheetFormat; ACondition: TsCFCondition;
  AValue1: Integer = MaxInt; AValue2: Integer = MaxInt);
var
  worksheet: TsWorksheet;
  workbook: TsWorkbook;
  row, col: Cardinal;
  tempFile: string;
  sollFMT: TsCellFormat;
  sollFmtIdx: Integer;
  sollRange: TsCellRange;
  actFMT: TsCellFormat;
  actFmtIdx: Integer;
  actRange: TsCellRange;
  actCondition: TsCFCondition;
  actValue1, actValue2: Variant;
  cf: TsConditionalFormat;
begin
  // Write out all test values
  workbook := TsWorkbook.Create;
  try
    workSheet:= workBook.AddWorksheet('CF');

    // Write test data: two rows with numbers 1..10
    for Row := 0 to 1 do
      for Col := 0 to 9 do
        worksheet.WriteNumber(row, col, col+1);

    // Write format used by the cells detected by conditional formatting
    InitFormatRecord(sollFmt);
    sollFmt.SetBackgroundColor(scRed);
    sollFmtIdx := workbook.AddCellFormat(sollFmt);

    // Write instruction for conditional formatting
    sollRange := Range(0, 0, 0, 8);
    if (AValue1 = MaxInt) and (AValue2 = MaxInt) then
      worksheet.WriteConditionalCellFormat(sollRange, ACondition, sollFmtIdx)
    else
    if (AValue2 = MaxInt) then
      worksheet.WriteConditionalCellFormat(sollRange, ACondition, AValue1, sollFmtIdx)
    else
      worksheet.WriteConditionalCellFormat(sollRange, ACondition, AValue1, AValue2, sollFmtIdx);

    // Save to file
    tempFile := NewTempFile;
    workBook.WriteToFile(tempFile, AFormat, true);
  finally
    workbook.Free;
  end;

  // Open the spreadsheet
  workbook := TsWorkbook.Create;
  try
    workbook.ReadFromFile(TempFile, AFormat);
    worksheet := GetWorksheetByName(workBook, 'CF');

    if worksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    // Check count of conditional formats
    CheckEquals(1, workbook.GetNumConditionalFormats, 'ConditionalFormat count mismatch.');

    // Read conditional format
    cf := Workbook.GetConditionalFormat(0);

    //Check range
    actRange := cf.CellRange;
    CheckEquals(sollRange.Row1, actRange.Row1, 'Conditional format range mismatch (Row1)');
    checkEquals(sollRange.Col1, actRange.Col1, 'Conditional format range mismatch (Col1)');
    CheckEquals(sollRange.Row2, actRange.Row2, 'Conditional format range mismatch (Row2)');
    checkEquals(sollRange.Col2, actRange.Col2, 'Conditional format range mismatch (Col2)');

    // Check rules count
    CheckEquals(1, cf.RulesCount, 'Conditional format rules count mismatch');

    // Check rules class
    CheckEquals(TsCFCellRule, cf.Rules[0].ClassType, 'Conditional format rule class mismatch');

    // Check condition
    actCondition := TsCFCellRule(cf.Rules[0]).Condition;
    CheckEquals(
      GetEnumName(TypeInfo(TsCFCondition), integer(ACondition)),
      GetEnumName(typeInfo(TsCFCondition), integer(actCondition)),
      'Conditional format condition mismatch.'
    );

    // Check 1st parameter
    actValue1 := TsCFCellRule(cf.Rules[0]).Operand1;
    if AValue1 <> MaxInt then
      CheckEquals(AValue1, Integer(actValue1), 'Conditional format parameter 1 mismatch')
    else
      CheckEquals(Integer(varEmpty), Integer(actValue1), 'Omitted parameter 1 detected.');

    // Check 2nd parameter
    actValue2 := TsCFCellRule(cf.Rules[0]).Operand2;
    if AValue2 <> MaxInt then
      CheckEquals(AValue2, actValue2, 'Conditional format parameter 2 mismatch')
    else
      CheckEquals(Integer(varEmpty), Integer(actValue2), 'Omitted parameter 2 detected.');

  finally
    workbook.Free;
  end;
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_Equal_Const;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcEqual, 5);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_NotEqual_Const;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcNotEqual, 5);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_GreaterThan_Const;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcGreaterThan, 5);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_LessThan_Const;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcLessThan, 5);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_GreaterEqual_Const;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcGreaterEqual, 5);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_LessEqual_Const;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcLessEqual, 5);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_Between_Const;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcBetween, 3, 7);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_NotBetween_Const;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcNotBetween, 3, 7);
end;

initialization
  RegisterTest(TSpreadWriteReadCFTests);

end.

