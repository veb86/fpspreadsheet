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
  Math, Variants,
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
      ACondition: TsCFCondition; AValue1, AValue2: Variant);
    procedure TestWriteRead_CF_Number(AFormat: TsSpreadsheetFormat;
      ACondition: TsCFCondition; AValue1: Variant);
    procedure TestWriteRead_CF_Number(AFormat: TsSpreadsheetFormat;
      ACondition: TsCFCondition);

  published
    procedure TestWriteRead_CF_Number_XLSX_Equal_Const;
    procedure TestWriteRead_CF_Number_XLSX_NotEqual_Const;
    procedure TestWriteRead_CF_Number_XLSX_GreaterThan_Const;
    procedure TestWriteRead_CF_Number_XLSX_LessThan_Const;
    procedure TestWriteRead_CF_Number_XLSX_GreaterEqual_Const;
    procedure TestWriteRead_CF_Number_XLSX_LessEqual_Const;
    procedure TestWriteRead_CF_Number_XLSX_Between_Const;
    procedure TestWriteRead_CF_Number_XLSX_NotBetween_Const;
    procedure TestWriteRead_CF_Number_XLSX_AboveAverage;
    procedure TestWriteRead_CF_Number_XLSX_BelowAverage;
    procedure TestWriteRead_CF_Number_XLSX_AboveEqualAverage;
    procedure TestWriteRead_CF_Number_XLSX_BelowEqualAverage;
    procedure TestWriteRead_CF_Number_XLSX_AboveAverage_2StdDev;
    procedure TestWriteRead_CF_Number_XLSX_BelowAverage_2StdDev;
    procedure TestWriteRead_CF_NUMBER_XLSX_Top3;
    procedure TestWriteRead_CF_NUMBER_XLSX_Top10Percent;
    procedure TestWriteRead_CF_NUMBER_XLSX_Bottom3;
    procedure TestWriteRead_CF_NUMBER_XLSX_Bottom10Percent;
    procedure TestWriteRead_CF_NUMBER_XLSX_BeginsWith;
    procedure TestWriteRead_CF_NUMBER_XLSX_EndsWith;
    procedure TestWriteRead_CF_NUMBER_XLSX_Contains;
    procedure TestWriteRead_CF_NUMBER_XLSX_NotContains;
    procedure TestWriteRead_CF_NUMBER_XLSX_Unique;
    procedure TestWriteRead_CF_NUMBER_XLSX_Duplicate;
    procedure TestWriteRead_CF_Number_XLSX_ContainsErrors;
    procedure TestWriteRead_CF_Number_XLSX_NotContainsErrors;
  end;

implementation

uses
  TypInfo;


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
  AFormat: TsSpreadsheetFormat; ACondition: TsCFCondition);
var
  dummy: variant;
begin
  VarClear(dummy);
  TestWriteRead_CF_NUMBER(AFormat, ACondition, dummy, dummy);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number(
  AFormat: TsSpreadsheetFormat; ACondition: TsCFCondition;
  AValue1: Variant);
var
  dummy: Variant;
begin
  VarClear(dummy);
  TestWriteRead_CF_NUMBER(AFormat, ACondition, AValue1, dummy);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number(
  AFormat: TsSpreadsheetFormat; ACondition: TsCFCondition;
  AValue1, AValue2: Variant);
const
  SHEET_NAME = 'CF';
  TEXTS: array[0..6] of String = ('abc', 'def', 'ghi', 'abc', 'jkl', 'akl', 'ab');
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
    workbook.Options := [boAutoCalc];
    workSheet:= workBook.AddWorksheet(SHEET_NAME);

    row := 0;
    for Col := 0 to High(TEXTS) do
      worksheet.WriteText(row, col, TEXTS[col]);

    row := 1;
    for col := 0 to 9 do
      worksheet.WriteNumber(row, col, col+1);
    worksheet.WriteFormula(row, col, '=1/0');

    // Write format used by the cells detected by conditional formatting
    InitFormatRecord(sollFmt);
    sollFmt.SetBackgroundColor(scRed);
    sollFmtIdx := workbook.AddCellFormat(sollFmt);

    // Write instruction for conditional formatting
    sollRange := Range(0, 0, 1, 10);
    if VarIsEmpty(AValue1) and VarIsEmpty(AValue2) then
      worksheet.WriteConditionalCellFormat(sollRange, ACondition, sollFmtIdx)
    else
    if VarIsEmpty(AValue2) then
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
    worksheet := GetWorksheetByName(workBook, SHEET_NAME);

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
    if not VarIsEmpty(AValue1) then
    begin
      if VarIsStr(AValue1) then
        CheckEquals(VarToStr(AValue1), VarToStr(actValue1), 'Conditional format parameter 1 mismatch')
      else if VarIsNumeric(AValue1) then
        CheckEquals(Double(AValue1), Double(actValue1), 'Conditional format parameter 1 mismatch')
      else
        raise Exception.Create('Unknown data type in variant');
    end else
      CheckEquals(true, VarIsEmpty(actValue1), 'Omitted parameter 1 detected.');

    // Check 2nd parameter
    actValue2 := TsCFCellRule(cf.Rules[0]).Operand2;
    if not (VarIsEmpty(AValue2) or VarIsNull(AValue2)) then
    begin
      if VarIsStr(AValue2) then
        CheckEquals(VarToStr(AValue2), VarToStr(actValue2), 'Conditional format parameter 2 mismatch')
      else
        CheckEquals(Double(AValue2), Double(actValue2),  'Conditional format parameter 2 mismatch');
    end else
      CheckEquals(true, VarIsEmpty(actValue2), 'Omitted parameter 2 detected.');

  finally
    workbook.Free;
    DeleteFile(tempFile);
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

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_AboveAverage;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcAboveAverage);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_BelowAverage;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcBelowAverage);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_AboveEqualAverage;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcAboveEqualAverage);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_BelowEqualAverage;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcBelowEqualAverage);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_AboveAverage_2StdDev;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcAboveAverage, 2.0);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_BelowAverage_2StdDev;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcBelowAverage, 2.0);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_NUMBER_XLSX_Top3;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcTop, 3);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_NUMBER_XLSX_Top10Percent;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcTopPercent, 10);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_NUMBER_XLSX_Bottom3;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcBottom, 3);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_NUMBER_XLSX_Bottom10Percent;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcBottomPercent, 10);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_NUMBER_XLSX_BeginsWith;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcBeginsWith, 'ab');
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_NUMBER_XLSX_EndsWith;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcEndsWith, 'kl');
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_NUMBER_XLSX_Contains;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcEndsWith, 'b');
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_NUMBER_XLSX_NotContains;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcEndsWith, 'b');
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_NUMBER_XLSX_Unique;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcUnique);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_NUMBER_XLSX_Duplicate;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcDuplicate);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_ContainsErrors;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcContainsErrors);
end;

procedure TSpreadWriteReadCFTests.TestWriteRead_CF_Number_XLSX_NotContainsErrors;
begin
  TestWriteRead_CF_Number(sfOOXML, cfcNotContainsErrors);
end;

initialization
  RegisterTest(TSpreadWriteReadCFTests);

end.

