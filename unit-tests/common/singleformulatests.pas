unit SingleFormulaTests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry, testsutility,
  fpstypes, fpspreadsheet, fpsexprparser,
  xlsbiff8;

type
  TFormulaTestKind = (ftkConstants, ftkCellConstant, ftkCells, ftkCellRange,
    ftkCellRangeSheet, ftkCellRangeSheetRange,
    ftkSortedNumbersASC, ftkSortedNumbersDESC);
  TWorksheetTestKind = (wtkRenameWorksheet, wtkDeleteWorksheet);

  { TSpreadDetailedFormulaFormula }
  TSpreadSingleFormulaTests = class(TTestCase)
  private
  protected
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestFormula(AFormula: String; AExpected: String;
      ATestKind: TFormulaTestKind; AFormat: TsSpreadsheetFormat;
      AExpectedFormula: String = '');
    procedure TestWorksheet(ATestKind: TWorksheetTestKind; ATestCase: Integer);
    procedure TestFormulaErrors(ATest: Integer);
    procedure TestInsDelRowCol(ATestIndex: Integer);
    procedure TestCountSumAvgIFS_Microsoft(AFormat: TsSpreadsheetFormat; AFormulaType: Integer);

  published
    procedure AddConst_BIFF2;
    procedure AddConst_BIFF5;
    procedure AddConst_BIFF8;
    procedure AddConst_OOXML;
    procedure AddConst_XML;
    procedure AddConst_ODS;

    procedure AddCells_BIFF2;
    procedure AddCells_BIFF5;
    procedure AddCells_BIFF8;
    procedure AddCells_OOXML;
    procedure AddCells_XML;
    procedure AddCells_ODS;

    procedure RoundConst1_ODS;
    procedure RoundConst2_ODS;
    procedure RoundCell1_ODS;
    procedure RoundCell2_ODS;

    procedure YearConst_BIFF8;
    procedure YearCell_BIFF8;
    procedure MonthConst_BIFF8;
    procedure MonthCell_BIFF8;
    procedure DayConst_BIFF8;
    procedure DayCell_BIFF8;

    procedure HourConst_BIFF8;
    procedure HourCell_BIFF8;
    procedure MinuteConst_BIFF8;
    procedure MinuteCell_BIFF8;
    procedure SecondConst_BIFF8;
    procedure SecondCell_BIFF8;

    procedure SumRange_BIFF2;
    procedure SumRange_BIFF5;
    procedure SumRange_BIFF8;
    procedure SumRange_OOXML;
    procedure SumRange_XML;
    procedure SumRange_ODS;

    procedure SumSheetRange_BIFF5;  // no 3d ranges for BIFF2
    procedure SumSheetRange_BIFF8;
    procedure SumSheetRange_OOXML;
    procedure SumSheetRange_XML;
    procedure SumSheetRange_ODS;

    procedure SumMultiSheetRange_BIFF5;
    procedure SumMultiSheetRange_BIFF8;
    procedure SumMultiSheetRange_OOXML;
    procedure SumMultiSheetRange_XML;
    procedure SumMultiSheetRange_ODS;

    procedure SumMultiSheetRange_FlippedCells_BIFF8;
    procedure SumMultiSheetRange_FlippedCells_OOXML;
    procedure SumMultiSheetRange_FlippedSheets_OOXML;
    procedure SumMultiSheetRange_FlippedSheetsAndCells_OOXML;
    procedure SumMultiSheetRange_FlippedCells_XML;
    procedure SumMultiSheetRange_FlippedSheets_XML;
    procedure SumMultiSheetRange_FlippedSheetsAndCells_XML;
    procedure SumMultiSheetRange_FlippedSheetsAndCells_ODS;

    procedure CountIFS_Microsoft_OOXML;
    procedure CountIFS_Microsoft_ODS;
    procedure SumIFS_Microsoft_OOXML;
    procedure SumIFS_Microsoft_ODS;
    procedure AverageIFS_Microsoft_OOXML;
    procedure AverageIFS_Microsoft_ODS;

    procedure IfConst_BIFF8;
    procedure IfConst_OOXML;
    procedure IfConst_XML;
    procedure IfConst_ODS;

    procedure IfConst_BIFF8_2;

    procedure CountIfRange_BIFF8;
    procedure CountIfRangeSheet_BIFF8;

    procedure SumIfRangeSheetSheet_BIFF8;

    procedure MatchColASC_BIFF8;
    procedure MatchColDESC_BIFF8;
    procedure MatchCol0_BIFF8;
    procedure MatchRowASC_BIFF8;
    procedure MatchRowDESC_BIFF8;

    procedure NonExistantSheet_BIFF5;
    procedure NonExistantSheet_BIFF8;
    procedure NonExistantSheet_OOXML;
    procedure NonExistantSheet_XML;
    procedure NonExistantSheet_ODS;

    procedure NonExistantSheetRange_BIFF5;
    procedure NonExistantSheetRange_BIFF8;
    procedure NonExistantSheetRange_OOXML;
    procedure NonExistantSheetRange_XML;
    procedure NonExistantSheetRange_ODS;

    procedure RenameWorksheet_Single;
    procedure RenameWorksheet_Multi_First;
    procedure RenameWorksheet_Multi_Inner;
    procedure RenameWorksheet_Multi_Last;

    procedure DeleteWorksheet_Single_BeforeRef;
    procedure DeleteWorksheet_Single_Ref;
    procedure DeleteWorksheet_Single_AfterRef;
    procedure DeleteWorksheet_Multi_Before;
    procedure DeleteWorksheet_Multi_First;
    procedure DeleteWorksheet_Multi_Inner;
    procedure DeleteWorksheet_Multi_Last;
    procedure DeleteWorksheet_Multi_After;
    procedure DeleteWorksheet_Multi_KeepFirst;
    procedure DeleteWorksheet_Multi_All;

    procedure Error_AddStringNumber;
    procedure Error_SubtractStringNumber;
    procedure Error_MultiplyStringNumber;
    procedure Error_DivideStringNumber;
    procedure Error_PowerStringNumber;
    procedure Error_SinString;
    procedure Error_SinStringAddNumber;
    procedure Error_Equal;
    procedure Error_NotEqual;
    procedure Error_Greater;
    procedure Error_Smaller;
    procedure Error_GreaterEqual;
    procedure Error_LessEqual;
    procedure Error_UnaryPlusString;
    procedure Error_UnaryMinusString;

    procedure Add_Number_NumString;
    procedure Equal_Number_NumString;
    procedure UnaryMinusNumString;
    
    procedure InsertRow_BeforeFormula;
    procedure InsertRow_AfterFormula;
    procedure InsertRow_AfterAll;
    procedure InsertCol_BeforeFormula;
    procedure InsertCol_AfterFormula;
    procedure InsertCol_AfterAll;
    
    procedure DeleteRow_BeforeFormula;
    procedure DeleteRow_AfterFormula;
    procedure DeleteRow_AfterAll;
    procedure DeleteCol_BeforeFormula;
    procedure DeleteCol_AfterFormula;
    procedure DeleteCol_AfterAll;
  end;

implementation

uses
 {$IFDEF FORMULADEBUG}
  LazLogger,
 {$ENDIF}
  typinfo, lazUTF8, fpsUtils;


{ TSpreadSingleFormulaTests }

procedure TSpreadSingleFormulaTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadSingleFormulaTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadSingleFormulaTests.TestFormula(AFormula: String;
  AExpected: String; ATestKind: TFormulaTestKind; AFormat: TsSpreadsheetFormat;
  AExpectedFormula: String = '');
const
  SHEET1 = 'Sheet1';
  SHEET2 = 'Sheet2';
  SHEET3 = 'Sheet3';
  SHEET4 = 'Sheet4';
  TESTCELL_ROW = 1;       // Cell with formula: C2
  TESTCELL_COL = 2;
var
  worksheet: TsWorksheet;
  othersheet: TsWorksheet;
  workbook: TsWorkbook;
  TempFile: string; //write xls/xml to this file and read back from it
  cell: PCell;
  actualformula: String;
  actualValue: String;
begin
  TempFile := GetTempFileName;
  if AExpectedFormula = '' then AExpectedFormula := AFormula;

  try
    // Create test workbook and write test formula and needed cells
    workbook := TsWorkbook.Create;
    try
      workbook.FormatSettings := ExprFormatSettings;
      workbook.Options := workbook.Options + [boCalcBeforeSaving, boAutoCalc];
      workSheet:= workBook.AddWorksheet(SHEET1);

      if ATestKind <> ftkConstants then begin
        // Write cells used by the formula
        worksheet.WriteNumber(2, 2, 1.0);   // C3
        worksheet.WriteNumber(3, 2, -2.0);  // C4
        worksheet.WriteNumber(4, 2, 1.5);   // C5
        worksheet.WriteNumber(2, 3, 15.0);  // D3

        worksheet.WriteDateTime( 9, 1, EncodeDate(2012, 2, 5), nfShortDate);    // B10
        worksheet.WriteDateTime(10, 1, EncodeTime(14, 20, 41, 0), nfLongTime);  // B11
      end;

      if ATestKind in [ftkCellRangeSheet, ftkCellRangeSheetRange] then begin
        otherSheet := Workbook.AddWorksheet(SHEET2);
        othersheet.WriteNumber(2, 2, 10.0);   // Sheet2!C3
        othersheet.WriteNumber(3, 2, -20.0);  // Sheet2!C4
        othersheet.WriteNumber(4, 2, 15.0);   // Sheet2!C5
        othersheet.WriteNumber(2, 3, 150.0);  // Sheet2!D5
      end;

      if ATestKind = ftkCellRangeSheetRange then begin
        otherSheet := Workbook.AddWorksheet(SHEET3);
        othersheet.WriteNumber(2, 2, 100.0);   // Sheet3!C3
        othersheet.WriteNumber(3, 2, -200.0);  // Sheet3!C4
        othersheet.WriteNumber(4, 2, 150.0);   // Sheet3!C5
        othersheet.WriteNumber(2, 3, 1500.0);  // Sheet3!D5
      end;

      if ATestkind = ftkSortedNumbersAsc then begin
        othersheet := Workbook.AddWorksheet(SHEET4);
        othersheet.WriteNumber(2, 2, 10.0);   // Sheet4!C3
        othersheet.WriteNumber(3, 2, 12.0);   // Sheet4!C4
        othersheet.WriteNumber(4, 2, 15.0);   // Sheet4!C5
        othersheet.WriteNumber(5, 2, 20.0);   // Sheet4!C6
        othersheet.WriteNumber(6, 2, 25.0);   // Sheet4!C7
        othersheet.WriteNumber(2, 3, 12.0);   // Sheet4!D3
        othersheet.WriteNumber(2, 4, 15.0);   // Sheet4!E3
        othersheet.WriteNumber(2, 5, 20.0);   // Sheet4!F3
        othersheet.WriteNumber(2, 6, 25.0);   // Sheet4!G3
      end else
      if ATestkind = ftkSortedNumbersDesc then begin
        othersheet := Workbook.AddWorksheet(SHEET4);
        othersheet.WriteNumber(2, 2, 25.0);   // Sheet4!C3
        othersheet.WriteNumber(3, 2, 20.0);   // Sheet4!C4
        othersheet.WriteNumber(4, 2, 15.0);   // Sheet4!C5
        othersheet.WriteNumber(5, 2, 12.0);   // Sheet4!C6
        othersheet.WriteNumber(6, 2, 10.0);   // Sheet4!C7
        othersheet.WriteNumber(2, 3, 20.0);   // Sheet4!D3
        othersheet.WriteNumber(2, 4, 15.0);   // Sheet4!E3
        othersheet.WriteNumber(2, 5, 12.0);   // Sheet4!F3
        othersheet.WriteNumber(2, 6, 10.0);   // Sheet4!G3
      end;

      // Write the formula
      cell := worksheet.WriteFormula(TESTCELL_ROW, TESTCELL_COL, AFormula);

      // Read formula before saving
      actualFormula := worksheet.ReadFormula(cell);
      CheckEquals(AExpectedFormula, actualFormula, 'Unsaved formula text mismatch');

      // Read calculated value before saving
      actualValue := worksheet.ReadAsText(TESTCELL_ROW, TESTCELL_COL);
      CheckEquals(AExpected, actualvalue, 'Unsaved calculated value mismatch');

      // Save
      workbook.WriteToFile(TempFile, AFormat, true);
    finally
      workbook.Free;
    end;

    // Read file
    workbook := TsWorkbook.Create;
    try
      workbook.FormatSettings := ExprFormatSettings;
      workbook.Options := workbook.Options + [boReadFormulas, boAutoCalc];
      workbook.ReadFromFile(TempFile, AFormat);
      worksheet := workbook.GetFirstWorksheet;

      // Read calculated formula value
      actualValue := worksheet.ReadAsText(TESTCELL_ROW, TESTCELL_COL);
      CheckEquals(AExpected, actualValue, 'Saved calculated value mismatch');

      cell := worksheet.FindCell(TESTCELL_ROW, TESTCELL_COL);
      actualformula := worksheet.Formulas.FindFormula(cell)^.Text;
      // When writing ranges are reconstructed in correct order -> compare against AExpectedFormula
      CheckEquals(AExpectedFormula, actualformula, 'Saved formula text mismatch.');
    finally
      workbook.Free;
    end;

  finally
    if FileExists(TempFile) then DeleteFile(TempFile);
  end;
end;

procedure TSpreadSingleFormulaTests.AddConst_BIFF2;
begin
  TestFormula('1+1', '2', ftkConstants, sfExcel2);
end;

procedure TSpreadSingleFormulaTests.AddConst_BIFF5;
begin
  TestFormula('1+1', '2', ftkConstants, sfExcel5);
end;

procedure TSpreadSingleFormulaTests.AddConst_BIFF8;
begin
  TestFormula('1+1', '2', ftkConstants, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.AddConst_OOXML;
begin
  TestFormula('1+1', '2', ftkConstants, sfOOXML);
end;

procedure TSpreadSingleFormulaTests.AddConst_XML;
begin
  TestFormula('1+1', '2', ftkConstants, sfExcelXML);
end;

procedure TSpreadSingleFormulaTests.AddConst_ODS;
begin
  TestFormula('1+1', '2', ftkConstants, sfOpenDocument);
end;

{---------------}

procedure TSpreadSingleFormulaTests.AddCells_BIFF2;
begin
  TestFormula('C3+C4', '-1', ftkCells, sfExcel2);
end;

procedure TSpreadSingleFormulaTests.AddCells_BIFF5;
begin
  TestFormula('C3+C4', '-1', ftkCells, sfExcel5);
end;

procedure TSpreadSingleFormulaTests.AddCells_BIFF8;
begin
  TestFormula('C3+C4', '-1', ftkCells, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.AddCells_OOXML;
begin
  TestFormula('C3+C4', '-1', ftkCells, sfOOXML);
end;

procedure TSpreadSingleFormulaTests.AddCells_XML;
begin
  TestFormula('C3+C4', '-1', ftkCells, sfExcelXML);
end;

procedure TSpreadSingleFormulaTests.AddCells_ODS;
begin
  TestFormula('C3+C4', '-1', ftkCells, sfOpenDocument);
end;

{ ------ }

procedure TSpreadSingleFormulaTests.RoundConst1_ODS;
begin
  TestFormula('ROUND(1234.56789,2)', '1234.57', ftkConstants, sfOpenDocument);
end;

procedure TSpreadSingleFormulaTests.RoundConst2_ODS;
begin
  TestFormula('ROUND(1234.56789,-2)', '1200', ftkConstants, sfOpenDocument);
end;

procedure TSpreadSingleFormulaTests.RoundCell1_ODS;
begin
  TestFormula('ROUND(1234.56789,C3)', '1234.6', ftkCells, sfOpenDocument);  // C3 = 1
end;

procedure TSpreadSingleFormulaTests.RoundCell2_ODS;
begin
  TestFormula('ROUND(1234.56789,C4)', '1200', ftkCells, sfOpenDocument);    // C4 = -2
end;

{ ------ }

procedure TSpreadSingleFormulaTests.YearConst_BIFF8;
var
  s: String;
begin
  s := FormatDateTime(DefaultFormatSettings.ShortDateFormat, EncodeDate(2012,2,5), ExprFormatSettings);
  TestFormula(Format('YEAR("%s")', [s]), '2012', ftkConstants, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.YearCell_BIFF8;
begin
  TestFormula('YEAR(B10)', '2012', ftkCells, sfExcel8);      // B10: 2012/02/05
end;

procedure TSpreadSingleFormulaTests.MonthConst_BIFF8;
var
  s: String;
begin
  s := FormatDateTime(DefaultFormatSettings.ShortDateFormat, EncodeDate(2012,2,5), DefaultFormatSettings);
  TestFormula(Format('MONTH("%s")', [s]), '2', ftkConstants, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.MonthCell_BIFF8;
begin
  TestFormula('MONTH(B10)', '2', ftkCells, sfExcel8);      // B10: 2012/02/05
end;

procedure TSpreadSingleFormulaTests.DayConst_BIFF8;
var
  s: String;
begin
  s := FormatDateTime(DefaultFormatSettings.ShortDateFormat, EncodeDate(2012,2,5), DefaultFormatSettings);
  TestFormula(Format('DAY("%s")', [s]), '5', ftkConstants, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.DayCell_BIFF8;
begin
  TestFormula('DAY(B10)', '5', ftkCells, sfExcel8);      // B10: 2012/02/05
end;

{ ----- }

procedure TSpreadSingleFormulaTests.HourConst_BIFF8;
var
  s: String;
begin
  s := FormatDateTime(ExprFormatSettings.LongTimeFormat, EncodeTime(14, 20, 41, 0), DefaultFormatSettings);
  TestFormula(Format('HOUR("%s")', [s]), '14', ftkConstants, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.HourCell_BIFF8;
begin
  TestFormula('HOUR(B11)', '14', ftkCells, sfExcel8);      // B11: 14:20:41
end;

procedure TSpreadSingleFormulaTests.MinuteConst_BIFF8;
var
  s: String;
begin
  s := FormatDateTime(ExprFormatSettings.LongTimeFormat, EncodeTime(14, 20, 41, 0), ExprFormatSettings);
  TestFormula(Format('MINUTE("%s")', [s]), '20', ftkConstants, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.MinuteCell_BIFF8;
begin
  TestFormula('MINUTE(B11)', '20', ftkCells, sfExcel8);      // B11: 14:20:41
end;

procedure TSpreadSingleFormulaTests.SecondConst_BIFF8;
var
  s: String;
begin
  s := FormatDateTime(ExprFormatSettings.LongTimeFormat, EncodeTime(14, 20, 41, 0), ExprFormatSettings);
  TestFormula(Format('SECOND("%s")', [s]), '41', ftkConstants, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.SecondCell_BIFF8;
begin
  TestFormula('SECOND(B11)', '41', ftkCells, sfExcel8);      // B11: 14:20:41
end;

{ ---- }

procedure TSpreadSingleFormulaTests.SumRange_BIFF2;
begin
  TestFormula('SUM(C3:C5)', '0.5', ftkCellRange, sfExcel2);
end;

procedure TSpreadSingleFormulaTests.SumRange_BIFF5;
begin
  TestFormula('SUM(C3:C5)', '0.5', ftkCellRange, sfExcel5);
end;

procedure TSpreadSingleFormulaTests.SumRange_BIFF8;
begin
  TestFormula('SUM(C3:C5)', '0.5', ftkCellRange, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.SumRange_OOXML;
begin
  TestFormula('SUM(C3:C5)', '0.5', ftkCellRange, sfOOXML);
end;

procedure TSpreadSingleFormulaTests.SumRange_XML;
begin
  TestFormula('SUM(C3:C5)', '0.5', ftkCellRange, sfExcelXML);
end;

procedure TSpreadSingleFormulaTests.SumRange_ODS;
begin
  TestFormula('SUM(C3:C5)', '0.5', ftkCellRange, sfOpenDocument);
end;

{ ---- }

procedure TSpreadSingleFormulaTests.SumSheetRange_BIFF5;
begin
  TestFormula('SUM(Sheet2!C3:C5)', '5', ftkCellRangeSheet, sfExcel5);
end;

procedure TSpreadSingleFormulaTests.SumSheetRange_BIFF8;
begin
  TestFormula('SUM(Sheet2!C3:C5)', '5', ftkCellRangeSheet, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.SumSheetRange_OOXML;
begin
  TestFormula('SUM(Sheet2!C3:C5)', '5', ftkCellRangeSheet, sfOOXML);
end;

procedure TSpreadSingleFormulaTests.SumSheetRange_XML;
begin
  TestFormula('SUM(Sheet2!C3:C5)', '5', ftkCellRangeSheet, sfExcelXML);
end;

procedure TSpreadSingleFormulaTests.SumSheetRange_ODS;
begin
  TestFormula('SUM(Sheet2!C3:C5)', '5', ftkCellRangeSheet, sfOpenDocument);
end;

{ ---- }

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_BIFF5;
begin
  TestFormula('SUM(Sheet2:Sheet3!C3:C5)', '55', ftkCellRangeSheetRange, sfExcel5);
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_BIFF8;
begin
  TestFormula('SUM(Sheet2:Sheet3!C3:C5)', '55', ftkCellRangeSheetRange, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_OOXML;
begin
  TestFormula('SUM(Sheet2:Sheet3!C3:C5)', '55', ftkCellRangeSheetRange, sfOOXML);
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_XML;
begin
  TestFormula('SUM(Sheet2:Sheet3!C3:C5)', '55', ftkCellRangeSheetRange, sfExcelXML);
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_ODS;
begin
  TestFormula('SUM(Sheet2:Sheet3!C3:C5)', '55', ftkCellRangeSheetRange, sfOpenDocument);
end;

{ --- }

{ Range formulas in which the parts are not ordered. They will be put into the
  correct order when then formula is written to the worksheet. --> the
  expected range must be in correct order. }
procedure TSpreadSingleFormulaTests.SumMultiSheetRange_FlippedSheetsAndCells_OOXML;
begin
  TestFormula('SUM(Sheet3:Sheet2!C5:C3)', '55', ftkCellRangeSheetRange, sfOOXML, 'SUM(Sheet2:Sheet3!C3:C5)');
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_FlippedSheetsAndCells_XML;
begin
  TestFormula('SUM(Sheet3:Sheet2!C5:C3)', '55', ftkCellRangeSheetRange, sfExcelXML, 'SUM(Sheet2:Sheet3!C3:C5)');
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_FlippedSheetsAndCells_ODS;
begin
  TestFormula('SUM(Sheet3:Sheet2!C5:C3)', '55', ftkCellRangeSheetRange, sfOpenDocument, 'SUM(Sheet2:Sheet3!C3:C5)');
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_FlippedCells_BIFF8;
begin
  // Upon writing the ranges are reconstructed for BIFF in correct order.
  TestFormula('SUM(Sheet2:Sheet3!C5:C3)', '55', ftkCellRangeSheetRange, sfExcel8, 'SUM(Sheet2:Sheet3!C3:C5)');
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_FlippedCells_OOXML;
begin
  TestFormula('SUM(Sheet2:Sheet3!C5:C3)', '55', ftkCellRangeSheetRange, sfOOXML, 'SUM(Sheet2:Sheet3!C3:C5)');
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_FlippedSheets_OOXML;
begin
  TestFormula('SUM(Sheet3:Sheet2!C3:C5)', '55', ftkCellRangeSheetRange, sfOOXML, 'SUM(Sheet2:Sheet3!C3:C5)');
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_FlippedCells_XML;
begin
  TestFormula('SUM(Sheet2:Sheet3!C5:C3)', '55', ftkCellRangeSheetRange, sfExcelXML, 'SUM(Sheet2:Sheet3!C3:C5)');
end;

procedure TSpreadSingleFormulaTests.SumMultiSheetRange_FlippedSheets_XML;
begin
  TestFormula('SUM(Sheet3:Sheet2!C3:C5)', '55', ftkCellRangeSheetRange, sfExcelXML, 'SUM(Sheet2:Sheet3!C3:C5)');
end;

{ --- }

{ Based on example from Microsoft site for SUMIFS:
  https://support.microsoft.com/en-us/office/sumifs-function-c9e748f5-7ea7-455d-9406-611cebce642b

  AFormulaType is 0 for COUNTIFS, 1 for SUMIFS, and 2 for AVERAGEIFS }
procedure TSpreadSingleFormulaTests.TestCountSumAvgIFS_Microsoft(AFormat: TsSpreadsheetFormat;
  AFormulaType: Integer);
const
  ROW1 = 10;
  ROW2 = 11;
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  value1, value2: Double;
  tempFile: String;
begin
  tempFile := GetTempFileName;

  book := TsWorkbook.Create;
  try
    sheet := book.AddWorksheet('Test');
    sheet.WriteText(0, 0, 'Quantity Sold');  sheet.WriteText(0, 1, 'Product');     sheet.WriteText(0, 2, 'SalesPerson');
    sheet.WriteNumber(1, 0, 5);              sheet.WriteText(1, 1, 'Apples');      sheet.WriteText(1, 2, 'Tom');
    sheet.WriteNumber(2, 0, 4);              sheet.WriteText(2, 1, 'Apples');      sheet.WriteText(2, 2, 'Sarah');
    sheet.WriteNumber(3, 0, 15);             sheet.WriteText(3, 1, 'Artichokes');  sheet.WriteText(3, 2, 'Tom');
    sheet.WriteNumber(4, 0, 3);              sheet.WriteText(4, 1, 'Artichokes');  sheet.WriteText(4, 2, 'Sarah');
    sheet.WriteNumber(5, 0, 22);             sheet.WriteText(5, 1, 'Bananas');     sheet.WriteText(5, 2, 'Tom');
    sheet.WriteNumber(6, 0, 12);             sheet.WriteText(6, 1, 'Bananas');     sheet.WriteText(6, 2, 'Sarah');
    sheet.WriteNumber(7, 0, 10);             sheet.WriteText(7, 1, 'Carrots');     sheet.WriteText(7, 2, 'Tom');
    sheet.WriteNumber(8, 0, 33);             sheet.WriteText(8, 1, 'Carrots');     sheet.WriteText(8, 2, 'Sarah');

    case AFormulaType of
      0: begin
           sheet.WriteFormula(ROW1, 0, 'COUNTIFS(B2:B9, "=A*", C2:C9, "Tom")');        // expected: 2
           sheet.WriteFormula(ROW2, 0, 'COUNTIFS(B2:B9, "<>Bananas", C2:C9, "Tom")');  // expected: 3
         end;
      1: begin
           sheet.WriteFormula(ROW1, 0, 'SUMIFS(A2:A9, B2:B9, "=A*", C2:C9, "Tom")');        // expected: 20
           sheet.WriteFormula(ROW2, 0, 'SUMIFS(A2:A9, B2:B9, "<>Bananas", C2:C9, "Tom")');  // expected: 30
         end;
      2: begin
           sheet.WriteFormula(ROW1, 0, 'AVERAGEIFS(A2:A9, B2:B9, "=A*", C2:C9, "Tom")');        // expected: 4.0
           sheet.WriteFormula(ROW2, 0, 'AVERAGEIFS(A2:A9, B2:B9, "<>Bananas", C2:C9, "Tom")');  // expected: 10.0
         end;
    end;

    sheet.CalcFormulas;

    value1 := sheet.ReadAsNumber(ROW1, 0);
    value2 := sheet.ReadAsNumber(ROW2, 0);
    case AFormulaType of
      0: begin    // Checking COUNTIFS
           CheckEquals(2, value1, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW1, 0));
           CheckEquals(3, value2, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW2, 0));
         end;
      1: begin    // Checking SUMIFS
           CheckEquals(20, value1, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW1, 0));
           CheckEquals(30, value2, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW2, 0));
         end;
      2: begin    // Checking AVERAGEIFS
           CheckEquals(10.0, value1, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW1, 0));
           CheckEquals(10.0, value2, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW2, 0));
         end;
    end;

    book.WriteToFile(tempFile, AFormat, true);
  finally
    book.Free;
  end;

  book := TsWorkbook.Create;
  try
    book.Options := [boReadFormulas, boAutoCalc];
    book.ReadFromFile(tempFile, AFormat);
    sheet := book.GetFirstWorksheet;

    value1 := sheet.ReadAsNumber(ROW1, 0);
    value2 := sheet.ReadAsNumber(ROW2, 0);
    case AFormulaType of
      0: begin    // Checking COUNTIFS
           CheckEquals(2, value1, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW1, 0));
           CheckEquals(3, value2, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW2, 0));
         end;
      1: begin    // Checking SUMIFS
           CheckEquals(20, value1, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW1, 0));
           CheckEquals(30, value2, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW2, 0));
         end;
      2: begin    // Checking AVERAGEIFS
           CheckEquals(10.0, value1, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW1, 0));
           CheckEquals(10.0, value2, 'Unsaved value mismatch in cell ' + CellNotation(sheet, ROW2, 0));
         end;
    end;
  finally
    book.Free;
  end;
end;

procedure TSpreadSingleFormulaTests.CountIFS_Microsoft_OOXML;
begin
  TestCountSumAvgIFS_Microsoft(sfOOXML, 0);
end;

procedure TSpreadSingleFormulaTests.CountIFS_Microsoft_ODS;
begin
  TestCountSumAvgIFS_Microsoft(sfOpenDocument, 0);
end;

procedure TSpreadSingleFormulaTests.SumIFS_Microsoft_OOXML;
begin
  TestCountSumAvgIFS_Microsoft(sfOOXML, 1);
end;

procedure TSpreadSingleFormulaTests.SumIFS_Microsoft_ODS;
begin
  TestCountSumAvgIFS_Microsoft(sfOpenDocument, 1);
end;

procedure TSpreadSingleFormulaTests.AverageIFS_Microsoft_OOXML;
begin
  TestCountSumAvgIFS_Microsoft(sfOOXML, 1);
end;

procedure TSpreadSingleFormulaTests.AverageIFS_Microsoft_ODS;
begin
  TestCountSumAvgIFS_Microsoft(sfOpenDocument, 1);
end;

{ --- }

procedure TSpreadSingleFormulaTests.IfConst_BIFF8;
begin
  TestFormula('IF(C3="A","is A","not A")', 'not A', ftkCellConstant, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.IfConst_OOXML;
begin
  TestFormula('IF(C3="A","is A","not A")', 'not A', ftkCellConstant, sfOOXML);
end;

procedure TSpreadSingleFormulaTests.IfConst_XML;
begin
  TestFormula('IF(C3="A","is A","not A")', 'not A', ftkCellConstant, sfExcelXML);
end;

procedure TSpreadSingleFormulaTests.IfConst_ODS;
begin
  TestFormula('IF(C3="A","is A","not A")', 'not A', ftkCellConstant, sfOpenDocument);
end;

{ --- }

procedure TSpreadSingleFormulaTests.IfConst_BIFF8_2;
begin
  TestFormula('IF(C3=1,"equal","different")', 'equal', ftkCellConstant, sfExcel8);
end;

{ --- }

procedure TSpreadSingleFormulaTests.CountIfRange_BIFF8;
begin
  TestFormula('COUNTIF(C3:C5,">1")', '1', ftkCellRange, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.CountIfRangeSheet_BIFF8;
begin
  TestFormula('COUNTIF(Sheet2!C3:C5,">10")', '1', ftkCellRangeSheet, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.SumIfRangeSheetSheet_BIFF8;
begin
  TestFormula('SUMIF(Sheet2!C3:C5,">10",Sheet3!C3:C5)', '150', ftkCellRangeSheetRange, sfExcel8);
end;

{ ---- }

procedure TSpreadSingleFormulaTests.MatchColASC_BIFF8;
begin                     //10,12,15,20,25
  TestFormula('MATCH(12.5,Sheet4!C3:C7,1)', '2', ftkSortedNumbersASC, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.MatchColDESC_BIFF8;
begin                     //25,20,15,12,10
  TestFormula('MATCH(12.5,Sheet4!C3:C7,-1)', '3', ftkSortedNumbersDESC, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.MatchCol0_BIFF8;
begin                   //10,12,15,20,25
  TestFormula('MATCH(12,Sheet4!C3:C7,0)', '2', ftkSortedNumbersASC, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.MatchRowASC_BIFF8;
begin
  TestFormula('MATCH(12,Sheet4!C3:G3,1)', '2', ftkSortedNumbersASC, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.MatchRowDESC_BIFF8;
begin
  TestFormula('MATCH(12,Sheet4!C3:G3,-1)', '4', ftkSortedNumbersDESC, sfExcel8);
end;

{ --- }

procedure TSpreadSingleFormulaTests.NonExistantSheet_BIFF5;
begin
  TestFormula('Missing!C3', '#REF!', ftkCellRangeSheet, sfExcel5, '#REF!');
end;

procedure TSpreadSingleFormulaTests.NonExistantSheet_BIFF8;
begin
  TestFormula('Missing!C3', '#REF!', ftkCellRangeSheet, sfExcel8, '#REF!');
end;

procedure TSpreadSingleFormulaTests.NonExistantSheet_OOXML;
begin
  TestFormula('Missing!C3', '#REF!', ftkCellRangeSheet, sfOOXML, '#REF!');
end;

procedure TSpreadSingleFormulaTests.NonExistantSheet_XML;
begin
  TestFormula('Missing!C3', '#REF!', ftkCellRangeSheet, sfExcelXML, '#REF!');
end;

procedure TSpreadSingleFormulaTests.NonExistantSheet_ODS;
begin
  TestFormula('Missing!C3', '#REF!', ftkCellRangeSheet, sfOpenDocument, '#REF!');
end;

{ --- }

procedure TSpreadSingleFormulaTests.NonExistantSheetRange_BIFF5;
begin
  TestFormula('SUM(Missing1:Missing2!C3)', '#REF!', ftkCellRangeSheet, sfExcel5, 'SUM(#REF!)');
end;

procedure TSpreadSingleFormulaTests.NonExistantSheetRange_BIFF8;
begin
  TestFormula('SUM(Missing1:Missing2!C3)', '#REF!', ftkCellRangeSheet, sfExcel8, 'SUM(#REF!)');
end;

procedure TSpreadSingleFormulaTests.NonExistantSheetRange_OOXML;
begin
  TestFormula('SUM(Missing1:Missing2!C3)', '#REF!', ftkCellRangeSheet, sfOOXML, 'SUM(#REF!)');
end;

procedure TSpreadSingleFormulaTests.NonExistantSheetRange_XML;
begin
  TestFormula('SUM(Missing1:Missing2!C3)', '#REF!', ftkCellRangeSheet, sfExcelXML, 'SUM(#REF!)');
end;

procedure TSpreadSingleFormulaTests.NonExistantSheetRange_ODS;
begin
  TestFormula('SUM(Missing1:Missing2!C3)', '#REF!', ftkCellRangeSheet, sfOpenDocument, 'SUM(#REF!)');
end;


{------------------------------------------------------------------------------}

{ ATestKind defines the action taken:
  - wtkRenameWorksheet
  - wtkDeleteWorksheet
}
procedure TSpreadSingleFormulaTests.TestWorksheet(ATestKind: TWorksheetTestKind;
  ATestCase: Integer);
const
  SHEET1 = 'Sheet1';
  SHEET2 = 'Sheet2';
  SHEET3 = 'Sheet3';
  SHEET4 = 'Sheet4';
  SHEET5 = 'Sheet5';
  SHEET6 = 'Sheet6';
  TESTCELL_ROW = 1;
  TESTCELL_COL = 1;
  ACTION_NAME: array [TWorksheetTestKind] of string = ('RENAME', 'DELETE');
var
  workbook: TsWorkbook;
  worksheet1: TsWorksheet;
  worksheet2: TsWorksheet;
  worksheet3: TsWorksheet;
  worksheet4: TsWorksheet;
  worksheet5: TsWorksheet;
  worksheet6: TsWorksheet;
  tempFile: String;
  formula: String;
  actualFormula: String;
  actualValue: String;
  expectedFormula: string;
  expectedValue: String;
  cell: PCell;
const
//  DO_SAVE: boolean = true;
  DO_SAVE: Boolean = false;
begin
  tempFile := GetTempFileName;

  try
    // Create test workbook and write test formula and needed cells
    workbook := TsWorkbook.Create;
    try
      workbook.Options := workbook.Options + [boCalcBeforeSaving, boAutoCalc];
      worksheet1 := workBook.AddWorksheet(SHEET1);
      worksheet2 := workbook.AddWorksheet(SHEET2);
      worksheet3 := workbook.AddWorksheet(SHEET3);
      worksheet4 := workbook.AddWorksheet(SHEET4);
      worksheet5 := workbook.AddWorksheet(SHEET5);
      worksheet6 := workbook.AddWorksheet(SHEET6);

      // Write cells used by the formula
      worksheet1.WriteNumber(2, 2, 1.0);   // C3
      worksheet1.WriteNumber(3, 2, -2.0);  // C4
      worksheet1.WriteNumber(4, 2, 1.5);   // C5
      worksheet1.WriteNumber(2, 3, 15.0);  // D3

      // No data in worksheet 2 - it is just a spacer

      worksheet3.WriteNumber(2, 2, 10.0);   // Sheet3!C3
      worksheet3.WriteNumber(3, 2, -20.0);  // Sheet3!C4
      worksheet3.WriteNumber(4, 2, 15.0);   // Sheet3!C5
      worksheet3.WriteNumber(2, 3, 150.0);  // Sheet3!D5

      worksheet4.WriteNumber(2, 2, 100.0);   // Sheet4!C3
      worksheet4.WriteNumber(3, 2, -200.0);  // Sheet4!C4
      worksheet4.WriteNumber(4, 2, 150.0);   // Sheet4!C5
      worksheet4.WriteNumber(2, 3, 1500.0);  // Sheet4!D5

      worksheet5.WriteNumber(2, 2, 1000.0);   // Sheet5!C3
      worksheet5.WriteNumber(3, 2, -2000.0);  // Sheet5!C4
      worksheet5.WriteNumber(4, 2, 1500.0);   // Sheet5!C5
      worksheet5.WriteNumber(2, 3, 15000.0);  // Sheet5!D5

      // No data in worksheet 6 - it is just a spacer

      // Write the formula
      case ATestKind of
        wtkRenameworksheet:
          case ATestCase of
            1:
              begin
                formula := 'Sheet3!C3';                // SINGLE_SHEET FORMULA
                expectedValue := '10';
              end;
            2..4:
              begin
                formula := 'SUM(Sheet3:Sheet5!C3)';    // MULTI-SHEET FORMULA
                expectedValue := '1110';
              end;
          end;
        wtkDeleteWorksheet:
          case ATestCase of
            1..3:
              begin
                formula := 'Sheet3!C3';
                expectedValue := '10';
              end;
            4..10:
              begin
                formula := 'SUM(Sheet3:Sheet5!C3)';    // MULTI-SHEET FORMULA
                expectedValue := '1110';
              end;
          end;
      end;
      cell := worksheet1.WriteFormula(TESTCELL_ROW, TESTCELL_COL, formula);
      workbook.CalcFormulas;

      // Read formula before action
      actualFormula := worksheet1.ReadFormula(cell);
      CheckEquals(formula, actualFormula,
        'Formula text mismatch before action "' + ACTION_NAME[ATestKind] + '"');

      // Read calculated value before action
      actualvalue := worksheet1.ReadAsText(TESTCELL_ROW, TESTCELL_COL);
      CheckEquals(expectedValue, actualvalue,
        'Calculated value mismatch before action "' + ACTION_NAME[ATestKind] + '"');

      // Action
      case ATestKind of
        // Renaming tests
        wtkRenameWorksheet:
          case ATestCase of
            1: begin // Rename sheet referred by single-sheet formula
                 worksheet3.Name := 'Table3';
                 expectedFormula := 'Table3!C3';
               end;
            2: begin  // Rename sheet referred by first of multi-sheet range
                 worksheet3.Name := 'Table';
                 expectedFormula := 'SUM(Table:Sheet5!C3)';
               end;
            3: begin  // Rename sheet referred by inner sheet of sheet range
                 worksheet4.Name := 'Table';
                 expectedFormula := 'SUM(Sheet3:Sheet5!C3)';
               end;
            4: begin  // Rename sheet referred by last sheet of sheet range
                 worksheet5.Name := 'Table';
                 expectedFormula := 'SUM(Sheet3:Table!C3)';
               end;
          end;

        // Deletion tests
        wtkDeleteWorksheet:
          case ATestCase of
            // Single-sheet tests
            1: begin  // Delete sheet before referenced sheet (Sheet3)
                 workbook.RemoveWorksheet(worksheet2);
                 expectedFormula := 'Sheet3!C3';
                 expectedValue := '10';
               end;
            2: begin  // Delete referenced sheet
                 workbook.RemoveWorksheet(worksheet3);
                 expectedFormula := '#REF!';
                 expectedValue := '#REF!';
               end;
            3: begin  // Delete sheet after referenced sheet
                 workbook.RemoveWorksheet(worksheet4);
                 expectedFormula := 'Sheet3!C3';
                 expectedValue := '10';
               end;
            // Range tests
            4: begin // Delete sheet before referenced range (Sheet3:Sheet5)
                 workbook.RemoveWorksheet(worksheet2);
                 expectedFormula := 'SUM(Sheet3:Sheet5!C3)';
                 expectedValue := '1110';
               end;
            5: begin  // Delete 1st sheet of referenced range
                 workbook.RemoveWorksheet(worksheet3);
                 expectedFormula := 'SUM(Sheet4:Sheet5!C3)';
                 expectedValue := '1100';
               end;
            6: begin  // Delete inner sheet of referenced range
                 workbook.RemoveWorksheet(worksheet4);
                 expectedformula := 'SUM(Sheet3:Sheet5!C3)';
                 expectedvalue := '1010';
               end;
            7: begin  // Delete last sheet of referenced range
                 workbook.RemoveWorksheet(worksheet5);
                 expectedformula := 'SUM(Sheet3:Sheet4!C3)';
                 expectedValue := '110';
               end;
            8: begin  // Delete sheet after referenced range
                 workbook.RemoveWorksheet(worksheet6);
                 expectedformula := 'SUM(Sheet3:Sheet5!C3)';
                 expectedValue := '1110';
               end;
            9: begin  // Delete all sheets expect first of range
                 workbook.RemoveWorksheet(worksheet4);
                 workbook.RemoveWorksheet(worksheet5);
                 expectedformula := 'SUM(Sheet3!C3)';
                 expectedValue := '10';
               end;
           10: begin  // Delete all sheets of referenced range
                 workbook.RemoveWorksheet(worksheet3);
                 workbook.RemoveWorksheet(worksheet4);
                 workbook.RemoveWorksheet(worksheet5);
                 expectedFormula := 'SUM(#REF!)';
                 expectedValue := '#REF!';
               end;
          end;
      end;

      workbook.CalcFormulas;
      if DO_SAVE then                          // For debugging...
        workbook.WriteToFile(tempFile, sfExcel8, true);

      // Read formula after action
      actualFormula := worksheet1.ReadFormula(cell);
      CheckEquals(expectedFormula, actualFormula,
        'Formula text mismatch after action "' + ACTION_NAME[ATestKind] + '"');

      // Read calculated value before action
      actualvalue := worksheet1.ReadAsText(TESTCELL_ROW, TESTCELL_COL);
      CheckEquals(expectedValue, actualvalue,
        'Calculated value mismatch after action "' + ACTION_NAME[ATestKind] + '"');

    finally
      workbook.Free;
    end;

  finally
    if DO_SAVE then
      DeleteFile(tempFile);
  end;
end;

procedure TSpreadSingleFormulaTests.RenameWorksheet_Single;
begin
  TestWorksheet(wtkRenameWorksheet, 1);
end;

procedure TSpreadSingleFormulaTests.RenameWorksheet_Multi_First;
begin
  TestWorksheet(wtkRenameWorksheet, 2);
end;

procedure TSpreadSingleFormulaTests.RenameWorksheet_Multi_Inner;
begin
  TestWorksheet(wtkRenameWorksheet, 3);
end;

procedure TSpreadSingleFormulaTests.RenameWorksheet_Multi_Last;
begin
  TestWorksheet(wtkRenameWorksheet, 4);
end;

procedure TSpreadSingleFormulaTests.DeleteWorksheet_Single_BeforeRef;
begin
  TestWorksheet(wtkDeleteWorksheet, 1);
end;

procedure TSpreadSingleFormulaTests.DeleteWorksheet_Single_Ref;
begin
  TestWorksheet(wtkDeleteWorksheet, 2);
end;

procedure TSpreadSingleFormulaTests.DeleteWorksheet_Single_AfterRef;
begin
  TestWorksheet(wtkDeleteWorksheet, 3);
end;

procedure TSpreadSingleFormulaTests.DeleteWorksheet_Multi_Before;
begin
  TestWorksheet(wtkDeleteWorksheet, 4);
end;

procedure TSpreadSingleFormulaTests.DeleteWorksheet_Multi_First;
begin
  TestWorksheet(wtkDeleteWorksheet, 5);
end;

procedure TSpreadSingleFormulaTests.DeleteWorksheet_Multi_Inner;
begin
  TestWorksheet(wtkDeleteWorksheet, 6);
end;

procedure TSpreadSingleFormulaTests.DeleteWorksheet_Multi_Last;
begin
  TestWorksheet(wtkDeleteWorksheet, 7);
end;

procedure TSpreadSingleFormulaTests.DeleteWorksheet_Multi_After;
begin
  TestWorksheet(wtkDeleteWorksheet, 8);
end;

procedure TSpreadSingleFormulaTests.DeleteWorksheet_Multi_KeepFirst;
begin
  TestWorksheet(wtkDeleteWorksheet, 9);
end;

procedure TSpreadSingleFormulaTests.DeleteWorksheet_Multi_All;
begin
  TestWorksheet(wtkDeleteWorksheet, 10);
end;


{ Formula errors }

procedure TSpreadSingleFormulaTests.TestFormulaErrors(ATest: Integer);
type
  TTestCase = record
    Formula: string;
    Expected: String;
  end;
const
  // Cell A1 is 'abc' (string), A2 is 1.0 (number), A3 is '1' (string)
  TestCases: array[0..17] of TTestCase = (
  {0}  (Formula: 'A1+A2';      Expected: '#VALUE!'),
       (Formula: 'A1-A2';      Expected: '#VALUE!'),
       (Formula: 'A1*A2';      Expected: '#VALUE!'),
       (Formula: 'A1/A2';      Expected: '#VALUE!'),
       (Formula: 'A1^A2';      Expected: '#VALUE!'),
  {5}  (Formula: 'sin(A1)';    Expected: '#VALUE!'),
       (Formula: 'sin(A1)+A2'; Expected: '#VALUE!'),
       (Formula: 'A1=A2';      Expected: 'FALSE'),
       (Formula: 'A1<>A2';     Expected: 'TRUE'),
       (Formula: 'A1>A2';      Expected: 'FALSE'),
  {10} (Formula: 'A1<A2';      Expected: 'FALSE'),
       (Formula: 'A1>=A2';     Expected: 'FALSE'),
       (Formula: 'A1<=A2';     Expected: 'FALSE'),
       (Formula: '+A1';        Expected: 'abc'),
       (Formula: '-A1';        Expected: '#VALUE!'),
  {15} (Formula: 'A2+A3';      Expected: '2'),
       (Formula: 'A2=A3';      Expected: 'TRUE'),
       (Formula: '-A3';        Expected: '-1')
  );

var
  book: TsWorkbook;
  sheet: TsWorksheet;
  s: String;
begin
  book := TsWorkbook.Create;
  try
    book.Options := book.Options + [boAutoCalc];
    sheet := book.AddWorksheet('Test');
    sheet.WriteText(0, 0, 'abc');   // A1 = 'abc'
    sheet.WriteNumber(1, 0, 1.0);   // A2 = 1.0
    sheet.WriteText(2, 0, '1');     // A2 = '1';
    sheet.WriteFormula(0, 1, TestCases[ATest].Formula);
    s := sheet.ReadAsText(0, 1);
    CheckEquals(TestCases[ATest].Expected, s, 'Error value match, formula "' + sheet.ReadFormula(0, 1) + '"');
  finally
    book.Free;
  end;
end;

procedure TSpreadSingleFormulaTests.Error_AddStringNumber;
begin
  TestFormulaErrors(0);
end;

procedure TSpreadSingleFormulaTests.Error_SubtractStringNumber;
begin
  TestFormulaErrors(1);
end;

procedure TSpreadSingleFormulaTests.Error_MultiplyStringNumber;
begin
  TestFormulaErrors(2);
end;

procedure TSpreadSingleFormulaTests.Error_DivideStringNumber;
begin
  TestFormulaErrors(3);
end;

procedure TSpreadSingleFormulaTests.Error_PowerStringNumber;
begin
  TestFormulaErrors(4);
end;

procedure TSpreadSingleFormulaTests.Error_SinString;
begin
  TestFormulaErrors(5);
end;

procedure TSpreadSingleFormulaTests.Error_SinStringAddNumber;
begin
  TestFormulaErrors(6);
end;

procedure TSpreadSingleFormulaTests.Error_Equal;
begin
  TestFormulaErrors(7);
end;

procedure TSpreadSingleFormulaTests.Error_NotEqual;
begin
  TestFormulaErrors(8);
end;

procedure TSpreadSingleFormulaTests.Error_Greater;
begin
  TestFormulaErrors(9);
end;

procedure TSpreadSingleFormulaTests.Error_Smaller;
begin
  TestFormulaErrors(10);
end;

procedure TSpreadSingleFormulaTests.Error_GreaterEqual;
begin
  TestFormulaErrors(11);
end;

procedure TSpreadSingleFormulaTests.Error_LessEqual;
begin
  TestFormulaErrors(12);
end;

procedure TSpreadSingleFormulaTests.Error_UnaryPlusString;
begin
  TestFormulaErrors(13);
end;

procedure TSpreadSingleFormulaTests.Error_UnaryMinusString;
begin
  TestFormulaErrors(14);
end;

procedure TSpreadSingleFormulaTests.Add_Number_NumString;
begin
  TestFormulaErrors(15);
end;

procedure TSpreadSingleFormulaTests.Equal_Number_NumString;
begin
  TestFormulaErrors(16);
end;

procedure TSpreadSingleFormulaTests.UnaryMinusNumString;
begin
  TestFormulaErrors(17);
end;


{-------------------------------------------------------------------------------
                             TestInsDelRowCol

  Inserts/deletes a row or column and checks whether the references used in 
  formulas are correctly adapted. 
  
  Note: other row/col tests are contained in unit colrowtests.pas
-------------------------------------------------------------------------------}
procedure TSpreadSingleFormulaTests.TestInsDelRowCol(ATestIndex: Integer);
const
  VALUE1 = 'abc-1';
  VALUE2 = 'abc-2';
var
  wb: TsWorkbook;
  ws1: TsWorksheet;
  ws2: TsWorksheet;
begin
  wb := TsWorkbook.Create;
  try
    wb.Options := wb.Options + [boAutoCalc];
    
    ws1 := wb.AddWorksheet('Sheet1');
    ws1.WriteText(6, 4, VALUE1);      // text 'abc-1' in cell E7 of sheet1
    ws1.WriteFormula(1, 1, '=E7');    // formula =E7 in cell B2
    
    ws2 := wb.AddWorksheet('Sheet2');
    ws2.WriteText(6, 4, VALUE2);      // text 'abc-2' in cell E7 of sheet2
    ws2.WriteFormula(1, 1, '=E7');    // formula =E7 in cell B2
    ws2.writeFormula(1, 2, '=Sheet1!E7');  // 3d formula in cell C2 referring to sheet1
    
    CheckEquals(VALUE1, ws1.ReadAsText(1, 1), 'Read initial value mismatch, cell Sheet1!B2');
    CheckEquals('E7', ws1.ReadFormula(1, 1), 'Read initial formula mismatch, cell Sheet1!B2');
    CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read initial value mismatch, cell Sheet2!B2');
    CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read initial formula mismatch, cell Sheet2!B2');
    CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read initial value mismatch, cell Sheet2!C2');
    CheckEquals('Sheet1!E7', ws2.ReadFormula(1, 2), 'Read initial formula mismatch, cell Sheet2!C2');
    
    case ATestIndex of
      0:  // Insert row in sheet1 before formula and referenced cell
        begin
          ws1.InsertRow(0);
          CheckEquals(VALUE1, ws1.ReadAsText(2, 1), 'Read value mismatch, cell Sheet1!B2');
          CheckEquals('E8', ws1.ReadFormula(2, 1), 'Read formula mismatch, cell Sheet1!B2');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!E8', ws2.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet2!C2');
        end;
      1:  // Insert row in sheet1 after formula, but before referenced cell
        begin
          ws1.InsertRow(3);
          CheckEquals(VALUE1, ws1.ReadAsText(1, 1), 'Read value mismatch, cell Sheet1!B2');
          CheckEquals('E8', ws1.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet1!B2');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!E8', ws2.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet2!C2');
        end;
      2:  // Insert row in sheet1 after formula and referenced cell
        begin
          ws1.InsertRow(10);
          CheckEquals(VALUE1, ws1.ReadAsText(1, 1), 'Read initial value mismatch, cell Sheet1!B2');
          CheckEquals('E7', ws1.ReadFormula(1, 1), 'Read initial formula mismatch, cell Sheet1!B2');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read initial value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read initial formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read initial value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!E7', ws2.ReadFormula(1, 2), 'Read initial formula mismatch, cell Sheet2!C2');
        end;
      
      10:  // Insert column in sheet1 before formula and referenced cell
        begin
          ws1.InsertCol(0);    
          CheckEquals(VALUE1, ws1.ReadAsText(1, 2), 'Read value mismatch, cell Sheet1!B2');
          CheckEquals('F7', ws1.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet1!B2');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!F7', ws2.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet2!C2');
        end;
      11:  // Insert column in sheet1 after formula, but before referenced cell
        begin
          ws1.InsertCol(3);    
          CheckEquals(VALUE1, ws1.ReadAsText(1, 1), 'Read value mismatch, cell Sheet1!B2');
          CheckEquals('F7', ws1.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet1!B2');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!F7', ws2.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet2!C2');
        end;
      12:  // Insert column in sheet1 after formula and referenced cell
        begin
          ws1.InsertCol(10);
          CheckEquals(VALUE1, ws1.ReadAsText(1, 1), 'Read initial value mismatch, cell Sheet1!B2');
          CheckEquals('E7', ws1.ReadFormula(1, 1), 'Read initial formula mismatch, cell Sheet1!B2');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read initial value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read initial formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read initial value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!E7', ws2.ReadFormula(1, 2), 'Read initial formula mismatch, cell Sheet2!C2');
        end;
      
      20:  // Delete row from sheet1 before formula and referenced cell
        begin
          ws1.DeleteRow(0);
          CheckEquals(VALUE1, ws1.ReadAsText(0, 1), 'Read value mismatch, cell Sheet1!B1');
          CheckEquals('E6', ws1.ReadFormula(0, 1), 'Read formula mismatch, cell Sheet1!B1');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!E6', ws2.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet2!C2');
        end;
      21:  // Delete row from sheet1 after formula, but before referenced cell
        begin
          ws1.DeleteRow(3);
          CheckEquals(VALUE1, ws1.ReadAsText(1, 1), 'Read value mismatch, cell Sheet1!B2');
          CheckEquals('E6', ws1.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet1!B2');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!E6', ws2.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet2!C2');
        end;
      22:  // Delete row from sheet1 after formula and referenced cell
        begin
          ws1.DeleteRow(10);
          CheckEquals(VALUE1, ws1.ReadAsText(1, 1), 'Read value mismatch, cell Sheet1!B2');
          CheckEquals('E7', ws1.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet1!21');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!E7', ws2.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet2!C2');
        end;
      
      30:  // Delete column from sheet1 before formula and referenced cell
        begin
          ws1.DeleteCol(0);
          CheckEquals(VALUE1, ws1.ReadAsText(1, 0), 'Read value mismatch, cell Sheet1!A2');
          CheckEquals('D7', ws1.ReadFormula(1, 0), 'Read formula mismatch, cell Sheet1!A2');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!D7', ws2.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet2!C2');
        end;
      31:  // Delete column from sheet1 after formula, but before referenced cell
        begin
          ws1.DeleteCol(3);
          CheckEquals(VALUE1, ws1.ReadAsText(1, 1), 'Read value mismatch, cell Sheet1!A2');
          CheckEquals('D7', ws1.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet1!A2');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!D7', ws2.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet2!C2');
        end;
      32:  // Delete column from sheet1 after formula and referenced cell
        begin
          ws1.DeleteCol(10);
          CheckEquals(VALUE1, ws1.ReadAsText(1, 1), 'Read value mismatch, cell Sheet1!B2');
          CheckEquals('E7', ws1.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet1!21');
          CheckEquals(VALUE2, ws2.ReadAsText(1, 1), 'Read value mismatch, cell Sheet2!B2');
          CheckEquals('E7', ws2.ReadFormula(1, 1), 'Read formula mismatch, cell Sheet2!B2');
          CheckEquals(VALUE1, ws2.ReadAsText(1, 2), 'Read value mismatch, cell Sheet2!C2');
          CheckEquals('Sheet1!E7', ws2.ReadFormula(1, 2), 'Read formula mismatch, cell Sheet2!C2');
        end;
    end;
          
    // wb.WriteToFile('test.xlsx', sfOOXML, true);    // Activate for looking at file, e.g. with Excel
  finally
    wb.Free;
  end;
end;

procedure TSpreadSingleFormulaTests.InsertRow_BeforeFormula;
begin
  TestInsDelRowCol(0);
end;

procedure TSpreadSingleFormulaTests.InsertRow_AfterFormula;
begin
  TestInsDelRowCol(1);
end;

procedure TSpreadSingleFormulaTests.InsertRow_AfterAll;
begin
  TestInsDelRowCol(2);
end;

procedure TSpreadSingleFormulaTests.InsertCol_BeforeFormula;
begin
  TestInsDelRowCol(10);
end;

procedure TSpreadSingleFormulaTests.InsertCol_AfterFormula;
begin
  TestInsDelRowCol(11);
end;

procedure TSpreadSingleFormulaTests.InsertCol_AfterAll;
begin
  TestInsDelRowCol(12);
end;

procedure TSpreadSingleFormulaTests.DeleteRow_BeforeFormula;
begin
  TestInsDelRowCol(20);
end;

procedure TSpreadSingleFormulaTests.DeleteRow_AfterFormula;
begin
  TestInsDelRowCol(21);
end;

procedure TSpreadSingleFormulaTests.DeleteRow_AfterAll;
begin
  TestInsDelRowCol(22);
end;

procedure TSpreadSingleFormulaTests.DeleteCol_BeforeFormula;
begin
  TestInsDelRowCol(30);
end;

procedure TSpreadSingleFormulaTests.DeleteCol_AfterFormula;
begin
  TestInsDelRowCol(31);
end;

procedure TSpreadSingleFormulaTests.DeleteCol_AfterAll;
begin
  TestInsDelRowCol(32);
end;


initialization
  // Register to include these tests in a full run
  RegisterTest(TSpreadSingleFormulaTests);


end.

