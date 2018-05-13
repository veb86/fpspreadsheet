unit SingleFormulaTests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testutils, testregistry,
  fpstypes, fpsallformats, fpspreadsheet, fpsexprparser,
  xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  TFormulaTestKind = (ftkConstants, ftkCellConstant, ftkCells, ftkCellRange,
    ftkCellRangeSheet, ftkCellRangeSheetRange);

  { TSpreadDetailedFormulaFormula }
  TSpreadSingleFormulaTests = class(TTestCase)
  private
  protected
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestFloatFormula(AFormula: String; AExpected: Double;
      ATestKind: TFormulaTestKind; AFormat: TsSpreadsheetFormat);

  published
    procedure AddConst_BIFF2;
    procedure AddConst_BIFF5;
    procedure AddConst_BIFF8;
    procedure AddConst_OOXML;
    procedure AddConst_ODS;

    procedure AddCells_BIFF2;
    procedure AddCells_BIFF5;
    procedure AddCells_BIFF8;
    procedure AddCells_OOXML;
    procedure AddCells_ODS;

    procedure SumRange_BIFF2;
    procedure SumRange_BIFF5;
    procedure SumRange_BIFF8;
    procedure SumRange_OOXML;

    procedure SumSheetRange_BIFF5;  // no 3d ranges for BIFF2
    procedure SumSheetRange_BIFF8;
    procedure SumSheetRange_OOXML;

  end;

implementation

uses
 {$IFDEF FORMULADEBUG}
  LazLogger,
 {$ENDIF}
  math, typinfo, lazUTF8, fpsUtils;


{ TSpreadExtendedFormulaTests }

procedure TSpreadSingleFormulaTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadSingleFormulaTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadSingleFormulaTests.TestFloatFormula(AFormula: String;
  AExpected: Double; ATestKind: TFormulaTestKind; AFormat: TsSpreadsheetFormat);
const
  SHEET1 = 'Sheet1';
  SHEET2 = 'Sheet2';
  SHEET3 = 'Sheet3';
  TESTCELL_ROW = 1;
  TESTCELL_COL = 2;
var
  worksheet: TsWorksheet;
  othersheet: TsWorksheet;
  workbook: TsWorkbook;
  TempFile: string; //write xls/xml to this file and read back from it
  cell: PCell;
  actualformula: String;
  actualValue: Double;
begin
  TempFile := GetTempFileName;

  // Create test workbook and write test formula and needed cells
  workbook := TsWorkbook.Create;
  try
    workbook.Options := workbook.Options + [boCalcBeforeSaving, boAutoCalc];
    workSheet:= workBook.AddWorksheet(SHEET1);

    if ATestKind <> ftkConstants then begin
      // Write cells used by the formula
      worksheet.WriteNumber(2, 2, 1.0);   // C3
      worksheet.WriteNumber(3, 2, -2.0);  // C4
      worksheet.WriteNumber(4, 2, 1.5);   // C5
      worksheet.WriteNumber(2, 3, 15.0);  // D3
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
      othersheet.WriteNumber(2, 2, 100.0);   // Sheet3C3
      othersheet.WriteNumber(3, 2, -200.0);  // Sheet3!C4
      othersheet.WriteNumber(4, 2, 150.0);   // Sheet3!C5
      othersheet.WriteNumber(2, 3, 1500.0);  // Sheet3!D5
    end;

    // Write the formula
    cell := worksheet.WriteFormula(TESTCELL_ROW, TESTCELL_COL, AFormula);

    // Read formula before saving
    actualFormula := cell^.Formulavalue;
    CheckEquals(AFormula, actualFormula, 'Unsaved formula text mismatch');

    // Read calculated value before saving
    actualvalue := worksheet.ReadAsNumber(TESTCELL_ROW, TESTCELL_COL);
    CheckEquals(AExpected, actualvalue, 'Unsaved calculated value mismatch');

    // Save
    workbook.WriteToFile(TempFile, AFormat, true);
  finally
    workbook.Free;
  end;

  // Read file
  workbook := TsWorkbook.Create;
  try
    workbook.Options := workbook.Options + [boReadFormulas, boAutoCalc];
    workbook.ReadFromFile(TempFile, AFormat);
    worksheet := workbook.GetFirstWorksheet;

    // Read calculated formula value
    actualvalue := worksheet.ReadAsNumber(TESTCELL_ROW, TESTCELL_COL);
    CheckEquals(AExpected, actualValue, 'Saved calculated value mismatch');

    cell := worksheet.FindCell(TESTCELL_ROW, TESTCELL_COL);
    actualformula := cell^.FormulaValue;
    CheckEquals(AFormula, actualformula, 'Saved formula text mismatch.');
  finally
    workbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadSingleFormulaTests.AddConst_BIFF2;
begin
  TestFloatFormula('1+1', 2, ftkConstants, sfExcel2);
end;

procedure TSpreadSingleFormulaTests.AddConst_BIFF5;
begin
  TestFloatFormula('1+1', 2, ftkConstants, sfExcel5);
end;

procedure TSpreadSingleFormulaTests.AddConst_BIFF8;
begin
  TestFloatFormula('1+1', 2, ftkConstants, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.AddConst_OOXML;
begin
  TestFloatFormula('1+1', 2, ftkConstants, sfOOXML);
end;

procedure TSpreadSingleFormulaTests.AddConst_ODS;
begin
  TestFloatFormula('1+1', 2, ftkConstants, sfOpenDocument);
end;

{---------------}

procedure TSpreadSingleFormulaTests.AddCells_BIFF2;
begin
  TestFloatFormula('C3+C4', -1.0, ftkCells, sfExcel2);
end;

procedure TSpreadSingleFormulaTests.AddCells_BIFF5;
begin
  TestFloatFormula('C3+C4', -1.0, ftkCells, sfExcel5);
end;

procedure TSpreadSingleFormulaTests.AddCells_BIFF8;
begin
  TestFloatFormula('C3+C4', -1.0, ftkCells, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.AddCells_OOXML;
begin
  TestFloatFormula('C3+C4', -1.0, ftkCells, sfOOXML);
end;

procedure TSpreadSingleFormulaTests.AddCells_ODS;
begin
  TestFloatFormula('C3+C4', -1.0, ftkCells, sfOpenDocument);
end;

{ ------ }

procedure TSpreadSingleFormulaTests.SumRange_BIFF2;
begin
  TestFloatFormula('SUM(C3:C5)', 0.5, ftkCellRange, sfExcel2);
end;

procedure TSpreadSingleFormulaTests.SumRange_BIFF5;
begin
  TestFloatFormula('SUM(C3:C5)', 0.5, ftkCellRange, sfExcel5);
end;

procedure TSpreadSingleFormulaTests.SumRange_BIFF8;
begin
  TestFloatFormula('SUM(C3:C5)', 0.5, ftkCellRange, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.SumRange_OOXML;
begin
  TestFloatFormula('SUM(C3:C5)', 0.5, ftkCellRange, sfOOXML);
end;

{ ---- }

procedure TSpreadSingleFormulaTests.SumSheetRange_BIFF5;
begin
  TestFloatFormula('SUM(Sheet2!C3:C5)', 5.0, ftkCellRangeSheet, sfExcel5);
end;

procedure TSpreadSingleFormulaTests.SumSheetRange_BIFF8;
begin
  TestFloatFormula('SUM(Sheet2!C3:C5)', 5.0, ftkCellRangeSheet, sfExcel8);
end;

procedure TSpreadSingleFormulaTests.SumSheetRange_OOXML;
begin
  TestFloatFormula('SUM(Sheet2!C3:C5)', 5.0, ftkCellRangeSheet, sfOOXML);
end;

{ ---- }


initialization
  // Register to include these tests in a full run
  RegisterTest(TSpreadSingleFormulaTests);

end.

