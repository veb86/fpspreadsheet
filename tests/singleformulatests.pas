unit SingleFormulaTests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, fpsexprparser,
  xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  TFormulaTestKind = (ftkConstants, ftkCellConstant, ftkCells, ftkCellRange,
    ftkCellRangeSheet, ftkCellRangeSheetRange);
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
    procedure SumRange_ODS;

    procedure SumSheetRange_BIFF5;  // no 3d ranges for BIFF2
    procedure SumSheetRange_BIFF8;
    procedure SumSheetRange_OOXML;
    procedure SumSheetRange_ODS;

    procedure SumMultiSheetRange_BIFF5;
    procedure SumMultiSheetRange_BIFF8;
    procedure SumMultiSheetRange_OOXML;
    procedure SumMultiSheetRange_ODS;

    procedure SumMultiSheetRange_FlippedCells_BIFF8;
    procedure SumMultiSheetRange_FlippedCells_OOXML;
    procedure SumMultiSheetRange_FlippedSheets_OOXML;
    procedure SumMultiSheetRange_FlippedSheetsAndCells_OOXML;
    procedure SumMultiSheetRange_FlippedSheetsAndCells_ODS;

    procedure NonExistantSheet_BIFF5;
    procedure NonExistantSheet_BIFF8;
    procedure NonExistantSheet_OOXML;
    procedure NonExistantSheet_ODS;

    procedure NonExistantSheetRange_BIFF5;
    procedure NonExistantSheetRange_BIFF8;
    procedure NonExistantSheetRange_OOXML;
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

  end;

implementation

uses
 {$IFDEF FORMULADEBUG}
  LazLogger,
 {$ENDIF}
  Math, typinfo, lazUTF8, fpsUtils;


{ TSpreadExtendedFormulaTests }

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
  TESTCELL_ROW = 1;
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
      workbook.FormatSettings.DecimalSeparator := '.';
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
      workbook.FormatSettings.DecimalSeparator := '.';
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

procedure TSpreadSingleFormulaTests.AddCells_ODS;
begin
  TestFormula('C3+C4', '-1', ftkCells, sfOpenDocument);
end;

{ ------ }

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




initialization
  // Register to include these tests in a full run
  RegisterTest(TSpreadSingleFormulaTests);

end.

