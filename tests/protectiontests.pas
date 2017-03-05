{ Protection tests
  These unit tests are writing out to and reading back from file.
}

unit protectiontests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, //xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadWriteReadProtectionTests }
  //Write to xls/xml file and read back
  TSpreadWriteReadProtectionTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteRead_WorkbookProtection(AFormat: TsSpreadsheetFormat;
      ACondition: Integer);
    procedure TestWriteRead_WorksheetProtection(AFormat: TsSpreadsheetFormat;
      ACondition: Integer);
    procedure TestWriteRead_CellProtection(AFormat: TsSpreadsheetFormat);
  published
    // Writes out protection & reads back.

    { OOXML protection tests }
    procedure TestWriteRead_OOXML_WorkbookProtection_None;
    procedure TestWriteRead_OOXML_WorkbookProtection_Struct;
    procedure TestWriteRead_OOXML_WorkbookProtection_Win;
    procedure TestWriteRead_OOXML_WorkbookProtection_StructWin;

    procedure TestWriteRead_OOXML_WorksheetProtection_None;
    procedure TestWriteRead_OOXML_WorksheetProtection_FormatCells;
    procedure TestWriteRead_OOXML_WorksheetProtection_FormatColumns;
    procedure TestWriteRead_OOXML_WorksheetProtection_FormatRows;
    procedure TestWriteRead_OOXML_WorksheetProtection_DeleteColumns;
    procedure TestWriteRead_OOXML_WorksheetProtection_DeleteRows;
    procedure TestWriteRead_OOXML_WorksheetProtection_InsertColumns;
    procedure TestWriteRead_OOXML_WorksheetProtection_InsertRows;
    procedure TestWriteRead_OOXML_WorksheetProtection_InsertHyperlinks;
    procedure TestWriteRead_OOXML_WorksheetProtection_Sheet;
    procedure TestWriteRead_OOXML_WorksheetProtection_Sort;
    procedure TestWriteRead_OOXML_WorksheetProtection_SelectLockedCells;
    procedure TestWriteRead_OOXML_WorksheetProtection_SelectUnlockedCells;
    procedure TestWriteRead_OOXML_WorksheetProtection_All;

    procedure TestWriteRead_OOXML_CellProtection;

  end;

implementation

const
  ProtectionSheet = 'Protection';

{ TSpreadWriteReadProtectionTests }

procedure TSpreadWriteReadProtectionTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadProtectionTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_WorkbookProtection(
  AFormat: TsSpreadsheetFormat; ACondition: Integer);
var
  MyWorkbook: TsWorkbook;
  TempFile: string; //write xls/xml to this file and read back from it
  expected, actual: TsWorkbookProtections;
  msg: String;
begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkBook.AddWorksheet(ProtectionSheet);
    case ACondition of
      0: expected := [];
      1: expected := [bpLockStructure];
      2: expected := [bpLockWindows];
      3: expected := [bpLockStructure, bpLockWindows];
    end;
    MyWorkbook.Protection := expected;
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    actual := MyWorkbook.Protection;
    if actual <> expected then begin
      msg := 'Test saved workbook protection mismatch: ';
      case ACondition of
        0: fail(msg + 'no protection');
        1: fail(msg + 'bpLockStructure');
        2: fail(msg + 'bpLockWindows');
        3: fail(msg + 'bpLockStructure, bpLockWindows');
      end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_WorksheetProtection(
  AFormat: TsSpreadsheetFormat; ACondition: Integer);
const
  ALL_SHEET_PROTECTIONS = [
    spFormatCells, spFormatColumns, spFormatRows,
    spDeleteColumns, spDeleteRows, spInsertColumns, spInsertRows,
    spInsertHyperlinks, spSort, spSelectLockedCells,
    spSelectUnlockedCells
  ];    // NOTE: spCells is handled separately
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  TempFile: string; //write xls/xml to this file and read back from it
  expected, actual: TsWorksheetProtections;
  msg: String;
begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorksheet := MyWorkBook.AddWorksheet(ProtectionSheet);
    case ACondition of
      0: expected := [];
      1: expected := [spFormatCells];
      2: expected := [spFormatColumns];
      3: expected := [spFormatRows];
      4: expected := [spDeleteColumns];
      5: expected := [spDeleteRows];
      6: expected := [spInsertColumns];
      7: expected := [spInsertHyperlinks];
      8: expected := [spInsertRows];
      9: expected := [spSort];
     10: expected := [spSelectLockedCells];
     11: expected := [spSelectUnlockedCells];
     12: expected := ALL_SHEET_PROTECTIONS;
    end;
    MyWorksheet.Protection := expected;
    MyWorksheet.Protect(true);
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    MyWorksheet := MyWorkbook.GetFirstWorksheet;
    if (ACondition > 0) and not MyWorksheet.IsProtected then
      fail(msg + 'Sheet protection not active');

    actual := MyWorksheet.Protection;
    if actual <> [] then actual := actual - [spCells];
    msg := 'Test saved worksheet protection mismatch: ';
    if actual <> expected then
      case ACondition of
        0: fail(msg + 'no protection');
        1: fail(msg + 'spFormatCells');
        2: fail(msg + 'spFormatColumns');
        3: fail(msg + 'spFormatRows');
        4: fail(msg + 'spDeleteColumns');
        5: fail(msg + 'spDeleteRows');
        6: fail(msg + 'spInsertColumns');
        7: fail(msg + 'spInsertHyperlinks');
        8: fail(msg + 'spInsertRows');
        9: fail(msg + 'spSort');
       10: fail(msg + 'spSelectLockedCells');
       11: fail(msg + 'spSelectUnlockedCells');
       12: fail(msg + 'all options');
      end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_CellProtection(
  AFormat: TsSpreadsheetFormat);
var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  cell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
  expected, actual: TsCellProtections;
  msg: String;
begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorksheet := MyWorkBook.AddWorksheet(ProtectionSheet);
    cell := Myworksheet.WriteText(0, 0, 'Protected');
    MyWorksheet.WriteCellProtection(cell, [cpLockCell]);
    cell := MyWorksheet.WriteText(1, 0, 'Not protected');
    MyWorksheet.WriteCellProtection(cell, []);
    cell := Myworksheet.WriteFormula(0, 1, '=A1');
    MyWorksheet.WriteCellProtection(cell, [cpLockCell, cpHideFormulas]);
    cell := MyWorksheet.WriteFormula(1, 1, '=A2');
    Myworksheet.WriteCellProtection(Cell, [cpHideFormulas]);
    MyWorksheet.Protect(true);
    // NOTE: FPSpreadsheet does not enforce these actions. They are only written
    // to the file for the Office application.
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    MyWorksheet := MyWorkbook.GetFirstWorksheet;
    msg := 'Test saved worksheet protection mismatch: ';
    if not MyWorksheet.IsProtected then begin
      fail(msg + 'Sheet protection not active');
      exit;
    end;

    cell := MyWorksheet.FindCell(0, 0);
    if cell = nil then
      fail(msg + 'Protected cell A1 not found.');
    actual := MyWorksheet.ReadCellProtection(cell);
    if actual <> [cpLockCell] then
      fail(msg + 'cell A1 protection = [cpLockCells]');

    cell := MyWorksheet.FindCell(1, 0);
    if cell = nil then
      fail(msg + 'Unprotected cell A2 not found.');
    actual := MyWorksheet.ReadCellProtection(cell);
    if actual <> [] then
      fail(msg + 'cell A2 protection = []');

    cell := Myworksheet.FindCell(0, 1);
    if cell = nil then
      fail(msg + 'Cell B1 not found.');
    actual := MyWorksheet.ReadCellProtection(cell);
    if actual <> [cpLockCell, cpHideFormulas] then
      fail(msg + 'cell B1 protection = [cpLockCells, cpHideFormulas]');

    cell := MyWorksheet.FindCell(1, 1);
    if cell = nil then
      fail(msg + 'Cell B2 protection = [cpHideFormulas]');

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;


{ Tests for OOXML file format }
procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorkbookProtection_None;
begin
  TestWriteRead_WorkbookProtection(sfOOXML, 0);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorkbookProtection_Struct;
begin
  TestWriteRead_WorkbookProtection(sfOOXML, 1);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorkbookProtection_Win;
begin
  TestWriteRead_WorkbookProtection(sfOOXML, 2);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorkbookProtection_StructWin;
begin
  TestWriteRead_WorkbookProtection(sfOOXML, 3);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_None;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 0);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_FormatCells;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 1);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_FormatColumns;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 2);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_FormatRows;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 3);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_DeleteColumns;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 4);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_DeleteRows;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 5);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_InsertColumns;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 6);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_InsertRows;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 7);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_InsertHyperlinks;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 8);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_Sheet;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 9);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_Sort;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 10);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_SelectLockedCells;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 11);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_SelectUnlockedCells;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 12);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_All;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 13);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_CellProtection;
begin
  TestWriteRead_CellProtection(sfOOXML);
end;

initialization
  RegisterTest(TSpreadWriteReadProtectionTests);

end.

