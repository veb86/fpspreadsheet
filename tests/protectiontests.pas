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

    { BIFF2 protection tests }
    procedure TestWriteRead_BIFF2_WorkbookProtection_None;
    procedure TestWriteRead_BIFF2_WorkbookProtection_Struct;
    procedure TestWriteRead_BIFF2_WorkbookProtection_Win;
    procedure TestWriteRead_BIFF2_WorkbookProtection_StructWin;

    procedure TestWriteRead_BIFF2_WorksheetProtection_Default;
    procedure TestWriteRead_BIFF2_WorksheetProtection_Objects;

    procedure TestWriteRead_BIFF2_CellProtection;

    { BIFF5 protection tests }
    procedure TestWriteRead_BIFF5_WorkbookProtection_None;
    procedure TestWriteRead_BIFF5_WorkbookProtection_Struct;
    procedure TestWriteRead_BIFF5_WorkbookProtection_Win;
    procedure TestWriteRead_BIFF5_WorkbookProtection_StructWin;

    procedure TestWriteRead_BIFF5_WorksheetProtection_Default;
    procedure TestWriteRead_BIFF5_WorksheetProtection_SelectLockedCells;
    procedure TestWriteRead_BIFF5_WorksheetProtection_SelectUnlockedCells;
    procedure TestWriteRead_BIFF5_WorksheetProtection_Objects;

    procedure TestWriteRead_BIFF5_CellProtection;

    { BIFF8 protection tests }
    procedure TestWriteRead_BIFF8_WorkbookProtection_None;
    procedure TestWriteRead_BIFF8_WorkbookProtection_Struct;
    procedure TestWriteRead_BIFF8_WorkbookProtection_Win;
    procedure TestWriteRead_BIFF8_WorkbookProtection_StructWin;

    procedure TestWriteRead_BIFF8_WorksheetProtection_Default;
    procedure TestWriteRead_BIFF8_WorksheetProtection_SelectLockedCells;
    procedure TestWriteRead_BIFF8_WorksheetProtection_SelectUnlockedCells;
    procedure TestWriteRead_BIFF8_WorksheetProtection_Objects;

    procedure TestWriteRead_BIFF8_CellProtection;

    { OOXML protection tests }
    procedure TestWriteRead_OOXML_WorkbookProtection_None;
    procedure TestWriteRead_OOXML_WorkbookProtection_Struct;
    procedure TestWriteRead_OOXML_WorkbookProtection_Win;
    procedure TestWriteRead_OOXML_WorkbookProtection_StructWin;

    procedure TestWriteRead_OOXML_WorksheetProtection_Default;
    procedure TestWriteRead_OOXML_WorksheetProtection_FormatCells;
    procedure TestWriteRead_OOXML_WorksheetProtection_FormatColumns;
    procedure TestWriteRead_OOXML_WorksheetProtection_FormatRows;
    procedure TestWriteRead_OOXML_WorksheetProtection_DeleteColumns;
    procedure TestWriteRead_OOXML_WorksheetProtection_DeleteRows;
    procedure TestWriteRead_OOXML_WorksheetProtection_InsertColumns;
    procedure TestWriteRead_OOXML_WorksheetProtection_InsertHyperlinks;
    procedure TestWriteRead_OOXML_WorksheetProtection_InsertRows;
    procedure TestWriteRead_OOXML_WorksheetProtection_Sort;
    procedure TestWriteRead_OOXML_WorksheetProtection_SelectLockedCells;
    procedure TestWriteRead_OOXML_WorksheetProtection_SelectUnlockedCells;
    procedure TestWriteRead_OOXML_WorksheetProtection_Objects;

    procedure TestWriteRead_OOXML_CellProtection;

    { ODS protection tests }
    procedure TestWriteRead_ODS_WorkbookProtection_None;
    procedure TestWriteRead_ODS_WorkbookProtection_Struct;
    //procedure TestWriteRead_ODS_WorkbookProtection_Win;
    //procedure TestWriteRead_ODS_WorkbookProtection_StructWin;

    procedure TestWriteRead_ODS_WorksheetProtection_Default;
    procedure TestWriteRead_ODS_WorksheetProtection_SelectLockedCells;
    procedure TestWriteRead_ODS_WorksheetProtection_SelectUnlockedCells;

    procedure TestWriteRead_ODS_CellProtection;

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
    spInsertHyperlinks, spSort, spObjects,
    spSelectLockedCells, spSelectUnlockedCells
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
    expected := DEFAULT_SHEET_PROTECTION;
    case ACondition of
      0: ;
      1: Exclude(expected, spFormatCells);
      2: Exclude(expected, spFormatColumns);
      3: Exclude(expected, spFormatRows);
      4: Exclude(expected, spDeleteColumns);
      5: Exclude(expected, spDeleteRows);
      6: Exclude(expected, spInsertColumns);
      7: Exclude(expected, spInsertHyperlinks);
      8: Exclude(expected, spInsertRows);
      9: Exclude(expected, spSort);
     10: Exclude(expected, spSelectLockedCells);
     11: Exclude(expected, spSelectUnlockedCells);
     12: Exclude(expected, spObjects);
    end;
    {
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
     12: expected := [spObjects];
     13: expected := ALL_SHEET_PROTECTIONS;
    end;
    }
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
 //   if actual <> [] then actual := actual - [spCells];
    msg := 'Test saved worksheet protection mismatch: ';
    if actual <> expected then
      case ACondition of
        0: fail(msg + 'default protection');
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
       12: fail(msg + 'spObjects');
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
    // A1 --> lock cell
    cell := Myworksheet.WriteText(0, 0, 'Protected');
    MyWorksheet.WriteCellProtection(cell, [cpLockCell]);
    // B1 --> not protected at all
    cell := MyWorksheet.WriteText(1, 0, 'Not protected');
    MyWorksheet.WriteCellProtection(cell, []);
    // A2 --> lock cell & hide formulas
    cell := Myworksheet.WriteFormula(0, 1, '=A1');
    MyWorksheet.WriteCellProtection(cell, [cpLockCell, cpHideFormulas]);
    // B2 --> hide formula only
    cell := MyWorksheet.WriteFormula(1, 1, '=A2');
    Myworksheet.WriteCellProtection(Cell, [cpHideFormulas]);
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


{------------------------------------------------------------------------------}
{                          Tests for BIFF2 file format                         }
{------------------------------------------------------------------------------}

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF2_WorkbookProtection_None;
begin
  TestWriteRead_WorkbookProtection(sfExcel2, 0);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF2_WorkbookProtection_Struct;
begin
  TestWriteRead_WorkbookProtection(sfExcel2, 1);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF2_WorkbookProtection_Win;
begin
  TestWriteRead_WorkbookProtection(sfExcel2, 2);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF2_WorkbookProtection_StructWin;
begin
  TestWriteRead_WorkbookProtection(sfExcel2, 3);
end;


procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF2_WorksheetProtection_Default;
begin
  TestWriteRead_WorksheetProtection(sfExcel2, 0);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF2_WorksheetProtection_Objects;
begin
  TestWriteRead_WorksheetProtection(sfExcel2, 12);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF2_CellProtection;
begin
  TestWriteRead_CellProtection(sfExcel2);
end;


{------------------------------------------------------------------------------}
{                          Tests for BIFF5 file format                         }
{------------------------------------------------------------------------------}

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF5_WorkbookProtection_None;
begin
  TestWriteRead_WorkbookProtection(sfExcel5, 0);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF5_WorkbookProtection_Struct;
begin
  TestWriteRead_WorkbookProtection(sfExcel5, 1);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF5_WorkbookProtection_Win;
begin
  TestWriteRead_WorkbookProtection(sfExcel5, 2);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF5_WorkbookProtection_StructWin;
begin
  TestWriteRead_WorkbookProtection(sfExcel5, 3);
end;


procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF5_WorksheetProtection_Default;
begin
  TestWriteRead_WorksheetProtection(sfExcel5, 0);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF5_WorksheetProtection_SelectLockedCells;
begin
  TestWriteRead_WorksheetProtection(sfExcel5, 10);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF5_WorksheetProtection_SelectUnlockedCells;
begin
  TestWriteRead_WorksheetProtection(sfExcel5, 11);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF5_WorksheetProtection_Objects;
begin
  TestWriteRead_WorksheetProtection(sfExcel5, 12);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF5_CellProtection;
begin
  TestWriteRead_CellProtection(sfExcel5);
end;


{------------------------------------------------------------------------------}
{                          Tests for BIFF8 file format                         }
{------------------------------------------------------------------------------}

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF8_WorkbookProtection_None;
begin
  TestWriteRead_WorkbookProtection(sfExcel8, 0);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF8_WorkbookProtection_Struct;
begin
  TestWriteRead_WorkbookProtection(sfExcel8, 1);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF8_WorkbookProtection_Win;
begin
  TestWriteRead_WorkbookProtection(sfExcel8, 2);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF8_WorkbookProtection_StructWin;
begin
  TestWriteRead_WorkbookProtection(sfExcel8, 3);
end;


procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF8_WorksheetProtection_Default;
begin
  TestWriteRead_WorksheetProtection(sfExcel8, 0);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF8_WorksheetProtection_SelectLockedCells;
begin
  TestWriteRead_WorksheetProtection(sfExcel8, 10);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF8_WorksheetProtection_SelectUnlockedCells;
begin
  TestWriteRead_WorksheetProtection(sfExcel8, 11);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF8_WorksheetProtection_Objects;
begin
  TestWriteRead_WorksheetProtection(sfExcel8, 12);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_BIFF8_CellProtection;
begin
  TestWriteRead_CellProtection(sfExcel8);
end;


{------------------------------------------------------------------------------}
{                          Tests for OOXML file format                         }
{------------------------------------------------------------------------------}

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

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_Default;
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

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_InsertHyperlinks;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 7);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_InsertRows;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 8);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_Sort;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 9);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_SelectLockedCells;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 10);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_SelectUnlockedCells;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 11);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_WorksheetProtection_Objects;
begin
  TestWriteRead_WorksheetProtection(sfOOXML, 12);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_OOXML_CellProtection;
begin
  TestWriteRead_CellProtection(sfOOXML);
end;


{------------------------------------------------------------------------------}
{                       Tests for OpenDocument file format                     }
{------------------------------------------------------------------------------}

procedure TSpreadWriteReadProtectionTests.TestWriteRead_ODS_WorkbookProtection_None;
begin
  TestWriteRead_WorkbookProtection(sfOpenDocument, 0);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_ODS_WorkbookProtection_Struct;
begin
  TestWriteRead_WorkbookProtection(sfOpenDocument, 1);
end;
{
procedure TSpreadWriteReadProtectionTests.TestWriteRead_ODS_WorkbookProtection_Win;
begin
  TestWriteRead_WorkbookProtection(sfOpenDocument, 2);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_ODS_WorkbookProtection_StructWin;
begin
  TestWriteRead_WorkbookProtection(sfOpenDocument, 3);
end;}

procedure TSpreadWriteReadProtectionTests.TestWriteRead_ODS_WorksheetProtection_Default;
begin
  TestWriteRead_WorksheetProtection(sfOpenDocument, 0);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_ODS_WorksheetProtection_SelectLockedCells;
begin
  TestWriteRead_WorksheetProtection(sfOpenDocument, 10);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_ODS_WorksheetProtection_SelectUnlockedCells;
begin
  TestWriteRead_WorksheetProtection(sfOpenDocument, 11);
end;

procedure TSpreadWriteReadProtectionTests.TestWriteRead_ODS_CellProtection;
begin
  TestWriteRead_CellProtection(sfOpenDocument);
end;

initialization
  RegisterTest(TSpreadWriteReadProtectionTests);

end.

