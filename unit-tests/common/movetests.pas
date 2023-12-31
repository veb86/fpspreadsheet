unit movetests;

{$mode objfpc}{$H+}

interface
{ Tests for copying cells
  NOTE: The code in these tests is very fragile because the test results are
  hard-coded. Any modification in "InitCopyData" must be carefully verified!
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8, fpsopendocument, {and a project requirement for lclbase for utf8 handling}
  testsutility;

type
  { TSpreadMoveTests }
  TSpreadMoveTests = class(TTestCase)
  private

  protected
    procedure Test_MoveCell(ATestKind: Integer);

  published
    procedure Test_MoveCell_Value;
    procedure Test_MoveCell_Format;
    procedure Test_MoveCell_Comment;
    procedure Test_MoveCell_Hyperlink;
    procedure Test_MoveCell_Formula_REL;
    procedure Test_MoveCell_Formula_ABS;
    procedure Test_MoveCell_FormulaRef_REL;
    procedure Test_MoveCell_FormulaRef_ABS;
    
    procedure Test_MoveCell_CircRef;
    procedure Test_MoveCell_OverwriteFormula;
    procedure Test_MoveCell_EmptyToValue;
    procedure Test_MoveCell_EmptyToFormula;
  end;

implementation

uses
  fpsutils;

const
  MoveTestSheet = 'Move';


{ TSpreadMoveTests }

{ In this test an occupied cell is moved to a different location. 
  ATestKind = 1: cell contains only a value
              2: cell contains also a format
              3: cell contains also a comment
              4: cell contains also a hyperlink
              5: cell contains a formula with relative reference to another cell
              6: like 5, but absolute reference.
              7: there is another cell with a formula pointing to the moved cell 
                 (relative reference). 
              8: like 7, but absolute reference. }
procedure TSpreadMoveTests.Test_MoveCell(ATestKind: Integer);
const
  SRC_ROW = 0;
  SRC_COL = 0;
  DEST_ROW = 11;
  DEST_COL = 6;
var
  worksheet: TsWorksheet;
  workbook: TsWorkbook;
  src_cell: PCell = nil;
  fmla_cell: PCell = nil;
  dest_cell: PCell = nil;
  nf: TsNumberFormat;
  nfs: String;
  hyperlink: PsHyperlink;
begin
  workbook := TsWorkbook.Create;
  try
    workbook.Options := workbook.Options + [boAutoCalc];

    worksheet := workBook.AddWorksheet(MoveTestSheet);
    
    // Prepare the worksheet in which a cell is moved:
    // The source cell is A1, and it is moved to G10
    src_cell := worksheet.WriteNumber(SRC_ROW, SRC_COL, 3.141592);      // A1
    case ATestKind of
      1: ;                                                  // just the value
      2: worksheet.WriteNumberFormat(src_cell, nfFixed, 2); // value + formatting
      3: worksheet.WriteComment(src_cell, 'test');          // value + comment
      4: worksheet.WriteHyperlink(src_cell, 'B2');          // value + hyperlink
      5: begin     // The test cell constains a formula pointing the B2
           worksheet.WriteNumber(1, 1, 3.141592);
           src_cell := worksheet.WriteFormula(SRC_ROW, SRC_COL, 'B2');
         end;  
      6: begin    // like 5, just absolute reference
           worksheet.WriteNumber(1, 1, 3.141592);
           src_cell := worksheet.WriteFormula(SRC_ROW, SRC_COL, '$B$2');
         end;  
      // In the two last tests, there is a formula pointing to the test cell.
      // It must be updated when the test cell is moved.
      7: fmla_cell := worksheet.WriteFormula(2, 2, 'A1');   // rel reference
      8: fmla_cell := worksheet.WriteFormula(2, 2, '$A$1'); // abs reference
    end;
    
    // Now move the cell
    worksheet.MoveCell(src_cell, DEST_ROW, DEST_COL);   // move cell (in A1) to G10
    
    // Check removal of source cell
    CheckEquals(true, worksheet.FindCell(SRC_ROW, SRC_COL) = nil, 'Source cell not removed');
    
    // Check existence of target cell
    dest_cell := worksheet.FindCell(DEST_ROW, DEST_COL);
    CheckEquals(true, dest_cell <> nil, 'Moved cell not found.');
    
    // Check value in target cell
    if not (ATestKind in [5, 6]) then
      CheckEquals(3.141592, worksheet.ReadAsNumber(dest_cell), 1E-9, 'Cell value mismatch');
    
    case ATestKind of
      1: ;
      2: begin
           worksheet.ReadNumFormat(dest_cell, nf, nfs);
           CheckEquals(Integer(nfFixed), Integer(nf), 'Number format mismatch');
           CheckEquals('0.00', nfs, 'Number format string mismatch');
         end;
      3: CheckEquals('test', worksheet.ReadComment(dest_cell), 'Comment mismatch');
      4: begin
           hyperlink := worksheet.FindHyperlink(dest_cell);
           CheckEquals(true, hyperlink <> nil, 'Hyperlink not found');
           CheckEquals('B2', hyperlink^.Target, 'hyperlink target mismatch');
         end;
      5: CheckEquals('B2', worksheet.ReadFormula(dest_cell), 'Moved formula mismatch');
      6: CheckEquals('$B$2', worksheet.ReadFormula(dest_cell), 'Moved formula mismatch');
      7: CheckEquals('G12', worksheet.ReadFormula(fmla_cell), 'Referencing formula mismatch');
      8: CheckEquals('$G$12', worksheet.ReadFormula(fmla_cell), 'Referencing formula mismatch');
    end;
    
  finally
    workbook.Free;
  end;
end;

{ Move cell with a number, no attached data. }
procedure TSpreadMoveTests.Test_MoveCell_Value;
begin
  Test_MoveCell(1);
end;

{ Move cell with a number, the cell contains a number format. }
procedure TSpreadMoveTests.Test_MoveCell_Format;
begin
  Test_MoveCell(2);
end;

{ Move cell with a number and comment. }
procedure TSpreadMoveTests.Test_MoveCell_Comment;
begin
  Test_MoveCell(3);
end;

{ Move cell with a number and hyperlink. }
procedure TSpreadMoveTests.Test_MoveCell_Hyperlink;
begin
  Test_MoveCell(4);
end;

{ Move cell with a formula (relative reference) }
procedure TSpreadMoveTests.Test_MoveCell_Formula_REL;
begin
  Test_MoveCell(5);
end;

{ Move cell with a formula (absolute reference). }
procedure TSpreadMoveTests.Test_MoveCell_Formula_ABS;
begin
  Test_MoveCell(6);
end;

{ Move cell with a number, a formula points to the cell with a relative reference. }
procedure TSpreadMoveTests.Test_MoveCell_FormulaRef_REL;
begin
  Test_MoveCell(7);
end;

{ Move cell with a number, a formula points to the cell with an absolute reference. }
procedure TSpreadMoveTests.Test_MoveCell_FormulaRef_ABS;
begin
  Test_MoveCell(8);
end;

{==============================================================================}

{ In the following test an occupied cell with a formula is moved to a location
  referenced by the formula. 
  This must result in a circular reference error. }
procedure TSpreadMoveTests.Test_MoveCell_CircRef;
const
  VALUE_CELL_ROW = 0;        // A1
  VALUE_CELL_COL = 0;
  FORMULA_CELL_ROW = 11;     // F10
  FORMULA_CELL_COL = 6;
var
  worksheet: TsWorksheet;
  workbook: TsWorkbook;
  formula_cell: PCell = nil;
  dest_cell: PCell = nil;
begin
  workbook := TsWorkbook.Create;
  try
    workbook.Options := workbook.Options + [boAutoCalc];

    worksheet := workBook.AddWorksheet(MoveTestSheet);
    
    // Prepare the worksheet in which a cell is moved.
    // The value cell is A1, the formula cell is F10 and it points to A1
    worksheet.WriteText(VALUE_CELL_ROW, VALUE_CELL_COL, 'abc');   // A1
    formula_cell := worksheet.WriteFormula(FORMULA_CELL_ROW, FORMULA_CELL_COL, 'A1');
    
    // Move the formula cell to overwrite the value cell
    try
      worksheet.MoveCell(formula_cell, VALUE_CELL_ROW, VALUE_CELL_COL);
      dest_cell := worksheet.FindCell(VALUE_CELL_ROW, VALUE_CELL_COL);
    except
    end;
    
    // The destination cell should contain a #REF! error
    CheckEquals(true, dest_cell^.ErrorValue = errIllegalRef, 'Circular reference not detected.');
    
  finally
    workbook.Free;
  end;
end;

{ In the following test an occupied cell with a value formula is moved to 
  a location with a formula cell pointing to the moved value cell.
  This operation must delete the formula after moving. }
procedure TSpreadMoveTests.Test_MoveCell_OverwriteFormula;
const
  VALUE_CELL_ROW = 0;        // A1
  VALUE_CELL_COL = 0;
  FORMULA_CELL_ROW = 11;     // F10
  FORMULA_CELL_COL = 6;
var
  worksheet: TsWorksheet;
  workbook: TsWorkbook;
  value_cell: PCell = nil;
  dest_cell: PCell = nil;
begin
  workbook := TsWorkbook.Create;
  try
    workbook.Options := workbook.Options + [boAutoCalc];

    worksheet := workBook.AddWorksheet(MoveTestSheet);
    
    // Prepare the worksheet in which a cell is moved.
    // The value cell is A1, the formula cell is F10 and it points to A1
    value_cell := worksheet.WriteText(VALUE_CELL_ROW, VALUE_CELL_COL, 'abc');   // A1
    worksheet.WriteFormula(FORMULA_CELL_ROW, FORMULA_CELL_COL, 'A1');
    
    // Move the value cell to overwrite the formula cell
    try
      worksheet.MoveCell(value_cell, FORMULA_CELL_ROW, FORMULA_CELL_COL);
      dest_cell := worksheet.FindCell(FORMULA_CELL_ROW, FORMULA_CELL_COL);
    except
    end;
    
    // The destination cell should not contain a formula any more.
    CheckEquals(false, HasFormula(dest_cell), 'Formula has not been removed.');
    // Check value at destination after moving
    CheckEquals('abc', worksheet.ReadAsText(dest_cell), 'Moved value mismatch.');
    // Check value at source after moving
    CheckEquals('', worksheet.ReadAsText(value_cell), 'Source value mismatch after moving.');
    
  finally
    workbook.Free;
  end;
end;
  
{ In the following test an empty cell is moved to a location with a value cell. 
  This operation must delete the value in the destination cell after moving. }
procedure TSpreadMoveTests.Test_MoveCell_EmptyToValue;
const
  SOURCE_CELL_ROW = 0;   // A1
  SOURCE_CELL_COL = 0;
  DEST_CELL_ROW = 2;     // C3
  DEST_CELL_COL = 2; 
var
  worksheet: TsWorksheet;
  workbook: TsWorkbook;
  src_cell: PCell = nil;
  dest_cell: PCell = nil;
begin
  workbook := TsWorkbook.Create;
  try
    workbook.Options := workbook.Options + [boAutoCalc];

    worksheet := workBook.AddWorksheet(MoveTestSheet);
    
    // Prepare the worksheet in which an empty cell is moved.
    src_cell := nil;     // A1
    dest_cell := worksheet.WriteText(DEST_CELL_ROW, DEST_CELL_COL, 'abc');   // C3
    
    // Move the source cell to overwrite the value cell
    try
      worksheet.MoveCell(src_cell, DEST_CELL_ROW, DEST_CELL_COL);
      dest_cell := worksheet.FindCell(DEST_CELL_ROW, DEST_CELL_COL);
    except
    end;
    
    // The destination cell should be empty.
    CheckEquals(true, dest_cell = nil, 'Destination cell nas not been deleted.');
    
  finally
    workbook.Free;
  end;
end;
  
{ In the following test an empty cell is moved to a location with a formula cell.
  This operation must delete the destination cell after moving. In particular, 
  there must not be a formula any more. }
procedure TSpreadMoveTests.Test_MoveCell_EmptyToFormula;
const
  SOURCE_CELL_ROW = 0;   // A1
  SOURCE_CELL_COL = 0;
  DEST_CELL_ROW = 2;     // C3
  DEST_CELL_COL = 2; 
var
  worksheet: TsWorksheet;
  workbook: TsWorkbook;
  src_cell: PCell = nil;
  dest_cell: PCell = nil;
begin
  workbook := TsWorkbook.Create;
  try
    workbook.Options := workbook.Options + [boAutoCalc];

    worksheet := workBook.AddWorksheet(MoveTestSheet);
    
    // Prepare the worksheet in which a cell is moved.
    // The value cell is A1, the formula cell is B2 and it points to A1
    src_cell := nil;     // A1
    dest_cell := worksheet.WriteFormula(DEST_CELL_ROW, DEST_CELL_COL, 'PI()');   // C3
    
    // Move the empty source cell to overwrite the formula cell
    try
      worksheet.MoveCell(src_cell, DEST_CELL_ROW, DEST_CELL_COL);
      dest_cell := worksheet.FindCell(DEST_CELL_ROW, DEST_CELL_COL);
    except
    end;
    
    // The destination cell should be empty.
    CheckEquals(false, HasFormula(dest_cell), 'Destination cell still contains a formula.');
    CheckEquals(true, dest_cell = nil, 'Destination cell has not been deleted.');
    
  finally
    workbook.Free;
  end;
end;
  
initialization
  RegisterTest(TSpreadMoveTests);

end.

