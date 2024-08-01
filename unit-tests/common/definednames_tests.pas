{ Defined names tests
  These unit tests are writing out to and reading back from file.
}

unit definednames_tests;

{$mode objfpc}{$H+}

interface

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  TDefinedNamesTestType = (dnNamedCellIsConst, dnNamedCellIsFormula, dnNamedRangeIsFormula);

type
  { TSpreadWriteReadDefinedNamesTests }
  //Write to xlsx/ods file and read back
  TSpreadWriteReadDefinedNamesTests = class(TTestCase)
  private
  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestWriteRead_DefinedNames(AFormat: TsSpreadsheetFormat;
      const ATestType: TDefinedNamesTestType);
  published
    // Writes out defined names & reads back.

    { OpenDocument defined tests }
    procedure TestWriteRead_ODS_Cell_Comment;
    procedure TestWriteRead_ODS_Cell_Formula;
    procedure TestWriteRead_ODS_Range_Formula;

    { OOXML comment tests }
    procedure TestWriteRead_XLSX_Cell_Comment;
    procedure TestWriteRead_XLSX_Cell_Formula;
    procedure TestWriteRead_XLSX_Range_Formula;

  end;

implementation


{ TSpreadWriteReadDefinedNamesTests }

procedure TSpreadWriteReadDefinedNamesTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadWriteReadDefinedNamesTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadWriteReadDefinedNamesTests.TestWriteRead_DefinedNames(
  AFormat: TsSpreadsheetFormat; const ATestType: TDefinedNamesTestType);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  expected, actual: Double;
  formula: String;
  col, row: Integer;
  cell: PCell;
  TempFile: string; //write xls/xml to this file and read back from it
begin
  TempFile := GetTempFileName;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet('Data');

    // Prepare the test cells: numbers 0, 10, 20, 30, 40 in cells A1..A5
    col := 0;
    for row := 0 to 4 do
      MyWorksheet.WriteNumber(row, col, row*10);

    // Add test cases
    case ATestType of
      dnNamedCellIsConst:
        MyWorksheet.DefinedNames.Add('const', 0, 0, 2, 0, 2, 0);   // expected: 20
      dnNamedCellIsFormula:
        begin
          MyWorksheet.DefinedNames.Add('formula', 0,0, 3,0, 3,0);
          MyWorksheet.WriteFormula(2, 1, '=formula');             // expected: 30
        end;
      dnNamedRangeIsFormula:
        begin
          MyWorksheet.DefinedNames.Add('values', 0, 0, 0, 0, 4, 0);  // expected: 100
          MyWorksheet.WriteFormula(2, 1, 'SUM(values)');
        end;
    end;

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := [boAutoCalc, boReadFormulas];
    MyWorkbook.ReadFromFile(TempFile, AFormat);
    MyWorksheet := GetWorksheetByName(MyWorkBook, 'Data');
    if MyWorksheet=nil then
      fail('Error in test code. Failed to get named worksheet');

    case ATestType of
      dnNamedCellIsConst:
        begin
          row := 2;
          col := 0;
          actual := MyWorksheet.ReadAsNumber(row, col);
          expected := 20;
          CheckEquals(expected, actual,
            'Test saved named cell mismatch, cell '+CellNotation(MyWorksheet, row, col));
        end;
      dnNamedCellIsFormula:
        begin
          row := 2;
          col := 1;
          cell := MyWorksheet.FindCell(row, col);
          formula := MyWorksheet.ReadFormulaAsString(cell);
          CheckEquals('formula', formula,
            'Test saved formula mismatch, cell ' + CellNotation(MyWorksheet, row, col));
          actual := MyWorksheet.ReadAsNumber(cell);
          expected := 30;
          CheckEquals(expected, actual,
            'Test saved named cell mismatch, cell '+CellNotation(MyWorksheet, row, col));
        end;
      dnNamedRangeIsFormula:
        begin
          row := 2;
          col := 1;
          cell := MyWorksheet.Findcell(row, col);
          formula := MyWorksheet.ReadFormulaAsString(cell);
          CheckEquals('SUM(values)', formula,
            'Test saved formula mismatch, cell ' + CellNotation(MyWorksheet, row, col));
          actual := MyWorksheet.ReadAsNumber(cell);
          expected := 100;
          CheckEquals(expected, actual,
            'Test saved named range mismatch, cell '+CellNotation(MyWorksheet, row, col));
        end;
    end;

  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;

{ Tests for Open Document file format }

procedure TSpreadWriteReadDefinedNamesTests.TestWriteRead_ODS_Cell_Comment;
begin
  TestWriteRead_DefinedNames(sfOpenDocument, dnNamedCellIsConst);
end;

procedure TSpreadWriteReadDefinedNamesTests.TestWriteRead_ODS_Cell_Formula;
begin
  TestWriteRead_DefinedNames(sfOpenDocument, dnNamedCellIsFormula);
end;

procedure TSpreadWriteReadDefinedNamesTests.TestWriteRead_ODS_Range_Formula;
begin
  TestWriteRead_DefinedNames(sfOpenDocument, dnNamedRangeIsFormula);
end;

{ Tests for XLSX file format }

procedure TSpreadWriteReadDefinedNamesTests.TestWriteRead_XLSX_Cell_Comment;
begin
  TestWriteRead_DefinedNames(sfOOXML, dnNamedCellIsConst);
end;

procedure TSpreadWriteReadDefinedNamesTests.TestWriteRead_XLSX_Cell_Formula;
begin
  TestWriteRead_DefinedNames(sfOOXML, dnNamedCellIsFormula);
end;

procedure TSpreadWriteReadDefinedNamesTests.TestWriteRead_XLSX_Range_Formula;
begin
  TestWriteRead_DefinedNames(sfOOXML, dnNamedRangeIsFormula);
end;


initialization
  RegisterTest(TSpreadWriteReadDefinedNamesTests);

end.

