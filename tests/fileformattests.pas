unit fileformattests;

{$mode objfpc}{$H+}

interface
{ Cell type tests
This unit tests writing the various cell data types out to and reading them 
back from files.
}

uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet,
  xlsbiff2, xlsbiff5, xlsbiff8, fpsOpenDocument,
  testsutility;

type
  { TSpreadFileFormatTests }
  // Write cell types to xls/xml file and read back
  TSpreadFileFormatTests = class(TTestCase)
  private

  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;
    procedure TestAutoDetect(AFormat: TsSpreadsheetFormat);

  published
    procedure TestAutoDetect_BIFF2;
    procedure TestAutoDetect_BIFF5;
    procedure TestAutoDetect_BIFF8;
    procedure TestAutoDetect_OOXML;
    procedure TestAutoDetect_ODS;
  end;

implementation

uses
  fpsReaderWriter;

const
  SheetName = 'FileFormat';


{ TSpreadFileFormatTests }

procedure TSpreadFileFormatTests.SetUp;
begin
  inherited SetUp;
end;

procedure TSpreadFileFormatTests.TearDown;
begin
  inherited TearDown;
end;

procedure TSpreadFileFormatTests.TestAutoDetect(AFormat: TsSpreadsheetFormat);
const
  EXPECTED_TEXT = 'abcefg';
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  row, col: Integer;
  MyCell: PCell;
  value: Boolean;
  TempFile: string; //write xls/xml to this file and read back from it
  actualText: String;
begin
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkSheet:= MyWorkBook.AddWorksheet(SheetName);

    // write any content to the file
    MyWorksheet.WriteText(0, 0, EXPECTED_TEXT);

    // Write workbook to file using format specified, but with wrong extension
    TempFile := ChangeFileExt(NewTempFile, '.abc');
    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    // Try to read file and detect format automatically
    try
      MyWorkbook.ReadFromFile(TempFile);
      // If the tests gets here the format was detected correctly.
      // Quickly check the cell content
      MyWorksheet := MyWorkbook.GetFirstWorksheet;
      actualText := MyWorksheet.ReadAsUTF8Text(0, 0);
      CheckEquals(EXPECTED_TEXT, actualText, 'Cell mismatch in A1');
    except
      fail('Cannot read file with format ' + GetSpreadFormatName(ord(AFormat)));
    end;
  finally
    MyWorkbook.Free;
    DeleteFile(TempFile);
  end;
end;


{ BIFF2 }
procedure TSpreadFileFormatTests.TestAutoDetect_BIFF2;
begin
  TestAutoDetect(sfExcel2);
end;

{ BIFF5 }
procedure TSpreadFileFormatTests.TestAutoDetect_BIFF5;
begin
  TestAutoDetect(sfExcel5);
end;

{ BIFF8 }
procedure TSpreadFileFormatTests.TestAutoDetect_BIFF8;
begin
  TestAutoDetect(sfExcel8);
end;

{ OOXML }
procedure TSpreadFileFormatTests.TestAutoDetect_OOXML;
begin
  TestAutoDetect(sfOOXML);
end;

{ ODS }
procedure TSpreadFileFormatTests.TestAutoDetect_ODS;
begin
  TestAutoDetect(sfOpenDocument);
end;


initialization
  RegisterTest(TSpreadFileFormatTests);

end.

