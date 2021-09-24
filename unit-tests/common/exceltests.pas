unit exceltests;

{$mode objfpc}{$H+}

interface

{ Tests which require Excel.
  Will be skipped if Excel is not available.
}

{$ifdef windows}
uses
  // Not using Lazarus package as the user may be working with multiple versions
  // Instead, add .. to unit search path
  Classes, SysUtils, fpcunit, testregistry,
  fpstypes, fpspreadsheet, xlsbiff2, xlsbiff5, xlsbiff8 {and a project requirement for lclbase for utf8 handling},
  testsutility;

type
  { TSpreadExcelTests }
  // Write "something" to Excel file and open it in Excel
  TSpreadExcelTests = class(TTestCase)
  private
    FExcelApp: OleVariant;
    FExcelAvail: Boolean;

  protected
    // Set up expected values:
    procedure SetUp; override;
    procedure TearDown; override;

    // Comments in several sheets (index of sheets specified)
    procedure Test_Comments(const ASheets: array of Integer; AFormat: TsSpreadsheetFormat);
    // Header/footer images in several sheets
    procedure Test_HeaderFooterImages(const ASheets: array of Integer;
      AddToFooter: Boolean; AFormat: TsSpreadsheetFormat);
    // Images in several sheets
    procedure Test_Images(const ASheets: array of Integer;
      AFormat: TsSpreadsheetFormat);
    // Comments, images, header/footer images, and hyperlinks
    procedure Test_All(const ASheets: array of Integer;
      AFormat: TsSpreadsheetFormat);

  published
    // OOXML test cases
    procedure Test_Comments0_OOXML;
    procedure Test_Comments01_OOXML;
    procedure Test_Comments02_OOXML;
    procedure Test_HeaderImages0_OOXML;
    procedure Test_HeaderImages01_OOXML;
    procedure Test_HeaderImages12_OOXML;
    procedure Test_FooterImages0_OOXML;
    procedure Test_FooterImages01_OOXML;
    procedure Test_FooterImages12_OOXML;
    procedure Test_Images0_OOXML;
    procedure Test_Images01_OOXML;
    procedure Test_Images12_OOXML;

    procedure Test_All0_OOXML;
    procedure Test_All01_OOXML;
    procedure Test_All12_OOXML;

  end;

{$endif}

implementation

{$ifdef windows}
uses
  TypInfo, comobj;

const
  SheetName = 'Excel-Test';


{ TSpreadExcelTexts }

procedure TSpreadExcelTests.SetUp;
begin
  inherited SetUp;
  try
    FExcelApp := CreateOleObject('Excel.Application');  // creates an Excel Object
    FExcelApp.Visible := False;         // hides Excel window
    FExcelApp.ScreenUpdating := False;  // turns off screen updating of Excel
    FExcelApp.DisplayAlerts := False;   // no excel warnings and error messages
    FExcelAvail := true;
  except
    FExcelAvail := false;
  end;
end;

procedure TSpreadExcelTests.TearDown;
begin
  FExcelApp.Quit;
  FExcelApp := UnAssigned;
  inherited TearDown;
end;

procedure TSpreadExcelTests.Test_Comments(
  const ASheets: array of Integer; AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  TempFile: string; //write xls/xml to this file and read back from it
  varFile: variant; // filename as variant for Excel
  i, j: Integer;
  book: OleVariant;
begin
  if not FExcelAvail then
  begin
    exit;
  end;

  TempFile := NewTempFile;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];

    for i:=0 to ASheets[High(ASheets)] do
    begin
      MyWorkSheet:= MyWorkBook.AddWorksheet(SheetName + IntToStr(i+1));
      for j:=0 to High(ASheets) do
        if ASheets[j] = i then
        begin
          MyWorksheet.WriteText(0, 0, 'This is a text');
          MyWorksheet.WriteComment(0, 0, 'This is a comment');
        end;
    end;

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet in Excel

  try
    varFile := TempFile;
    book := FExcelApp.Workbooks.Open(varFile);
    book.Close;
  except
    Fail('Excel fails to open the file.');
  end;

  DeleteFile(TempFile);
end;

procedure TSpreadExcelTests.Test_HeaderFooterImages(
  const ASheets: array of Integer; AddToFooter: Boolean;
  AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  TempFile: string; // write xls/xml to this file and read back from it
  varFile: variant; // filename as variants for Excel
  i, j: Integer;
  book: OleVariant;
begin
  if not FExcelAvail then
  begin
    exit;
  end;

  TempFile := NewTempFile;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];

    for i:=0 to ASheets[High(ASheets)] do
    begin
      MyWorkSheet:= MyWorkBook.AddWorksheet(SheetName + IntToStr(i+1));
      for j:=0 to High(ASheets) do
        if ASheets[j] = i then
        begin
          MyWorksheet.WriteText(0, 0, 'This is a text');
          if AddToFooter then
            MyWorksheet.PageLayout.AddFooterImage(1, hfsCenter, 'lazarus32x32.png')
          else
            MyWorksheet.PageLayout.AddHeaderImage(1, hfsCenter, 'lazarus32x32.png');
        end;
    end;

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet in Excel

  try
    varFile := TempFile;
    book := FExcelApp.Workbooks.Open(varFile);
    book.Close;
  except
    Fail(Format('Excel fails to open file "%s".', [TempFile]));
  end;

  DeleteFile(TempFile);
end;

procedure TSpreadExcelTests.Test_Images(const ASheets: array of Integer;
  AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  TempFile: string; // write xls/xml to this file and read back from it
  varFile: variant; // filename as variants for Excel
  i, j: Integer;
  book: OleVariant;
begin
  if not FExcelAvail then
  begin
    exit;
  end;

  TempFile := NewTempFile;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];

    for i:=0 to ASheets[High(ASheets)] do
    begin
      MyWorkSheet:= MyWorkBook.AddWorksheet(SheetName + IntToStr(i+1));
      for j:=0 to High(ASheets) do
        if ASheets[j] = i then
        begin
          MyWorksheet.WriteText(0, 0, 'This is a text');
          MyWorksheet.WriteImage(0, 1, 'lazarus32x32.png')
        end;
    end;

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet in Excel

  try
    varFile := TempFile;
    book := FExcelApp.Workbooks.Open(varFile);
    book.Close;
  except
    Fail(Format('Excel fails to open file "%s".', [TempFile]));
  end;

  DeleteFile(TempFile);
end;

procedure TSpreadExcelTests.Test_All(const ASheets: array of Integer;
  AFormat: TsSpreadsheetFormat);
var
  MyWorksheet: TsWorksheet;
  MyWorkbook: TsWorkbook;
  TempFile: string; // write xls/xml to this file and read back from it
  varFile: variant; // filename as variants for Excel
  i, j: Integer;
  book: OleVariant;
begin
  if not FExcelAvail then
  begin
    exit;
  end;

  TempFile := NewTempFile;

  MyWorkbook := TsWorkbook.Create;
  try
    MyWorkbook.Options := MyWorkbook.Options + [boCalcBeforeSaving];

    for i:=0 to ASheets[High(ASheets)] do
    begin
      MyWorkSheet:= MyWorkBook.AddWorksheet(SheetName + IntToStr(i+1));
      for j:=0 to High(ASheets) do
        if ASheets[j] = i then
        begin
          MyWorksheet.WriteText(0, 0, 'This is a text');
          MyWorksheet.WriteComment(0, 0, 'This is a comment');
          MyWorksheet.WriteHyperlink(0, 0, 'http://www.lazarus-ide.org/');
          MyWorksheet.WriteImage(0, 1, 'lazarus32x32.png');
          MyWorksheet.PageLayout.AddFooterImage(1, hfsCenter, 'lazarus32x32.png');
          MyWorksheet.PageLayout.AddHeaderImage(1, hfsCenter, 'lazarus32x32.png');
        end;
    end;

    MyWorkBook.WriteToFile(TempFile, AFormat, true);
  finally
    MyWorkbook.Free;
  end;

  // Open the spreadsheet in Excel

  try
    varFile := TempFile;
    book := FExcelApp.Workbooks.Open(varFile);
    book.Close;
  except
    Fail(Format('Excel fails to open file "%s".', [TempFile]));
  end;

  DeleteFile(TempFile);
end;


{ OOXML }

procedure TSpreadExcelTests.Test_Comments0_OOXML;
begin
  Test_Comments([0], sfOOXML);
end;

procedure TSpreadExcelTests.Test_Comments01_OOXML;
begin
  Test_Comments([0,1], sfOOXML);
end;

procedure TSpreadExcelTests.Test_Comments02_OOXML;
begin
  Test_Comments([0,2], sfOOXML);
end;


procedure TSpreadExcelTests.Test_HeaderImages0_OOXML;
begin
  Test_HeaderFooterImages([0], false, sfOOXML);
end;

procedure TSpreadExcelTests.Test_HeaderImages01_OOXML;
begin
  Test_HeaderFooterImages([0,1], false, sfOOXML);
end;

procedure TSpreadExcelTests.Test_HeaderImages12_OOXML;
begin
  Test_HeaderFooterImages([1,2], false, sfOOXML);
end;


procedure TSpreadExcelTests.Test_FooterImages0_OOXML;
begin
  Test_HeaderFooterImages([0], true, sfOOXML);
end;

procedure TSpreadExcelTests.Test_FooterImages01_OOXML;
begin
  Test_HeaderFooterImages([0,1], true, sfOOXML);
end;

procedure TSpreadExcelTests.Test_FooterImages12_OOXML;
begin
  Test_HeaderFooterImages([1,2], true, sfOOXML);
end;


procedure TSpreadExcelTests.Test_Images0_OOXML;
begin
  Test_Images([0], sfOOXML);
end;

procedure TSpreadExcelTests.Test_Images01_OOXML;
begin
  Test_Images([0,1], sfOOXML);
end;

procedure TSpreadExcelTests.Test_Images12_OOXML;
begin
  Test_Images([1,2], sfOOXML);
end;


procedure TSpreadExcelTests.Test_All0_OOXML;
begin
  Test_All([0], sfOOXML);
end;

procedure TSpreadExcelTests.Test_All01_OOXML;
begin
  Test_All([0,1], sfOOXML);
end;

procedure TSpreadExcelTests.Test_All12_OOXML;
begin
  Test_All([1,2], sfOOXML);
end;


initialization
  RegisterTest(TSpreadExcelTests);

{$endif}

end.

