{ This example uses the "ignoreFormula" workbook option to create an ods
  file with an external reference.

  NOTE: The external reference is not calculated. This will happen when
  LibreOffice Calc loads the file. When the file is closed in LOCalc
  confirmation must be given to save the file because it has been changed
  by LOCalc.

  This method does not work with Excel because it writes an additonal
  folder and xml files for external links. }

program demo_ignore_formula;

{$mode objfpc}{$H+}

{$DEFINE ODS}
{.$DEFINE XLSX}   // <---- NOT WORKING

uses
  SysUtils, FileUtil,
  fpsTypes, fpsUtils, fpSpreadsheet, fpsOpenDocument, xlsxOOXML;

const
  {$IFDEF ODS}
  FILE_FORMAT = sfOpenDocument;
  MASTER_FILE = 'master.ods';
  EXTERNAL_FILE = 'external.ods';
  {$ENDIF}
  {$IFDEF XLSX}
  FILE_FORMAT = sfOOXML;
  MASTER_FILE = 'master.xlsx';
  EXTERNAL_FILE = 'external.xlsx';
  {$ENDIF}
  EXTERNAL_SHEET = 'Sheet';
  CELL1 = 'A1';
  CELL2 = 'B1';


var
  book: TsWorkbook;
  sheet: TsWorksheet;
  cell: PCell;

  // example for an external ods reference:
  // ='file:///D:/fpspreadsheet/examples/other/external.ods'#$Sheet.A1
  function ODS_ExtRef(AFilename, ASheetName, ACellAddr: String): String;
  var
    i: Integer;
  begin
    Result := ExpandFileName(AFileName);
    for i:=1 to Length(Result) do
      if Result[i] = '\' then Result[i] := '/';
    Result := Format('''file:///%s''#$%s.%s', [
      Result, ASheetName, ACellAddr
    ]);
  end;

  // example for an external xlsx reference:
  // =[external.xlsx]Sheet!$A$1
  function XLSX_ExtRef(AFilename, ASheetName, ACellAddr: String): String;
  var
    r, c: Cardinal;
    flags: TsRelFlags;
  begin
    ParseCellString(ACellAddr, r, c, flags);
    Result := Format('[%s]%s!%s', [
      ExtractFileName(AFileName), ASheetName, GetCellString(r, c, [])
    ]);
  end;

  function ExtRef(AFileName, ASheetName, ACellAddr: String): String;
  begin
    {$IFDEF ODS}
    Result := ODS_ExtRef(AFileName, ASheetName, ACellAddr);
    {$ENDIF}
    {$IFDEF XLSX}
    Result := XLSX_ExtRef(AFilename, ASheetName, ACellAddr);
    {$ENDIF}
  end;

begin
  // Write external file
  book := TsWorkbook.Create;
  try
    sheet := book.AddWorksheet(EXTERNAL_SHEET);

    cell := sheet.GetCell(CELL1);
    sheet.WriteNumber(cell, 1000.0);

    cell := sheet.GetCell(CELL2);
    sheet.WriteText(cell, 'Hallo');

    book.WriteToFile(EXTERNAL_FILE, FILE_FORMAT, true);
  finally
    book.Free;
  end;

  // Write ods and xlsx master files
  book := TsWorkbook.Create;
  try
    // Instruct fpspreadsheet to leave the formula alone.
    book.Options := book.Options + [boIgnoreFormulas];
    sheet := book.AddWorksheet('Sheet');

    // Write external references
    sheet.WriteFormula(0, 0, ExtRef(EXTERNAL_FILE, EXTERNAL_SHEET, CELL1));
    sheet.WriteFormula(1, 0, ExtRef(EXTERNAL_FILE, EXTERNAL_SHEET, CELL2));
    book.WriteToFile(MASTER_FILE, FILE_FORMAT, true);

  finally
    book.Free;
  end;
end.

