program demo_search;

{$mode objfpc}{$H+}

uses
  {$IFDEF UNIX}{$IFDEF UseCThreads}
  cthreads,
  {$ENDIF}{$ENDIF}
  SysUtils, Classes, TypInfo,
  fpsTypes, fpSpreadsheet, fpsSearch, fpsUtils;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  s: String;
  searchParams: TsSearchParams;
  rowFound, colFound: Cardinal;
  worksheetFound: TsWorksheet;

begin
  workbook := TsWorkbook.Create;
  try
    worksheet := workbook.AddWorksheet('Test');

    worksheet.WriteNumber(0, 0, 10);        // A1
    worksheet.WriteNumber(1, 0, 2);         // A2
    worksheet.WriteNumber(2, 0, 5);         // A3  <---
    worksheet.WriteNumber(3, 0, 1);         // A4
    worksheet.WriteNumber(4, 0, 5);         // A5  <---
    worksheet.WriteNumber(5, 0, 3);         // A6
    worksheet.WriteNumber(0, 1, 5);         // B1  <---

    worksheet.WriteComment(0, 0, '5');
    worksheet.WriteComment(1, 0, '2');

    searchParams := InitSearchParams('5', [soEntireDocument]);

    // Create search engine and execute search
    with TsSearchEngine.Create(workbook) do begin
      if FindFirst(searchParams, worksheetFound, rowFound, colFound) then begin
        WriteLn('First "', searchparams.SearchText, '" found in cell ',
          GetCellString(rowFound, colFound), ' of worksheet ', worksheetFound.Name);
        while FindNext(searchParams, worksheetFound, rowFound, colFound) do
          WriteLn('Next "', searchParams.SearchText, '" found in cell ',
            GetCellString(rowFound, colFound), ' of worksheet ', worksheetFound.Name);
      end;
      Free;
    end;

    // Now search in comments
    Include(searchparams.Options, soSearchInComment);
    with TsSearchEngine.Create(workbook) do begin
      if FindFirst(searchParams, worksheetFound, rowFound, colFound) then begin
        WriteLn('First "', searchparams.SearchText, '" found in comment of cell ',
          GetCellString(rowFound, colFound), ' of worksheet ', worksheetFound.Name);
        while FindNext(searchParams, worksheetFound, rowFound, colFound) do
          WriteLn('Next "', searchParams.SearchText, '" found in comment of cell ',
            GetCellString(rowFound, colFound), ' of worksheet ', worksheetFound.Name);
      end;
      Free;
    end;

  finally
    workbook.Free;
  end;

  {$IFDEF MSWINDOWS}
  WriteLn;
  WriteLn('Press ENTER to quit...');
  ReadLn;
  {$ENDIF}
end.

