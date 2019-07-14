{-------------------------------------------------------------------------------
                           excelxmlread.lpr

Demonstrates how to read an Excel 2003 xml file using the fpspreadsheet library
-------------------------------------------------------------------------------}
program excelxmlread;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, LazUTF8, fpsTypes, fpspreadsheet, xlsxml, fpsutils;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  inputFilename: string;
  dir: string;
  i: Integer;
  cell: PCell;

{$R *.res}

begin
  // Open the input file
  dir := ExtractFilePath(ParamStr(0));
//  inputFileName := dir + 'test.xml';
  inputFileName := dir + 'datatypes.xml';

  if not FileExists(inputFileName) then begin
    WriteLn('Input file ', inputFileName, ' does not exist. Please run excelxmlwrite first.');
    Halt;
  end;
  WriteLn('Opening input file ', inputFilename);

  // Create the spreadsheet
  workbook := TsWorkbook.Create;
  try
    workbook.Options := workbook.Options + [boReadFormulas];
    workbook.ReadFromFile(inputFilename, sfExcelXML);

    WriteLn('The workbook contains ', workbook.GetWorksheetCount, ' sheets.');
    WriteLn;

    // Write all cells with contents to the console
    for i:=0 to workbook.GetWorksheetCount-1 do begin
      worksheet := workbook.GetWorksheetByIndex(i);
      WriteLn('');
      WriteLn('Contents of the worksheet "', worksheet.Name, '":');
      WriteLn('');

      for cell in worksheet.Cells do
      begin
        Write(' ',
              ' Row: ', cell^.Row,
              ' Col: ', cell^.Col,
              ' Type: ', cell^.ContentType,
              ' Value: ', UTF8ToConsole(worksheet.ReadAsText(cell^.Row, cell^.Col))
        );
        if HasFormula(cell) then
          WriteLn(' Formula: ', workSheet.ReadFormulaAsString(cell))
        else
          WriteLn;
      end;
    end;

  finally
    // Finalization
    workbook.Free;
  end;

  {$IFDEF WINDOWS}
  WriteLn;
  WriteLn('Press ENTER to quit...');
  ReadLn;
  {$ENDIF}
end.

