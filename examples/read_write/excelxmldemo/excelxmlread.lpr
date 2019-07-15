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
  inputFileName := dir + 'test.xml';

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

      WriteLn;
      WriteLn('Printer settings/Page layout');
      WriteLn('  Page width: ', worksheet.PageLayout.PageWidth:0:1, ' mm');
      WriteLn('  Page height: ', worksheet.PageLayout.PageHeight:0:1, ' mm');
      WriteLn('  Orientation: ', worksheet.PageLayout.Orientation);
      WriteLn('  Left margin: ', worksheet.PageLayout.LeftMargin:0:1, ' mm');
      WriteLn('  Right margin: ', worksheet.PageLayout.RightMargin:0:1, ' mm');
      WriteLn('  Top margin: ', worksheet.PageLayout.TopMargin:0:1, ' mm');
      WriteLn('  Bottom margin: ', worksheet.PageLayout.BottomMargin:0:1, ' mm');
      WriteLn('  Header margin: ', worksheet.PageLayout.HeaderMargin:0:1, ' mm');
      WriteLn('  Header text: ', worksheet.PageLayout.Headers[0]);
      WriteLn('  Footer margin: ', worksheet.PageLayout.FooterMargin:0:1, ' mm');
      WriteLn('  Footer text: ', worksheet.PageLayout.Footers[0]);
      WriteLn('  Scaling factor: ', worksheet.PageLayout.ScalingFactor, ' %');
      WriteLn('  Start page number: ', worksheet.PageLayout.StartPageNumber);
      Write('  Options: ');
      if (poPrintGridLines in worksheet.PageLayout.Options) then Write('GridLines ');
      if (poMonochrome in worksheet.PageLayout.Options) then Write('Black&White ');
      if (poDraftQuality in worksheet.PageLayout.Options) then Write('Draft ');
      if (poPrintHeaders in worksheet.PageLayout.Options) then Write('Headers ');
      if (poCommentsAtEnd in worksheet.Pagelayout.Options) then Write('CommentsAtEnd ');
      if (poPrintCellComments in worksheet.PageLayout.Options) then Write('CellComments ');
      if (poHorCentered in worksheet.PageLayout.Options) then Write('HorCentered ');
      if (poVertCentered in worksheet.PageLayout.Options) then Write('VertCentered ');
      if (poPrintPagesByRows in worksheet.PageLayout.Options) then Write('PagesByRows ');
      if (poFitPages in worksheet.PageLayout.Options) then Write('FitPage' );
      WriteLn;
      WriteLn('  Fit height to pages: ', worksheet.Pagelayout.FitHeightToPages);
      WriteLn('  Fit width to pages: ', worksheet.PageLayout.FitWidthToPages);
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

