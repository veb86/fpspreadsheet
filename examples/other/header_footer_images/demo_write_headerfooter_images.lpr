program demo_write_headerfooter_images;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpstypes, fpspreadsheet, fpsallformats, fpsutils,
  fpsPageLayout;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
  cell: PCell;
  i, r, c: Integer;

const
  image1 = '../../images/components/TSWORKBOOKSOURCE.png';
  image2 = '../../images/components/TSWORKSHEETGRID.png';
  image3 = '../../images/components/TSCELLEDIT.png';

begin
  Writeln('Starting program "demo_write_headerfooter_images"...');
  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorksheet := MyWorkbook.AddWorksheet('Sheet 1');
    MyWorksheet.WriteText(0, 0, 'The header of this sheet contains an image');
    MyWorksheet.PageLayout.HeaderMargin := 10;
    MyWorksheet.Pagelayout.TopMargin := 30;     // the header is 20 mm high
    MyWorksheet.PageLayout.Headers[HEADER_FOOTER_INDEX_ALL] := '&CHeader with image!';
    MyWorksheet.PageLayout.AddHeaderImage(HEADER_FOOTER_INDEX_ALL, hfsLeft, image1);

    MyWorksheet := MyWorkbook.AddWorksheet('Sheet 2');
    MyWorksheet.WriteText(0, 0, 'The footer of this sheet contains an image');
    MyWorksheet.PageLayout.Footers[HEADER_FOOTER_INDEX_ALL] := '&CFooter with image!';
    MyWorksheet.PageLayout.AddFooterImage(HEADER_FOOTER_INDEX_ALL, hfsRight, image2);

    // Save the spreadsheet to files
    MyDir := ExtractFilePath(ParamStr(0));
    MyWorkbook.WriteToFile(MyDir + 'hfimg.xlsx', sfOOXML, true);
    MyWorkbook.WriteToFile(MyDir + 'hfimg.ods', sfOpenDocument, true);

//  MyWorkbook.WriteToFile(MyDir + 'hfimg.xls', sfExcel8, true);
//  MyWorkbook.WriteToFile(MyDir + 'hfimg5.xls', sfExcel5, true);
//  MyWorkbook.WriteToFile(MyDir + 'hfimg2.xls', sfExcel2, true);

    if MyWorkbook.ErrorMsg <> '' then
      WriteLn(MyWorkbook.ErrorMsg);

    WriteLn('Finished.');
    WriteLn('Please open the files "hfimg.*" in your spreadsheet program.');
   {$ifdef WINDOWS}
    WriteLn('Press ENTER to close this program...');
    ReadLn;
   {$endif}

  finally
    MyWorkbook.Free;
  end;
end.

