program demo_write_headerfooter_images;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpstypes, fpspreadsheet, fpsallformats, fpsutils,
  fpsPageLayout;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;

const
  image1 = '../../../images/components/TSWORKBOOKSOURCE.png';
  image2 = '../../../images/components/TSWORKSHEETGRID.png';

begin
  Writeln('Starting program "demo_write_headerfooter_images"...');
  
  if not FileExists(image1) then
  begin
    WriteLn(ExpandFilename(image1) + ' not found.');
    Halt;
  end;
    
  if not FileExists(image2) then
  begin
    WriteLn(ExpandFilename(image2) + ' not found.');
    Halt;
  end;

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
    MyWorksheet.PageLayout.Footers[HEADER_FOOTER_INDEX_ALL] := '&CFooter with image, scaled by factor 2!';
    MyWorksheet.PageLayout.AddFooterImage(HEADER_FOOTER_INDEX_ALL, hfsRight, image2, 2.0, 2.0);

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

    if ParamCount = 0 then
    begin
      {$IFDEF MSWINDOWS}
      WriteLn('Press [ENTER] to close this program...');
      ReadLn;
      {$ENDIF}
    end;

  finally
    MyWorkbook.Free;
  end;
end.

