program demo_read_headerfooter_images;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpstypes, fpspreadsheet, fpsallformats, fpsutils,
  fpsImages, fpsPageLayout;

var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  i: Integer;
  hfs: TsHeaderFooterSectionIndex;
  img: TsEmbeddedObj;

const
  FILE_NAME = 'hfimg';
  
  function hfsToStr(idx: TsHeaderFooterSectionIndex): String;
  begin
    case idx of
      hfsLeft: Result := 'left';
      hfsCenter: Result := 'center';
      hfsRight: Result := 'right';
    end;
  end;

begin
  Writeln('Starting program "demo_read_headerfooter_images"...');
  WriteLn;

  // Create the spreadsheet
  workbook := TsWorkbook.Create;
  try
    workbook.ReadFromFile(FILE_NAME + '.xlsx', sfOOXML);
    
    for i := 0 to workbook.GetWorksheetCount-1 do 
    begin
      worksheet := workbook.GetWorksheetByIndex(i);
      WriteLn('Worksheet "', worksheet.Name, '", PageLayout:');
      WriteLn('  HeaderMargin: ', worksheet.Pagelayout.HeaderMargin:6:1, ' mm');
      WriteLn('  TopMargin   : ', worksheet.Pagelayout.TopMargin:6:1, ' mm');
      WriteLn('  Headers:');
      WriteLn('    first page: ', worksheet.PageLayout.Headers[0]);
      WriteLn('    odd or all pages: ', worksheet.PageLayout.Headers[1]);
      WriteLn('    even pages:', worksheet.PageLayout.Headers[2]);
      for hfs in TsHeaderFooterSectionIndex do
      begin
        Write('    Images ', hfsToStr(hfs), ': '); 
        if worksheet.PageLayout.HeaderImages[hfs].Index = -1 then
          WriteLn('(none)')
        else
        begin
          img := workbook.GetEmbeddedObj(worksheet.PageLayout.HeaderImages[hfs].Index);
          if img = nil then
            WriteLn('image not found')
          else
            WriteLn(img.FileName);
        end;
      end;
      WriteLn('  Footers:');
      WriteLn('    first page: ', worksheet.PageLayout.Footers[0]);
      WriteLn('    odd or all pages: ', worksheet.PageLayout.Footers[1]);
      WriteLn('    even pages: ', worksheet.PageLayout.Footers[2]);
      for hfs in TsHeaderFooterSectionIndex do
      begin
        Write  ('    Images ', hfsToStr(hfs), ': '); 
        if worksheet.PageLayout.FooterImages[hfs].Index = -1 then
          WriteLn('(none)')
        else
        begin
          img := workbook.GetEmbeddedObj(worksheet.PageLayout.FooterImages[hfs].Index);
          if img = nil then
            WriteLn('image not found')
          else
            WriteLn(img.FileName);
        end;
      end;
      WriteLn;
    end;
    
    if workbook.ErrorMsg <> '' then
      WriteLn(workbook.ErrorMsg);

    WriteLn('Finished.');
    if ParamCount = 0 then
    begin
      {$IFDEF MSWINDOWS}
      WriteLn('Press [ENTER] to close this program...');
      ReadLn;
      {$ENDIF}
    end;

  finally
    workbook.Free;
  end;
end.

