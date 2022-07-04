program demo_read_images;

// Activate one of these defines
{$DEFINE USE_XLSX}
{.$DEFINE USE_OPENDOCUMENT}

uses
  SysUtils, fpspreadsheet, fpstypes, fpsutils, fpsimages, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'img';
  
var
  workbook: TsWorkbook;
  worksheet: TsWorksheet;
  myDir: String;
  i, j: Integer;
  embobj: TsEmbeddedObj;
  img: PsImage;
  
begin
  WriteLn('Starting program "demo_read_images"...');
  WriteLn;
  
  // Create a spreadsheet
  workbook := TsWorkbook.Create;
  try
    // Read spreadsheet file
    myDir := ExtractFilePath(ParamStr(0));
    {$IFDEF USE_XLSX}
    workbook.ReadFromFile(myDir + FILE_NAME + '.xlsx', sfOOXML);
    {$ENDIF}
    {$IFDEF USE_OPENDOCUMENT}
    workbook.ReadFromFile(myDir + FILE_NAME + '.ods', sfOpenDocument);
    {$ENDIF}
    
    // Get worksheets
    for i := 0 to workbook.GetWorksheetCount-1 do
    begin
      worksheet := workbook.GetWorksheetByIndex(i);
      WriteLn('Worksheet "' + worksheet.Name + '" contains ' + IntToStr(worksheet.GetImageCount) + ' image(s).');
      // Read out images and save as separate files
      for j := 0 to worksheet.GetImageCount-1 do
      begin
        img := worksheet.GetPointerToImage(j);
        embObj := workbook.GetEmbeddedObj(img^.Index);
        WriteLn('  Image Index=', img^.Index, 
          ', Cell=', GetCellString(img^.Row, img^.Col),
          ', File=', ExtractFileName(embObj.FileName), 
          ', Width=', embobj.ImageWidth:0:1, 'mm',
          ', Height=', embObj.ImageHeight:0:1,'mm',
          ', ScaleX=', img^.ScaleX:0:2,
          ', ScaleY=', img^.ScaleY:0:2
        );
        if embObj.FileName <> '' then
          embobj.Stream.SaveToFile(ExtractFileName(embobj.FileName));
      end;
      WriteLn;
    end;
    WriteLn('Finished.');
    
  finally
    workbook.Free;
  end;
  
  if ParamCount = 0 then
  begin
    {$IFDEF MSWINDOWS}
    WriteLn;
    WriteLn('Press [ENTER] to quit...');
    ReadLn;
    {$ENDIF}
  end;  
end.

