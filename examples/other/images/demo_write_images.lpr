program demo_write_images;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpstypes, fpspreadsheet, fpsallformats, fpsutils,
  fpsPageLayout;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
  i: Integer;

const
  image1 = '../../images/components/TSWORKBOOKSOURCE.png';
  image2 = '../../images/components/TSWORKSHEETGRID.png';
  image3 = '../../images/components/TSCELLEDIT.png';

begin
  Writeln('Starting program "demo_write_images"...');

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  try
    MyWorksheet := MyWorkbook.AddWorksheet('Sheet 1');
               {
    for i:=0 to 20 do
      MyWorksheet.WriteRowHeight(i, 4.5, suMillimeters);
//    MyWorksheet.DefaultRowHeight := 4.5;  // millimeters
                }
    MyWorksheet.WriteText(0, 0, 'There are images in cells A3 and B3'); //
    // These images are offset by 1mm in both directions from the top/left cell edge
    MyWorksheet.WriteImage(2, 0, image1, 1.0, 1.0, 2.0, 2.0);   // This image is magnified by factor 2
    MyWorksheet.WriteImage(2, 1, image2, 1.0, 1.0);

    MyWorksheet := MyWorkbook.AddWorksheet('Sheet 2');
    MyWorksheet.WriteText(0, 0, 'There is an image in cell B3');
    MyWorksheet.WriteImage(2, 1, image3);

    // Save the spreadsheet to files
    MyDir := ExtractFilePath(ParamStr(0));
    MyWorkbook.WriteToFile(MyDir + 'img.xlsx', sfOOXML, true);
    MyWorkbook.WriteToFile(MyDir + 'img.ods', sfOpenDocument, true);
//  MyWorkbook.WriteToFile(MyDir + 'img.xls', sfExcel8, true);
//  MyWorkbook.WriteToFile(MyDir + 'img5.xls', sfExcel5, true);
//  MyWorkbook.WriteToFile(MyDir + 'img2.xls', sfExcel2, true);

    if MyWorkbook.ErrorMsg <> '' then
      WriteLn(MyWorkbook.ErrorMsg);

    WriteLn('Finished.');
    WriteLn('Please open the files "img.*" in your spreadsheet program.');
   {$ifdef WINDOWS}
    WriteLn('Press ENTER to close this program...');
    ReadLn;
   {$ENDIF}

  finally
    MyWorkbook.Free;
  end;
end.

