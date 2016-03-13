program demo_write_images;

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
  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;
  MyWorkbook.Options := [boFileStream];

  MyWorksheet := MyWorkbook.AddWorksheet('Sheet 1');
  MyWorksheet.DefaultRowHeight := 1.2;
  MyWorksheet.WriteText(0, 0, 'There are images in cells A3 and B3'); //
  MyWorksheet.WriteImage(2, 0, image1, 1.0, 1.0, 2.0, 2.0);
  MyWorksheet.WriteImage(2, 1, image2, 1.0, 1.0);
                                         {
  MyWorksheet := MyWorkbook.AddWorksheet('Sheet 2');
  MyWorksheet.WriteText(0, 0, 'There is an image in cell B3');
  MyWorksheet.WriteImage(2, 1, image3);
//  MyWorksheet.WriteImage(0, 2, 'D:\Prog_Lazarus\svn\lazarus-ccr\components\fpspreadsheet\examples\read_write\ooxmldemo\laz_open.png');
//  MyWorksheet.WriteHyperlink(0, 0, 'http://www.chip.de');
//  MyWorksheet.PageLayout.AddHeaderImage(1, hfsLeft, 'D:\Prog_Lazarus\svn\lazarus-ccr\components\fpspreadsheet\examples\read_write\ooxmldemo\laz_open.png');
//  MyWorksheet.PageLayout.Headers[1] := '&LThis is a header&R&G';
                                          }
  // Save the spreadsheet to a file
  MyDir := ExtractFilePath(ParamStr(0));
  MyWorkbook.WriteToFile(MyDir + 'img.xlsx', sfOOXML, true);
  MyWorkbook.WriteToFile(MyDir + 'img.ods', sfOpenDocument, true);
//  MyWorkbook.WriteToFile(MyDir + 'img.xls', sfExcel8, true);
//  MyWorkbook.WriteToFile(MyDir + 'img5.xls', sfExcel5, true);
//  MyWorkbook.WriteToFile(MyDir + 'img2.xls', sfExcel2, true);

  if MyWorkbook.ErrorMsg <> '' then
  begin
    WriteLn(MyWorkbook.ErrorMsg);
  end;

  MyWorkbook.Free;
end.

