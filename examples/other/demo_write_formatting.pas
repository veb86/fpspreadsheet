{
test_write_formatting.pas

Demonstrates how to write an Excel 8+ file using the fpspreadsheet library

Adds formatting to the file

AUTHORS: Felipe Monteiro de Carvalho
}
program demo_write_formatting;

{$mode delphi}{$H+}

uses
  Classes, SysUtils, fpsTypes, fpspreadsheet, xlsbiff8, fpsopendocument,
  fpscell, laz_fpspreadsheet;

var
  MyWorkbook: TsWorkbook;
  MyWorksheet: TsWorksheet;
  MyDir: string;
  MyCell: PCell;

procedure WriteFirstWorksheet();
begin
  MyWorksheet := MyWorkbook.AddWorksheet('Worksheet1');

  // Write some cells
  MyWorksheet.WriteText(1, 0, 'Border'); // A2

  MyWorksheet.WriteText(1, 1, '[]');     // B2
  MyCell := MyWorksheet.GetCell(1, 1);
  MyCell.Border := [];

  MyWorksheet.WriteText(1, 3, '[N]');    // D2
  MyCell := MyWorksheet.GetCell(1, 3);
  MyCell^.Border := [cbNorth];

  MyWorksheet.WriteText(1, 5, '[W]');    // F2
  MyCell := MyWorksheet.GetCell(1, 5);
  MyCell^.Border := [cbWest];

  MyWorksheet.WriteText(1, 7, '[E]');    // H2
  MyCell := MyWorksheet.GetCell(1, 7);
  MyCell^.Border := [cbEast];

  MyWorksheet.WriteText(1, 9, '[S]');    // J2
  MyCell := MyWorksheet.GetCell(1, 9);
  MyCell^.Border := [cbSouth];

  MyWorksheet.WriteText(3, 1, '[N,W]');  // B4
  MyCell := MyWorksheet.GetCell(3, 1);
  MyCell^.Border := [cbNorth, cbWest];

  MyWorksheet.WriteText(3, 3, '[N,E]');  // D4
  MyCell := MyWorksheet.GetCell(3, 3);
  MyCell^.Border := [cbNorth, cbEast];

  MyWorksheet.WriteText(3, 5, '[N,S]');  // F4
  MyCell := MyWorksheet.GetCell(3, 5);
  MyCell^.Border := [cbNorth, cbSouth];

  MyWorksheet.WriteText(3, 7, '[W,E]');  // H4
  MyCell := MyWorksheet.GetCell(3, 7);
  MyCell^.Border := [cbWest, cbEast];

  MyWorksheet.WriteText(3, 9, '[W,S]');  // J4
  MyCell := MyWorksheet.GetCell(3, 9);
  MyCell^.Border := [cbWest, cbSouth];

  MyWorksheet.WriteText(3, 11, '[E,S]'); // L4
  MyCell := MyWorksheet.GetCell(3, 11);
  MyCell^.Border := [cbEast, cbSouth];

  MyWorksheet.WriteText(5, 1, '[N,W,E]');// B6
  MyCell := MyWorksheet.GetCell(5, 1);
  MyCell^.Border := [cbNorth, cbWest, cbEast];

  MyWorksheet.WriteText(5, 3, '[N,W,S]');// D6
  MyCell := MyWorksheet.GetCell(5, 3);
  MyCell^.Border := [cbNorth, cbWest, cbSouth];

  MyWorksheet.WriteText(5, 5, '[N,E,S]');// F6
  MyCell := MyWorksheet.GetCell(5, 5);
  MyCell^.Border := [cbNorth, cbEast, cbSouth];

  MyWorksheet.WriteText(5, 7, '[W,E,S]');// H6
  MyCell := MyWorksheet.GetCell(5, 7);
  MyCell^.Border := [cbWest, cbEast, cbSouth];

  MyWorksheet.WriteText(5, 9, '[N,W,E,S]');// J6
  MyCell := MyWorksheet.GetCell(5, 9);
  MyCell^.Border := [cbNorth, cbWest, cbEast, cbSouth];
  MyCell^.BackgroundColor := scGreen;
end;

procedure WriteSecondWorksheet();
begin
  MyWorksheet := MyWorkbook.AddWorksheet('Worksheet2');

  // Write some cells

  // Line 1

  MyWorksheet.WriteText(1, 1, 'Relatório');
  MyCell := MyWorksheet.GetCell(1, 1);
  MyCell^.Border := [cbNorth, cbWest, cbSouth];
  MyCell^.BackgroundColor := scGrey20pct;

  MyWorksheet.WriteText(1, 2, ' ');
  MyCell := MyWorksheet.GetCell(1, 2);
  MyCell^.Border := [cbNorth, cbEast, cbSouth];
  MyCell^.BackgroundColor := scGrey20pct;

  // Line 2

  MyWorksheet.WriteText(2, 1, 'Compras');
  MyCell := MyWorksheet.GetCell(2, 1);
  MyCell^.Border := [cbWest];
  MyCell^.BackgroundColor := scGrey10pct;

  MyWorksheet.WriteText(2, 2, 'R$ 20');
  MyCell := MyWorksheet.GetCell(2, 2);
  MyCell^.Border := [cbEast];
  MyCell^.BackgroundColor := scGrey10pct;

  // Line 3

  MyWorksheet.WriteText(3, 1, 'Total:');
  MyCell := MyWorksheet.GetCell(3, 1);
  MyCell^.Border := [cbWest, cbSouth];
  MyCell^.BackgroundColor := scGrey10pct;

  MyWorksheet.WriteText(3, 2, 'R$ 20');
  MyCell := MyWorksheet.GetCell(3, 2);
  MyCell^.Border := [cbEast, cbSouth];
  MyCell^.BackgroundColor := scGrey10pct;
end;

const
  TestFile='test3.odt';

{$R *.res}

begin
  writeln('Starting program.');
  MyDir := ExtractFilePath(ParamStr(0));

  // Create the spreadsheet
  MyWorkbook := TsWorkbook.Create;

  WriteFirstWorksheet();
  WriteSecondWorksheet();

  // Save the spreadsheet to a file
//  MyWorkbook.WriteToFile(MyDir + 'test3.xls', sfExcel8, False);
  MyWorkbook.WriteToFile(MyDir + TestFile, sfOpenDocument, False);
  MyWorkbook.Free;
  writeln('Finished. Please open "'+TestFile+'" in your spreadsheet program.');
end.

