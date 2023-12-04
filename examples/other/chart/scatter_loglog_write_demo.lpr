program scatter_loglog_write_demo;

{$mode objfpc}{$H+}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'scatter_loglog';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsScatterSeries;
begin
  book := TsWorkbook.Create;
  try
    // Worksheet
    sheet := book.AddWorksheet('test');

    // Enter data
    sheet.WriteText(0, 0, 'Data');
    sheet.WriteFont(0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText  (2, 0, 'x');  sheet.Writetext   (2, 1, 'y');
    sheet.WriteNumber(3, 0, 0.1);  sheet.WriteFormula(3, 1, 'A4^2');
    sheet.WriteNumber(4, 0, 0.8);  sheet.WriteFormula(4, 1, 'A5^2');
    sheet.WriteNumber(5, 0, 1.9);  sheet.WriteFormula(5, 1, 'A6^2');
    sheet.WriteNumber(6, 0, 4.6);  sheet.WriteFormula(6, 1, 'A7^2');
    sheet.WriteNumber(7, 0, 8.3);  sheet.WriteFormula(7, 1, 'A8^2');
    sheet.WriteNumber(8, 0,15.9);  sheet.WriteFormula(8, 1, 'A9^2');
    sheet.WriteNumber(9, 0,25.6);  sheet.WriteFormula(9, 1, 'A10^2');
    sheet.WriteNumber(10,0,68.3);  sheet.WriteFormula(10,1, 'A11^2');

    // Create chart: left/top in cell D4, 150 mm x 100 mm
    ch := book.AddChart(sheet, 2, 3, 150, 100);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Legend.Border.Style := clsNoLine;
    ch.XAxis.Logarithmic := true;
    ch.YAxis.Logarithmic := true;

    // Add scatter series
    ser := TsScatterSeries.Create(ch);

    // Series properties
    ser.SetTitleAddr(0, 0);
    ser.SetXRange(3, 0, 10, 0);
    ser.SetYRange(3, 1, 10, 1);
    ser.ShowLines := true;
    ser.ShowSymbols := true;
    ser.Symbol := cssCircle;

    {
    book.WriteToFile(FILE_NAME + '.xlsx', true);   // Excel fails to open the file
    WriteLn('Data saved with chart to ', FILE_NAME, '.xlsx');
    }

    book.Options := book.Options + [boCalcBeforeSaving];
    book.WriteToFile(FILE_NAME + '.ods', true);
    WriteLn('Data saved with chart to ', FILE_NAME, '.ods');
  finally
    book.Free;
  end;
end.

