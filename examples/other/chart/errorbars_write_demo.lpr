program regressionchart_write_demo;

{$mode objfpc}{$H+}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'errorbars';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsScatterSeries;
  fn: String;
  rotated: Boolean;
begin
  fn := FILE_NAME;

  book := TsWorkbook.Create;
  try
    // Worksheet
    sheet := book.AddWorksheet('errorbars_test');

    // Enter data
    sheet.WriteText(0, 0, 'Data');
    sheet.WriteFont(0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText  (2, 0, 'x');  sheet.Writetext  (2, 1, 'y');
    sheet.WriteNumber(3, 0, 1.1);  sheet.WriteNumber(3, 1,  9.0);
    sheet.WriteNumber(4, 0, 1.9);  sheet.WriteNumber(4, 1, 20.5);
    sheet.WriteNumber(5, 0, 2.5);  sheet.WriteNumber(5, 1, 24.5);
    sheet.WriteNumber(6, 0, 3.1);  sheet.WriteNumber(6, 1, 33.2);
    sheet.WriteNumber(7, 0, 5.2);  sheet.WriteNumber(7, 1, 49.4);
    sheet.WriteNumber(8, 0, 6.8);  sheet.WriteNumber(8, 1, 71.3);

    // Create chart: left/top in cell D4, 150 mm x 100 mm
    ch := book.AddChart(sheet, 2, 3, 150, 100);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Legend.Border.Style := clsNoLine;

    // Add scatter series
    ser := TsScatterSeries.Create(ch);

    // Series properties
    ser.SetTitleAddr(0, 0);
    ser.SetXRange(3, 0, 8, 0);
    ser.SetYRange(3, 1, 8, 1);
    ser.ShowLines := false;
    ser.ShowSymbols := true;
    ser.Symbol := cssCircle;

    ser.XErrorBars.Visible := true;
    ser.XErrorBars.Kind := cebkConstant;
    ser.XErrorBars.ValuePos := 0.5;
    ser.XErrorBars.ValueNeg := 0.5;
    ser.XErrorBars.Line.Color := scRed;

    ser.YErrorBars.Visible := true;
    ser.YErrorBars.Kind := cebkPercentage;
    ser.YErrorBars.ValuePos := 10;  // percent
    ser.YErrorBars.ValueNeg := 10;  // percent
    ser.YErrorBars.Line.Color := scRed;

    {
    book.WriteToFile(fn + '.xlsx', true);   // Excel fails to open the file
    WriteLn('Data saved with chart to ', fn, '.xlsx');
    }

    book.WriteToFile(fn + '.ods', true);
    WriteLn('Data saved with chart to ', fn, '.ods');
  finally
    book.Free;
  end;
end.

