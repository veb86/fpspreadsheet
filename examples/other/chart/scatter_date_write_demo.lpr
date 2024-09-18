program scatter_date_write_demo;

{$mode objfpc}{$H+}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'scatter-date';
  FMT = 'yyyy-mm-dd';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsScatterSeries;
  dir, fn: String;

begin
  dir := 'files/';
  ForceDirectories(dir);
  fn := FILE_NAME;

  book := TsWorkbook.Create;
  try
    // Worksheet
    sheet := book.AddWorksheet('scatter_date');

    // Enter data
    sheet.WriteText( 0, 0, 'Data');
    sheet.WriteFont( 0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText( 2, 0,  'x');
    sheet.WriteText( 2, 1, 'y');

    sheet.WriteDateTime( 3, 0, EncodeDate(2024, 1, 1), FMT);  sheet.WriteNumber( 3, 1, 12.4);
    sheet.WriteDateTime( 4, 0, EncodeDate(2024, 2,15), FMT);  sheet.WriteNumber( 4, 1, 18.8);
    sheet.WriteDateTime( 5, 0, EncodeDate(2024, 6,20), FMT);  sheet.WriteNumber( 5, 1, 21.3);
    sheet.WriteDateTime( 6, 0, EncodeDate(2024, 7, 9), FMT);  sheet.WriteNumber( 6, 1, 20.5);
    sheet.WriteDateTime( 7, 0, EncodeDate(2024, 8,21), FMT);  sheet.WriteNumber( 7, 1, 22.9);
    sheet.WriteDateTime( 8, 0, EncodeDate(2024, 8,31), FMT);  sheet.WriteNumber( 8, 1, 19.4);
    sheet.WriteDateTime( 9, 0, EncodeDate(2024,11, 3), FMT);  sheet.WriteNumber( 9, 1, 17.7);
    sheet.WriteDateTime(10, 0, EncodeDate(2024,12,28), FMT);  sheet.WriteNumber(10, 1, 12.9);

    // Create chart: left/top in cell D4, 160 mm x 100 mm
    ch := sheet.AddChart(160, 100, 2, 2);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Legend.Border.Style := clsNoLine;
    ch.xAxis.LabelFormatDateTime := FMT;
    ch.xAxis.DateTime := true;

    // Add scatter series
    ser := TsScatterSeries.Create(ch);

    // Series properties
    ser.SetTitleAddr(0, 0);        // A1
    ser.SetXRange(3, 0, 10, 0);    // A4:A11
    ser.SetYRange(3, 1, 10, 1);    // B4:B11
    ser.ShowLines := true;
    ser.ShowSymbols := true;
    ser.Symbol := cssCircle;
    ser.SymbolFill.Style := cfsSolid;
    ser.SymbolFill.Color := ChartColor(scRed);
    ser.SymbolBorder.Style := clsNoLine;
//    ser.Line.Style := clsDash;
    ser.Line.Width := 0.5;  // mm

    book.WriteToFile(dir + fn + '.xlsx', true);
    WriteLn('... ', fn + '.xlsx');

    book.Options := book.Options + [boCalcBeforeSaving];
    book.WriteToFile(dir + fn + '.ods', true);
    WriteLn('... ', fn + '.ods');
  finally
    book.Free;
  end;
end.

