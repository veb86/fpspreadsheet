program regressionchart_write_demo;

{$mode objfpc}{$H+}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'regression';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsScatterSeries;
  dir, fn: String;
  rotated: Boolean;
begin
  fn := FILE_NAME;

  rotated := (ParamCount >= 1) and (lowercase(ParamStr(1)) = 'rotated');
  if rotated then
    fn := fn + '-rotated';

  dir := ExtractFilePath(ParamStr(0)) + 'files/';
  ForceDirectories(dir);

  book := TsWorkbook.Create;
  try
    // Worksheet
    sheet := book.AddWorksheet('regression_test');

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
    ch := book.AddChart(sheet, 150, 100, 2, 3);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Legend.Border.Style := clsNoLine;
    ch.RotatedAxes := rotated;

    // Add scatter series
    ser := TsScatterSeries.Create(ch);

    // Series properties
    ser.SetTitleAddr(0, 0);
    ser.SetXRange(3, 0, 8, 0);
    ser.SetYRange(3, 1, 8, 1);
    ser.ShowLines := false;
    ser.ShowSymbols := true;
    ser.Symbol := cssCircle;
    ser.SymbolFill.Color := ChartColor(scRed);
    ser.SymbolFill.Style := cfsSolid;
    ser.Trendline.Title := 'Fit curve';
    ser.Trendline.TrendlineType := tltPolynomial; //tltLinear;
    ser.Trendline.ExtrapolateForwardBy := 10;
    ser.Trendline.ExtrapolateBackwardBy := 10;
    ser.Trendline.Line.Color := ChartColor(scRed);
    ser.Trendline.Line.Style := clsDash;
    ser.Trendline.ForceYIntercept := true;  // not used by logarithmic, power
    ser.Trendline.YInterceptValue := 10.0;  // dto.
    ser.Trendline.PolynomialDegree := 2;
    ser.Trendline.DisplayEquation := true;
    ser.Trendline.DisplayRSquare := true;
    ser.Trendline.Equation.XName := 'X';
    ser.Trendline.Equation.YName := 'Y';
    ser.Trendline.Equation.Border.Style := clsSolid;
    ser.Trendline.Equation.Border.Color := ChartColor(scGray);
    ser.Trendline.Equation.Fill.Style := cfsSolid;
    ser.Trendline.Equation.Fill.Color := ChartColor(scSilver);
    ser.Trendline.Equation.NumberFormat := '0.000';

    // Fine-tuning the position of the trendline result box is not very
    // practical and error-prone because its is measure relative to the
    // top/left corner of the chart, but we don't know where the plotarea is.

    //ser.Trendline.Equation.Top := 5;
    //ser.Trendline.Equation.Left := 5;

    book.WriteToFile(dir + fn + '.xlsx', true);
    WriteLn('... ', fn + '.xlsx');

    book.WriteToFile(dir + fn + '.ods', true);
    WriteLn('... ', fn + '.ods');
  finally
    book.Free;
  end;
end.

