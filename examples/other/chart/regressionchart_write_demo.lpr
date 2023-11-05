program regressionchart_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;
var
  b: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsScatterSeries;
begin
  b := TsWorkbook.Create;
  try
    // worksheet
    sheet := b.AddWorksheet('regression_test');

    // Enter data
    sheet.WriteText(0, 0, 'Data');
    sheet.WriteFont(0, 0, '', 12, [fssBold], scBlack);
    sheet.WriteText  (2, 0, 'x');  sheet.Writetext  (2, 1, 'y');
    sheet.WriteNumber(3, 0, 1.1);  sheet.WriteNumber(3, 1, 0.9);
    sheet.WriteNumber(4, 0, 1.9);  sheet.WriteNumber(4, 1, 2.05);
    sheet.WriteNumber(5, 0, 2.5);  sheet.WriteNumber(5, 1, 2.45);
    sheet.WriteNumber(6, 0, 3.1);  sheet.WriteNumber(6, 1, 3.3);
    sheet.WriteNumber(7, 0, 5.2);  sheet.WriteNumber(7, 1, 4.9);
    sheet.WriteNumber(8, 0, 6.8);  sheet.WriteNumber(8, 1, 7.1);        // sheet.WriteChartColor(8, 2, $FF8080);

    // Create chart: left/top in cell D4, 120 mm x 100 mm
    ch := b.AddChart(sheet, 2, 3, 120, 100);

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
    ser.Regression.Title := 'Fit curve';
    ser.Regression.RegressionType := rtPolynomial; //rtLinear;
    ser.Regression.ExtrapolateForwardBy := 10;
    ser.Regression.ExtrapolateBackwardBy := 10;
    ser.Regression.Line.Color := scRed;
    ser.Regression.Line.Style := clsDash;
    ser.Regression.ForceYIntercept := true;  // not used by logarithmic, power
    ser.Regression.YInterceptValue := 1.0;
    ser.Regression.PolynomialDegree := 2;
    ser.Regression.DisplayEquation := true;
    ser.Regression.DisplayRSquare := true;
    ser.Regression.Equation.XName := 'X';
    ser.Regression.Equation.YName := 'Y';
    ser.Regression.Equation.Border.Style := clsSolid;
    ser.Regression.Equation.Border.Color := scRed;
    ser.Regression.Equation.Fill.Style := cfsSolid;
    ser.Regression.Equation.Fill.Color := scSilver;
    ser.Regression.Equation.NumberFormat := '0.000';
    //ser.Regression.Equation.Top := 5;
    //ser.Regression.Equation.Left := 5;

    b.WriteToFile('regression.xlsx', true);   // Excel fails to open the file
    b.WriteToFile('regression.ods', true);
  finally
    b.Free;
  end;
end.

