program barchart_date_write_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils,
  fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;

const
  FILE_NAME = 'bars_date';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  ch: TsChart;
  ser: TsChartSeries;
  dir, fn: String;
  i: Integer;
begin
  fn := FILE_NAME;
  dir := 'files/';
  ForceDirectories(dir);

  book := TsWorkbook.Create;
  try
    // worksheet
    sheet := book.AddWorksheet('bar_series');

    // Enter data
    sheet.WriteText( 0, 0, 'Sales');
    sheet.WriteFont( 0, 0, '', 12, [fssBold], scBlack);

    sheet.WriteText( 2, 0, '');
    sheet.WriteText( 2, 1, 'Product A');
    sheet.WriteText( 2, 2, 'Product B');

    sheet.WriteDateTime( 3, 0, EncodeDate(2024,1,1), 'mmm yyyy' );   sheet.WriteNumber( 3, 1, 12);   sheet.WriteNumber( 3, 2, 15);
    sheet.WriteDateTime( 4, 0, EncodeDate(2024,2,1), 'mmm yyyy' );   sheet.WriteNumber( 4, 1, 11);   sheet.WriteNumber( 4, 2, 13);
    sheet.WriteDateTime( 5, 0, EncodeDate(2024,3,1), 'mmm yyyy' );   sheet.WriteNumber( 5, 1, 16);   sheet.WriteNumber( 5, 2, 11);
    sheet.WriteDateTime( 6, 0, EncodeDate(2024,4,1), 'mmm yyyy' );   sheet.WriteNumber( 6, 1, 18);   sheet.WriteNumber( 6, 2, 11);
    sheet.WriteDateTime( 7, 0, EncodeDate(2024,5,1), 'mmm yyyy' );   sheet.WriteNumber( 7, 1, 16);   sheet.WriteNumber( 7, 2,  7);
    sheet.WriteDateTime( 8, 0, EncodeDate(2024,6,1), 'mmm yyyy' );   sheet.WriteNumber( 8, 1, 10);   sheet.WriteNumber( 8, 2, 17);

    // Create chart: left/top in cell D4, 160 mm x 100 mm
    ch := sheet.AddChart(160, 100, 2, 3);

    // Chart properties
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'Sales';
    ch.Title.Font.Style := [fssBold];
    ch.Legend.Border.Style := clsNoLine;
    //ch.XAxis.DateTime := true;                        // Switches the axis to date/time labels, not needed for Excel
    //ch.XAxis.LabelFormatDateTime := 'mm yyyy';        // Defines the label format of the dates
    ch.XAxis.Title.Caption := '';
    ch.YAxis.Title.Caption := 'Total';
    ch.YAxis.AxisLine.Color := ChartColor(scSilver);
    ch.YAxis.MajorTicks := [];
    ch.BarGapWidthPercent := 75;

    // Add 1st bar series ("Product A")
    ser := TsBarSeries.Create(ch);
    ser.SetTitleAddr(2, 1);             // series 1, title in cell B3
    ser.SetLabelRange(3, 0, 8, 0);      // series 1, x labels in A4:A11
    ser.SetYRange(3, 1, 8, 1);          // series 1, y values in B4:B11
    ser.Line.Color := ChartColor(scDarkRed);
    ser.Fill.Style := cfsSolidHatched;
    ser.Fill.Hatch := ch.Hatches.AddLineHatch('Crossed', chsDouble, ChartColor(scDarkRed), 2, 0.1, 45);
    ser.Fill.Color := ChartColor(scRed);
    ser.DataLabels := [cdlValue];        // Show sales as datapoint labels

    // Add 2nd bar series ("Product B")
    ser := TsBarSeries.Create(ch);
    ser.SetTitleAddr(2, 2);              // series 2, title in cell C3
    ser.SetLabelRange(3, 0, 8, 0);      // series 2, x labels in A4:A11
    ser.SetYRange(3, 2, 8, 2);          // series 2, y values in C4:C11
    ser.Line.Color := ChartColor(scDarkBlue);
    ser.Fill.Style := cfsSolidHatched;
    ser.Fill.Hatch := ch.Hatches.AddLineHatch('Forward', chsSingle, ChartColor(scWhite), 1.5, 0.1, 45);
    ser.Fill.Color := ChartColor(scBlue);
    ser.DataLabels := [cdlValue];        // Show sales as datapoint labels

    book.WriteToFile(dir + fn + '.xlsx', true);
    WriteLn('... ', fn + '.xlsx');

    book.WriteToFile(dir + fn + '.ods', true);
    WriteLn('... ', fn + '.ods');
  finally
    book.Free;
  end;
end.

