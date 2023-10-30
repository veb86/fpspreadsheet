program write_chart_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils, fpspreadsheet, fpstypes, fpsUtils, fpschart, xlsxooxml, fpsopendocument;
const
//  SERIES_CLASS: TsChartSeriesClass = TsAreaSeries;
//  SERIES_CLASS: TsChartSeriesClass = TsBarSeries;
//  SERIES_CLASS: TsChartSeriesClass = TsBubbleSeries;
//  SERIES_CLASS: TsChartSeriesClass = TsLineSeries;
//  SERIES_CLASS: TsChartSeriesClass = TsScatterSeries;
//  SERIES_CLASS: TsChartSeriesClass = TsRadarSeries;
  SERIES_CLASS: TsChartSeriesClass = TsPieSeries;
  r1 = 1;
  r2 = 8;
  FILL_COLORS: array[0..r2-r1] of TsColor = (scRed, scGreen, scBlue, scYellow, scMagenta, scSilver, scBlack, scOlive);
var
  b: TsWorkbook;
  sheet1, sheet2, sheet3: TsWorksheet;
  ch: TsChart;
  ser: TsChartSeries;
  i: Integer;
begin
  b := TsWorkbook.Create;
  try
    // 1st sheet
    sheet1 := b.AddWorksheet('test1');
    sheet1.WriteText(0, 1, '1+sin(x)');
    sheet1.WriteText(0, 2, '1+sin(x/2)');
    sheet1.WriteText(0, 3, 'Bubble Radius');
    sheet1.WriteText(0, 4, 'Fill Color');
    sheet1.WriteText(0, 5, 'Border Color');
    for i := r1 to r2-1 do
    begin
      // x values or labels
      sheet1.WriteNumber(i, 0, i-1);
      // 1st series y values
      sheet1.WriteNumber(i, 1, 1+sin(i-1));
      // 2nd series y values
      sheet1.WriteNumber(i, 2, 1+sin((i-1)/2));
      // Bubble radii
      sheet1.WriteNumber(i, 3, i*i);
      // Fill colors
      sheet1.WriteNumber(i, 4, FlipColorBytes(FILL_COLORS[i-r1]));  // !! ODS need red and blue channels exchanged !!
      // Border colors
      sheet1.WriteNumber(i, 5, FlipColorBytes(FILL_COLORS[r2-i]));
    end;
    sheet1.WriteNumber(r2, 0, 9);
    sheet1.WriteNumber(r2, 1, 2);
    sheet1.WriteNumber(r2, 2, 2.5);
    sheet1.WriteNumber(r2, 3, r2*r2);

    // Create chart
    ch := b.AddChart(sheet1, 4, 6, 160, 100);

    // Add first series (type depending on SERIES_CLASS)
    ser := SERIES_CLASS.Create(ch);
    ser.SetTitleAddr(0, 1);
    ser.SetLabelRange(r1, 0, r2, 0);
    ser.SetXRange(r1, 0, r2, 0);     // is used only by scatter series
    ser.SetYRange(r1, 1, r2, 1);
    ser.Line.Color := scBlue;
    ser.Fill.FgColor := scBlue;
    ser.SetFillColorRange(r1, 4, r2, 4);
    ser.DataLabels := [cdlPercentage, cdlSymbol];
    if (ser is TsLineSeries) then
    begin
      TsLineSeries(ser).ShowSymbols := true;
      TsLineSeries(ser).Symbol := cssCircle;
    end;
    if (ser is TsBubbleSeries) then
    begin
      TsBubbleSeries(ser).SetXRange(r1, 0, r2, 0);
      TsBubbleSeries(ser).SetYRange(r1, 2, r2, 2);
      TsBubbleSeries(ser).SetBubbleRange(r1, 3, r2, 3);
    end;

    if SERIES_CLASS <> TsBubbleSeries then
    begin
      // Add second series
      ser := SERIES_CLASS.Create(ch);
  //    ser := TsBarSeries.Create(ch);
      ser.SetTitleAddr(0, 2);
      ser.SetLabelRange(r1, 0, r2, 0);
      ser.SetXRange(r1, 0, r2, 0);
      ser.SetYRange(r1, 2, r2, 2);
      ser.Line.Color := scRed;
      ser.Fill.FgColor := scRed;
    end;

    {$IFDEF DARK_MODE}
    ch.Background.FgColor := scBlack;
    ch.Border.Color := scWhite;
    ch.PlotArea.Background.FgColor := $1F1F1F;
    {$ELSE}
    ch.Background.FgColor := scWhite;
    ch.Border.Color := scBlack;
    ch.PlotArea.Background.FgColor := $F0F0F0;
    {$ENDIF}
    // Background and wall working
    ch.Background.Style := fsSolidFill;
    ch.Border.Style := clsSolid;
    ch.PlotArea.Background.Style := fsSolidFill;
    //ch.RotatedAxes := true;
    //ch.StackMode := csmStackedPercentage;
    //ch.Interpolation := ciCubicSpline;

    ch.XAxis.ShowLabels := true;
    ch.XAxis.LabelFont.Size := 9;
    ch.XAxis.LabelFont.Color := scRed;
    //ch.XAxis.LabelFont.Style := [fssStrikeout];
    ch.XAxis.AxisLine.Color := scRed;
    ch.XAxis.Caption := 'This is the x axis';
    ch.XAxis.CaptionFont.Color := scRed;
    ch.XAxis.CaptionFont.Size := 12;
    //ch.XAxis.Inverted := true;
    ch.XAxis.MajorGridLines.Color := scRed;
    ch.XAxis.MinorGridLines.Color := scBlue;
    ch.XAxis.MajorGridLines.Style := clsNoLine; //Solid;
    ch.XAxis.MinorGridLines.Style := clsNoLine; //Solid;
    ch.XAxis.Position := capStart;

    ch.YAxis.ShowLabels := true;
    ch.YAxis.LabelFont.Size := 8;
    ch.YAxis.LabelFont.Color := scBlue;
    ch.YAxis.AxisLine.Color := scBlue;
    ch.YAxis.Caption := 'This is the y axis';
    ch.YAxis.CaptionFont.Color := scBlue;
    ch.YAxis.CaptionFont.Size := 12;
    //ch.YAxis.LabelRotation := 90;
    //ch.YAxis.CaptionRotation := 90;
    ch.YAxis.Min := -5;
    ch.yAxis.Max := 5;
    ch.YAxis.AutomaticMin := false;
    ch.YAxis.AutomaticMax := false;
    ch.YAxis.MajorGridLines.Color := scBlue;
    ch.YAxis.MajorGridLines.Style := clsLongDash; //clsSolid;
    ch.YAxis.MajorGridLines.Width := 0.5;  // mm
//    ch.YAxis.MinorGridLines.Style := clsLongDashDot; //Dash; //clsSolid;

    ch.Title.Caption := 'HALLO';
    ch.Title.Visible := true;
    ch.Title.Font.Color := scMagenta;
    ch.Title.Font.Size := 20;
    ch.Title.Font.Style := [fssBold];

    ch.SubTitle.Caption := 'hallo';
    ch.SubTitle.Visible := true;

    // Legend working
    ch.Legend.Font.Size := 12;
    ch.Legend.Font.Color := scBlue;
    ch.Legend.Border.Width := 0.3; // mm
    ch.Legend.Border.Color := scGray;
    ch.Legend.Background.FgColor := $F0F0F0;
    ch.Legend.Background.Style := fsSolidFill;
    //ch.Legend.CanOverlapPlotArea := true;
    ch.Legend.Position := lpBottom;

    // 2nd sheet
    sheet2 := b.AddWorksheet('test2');
    sheet2.WriteText(0, 0, 'abc');

    // 3rd sheet
    sheet3 := b.AddWorksheet('test3');
    sheet3.WriteText(0, 1, 'cos(x)');
    sheet3.WriteText(0, 2, 'sin(x)');
    for i := 1 to 7 do
    begin
      sheet3.WriteNumber(i, 0, i-1);
      sheet3.WriteNumber(i, 1, cos(i-1), nfFixed, 2);
      sheet3.WriteNumber(i, 2, sin(i-1), nfFixed, 2);
    end;

    ch := b.AddChart(sheet3, 1, 3, 125, 95);
    ser := TsLineSeries.Create(ch);
    ser.SetTitleAddr(0, 1);
    ser.SetLabelRange(1, 0, 7, 0);
    ser.SetYRange(1, 1, 7, 1);
    ser := TsLineSeries.Create(ch);
    ser.SetTitleAddr(0, 2);
    ser.SetLabelRange(1, 0, 7, 0);
    ser.SetYRange(1, 2, 7, 2);
    ch.Border.Style := clsNoLine;
    ch.Title.Caption := 'HALLO';
    ch.Title.Visible := true;
    ch.SubTitle.Caption := 'hallo';
    ch.Subtitle.Visible := true;
    ch.XAxis.MajorGridLines.Style := clsSolid; //NoLine;
    ch.XAxis.MinorGridLines.Style := clsNoLine;
    ch.YAxis.MajorGridLines.Style := clsNoLine;
    ch.YAxis.MinorGridLines.Style := clsNoLine;
    ch.YAxis.CaptionRotation := 0;
    ch.XAxis.CaptionFont.Size := 18;
    ch.YAxis.CaptionFont.Size := 18;
    ch.XAxis.LabelFont.Style := [fssItalic];
    ch.YAxis.LabelFont.Style := [fssItalic];
    ch.YAxis.MajorTicks := [catInside, catOutside];
    ch.YAxis.MinorTicks := [catOutside];

    b.WriteToFile('test.xlsx', true);   // Excel fails to open the file
    b.WriteToFile('test.ods', true);
  finally
    b.Free;
  end;

  WriteLn;
  Write('Press ENTER to close...');
 // ReadLn;
end.

