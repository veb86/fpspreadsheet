program write_chart_demo;

{.$DEFINE DARK_MODE}

uses
  SysUtils, fpspreadsheet, fpstypes, fpschart, xlsxooxml, fpsopendocument;
var
  b: TsWorkbook;
  sh1, sh2, sh3: TsWorksheet;
  ch: TsChart;
  ser: TsChartSeries;
  i: Integer;
begin
  b := TsWorkbook.Create;
  try
    // 1st sheet
    sh1 := b.AddWorksheet('test1');
    sh1.WriteText(0, 1, 'sin(x)');
    sh1.WriteText(0, 2, 'sin(x/2)');
    for i := 1 to 7 do
    begin
      sh1.WriteNumber(i, 0, i-1);
      sh1.WriteNumber(i, 1, sin(i-1));
      sh1.WriteNumber(i, 2, sin((i-1)/2));
    end;

    ch := b.AddChart(sh1, 4, 4, 160, 100);
    ser := TsLineSeries.Create(ch);
    ser.SetTitleAddr(0, 1);
    ser.SetLabelRange(1, 0, 7, 0);
    ser.SetYRange(1, 1, 7, 1);
    ser.Line.Color := scBlue;
    TsLineSeries(ser).ShowSymbols := true;
    TsLineSeries(ser).Symbol := cssCircle;

    ser := TsLineSeries.Create(ch);
    ser.SetTitleAddr(0, 2);
    ser.SetLabelRange(1, 0, 7, 0);
    ser.SetYRange(1, 2, 7, 2);
    ser.Line.Color := scRed;
    TsLineSeries(ser).ShowSymbols := true;
    TsLineSeries(ser).Symbol := cssDiamond;

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

    ch.XAxis.ShowLabels := true;
    ch.XAxis.LabelFont.Size := 8;
    ch.XAxis.LabelFont.Color := scRed;
    ch.XAxis.LabelFont.Style := [fssStrikeout];
    ch.XAxis.AxisLine.Color := scRed;
    ch.XAxis.Caption := 'This is the x axis';
    ch.XAxis.CaptionFont.Color := scRed;
    ch.XAxis.CaptionFont.Size := 12;
    ch.XAxis.Inverted := true;
    ch.XAxis.MajorGridLines.Color := scRed;
    ch.XAxis.MinorGridLines.Color := scBlue;
    ch.XAxis.MajorGridLines.Style := clsNoLine;//Solid;
    ch.XAxis.MinorGridLines.Style := clsNoLine; //Solid;

    ch.YAxis.ShowLabels := true;
    ch.YAxis.LabelFont.Size := 8;
    ch.YAxis.LabelFont.Color := scBlue;
    ch.YAxis.AxisLine.Color := scBlue;
    ch.YAxis.Caption := 'This is the y axis';
    ch.YAxis.CaptionFont.Color := scBlue;
    ch.YAxis.CaptionFont.Size := 12;
    ch.YAxis.LabelRotation := 90;
    ch.YAxis.CaptionRotation := 90;
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
    ch.Legend.Background.FgColor := scSilver;
    ch.Legend.Background.Style := fsSolidFill;

    // 2nd sheet
    sh2 := b.AddWorksheet('test2');

    // 3rd sheet
    sh3 := b.AddWorksheet('test3');
    sh3.WriteText(0, 1, 'cos(x)');
    sh3.WriteText(0, 2, 'sin(x)');
    for i := 1 to 7 do
    begin
      sh3.WriteNumber(i, 0, i-1);
      sh3.WriteNumber(i, 1, cos(i-1), nfFixed, 2);
      sh3.WriteNumber(i, 2, sin(i-1), nfFixed, 2);
    end;

    ch := b.AddChart(sh3, 1, 3, 125, 95);
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

    b.WriteToFile('test.xlsx', true);   // Excel fails to open the file
    b.WriteToFile('test.ods', true);
  finally
    b.Free;
  end;

  WriteLn;
  Write('Press ENTER to close...');
  ReadLn;
end.

