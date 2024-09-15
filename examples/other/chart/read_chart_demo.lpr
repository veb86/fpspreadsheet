program read_chart_demo;

uses
  SysUtils, TypInfo,
  fpSpreadsheet, fpsTypes, fpsUtils, fpsChart, fpsOpenDocument;

function GetFontStr(AFont: TsFont): String;
begin
  Result := Format('Name="%s", Size=%.0f, Style=%s, Color=%.6x', [
    AFont.FontName,
    AFont.Size,
    SetToString(PTypeInfo(TypeInfo(TsFontStyles)), integer(AFont.Style), True),
    AFont.Color
  ]);
end;

function GetFillStr(AFill: TsChartFill): String;
begin
  Result := Format('Style=%s, Color=%.6x, Gradient=%d, Hatch=%d, Transparency=%.2f', [
    GetEnumName(TypeInfo(TsChartFillStyle), ord(AFill.Style)),
    AFill.Color.Color, AFill.Gradient, AFill.Hatch, AFill.Color.Transparency
  ]);
end;

function GetLineStr(ALine: TsChartLine): String;
var
  s: String;
begin

  if ALine.Style = -1 then
    s := 'solid'
  else if ALine.Style = -2 then
    s := 'noLine'
  else if ALine.Style = clsFineDot then
    s := 'fine-dot'
  else if ALine.Style = clsDot then
    s := 'dot'
  else if ALine.Style = clsDash then
    s := 'dash'
  else if ALine.Style = clsDashDot then
    s := 'dash-dot'
  else if ALine.Style = clsLongDash then
    s := 'long dash'
  else if ALine.Style = clsLongDashDot then
    s := 'long dash-dot'
  else if ALine.Style = clsLongDashDotDot then
    s := 'long dash-dot-dot'
  else
    s := 'custom #' + IntToStr(ALine.Style);

  Result := Format('Style=%s, Width=%.0fmm, Color=%.6x, Transparency=%.2f', [
    s, ALine.Width, ALine.Color.Color, ALine.Color.Transparency
  ]);
end;

function GetRangeStr(ARange: TsChartRange): String;
begin
  with ARange do
    Result := GetCellRangeString(Sheet1, Sheet2, Row1, Col1, Row2, Col2, rfAllRel, false);
  if Result = '' then
    Result := '(none)';
end;

function GetCellAddrStr(ACellAddr: TsChartCellAddr): String;
begin
  with ACellAddr do
    Result := GetCellRangeString(Sheet, Sheet, Row, Col, Row, Col, rfAllRel, false);
end;

const
  FOLDER = 'files/';

// Uncomment one of the following FILE_NAMEs

//  FILE_NAME = 'test.ods';
//  FILE_NAME = 'area.ods';
//  FILE_NAME = 'bars.ods';
  FILE_NAME = 'regression.ods';
//  FILE_NAME = 'pie.ods';
//  FILE_NAME = 'radar.ods';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  chart: TsChart;
  series: TsChartSeries;
  trendline: TsChartTrendline;
  i, j: Integer;
begin
  FormatSettings.DecimalSeparator := '.';

  book := TsWorkbook.Create;
  try
    book.ReadFromFile(FOLDER + FILE_NAME);
    for i := 0 to book.GetChartCount-1 do
    begin
      chart := book.GetChartByIndex(i);
      sheet := TsWorksheet(chart.Worksheet);

      WriteLn('--------------------------------------------------------------------------------');
      WriteLn('Chart "', chart.Name, '":');
      WriteLn('  Worksheet "', sheet.Name, '", ',
        'row:', chart.Row, ' (+',chart.OffsetY:0:0, 'mm) ',
        'col:', chart.Col, ' (+',chart.OffsetX:0:0, 'mm) ',
        'width:', chart.Width:0:0, 'mm height:', chart.Height:0:0,  'mm');

      Write('  LINE STYLES: ');
      for j := 0 to chart.LineStyles.Count-1 do
        Write('"', chart.GetLineStyle(j).Name, '" ');
      WriteLn;

      WriteLn  ('  HATCH STYLES: ');
      for j := 0 to chart.Hatches.Count-1 do
        WriteLn('    ', j, ': "', chart.Hatches[j].Name, '" ',
          GetEnumName(TypeInfo(TsChartHatchStyle), ord(chart.Hatches[j].Style)), ' ',
          'PatternColor:', IntToHex(chart.Hatches[j].PatternColor.Color, 6), ' ',
          'Distance:', chart.Hatches[j].PatternWidth:0:0, 'mm ',
          'Angle:', chart.Hatches[j].PatternAngle:0:0, 'deg ');

      WriteLn  ('  GRADIENT STYLES: ');
      for j := 0 to chart.Gradients.Count-1 do
        WriteLn('    ', j, ': "', chart.Gradients[j].Name, '" ',
          GetEnumName(TypeInfo(TsChartGradientStyle), ord(chart.Gradients[j].Style)), ' ',
          'StartColor:', IntToHex(chart.Gradients[j].StartColor.Color, 6), ' ',
          'EndColor:', IntToHex(chart.Gradients[j].EndColor.Color, 6), ' ',
//          'StartIntensity:', chart.Gradients[j].StartIntensity*100:0:0, '% ',
//          'EndIntensity:', chart.Gradients[j].EndIntensity*100:0:0, '% ',
          'Border:', chart.Gradients[j].Border*100:0:0, '% ',
          'Angle:', chart.Gradients[j].Angle:0:0, 'deg ',
          'CenterX:', chart.Gradients[j].CenterX*100:0:0, '% ',
          'CenterY:', chart.Gradients[j].CenterY*100:0:0, '% ');

      WriteLn;
      WriteLn('  CHART BORDER        ', GetLineStr(chart.Border));
      WriteLn('  CHART BACKGROUND    ',GetFillStr(chart.Background));
      WriteLn;

      WriteLn('  CHART LEGEND        Position=', GetEnumName(TypeInfo(TsChartLegendPosition), ord(chart.Legend.Position)),
                                  ', CanOverlapPlotArea=', chart.Legend.CanOverlapPlotArea);
      WriteLn('                      Background: ', GetFillStr(chart.Legend.Background));
      WriteLn('                      Border: ', GetLineStr(chart.Legend.Border));
      WriteLn('                      Font: ', GetFontStr(chart.Legend.Font));
      WriteLn;

      WriteLn('  CHART TITLE         Caption="', StringReplace(chart.Title.Caption, FPS_LINE_ENDING, '\n', [rfReplaceAll]), '"',
                                  ', Rotation=', chart.Title.RotationAngle);
      WriteLn('                      Background: ', GetFillStr(chart.Title.Background));
      WriteLn('                      Border: ', GetLineStr(chart.Title.Border));
      WriteLn('                      Font: ', GetFontStr(chart.Title.Font));
      WriteLn;

      WriteLn('  CHART SUBTITLE      Caption="', StringReplace(chart.Subtitle.Caption, FPS_LINE_ENDING, '\n', [rfReplaceAll]), '"',
                                  ', Rotation=', chart.Subtitle.RotationAngle);
      WriteLn('                      Background: ', GetFillStr(chart.Subtitle.Background));
      WriteLn('                      Border: ', GetLineStr(chart.SubTitle.Border));
      WriteLn('                      Font: ', GetFontStr(chart.Subtitle.Font));
      WriteLn;

      WriteLn('  CHART X AXIS        Visible=', chart.XAxis.Visible);
      WriteLn('    TITLE             Caption="', StringReplace(chart.XAxis.Title.Caption, FPS_LINE_ENDING, '\n', [rfReplaceAll]), '"');
      WriteLn('                      Visible=', chart.XAxis.Title.Visible, ', Rotation=', chart.XAxis.Title.RotationAngle);
      WriteLn('                      Font: ', GetFontStr(chart.XAxis.Title.Font));
      WriteLn('    CATEGORIES        ', GetRangeStr(chart.XAxis.CategoryRange));
      WriteLn('    RANGE             AutomaticMin=', chart.XAxis.AutomaticMin, ', Minimum=', chart.XAxis.Min:0:3);
      WriteLn('                      AutomaticMax=', chart.XAxis.AutomaticMax, ', Maximum=', chart.XAxis.Max:0:3);
      WriteLn('    LABELS            Format="', chart.XAxis.LabelFormat, '"');
      WriteLn('    POSITION          ', GetEnumName(TypeInfo(TsChartAxisPosition), ord(chart.XAXis.Position)),
                                  ', Value=', chart.XAxis.PositionValue:0:3);
      WriteLn('    AXIS TICKS:       Major interval=', chart.XAxis.MajorInterval:0:2,
                                  ', Major ticks=', SetToString(PTypeInfo(TypeInfo(TsChartAxisTicks)), integer(chart.XAxis.MajorTicks), True));
      WriteLn('                      Minor count=', chart.XAxis.MinorCount,
                                  ', Minor ticks=', SetToString(PTypeInfo(TypeInfo(TsChartAxisTicks)), integer(chart.XAxis.MinorTicks), True));
      WriteLn('    AXIS LINE         ', GetLineStr(chart.XAxis.AxisLine));
      WriteLn('    MAJOR GRID        ', GetLineStr(chart.XAxis.MajorGridLines));
      WriteLn('    MINOR GRID        ', GetLineStr(chart.XAxis.MinorGridLines));

      WriteLn;
      WriteLn('  CHART Y AXIS        Visible=', chart.YAxis.Visible);
      WriteLn('    TITLE             Caption="', StringReplace(chart.YAxis.Title.Caption, FPS_LINE_ENDING, '\n', [rfReplaceAll]), '"');
      WriteLn('                      Visible=', chart.YAxis.Title.Visible, ', Rotation: ', chart.YAxis.Title.RotationAngle);
      WriteLn('                      Font: ', GetFontStr(chart.YAxis.Title.Font));
      WriteLn('    RANGE             AutomaticMin=', chart.YAxis.AutomaticMin, ', Minimum=', chart.YAxis.Min:0:3);
      WriteLn('                      AutomaticMax=', chart.YAxis.AutomaticMax, ', Maximum=', chart.YAxis.Max:0:3);
      WriteLn('    LABELS            Format="', chart.YAxis.LabelFormat, '", FormatPercent="', chart.YAxis.LabelFormatPercent,'"');
      WriteLn('    POSITION          ', GetEnumName(TypeInfo(TsChartAxisPosition), ord(chart.YAXis.Position)),
                                  ', Value:', chart.YAxis.PositionValue:0:3);
      WriteLn('    AXIS TICKS        Major interval=', chart.YAxis.MajorInterval:0:2,
                                  ', Major ticks=', SetToString(PTypeInfo(TypeInfo(TsChartAxisTicks)), integer(chart.YAxis.MajorTicks), True));
      WriteLn('                      Minor count=', chart.YAxis.MinorCount,
                                  ', Minor ticks=', SetToString(PTypeInfo(TypeInfo(TsChartAxisTicks)), integer(chart.YAxis.MinorTicks), True));
      WriteLn('    AXIS LINE         ', GetLineStr(chart.YAxis.AxisLine));
      WriteLn('    MAJOR GRID        ', GetLineStr(chart.YAxis.MajorGridLines));
      WriteLn('    MINOR GRID        ', GetLineStr(chart.YAxis.MinorGridLines));

      for j := 0 to chart.Series.Count-1 do
      begin
        series := chart.Series[j];
        WriteLn;
        WriteLn(  '  SERIES #', j, ': ', series.ClassName);
        WriteLn(  '    TITLE:            ', GetCellAddrStr(series.TitleAddr));
        WriteLn(  '    LABEL RANGE:      ', GetRangeStr(series.LabelRange), ', Format="', series.LabelFormat, '"');
        if (series is TsScatterSeries) or (series is TsBubbleSeries) then
          WriteLn('    X RANGE:          ', GetRangeStr(series.XRange));
        WriteLn(  '    Y RANGE:          ', GetRangeStr(series.YRange));
        WriteLn(  '    FILL COLOR RANGE: ', GetRangeStr(series.FillColorRange));
        WriteLn(  '    LINE COLOR RANGE: ', GetRangeStr(series.LineColorRange));
        if series is TsBubbleSeries then
          WriteLn('    BUBBLE RANGE:     ', GetRangeStr(TsBubbleSeries(series).BubbleRange));

        if series is TsLineSeries then with TsLineSeries(series) do
        begin
          Write(  '    SYMBOLS:          ');
          if ShowSymbols then
            WriteLn('Symbol=', GetEnumName(TypeInfo(TsChartSeriesSymbol), ord(Symbol)),
                    ', Width=', SymbolWidth:0:1, 'mm',
                    ', Height=', SymbolHeight:0:1, 'mm')
          else
            WriteLn('none');
        end;

        WriteLn(  '    FILL:             ', GetFillStr(series.Fill));
        WriteLn(  '    LINES:            ', GetLineStr(series.Line));

        if (series is TsScatterSeries) and (TsScatterSeries(series).TrendLine.TrendLineType <> tltNone) then
        begin
          trendline := TsScatterSeries(series).TrendLine;
          with trendline do
          begin
            Write('    TRENDLINE:        ');
            Write(  'Type=', GetEnumName(TypeInfo(TsTrendlineType), ord(TrendlineType)));
            if TrendlineType = tltPolynomial then
              Write( ', PolynomialDegree=', PolynomialDegree);
            Write(   ', ForceYIntercept=', ForceYIntercept);
            if ForceYIntercept then
              Write( ', YInterceptValue=', YInterceptValue:0:2);
            WriteLn;
            WriteLn('                      ExtrapolateForwardBy=', ExtrapolateForwardBy:0:2,
                                        ', ExtrapolateBackwardBy=', ExtrapolateBackwardBy:0:2);
            WriteLn('                      DisplayEquation=', DisplayEquation,
                                        ', DisplayRSquare=', DisplayRSquare);
          end;
          if (trendline.DisplayEquation or trendline.DisplayRSquare) then
          begin
            with trendline.Equation do
            begin
              WriteLn('    TRENDLINE EQU:    XName="', XName,'", YName="', YName,'", Number format="', NumberFormat, '"');
              WriteLn('                      FONT:   ', GetFontStr(trendline.Equation.Font));
              WriteLn('                      FILL:   ', GetFillStr(trendline.Equation.Fill));
              WriteLn('                      BORDER: ', GetLineStr(trendline.Equation.Border));
            end;
          end;

        end;

      end;
    end;

  finally
    book.Free;
  end;

  WriteLn;
  Write('Press ENTER to close...');
  ReadLn;
end.

