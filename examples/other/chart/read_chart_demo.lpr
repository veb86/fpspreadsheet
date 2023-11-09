program read_chart_demo;

uses
  SysUtils, TypInfo,
  fpSpreadsheet, fpsTypes, fpsChart, fpsOpenDocument;

const
  FILE_NAME = 'test.ods';
//  FILE_NAME = 'area.ods';
var
  book: TsWorkbook;
  sheet: TsWorksheet;
  chart: TsChart;
  i, j: Integer;
begin
  book := TsWorkbook.Create;
  try
    book.ReadFromFile(FILE_NAME);
    for i := 0 to book.GetChartCount-1 do
    begin
      chart := book.GetChartByIndex(i);
      sheet := book.GetWorksheetByIndex(chart.SheetIndex);

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
          'LineColor:', IntToHex(chart.Hatches[j].LineColor, 6), ' ',
          'Distance:', chart.Hatches[j].LineDistance:0:0, 'mm ',
          'Angle:', chart.Hatches[j].LineAngle:0:0, 'deg ',
          'Filled:', chart.Hatches[j].Filled);

      WriteLn  ('  GRADIENT STYLES: ');
      for j := 0 to chart.Gradients.Count-1 do
        WriteLn('    ', j, ': "', chart.Gradients[j].Name, '" ',
          GetEnumName(TypeInfo(TsChartGradientStyle), ord(chart.Gradients[j].Style)), ' ',
          'StartColor:', IntToHex(chart.Gradients[j].StartColor, 6), ' ',
          'EndColor:', IntToHex(chart.Gradients[j].EndColor, 6), ' ',
//          'StartIntensity:', chart.Gradients[j].StartIntensity*100:0:0, '% ',
//          'EndIntensity:', chart.Gradients[j].EndIntensity*100:0:0, '% ',
          'Border:', chart.Gradients[j].Border*100:0:0, '% ',
          'Angle:', chart.Gradients[j].Angle:0:0, 'deg ',
          'CenterX:', chart.Gradients[j].CenterX*100:0:0, '% ',
          'CenterY:', chart.Gradients[j].CenterY*100:0:0, '% ');

      WriteLn;
      WriteLn('  CHART BORDER');
      WriteLn('    Style:', chart.Border.Style,
                 ' Width:', chart.Border.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.Border.Color, 6),
                 ' Transparency:', chart.Border.Transparency:0:2);

      WriteLn;
      WriteLn('  CHART BACKGROUND');
      WriteLn('    Style:', GetEnumName(TypeInfo(TsChartFillStyle), ord(chart.Background.Style)),
                 ' Color:', IntToHex(chart.background.Color, 6),
                 ' Gradient:', chart.Background.Gradient,
                 ' Hatch:', chart.Background.Hatch,
                 ' Transparency:', chart.Background.Transparency:0:2);
      WriteLn;
      WriteLn('  CHART LEGEND');
      WriteLn('    Position: ', GetEnumName(TypeInfo(TsChartLegendPosition), ord(chart.Legend.Position)),
                 ' CanOverlapPlotArea:', chart.Legend.CanOverlapPlotArea);
      WriteLn('    Background: Style:', GetEnumName(TypeInfo(TsChartFillStyle), ord(chart.Legend.Background.Style)),
                 ' Color:', IntToHex(chart.Legend.Background.Color, 6),
                 ' Gradient:', chart.Legend.Background.Gradient,
                 ' Hatch:', chart.Legend.Background.Hatch,
                 ' Transparency:', chart.Legend.Background.Transparency);
      WriteLn('    Border: Style:', chart.Legend.Border.Style,
                 ' Width:', chart.Legend.Border.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.Legend.Border.Color, 6),
                 ' Transparency:', chart.Legend.Border.Transparency:0:2);
      WriteLn('    Font: "', chart.Legend.Font.FontName, '" Size:', chart.Legend.Font.Size:0:0,
                 ' Style:', SetToString(PTypeInfo(TypeInfo(TsFontStyles)), integer(chart.Legend.Font.Style), True),
                 ' Color:', IntToHex(chart.Legend.Font.Color, 6));

      WriteLn;
      WriteLn('  CHART TITLE');
      WriteLn('    Caption: "', StringReplace(chart.Title.Caption, FPS_LINE_ENDING, '\n', [rfReplaceAll]), '"',
                 ' Rotation: ', chart.Title.RotationAngle);
      WriteLn('    Background: Style:', GetEnumName(TypeInfo(TsChartFillStyle), ord(chart.Title.Background.Style)),
                 ' Color:', IntToHex(chart.Title.Background.Color, 6),
                 ' Gradient:', chart.Title.Background.Gradient,
                 ' Hatch:', chart.Title.Background.Hatch,
                 ' Transparency:', chart.Title.Background.Transparency);
      WriteLn('    Border: Style:', chart.Title.Border.Style,
                 ' Width:', chart.Title.Border.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.Title.Border.Color, 6),
                 ' Transparency:', chart.Title.Border.Transparency:0:2);
      WriteLn('    Font: "', chart.Title.Font.FontName, '" Size:', chart.Title.Font.Size:0:0,
                 ' Style:', SetToString(PTypeInfo(TypeInfo(TsFontStyles)), integer(chart.Title.Font.Style), True),
                 ' Color:', IntToHex(chart.Title.Font.Color, 6));

      WriteLn;
      WriteLn('  CHART SUBTITLE');
      WriteLn('    Caption: "', StringReplace(chart.Subtitle.Caption, FPS_LINE_ENDING, '\n', [rfReplaceAll]), '"',
                 ' Rotation: ', chart.Subtitle.RotationAngle);
      WriteLn('    Background: Style:', GetEnumName(TypeInfo(TsChartFillStyle), ord(chart.Subtitle.Background.Style)),
                 ' Color:', IntToHex(chart.Subtitle.Background.Color, 6),
                 ' Gradient:', chart.Subtitle.Background.Gradient,
                 ' Hatch:', chart.Subtitle.Background.Hatch,
                 ' Transparency:', chart.Subtitle.Background.Transparency);
      WriteLn('    Border: Style:', chart.Subtitle.Border.Style,
                 ' Width:', chart.Subtitle.Border.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.Subtitle.Border.Color, 6),
                 ' Transparency:', chart.Subtitle.Border.Transparency:0:2);
      WriteLn('    Font: "', chart.Subtitle.Font.FontName, '" Size:', chart.Subtitle.Font.Size:0:0,
                 ' Style:', SetToString(PTypeInfo(TypeInfo(TsFontStyles)), integer(chart.Subtitle.Font.Style), True),
                 ' Color:', IntToHex(chart.Subtitle.Font.Color, 6));

      WriteLn;
      WriteLn('  CHART X AXIS');
      WriteLn('    VISIBLE:', chart.YAxis.Visible);
      WriteLn('    TITLE: Caption: "', StringReplace(chart.XAxis.Title.Caption, FPS_LINE_ENDING, '\n', [rfReplaceAll]), '"',
                 ' Visible: ', chart.XAxis.Title.Visible,
                 ' Rotation: ', chart.XAxis.Title.RotationAngle,
                 ' Font: "', chart.XAxis.Title.Font.FontName, '" Size:', chart.XAxis.Title.Font.Size:0:0,
                 ' Style:', SetToString(PTypeInfo(TypeInfo(TsFontStyles)), integer(chart.XAxis.Title.Font.Style), True),
                 ' Color:', IntToHex(chart.XAxis.Title.Font.Color, 6));
      WriteLn('    RANGE: AutomaticMin:', chart.XAxis.AutomaticMin, ' Minimum: ', chart.XAxis.Min:0:3,
                 ' AutomaticMax:', chart.XAxis.AutomaticMax, ' Maximum: ', chart.XAxis.Max:0:3);
      WriteLn('    POSITION: ', GetEnumName(TypeInfo(TsChartAxisPosition), ord(chart.XAXis.Position)),
                 ' Value:', chart.XAxis.PositionValue:0:3);
      WriteLn('    AXIS TICKS: Major interval:', chart.XAxis.MajorInterval:0:2,
                 ' Major ticks:', SetToString(PTypeInfo(TypeInfo(TsChartAxisTicks)), integer(chart.XAxis.MajorTicks), True),
                 ' Minor count:', chart.XAxis.MinorCount,
                 ' Minor ticks:', SetToString(PTypeInfo(TypeInfo(TsChartAxisTicks)), integer(chart.XAxis.MinorTicks), True));
      WriteLn('    AXIS LINE: Style:', chart.XAxis.AxisLine.Style,
                 ' Width:', chart.XAxis.AxisLine.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.XAxis.AxisLine.Color, 6),
                 ' Transparency:', chart.XAxis.AxisLine.Transparency:0:2);
      WriteLn('    MAJOR GRID: Style:', chart.XAxis.MajorGridLines.Style,
                 ' Width:', chart.XAxis.MajorGridLines.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.XAxis.MajorGridLines.Color, 6),
                 ' Transparency:', chart.XAxis.MajorGridLines.Transparency:0:2);
      WriteLn('    MINOR GRID: Style:', chart.XAxis.MinorGridLines.Style,
                 ' Width:', chart.XAxis.MinorGridLines.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.XAxis.MinorGridLines.Color, 6),
                 ' Transparency:', chart.XAxis.MinorGridLines.Transparency:0:2);

      WriteLn;
      WriteLn('  CHART Y AXIS:');
      WriteLn('    VISIBLE:', chart.YAxis.Visible);
      WriteLn('    TITLE: Caption: "', StringReplace(chart.YAxis.Title.Caption, FPS_LINE_ENDING, '\n', [rfReplaceAll]), '"',
                 ' Visible: ', chart.YAxis.Title.Visible,
                 ' Rotation: ', chart.YAxis.Title.RotationAngle,
                 ' Font: "', chart.YAxis.Title.Font.FontName, '" Size:', chart.YAxis.Title.Font.Size:0:0,
                 ' Style:', SetToString(PTypeInfo(TypeInfo(TsFontStyles)), integer(chart.YAxis.Title.Font.Style), True),
                 ' Color:', IntToHex(chart.YAxis.Title.Font.Color, 6));
      WriteLn('    RANGE: AutomaticMin:', chart.YAxis.AutomaticMin, ' Minimum: ', chart.YAxis.Min:0:3,
                 ' AutomaticMax:', chart.YAxis.AutomaticMax, ' Maximum: ', chart.YAxis.Max:0:3);
      WriteLn('    POSITION: ', GetEnumName(TypeInfo(TsChartAxisPosition), ord(chart.YAXis.Position)),
                 ' Value:', chart.YAxis.PositionValue:0:3);
      WriteLn('    AXIS TICKS: Major interval:', chart.YAxis.MajorInterval:0:2,
                 ' Major ticks:', SetToString(PTypeInfo(TypeInfo(TsChartAxisTicks)), integer(chart.YAxis.MajorTicks), True),
                 ' Minor count:', chart.YAxis.MinorCount,
                 ' Minor ticks:', SetToString(PTypeInfo(TypeInfo(TsChartAxisTicks)), integer(chart.YAxis.MinorTicks), True));
      WriteLn('    AXIS LINE: Style:', chart.YAxis.AxisLine.Style,
                 ' Width:', chart.YAxis.AxisLine.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.YAxis.AxisLine.Color, 6),
                 ' Transparency:', chart.YAxis.AxisLine.Transparency:0:2);
      WriteLn('    MAJOR GRID: Style:', chart.YAxis.MajorGridLines.Style,
                 ' Width:', chart.YAxis.MajorGridLines.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.YAxis.MajorGridLines.Color, 6),
                 ' Transparency:', chart.YAxis.MajorGridLines.Transparency:0:2);
      WriteLn('    MINOR GRID: Style:', chart.YAxis.MinorGridLines.Style,
                 ' Width:', chart.YAxis.MinorGridLines.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.YAxis.MinorGridLines.Color, 6),
                 ' Transparency:', chart.YAxis.MinorGridLines.Transparency:0:2);
    end;

  finally
    book.Free;
  end;

  WriteLn;
  Write('Press ENTER to close...');
  ReadLn;
end.

