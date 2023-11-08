program read_chart_demo;

uses
  SysUtils, TypInfo,
  fpSpreadsheet, fpsTypes, fpsChart, fpsOpenDocument;

const
//  FILE_NAME = 'test.ods';
  FILE_NAME = 'area.ods';
var
  b: TsWorkbook;
  sheet: TsWorksheet;
  chart: TsChart;
  i, j: Integer;
begin
  b := TsWorkbook.Create;
  try
    b.ReadFromFile(FILE_NAME);
    for i := 0 to b.GetChartCount-1 do
    begin
      chart := b.GetChartByIndex(i);
      sheet := b.GetWorksheetByIndex(chart.SheetIndex);
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
                 ' Hatch:', chart.Background.Hatch);
      WriteLn;
      WriteLn('  CHART LEGEND');
      WriteLn('    Position: ', GetEnumName(TypeInfo(TsChartLegendPosition), ord(chart.Legend.Position)),
                 ' CanOverlapPlotArea:', chart.Legend.CanOverlapPlotArea);
      WriteLn('    Background: Style:', GetEnumName(TypeInfo(TsChartFillStyle), ord(chart.Legend.Background.Style)),
                 ' Color:', IntToHex(chart.Legend.Background.Color, 6),
                 ' Gradient:', chart.Legend.Background.Gradient,
                 ' Hatch:', chart.Legend.Background.Hatch);
      WriteLn('    Border: Style:', chart.Legend.Border.Style,
                 ' Width:', chart.Legend.Border.Width:0:0, 'mm',
                 ' Color:', IntToHex(chart.Legend.Border.Color, 6),
                 ' Transparency:', chart.Legend.Border.Transparency:0:2);
      WriteLn('    Font: "', chart.Legend.Font.FontName, '" Size:', chart.Legend.Font.Size:0:0,
                 ' Style:', SetToString(PTypeInfo(TypeInfo(TsFontStyles)), integer(chart.Legend.Font.Style), True),
                 ' Color:', IntToHex(chart.Legend.Font.Color, 6));
    end;

  finally
    b.Free;
  end;

  WriteLn;
  Write('Press ENTER to close...');
  ReadLn;
end.

