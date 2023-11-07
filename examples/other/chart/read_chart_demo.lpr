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
      WriteLn('  in worksheet "', sheet.Name, '", ',
        'row:', chart.Row, ' (+',chart.OffsetY:0:0, 'mm) ',
        'col:', chart.Col, ' (+',chart.OffsetX:0:0, 'mm) ',
        'width:', chart.Width:0:0, 'mm height:', chart.Height:0:0,  'mm');

      Write('  Line styles: ');
      for j := 0 to chart.LineStyles.Count-1 do
        Write('"', chart.GetLineStyle(j).Name, '" ');
      WriteLn;

      WriteLn  ('  Hatch styles: ');
      for j := 0 to chart.Hatches.Count-1 do
        WriteLn('                "', chart.Hatches[j].Name, '" ',
          GetEnumName(TypeInfo(TsChartHatchStyle), ord(chart.Hatches[j].Style)), ' ',
          'Line color: ', IntToHex(chart.Hatches[j].LineColor, 6), ' ',
          'Distance: ', chart.Hatches[j].LineDistance:0:0, 'mm ',
          'Angle: ', chart.Hatches[j].LineAngle:0:0, 'deg ',
          'Filled: ', chart.Hatches[j].Filled);
    end;

  finally
    b.Free;
  end;

  WriteLn;
  Write('Press ENTER to close...');
  ReadLn;
end.

