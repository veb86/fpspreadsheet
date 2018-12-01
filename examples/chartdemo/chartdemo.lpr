program chartdemo;

uses
  fpspreadsheet, fpsutils, fpstypes, fpschart;

var
  wb: TsWorkbook;
  sh: TsWorksheet;
  chart: TsChart;
  ser: TsChartSeries;
  idx: Integer;

begin
  wb := TsWorkbook.Create;
  try
    ws := wb.AddWorksheet('Test');
    // x values
    ws.WriteNumber(0, 0, 1.0);
    ws.WriteNumber(1, 0, 2.1);
    ws.WriteNumber(2, 0, 2.9);
    ws.WriteNumber(3, 0, 4.15);
    ws.WriteNumber(4, 0, 5.05);
    // y values
    ws.WriteNumber(0, 1, 10.0);
    ws.WriteNumber(1, 1, 12.0);
    ws.WriteNumber(2, 1,  9.0);
    ws.WriteNumber(3, 1,  7.5);
    ws.WriteNumber(4, 1, 11.2);

    idx := ws.WriteChart(0, 0, 12, 9);
    chart := ws.GetChart(idx);
    ser := TsLineSeries.Create(chart);
    ser.XRange := Range(0, 0, 4, 0);
    ser.YRange := Range(0, 1, 4, 1);
    ser.Title := 'Scatter series';
    ser.ShowSymbols := true;
    ser.ShowLines := true;

    chart.AddSeries(
  finally
    wb.Free;
  end;
end.

