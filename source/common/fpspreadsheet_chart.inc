{ Included by fpspreadsheet.pas }

{===============================================================================
                         Worksheet chart support
===============================================================================}

{$ifdef FPS_CHARTS}
{@@ ----------------------------------------------------------------------------
  Adds a chart to the worksheet.

  @param  AWidth   Width of the chart, in workbook units
  @param  AHeight  Height of the chart, in workbook units
  @param  ARow     Index of the row in which the top of the chart should be positioned.
  @param  ACol     Index of the column in which the left side of the chart should be positioned.
  @param  AOffsetX Distance of the chart's left side from the left side of column ACol, in workbook units
  @param  AOffsetY Distance of the chart's top side from the top side of row ARow, in workbook units
  @returns  The instance of the created chart
-------------------------------------------------------------------------------}
function TsWorksheet.AddChart(AWidth, AHeight: Double; ARow, ACol: Cardinal;
  AOffsetX: Double = 0.0; AOffsetY: Double = 0.0): TsChart;
begin
  Result := Workbook.AddChart(self, AWidth, AHeight, ARow, ACol, AOffsetX, AOffsetY);
end;

{@@ ----------------------------------------------------------------------------
  Determines the count of charts in this worksheet
-------------------------------------------------------------------------------}
function TsWorksheet.GetChartCount: Integer;
var
  i: Integer;
  chart: TsChart;
begin
  Result := 0;
  for i := 0 to Workbook.GetChartCount-1 do
  begin
    chart := Workbook.GetChartByIndex(i);
    if chart.Worksheet = self then inc(Result);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Collects all charts which are linked to this worksheet
-------------------------------------------------------------------------------}
function TsWorksheet.GetCharts: TsChartArray;
var
  i, j, n: Integer;
  chart: TsChart;
begin
  Result := nil;
  n := Workbook.GetChartCount;
  j := 0;
  SetLength(Result, n);
  for i := 0 to n-1 do
  begin
    chart := Workbook.GetChartByIndex(i);
    if chart.Worksheet = self then
    begin
      Result[j] := chart;
      inc(j);
    end;
  end;
  SetLength(Result, j);
end;

{@@ ----------------------------------------------------------------------------
  Destroys all charts from the worksheet and removes them from the workbook's
  internal FCharts list.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveAllCharts;
var
  i: Integer;
begin
  for i := Workbook.FCharts.Count-1 downto 0 do
    if Workbook.GetChartByIndex(i).Worksheet = Self then
      Workbook.FCharts.Delete(i);    // This destroys the chart
end;

{@@ ----------------------------------------------------------------------------
  Removes the specified chart from the worksheet, destroys it and removes it
  from the workbook's internal FCharts list.
-------------------------------------------------------------------------------}
procedure TsWorksheet.RemoveChart(AChart: TsChart);
var
  idx: Integer;
begin
  if AChart.Worksheet <> self then
    exit;

  idx := Workbook.FCharts.IndexOf(AChart);
  if idx > -1 then Workbook.FCharts.Delete(idx);  // This destroys the chart
end;
{$endif}


{===============================================================================
                          Workbook chart support
===============================================================================}

{@@ ----------------------------------------------------------------------------
  Creates a chart object with its top/left corner in the specified row/colum and
  having the specified width and height (in workbook units).
  Inserts the chart in the internal FCharts list of the workbook and returns
  the chart instance.

  @param  ASheet   Worksheet into which the chart will be inserted.
  @param  AWidth   Width of the chart, in workbook units (usually millimeters)
  @param  AHeight  Height of the chart, in workbook units
  @param  ARow     Index of the row in which the top of the chart should be positioned.
  @param  ACol     Index of the column in which the left side of the chart should be positioned.
  @param  AOffsetX Distance of the chart's left side from the left side of column ACol, in workbook units
  @param  AOffsetY Distance of the chart's top side from the top side of row ARow, in workbook units
  @returns  The instance of the created chart
-------------------------------------------------------------------------------}
function TsWorkbook.AddChart(ASheet: TsBasicWorksheet; AWidth, AHeight: Double;
  ARow, ACol: Cardinal; AOffsetX: Double = 0.0; AOffsetY: Double = 0.0): TsChart;
begin
  Result := TsChart.Create;
  if (ASheet = nil) then
    raise Exception.Create('To do: Insert chart as new ChartSheet');
  Result.Workbook := self;
  Result.Worksheet := ASheet;
  Result.Row := ARow;
  Result.Col := ACol;
  Result.OffsetX := AOffsetX;
  Result.OffsetY := AOffsetY;
  Result.Width := AWidth;
  Result.Height := AHeight;
  Result.Index := FCharts.Add(Result);
end;

{@@ ----------------------------------------------------------------------------
  Returns the chart having the given index in the workbook's chart list
-------------------------------------------------------------------------------}
function TsWorkbook.GetChartByIndex(AIndex: Integer): TsChart;
begin
  if (AIndex >= 0) and (AIndex < FCharts.Count) then
    Result := FCharts[AIndex]
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Returns the number of charts embedded on this sheet
-------------------------------------------------------------------------------}
function TsWorkbook.GetChartCount: Integer;
begin
  Result := FCharts.Count;
end;

{@@ ----------------------------------------------------------------------------
  Returns the index of the given chart in the workbook's chart list
-------------------------------------------------------------------------------}
function TsWorkbook.GetChartIndex(AChart: TsChart): Integer;
begin
  for Result := 0 to FCharts.Count-1 do
    if FCharts[Result] = AChart then
      exit;
  Result := -1;
end;

