{ Included by fpspreadsheet.pas }

{ Chart support }

{@@ ----------------------------------------------------------------------------
  Creates a chart object with its top/left corner in the specified row/colum and
  having the specified width (in workbook units).
  Inserts the chart in the FCharts list of the workbook and returns the chart
  instance.
-------------------------------------------------------------------------------}
function TsWorkbook.AddChart(ASheet: TsBasicWorksheet; ARow, ACol: Cardinal;
  AWidth, AHeight: Double; AOffsetX: Double = 0.0; AOffsetY: Double = 0.0): TsChart;
begin
  Result := TsChart.Create;
  if (ASheet = nil) then
    raise Exception.Create('To do: Insert chart as new ChartSheet');
  Result.Workbook := self;
  Result.SheetIndex := GetWorksheetIndex(ASheet);
  Result.Row := ARow;
  Result.Col := ACol;
  Result.OffsetX := AOffsetX;
  Result.OffsetY := AOffsetY;
  Result.Width := AWidth;
  Result.Height := AHeight;
  Result.Index := FCharts.Add(Result);
end;

{@@ ----------------------------------------------------------------------------
  Returns the chart having the given index in the worksheet's chart list
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

