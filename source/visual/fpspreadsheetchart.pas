{ fpspreadsheetchart.pas }

{@@ ----------------------------------------------------------------------------
Chart data source designed to work together with TChart from Lazarus
to display the data and with FPSpreadsheet to load data.

AUTHORS: Felipe Monteiro de Carvalho, Werner Pamler

LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
         distribution, for details about the license.
-------------------------------------------------------------------------------}

unit fpspreadsheetchart;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, LResources, Forms, Controls, Graphics, Dialogs,
  // TChart
  TATypes, TATextElements, TAChartUtils, TACustomSource, TACustomSeries, TASeries,
  TAChartAxisUtils, TAChartAxis, TALegend, TAGraph,
  // FPSpreadsheet Visual
  fpSpreadsheetCtrls, fpSpreadsheetGrid, fpsVisualUtils,
  // FPSpreadsheet
  fpsTypes, fpSpreadsheet, fpsUtils, fpsChart;

type

  {@@ Chart data source designed to work together with TChart from Lazarus
    to display the data.

    The data can be loaded from a TsWorksheetGrid Grid component or
    directly from a TsWorksheet FPSpreadsheet Worksheet }

  { TsWorkbookChartSource }

  TsXYLRange = (rngX, rngY, rngLabel, rngColor);

  TsWorkbookChartSource = class(TCustomChartSource, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FWorksheets: array[TsXYLRange] of TsWorksheet;
    FRangeStr: array[TsXYLRange] of String;
    FRanges: array[TsXYLRange] of TsCellRangeArray;
    FPointsNumber: Cardinal;
    FTitleCol, FTitleRow: Cardinal;
    FTitleSheetName: String;
    function GetRange(AIndex: TsXYLRange): String;
    function GetTitle: String;
    function GetWorkbook: TsWorkbook;
    procedure GetXYItem(ARangeIndex:TsXYLRange; APointIndex: Integer;
      out ANumber: Double; out AText: String);
    procedure SetRange(AIndex: TsXYLRange; const AValue: String);
    procedure SetRangeFromChart(AIndex: TsXYLRange; const ARange: TsChartRange);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    FCurItem: TChartDataItem;
    function BuildRangeStr(AIndex: TsXYLRange; AListSeparator: char = #0): String;
    function CountValues(AIndex: TsXYLRange): Integer;
    function GetCount: Integer; override;
    function GetItem(AIndex: Integer): PChartDataItem; override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure Prepare; overload;
    procedure Prepare(AIndex: TsXYLRange); overload;
    procedure SetYCount(AValue: Cardinal); override;
  public
    destructor Destroy; override;
    procedure Reset;
    procedure SetColorRange(ARange: TsChartRange);
    procedure SetLabelRange(ARange: TsChartRange);
    procedure SetXRange(ARange: TsChartRange);
    procedure SetYRange(ARange: TsChartRange);
    procedure SetTitleAddr(Addr: TsChartCellAddr);
    property PointsNumber: Cardinal read FPointsNumber;
    property Workbook: TsWorkbook read GetWorkbook;
  public
    // Interface to TsWorkbookSource
    procedure ListenerNotification(AChangedItems: TsNotificationItems; AData: Pointer = nil);
    procedure RemoveWorkbookSource;
  published
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
    property ColorRange: String index rngColor read GetRange write SetRange;
    property LabelRange: String index rngLabel read GetRange write SetRange;
    property XRange: String index rngX read GetRange write SetRange;
    property YRange: String index rngY read GetRange write SetRange;
    property Title: String read GetTitle;
  end;

  {@@ Link between TAChart and the fpspreadsheet chart class }

  { TsWorkbookChartLink }

  TsWorkbookChartLink = class(TComponent, IsSpreadsheetControl)
  private
    FChart: TChart;
    FWorkbookSource: TsWorkbookSource;
    FWorkbook: TsWorkbook;
    FWorkbookChartIndex: Integer;
    procedure SetChart(AValue: TChart);
    procedure SetWorkbookChartIndex(AValue: Integer);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);

  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

    procedure AddSeries(ASeries: TsChartSeries);
    procedure FixAreaSeries(AWorkbookChart: TsChart);
    procedure ClearChart;
    function GetWorkbookChart: TsChart;
    procedure UpdateChartAxis(AWorkbookAxis: TsChartAxis);
    procedure UpdateChartAxisLabels(AWorkbookChart: TsChart);
    procedure UpdateChartBackground(AWorkbookChart: TsChart);
    procedure UpdateBarSeries(AWorkbookChart: TsChart);
    procedure UpdateChartBrush(AWorkbookFill: TsChartFill; ABrush: TBrush);
    procedure UpdateChartLegend(AWorkbookLegend: TsChartLegend; ALegend: TChartLegend);
    procedure UpdateChartPen(AWorkbookLine: TsChartLine; APen: TPen);
    procedure UpdateChartTitle(AWorkbookTitle: TsChartText; AChartTitle: TChartTitle);

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure UpdateChart;

    { Interfacing with WorkbookSource}
    procedure ListenerNotification(AChangedItems: TsNotificationItems; AData: Pointer = nil);
    procedure RemoveWorkbookSource;

  published
    property Chart: TChart read FChart write SetChart;
    property WorkbookChartIndex: Integer read FWorkbookChartIndex write SetWorkbookChartIndex;
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;

  end;


implementation

uses
  Math;

function mmToPx(mm: Double; ppi: Integer): Integer;
begin
  Result := round(mmToIn(mm * ppi));
end;

{------------------------------------------------------------------------------}
{                             TsWorkbookChartSource                            }
{------------------------------------------------------------------------------}

{@@ ----------------------------------------------------------------------------
  Destructor of the WorkbookChartSource.
  Removes itself from the WorkbookSource's listener list.
-------------------------------------------------------------------------------}
destructor TsWorkbookChartSource.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited Destroy;
end;

{@@ ----------------------------------------------------------------------------
  Constructs the range string from the stored internal information. Is needed
  to have the worksheet name in the range string in order to make the range
  string unique.
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.BuildRangeStr(AIndex: TsXYLRange;
  AListSeparator: Char = #0): String;
var
  L: TStrings;
  range: TsCellRange;
begin
  if (Workbook = nil) or (FWorksheets[AIndex] = nil) or (Length(FRanges) = 0) then
    exit('');

  L := TStringList.Create;
  try
    if AListSeparator = #0 then
      L.Delimiter := Workbook.FormatSettings.ListSeparator
    else
      L.Delimiter := AListSeparator;
    L.StrictDelimiter := true;
    for range in FRanges[AIndex] do
      L.Add(GetCellRangeString(range, rfAllRel, true));
    Result := FWorksheets[AIndex].Name + SHEETSEPARATOR + L.DelimitedText;
  finally
    L.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Counts the number of x or y values contained in the x/y ranges

  @param   AIndex   Identifies whether values in the x or y ranges are counted.
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.CountValues(AIndex: TsXYLRange): Integer;
var
  range: TsCellRange;
begin
  Result := 0;
  for range in FRanges[AIndex] do
  begin
    if range.Col1 = range.Col2 then
      inc(Result, range.Row2 - range.Row1 + 1)
    else
    if range.Row1 = range.Row2 then
      inc(Result, range.Col2 - range.Col1 + 1)
    else
      raise Exception.Create('x/y ranges can only be 1 column wide or 1 row high.');
  end;
end;

{@@ ----------------------------------------------------------------------------
  Inherited ChartSource method telling the series how many data points are
  available
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.GetCount: Integer;
begin
  Result := FPointsNumber;
end;

{@@ ----------------------------------------------------------------------------
  Main ChartSource method called from the series requiring data for plotting.
  Retrieves the data from the workbook.

  @param   AIndex   Index of the data point in the series.
  @return  Pointer to a TChartDataItem record containing the x and y coordinates,
           the data point mark text, and the individual data point color.
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.GetItem(AIndex: Integer): PChartDataItem;
var
  dummyNumber: Double;
  dummyString: String;
  tmpLabel: String;
begin
  if FRanges[rngX] <> nil then
    GetXYItem(rngX, AIndex, FCurItem.X, tmpLabel)
  else
    FCurItem.X := AIndex;

  GetXYItem(rngY, AIndex, FCurItem.Y, dummyString);

  GetXYItem(rngLabel, AIndex, dummyNumber, FCurItem.Text);
  if FCurItem.Text = '' then FCurItem.Text := tmpLabel;

  if FRanges[rngColor] <> nil then
  begin
    GetXYItem(rngColor, AIndex, dummyNumber, dummyString);
    FCurItem.Color := round(dummyNumber);
  end else
    FCurItem.Color := clDefault;

  Result := @FCurItem;
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the cell range used for x or y coordinates or x labels

  @param   AIndex   Determines whether the methods deals with x, y values or
                    vakze labels.
  @return  An Excel string containing workbookname and cell block(s) in A1
           notation. Multiple blocks are separated by the ListSeparator defined
           by the workbook's FormatSettings.
-------------------------------------------------------------------------------}
function TsWorkbookChartsource.GetRange(AIndex: TsXYLRange): String;
begin
  Result := FRangeStr[AIndex];
end;

function TsWorkbookChartSource.GetTitle: String;
var
  sheet: TsWorksheet;
begin
  Result := '';
  if FWorkbookSource = nil then
    exit;
  sheet := FWorkbookSource.Workbook.GetWorksheetByName(FTitleSheetName);
  if sheet <> nil then
    Result := sheet.ReadAsText(FTitleRow, FTitleCol);
end;

{@@ ----------------------------------------------------------------------------
  Getter method for the linked workbook
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := WorkbookSource.Workbook
  else
    Result := nil;
//  FWorkbook := Result;
end;

{@@ ----------------------------------------------------------------------------
  Helper method the prepare the information required for the series data point.

  @param  ARangeIndex  Identifies whether the method retrieves the x or y
                       coordinate, or the label text
  @param  APointIndex  Index of the data point for which the data are required
  @param  ANumber      (output) x or y coordinate of the data point
  @param  AText        Data point marks label text
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.GetXYItem(ARangeIndex:TsXYLRange;
  APointIndex: Integer; out ANumber: Double; out AText: String);
var
  range: TsCellRange;
  idx: Integer;
  len: Integer;
  row, col: Cardinal;
  cell: PCell;
begin
  ANumber := NaN;
  AText := '';
  if FRanges[ARangeIndex] = nil then
    exit;
  if FWorksheets[ARangeIndex] = nil then
    exit;

  cell := nil;
  idx := 0;

  for range in FRanges[ARangeIndex] do
  begin
    if (range.Col1 = range.Col2) then  // vertical range
    begin
      len := range.Row2 - range.Row1 + 1;
      if (APointIndex >= idx) and (APointIndex < idx + len) then
      begin
        row := longint(range.Row1) + APointIndex - idx;
        col := range.Col1;
        break;
      end;
      inc(idx, len);
    end else  // horizontal range
    if (range.Row1 = range.Row2) then
    begin
      len := longint(range.Col2) - range.Col1 + 1;
      if (APointIndex >= idx) and (APointIndex < idx + len) then
      begin
        row := range.Row1;
        col := longint(range.Col1) + APointIndex - idx;
        break;
      end;
    end else
      raise Exception.Create('Ranges can only be 1 column wide or 1 row high');
  end;

  cell := FWorksheets[ARangeIndex].FindCell(row, col);

  if cell <> nil then
    case cell^.ContentType of
      cctUTF8String:
        begin
          ANumber := APointIndex;
          AText := FWorksheets[ARangeIndex].ReadAsText(cell);
        end;
      else
        ANumber := FWorksheets[ARangeIndex].ReadAsNumber(cell);
        AText := '';
    end;
end;

{@@ ----------------------------------------------------------------------------
  Notification message received from the WorkbookSource telling which
  spreadsheet item has changed.
  Responds to workbook changes by reading the worksheet names into the tabs,
  and to worksheet changes by selecting the tab corresponding to the selected
  worksheet.

  @param  AChangedItems  Set with elements identifying whether workbook,
                         worksheet, cell content or cell formatting has changed
  @param  AData          Additional data, not used here

  @see    TsNotificationItem
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
var
  ir: Integer;
  cell: PCell;
  ResetDone: Boolean;
  rng: TsXYLRange;
begin
  Unused(AData);

  // Workbook has been successfully loaded, all sheets are ready
  if (lniWorkbook in AChangedItems) then
    Prepare;

  // Used worksheet has been renamed?
  if (lniWorksheetRename in AChangedItems) then
    for rng in TsXYLRange do
      if TsWorksheet(AData) = FWorksheets[rng] then begin
        FRangeStr[rng] := BuildRangeStr(rng);
        Prepare(rng);
      end;

  // Used worksheet will be deleted?
  if (lniWorksheetRemoving in AChangedItems) then
    for rng in TsXYLRange do
      if TsWorksheet(AData) = FWorksheets[rng] then begin
        FWorksheets[rng] := nil;
        FRangeStr[rng] := BuildRangeStr(rng);
        Prepare(rng);
      end;

  // Cell changes: Enforce recalculation of axes if modified cell is within the
  // x or y range(s).
  if (lniCell in AChangedItems) and (Workbook <> nil) then
  begin
    cell := PCell(AData);
    if (cell <> nil) then begin
      ResetDone := false;
      for rng in TsXYLRange do
        for ir:=0 to High(FRanges[rng]) do
        begin
          if FWorksheets[rng].CellInRange(cell^.Row, cell^.Col, FRanges[rng, ir]) then
          begin
            Reset;
            ResetDone := true;
            break;
          end;
        if ResetDone then break;
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification: The ChartSource is notified that the
  WorkbookSource is being removed.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.Notification(AComponent: TComponent;
  Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) and (AComponent = FWorkbookSource) then
    SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Parses the x and y cell range strings and extracts internal information
  (worksheet used, cell range coordinates)
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.Prepare;
begin
  Prepare(rngColor);
  Prepare(rngLabel);
  Prepare(rngX);
  Prepare(rngY);
end;

{@@ ----------------------------------------------------------------------------
  Parses the range string of the data specified by AIndex and extracts internal
  information (worksheet used, cell range coordinates)

  @param  AIndex   Identifies whether x or y or label or color cell ranges are
                   analyzed
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.Prepare(AIndex: TsXYLRange);
var
  range: TsCellRange;
begin
  if (Workbook = nil) or (FRangeStr[AIndex] = '') then
  begin
    FWorksheets[AIndex] := nil;
    SetLength(FRanges[AIndex], 0);
    if AIndex = rngY then
      FPointsNumber := 0;
    Reset;
    exit;
  end;

  if Workbook.TryStrToCellRanges(FRangeStr[AIndex], FWorksheets[AIndex], FRanges[AIndex])
  then begin
    for range in FRanges[AIndex] do
      if (range.Col1 <> range.Col2) and (range.Row1 <> range.Row2) then
        raise Exception.Create('x/y ranges can only be 1 column wide or 1 row high');
    FPointsNumber := Max(CountValues(rngX), CountValues(rngY));
    // If x and y ranges are of different size empty data points will be plotted.
    Reset;
    // Make sure to include worksheet name in RangeString.
    FRangeStr[AIndex] := BuildRangeStr(AIndex);
  end else
  if (Workbook.GetWorksheetCount > 0) then begin
    if FWorksheets[AIndex] = nil then
      exit;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Removes the link of the ChartSource to the WorkbookSource.
  Required before destruction.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Resets internal buffers and notfies chart elements of the changes,
  in particular, enforces recalculation of axis limits
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.Reset;
begin
  InvalidateCaches;
  Notify;
end;

procedure TsWorkbookChartSource.SetColorRange(ARange: TsChartRange);
begin
  SetRangeFromChart(rngColor, ARange);
end;

procedure TsWorkbookChartSource.SetLabelRange(ARange: TsChartRange);
begin
  SetRangeFromChart(rngLabel, ARange);
end;

{@@ ----------------------------------------------------------------------------
  Shared method to set the cell ranges for x, y, labels or colors directly from
  the chart ranges.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetRangeFromChart(AIndex: TsXYLRange;
  const ARange: TsChartRange);
begin
  if ARange.Sheet1 <> ARange.Sheet2 then
    raise Exception.Create('A chart cell range can only be from a single worksheet.');
  SetLength(FRanges[AIndex], 1);
  FRanges[AIndex,0].Row1 := ARange.Row1;  // FIXME: Assuming here single-block range !!!
  FRanges[AIndex,0].Col1 := ARange.Col1;
  FRanges[AIndex,0].Row2 := ARange.Row2;
  FRanges[AIndex,0].Col2 := ARange.Col2;
  FWorksheets[AIndex] := FworkbookSource.Workbook.GetWorksheetByName(ARange.Sheet1);
  if AIndex in [rngX, rngY] then
    FPointsNumber := Max(CountValues(rngX), CountValues(rngY));
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the cell range used for x or y data (or labels) in the chart
  If it does not contain the worksheet name the currently active worksheet of
  the WorkbookSource is assumed.

  @param   AIndex     Distinguishes whether the method deals with x, y or
                      label ranges.
  @param   AValue     String in Excel syntax containing the cell range to be
                      used for x or y (depending on AIndex). Can contain multiple
                      cell blocks which must be separator by the ListSeparator
                      character defined in the Workbook's FormatSettings.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetRange(AIndex: TsXYLRange;
  const AValue: String);
begin
  FRangeStr[AIndex] := AValue;
  Prepare;
end;

procedure TsWorkbookChartSource.SetTitleAddr(Addr: TsChartCellAddr);
begin
  FTitleRow := Addr.Row;
  FTitleCol := Addr.Col;
  FTitleSheetName := Addr.GetSheetName;
end;

procedure TsWorkbookChartSource.SetXRange(ARange: TsChartRange);
begin
  SetRangeFromChart(rngX, ARange);
end;

procedure TsWorkbookChartSource.SetYRange(ARange: TsChartRange);
begin
  SetRangeFromChart(rngY, ARange);
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the WorkbookSource
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
    FWorkbookSource.AddListener(self);
//  FWorkbook := GetWorkbook;
  ListenerNotification([lniWorkbook, lniWorksheet]);
  Prepare;
end;

{@@ ----------------------------------------------------------------------------
  Inherited ChartSource method telling the series how many y values are used.
  Currently we support only single valued data (YCount = 1, by default).
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetYCount(AValue: Cardinal);
begin
  FYCount := AValue;
end;


{------------------------------------------------------------------------------}
{                             TsWorkbookChartLink                              }
{------------------------------------------------------------------------------}

constructor TsWorkbookChartLink.Create(AOwner: TComponent);
begin
  inherited;
  FWorkbookChartIndex := -1;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the WorkbookChartLink.
  Removes itself from the WorkbookSource's listener list.
-------------------------------------------------------------------------------}
destructor TsWorkbookChartLink.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  inherited;
end;

procedure TsWorkbookChartLink.AddSeries(ASeries: TsChartSeries);
const
  POINTER_STYLES: array[TsChartSeriesSymbol] of TSeriesPointerstyle = (
    psRectangle,
    psDiamond,
    psTriangle,
    psDownTriangle,
    psLeftTriangle,
    psRightTriangle,
    psCircle,
    psStar,
    psDiagCross,
    psCross,
    psFullStar
  );
var
  src: TsWorkbookChartSource;
  ser: TChartSeries;
  ppi: Integer;
begin
  src := TsWorkbookChartSource.Create(self);
  src.WorkbookSource := FWorkbookSource;
  if not ASeries.LabelRange.IsEmpty then src.SetLabelRange(ASeries.LabelRange);
  if not ASeries.XRange.IsEmpty then src.SetXRange(ASeries.XRange);
  if not ASeries.YRange.IsEmpty then src.SetYRange(ASeries.YRange);
  if not ASeries.FillColorRange.IsEmpty then src.SetColorRange(ASeries.FillColorRange);

  ppi := GetParentForm(FChart).PixelsPerInch;
  case ASeries.ChartType of
    ctBar:
      begin
        ser := TBarSeries.Create(FChart);
        UpdateChartBrush(ASeries.Fill, TBarSeries(ser).BarBrush);
        UpdateChartPen(ASeries.Line, TBarSeries(ser).BarPen);
      end;
    ctLine, ctScatter:
      begin
        ser := TLineSeries.Create(FChart);
        UpdateChartPen(ASeries.Line, TLineSeries(ser).LinePen);
        TLineSeries(ser).ShowLines := ASeries.Line.Style <> clsNoLine;
        TLineSeries(ser).ShowPoints := TsLineSeries(ASeries).ShowSymbols;
        if TLineSeries(ser).ShowPoints then
        begin
          UpdateChartBrush(ASeries.Fill, TLineSeries(ser).Pointer.Brush);
          TLineSeries(ser).Pointer.Pen.Color := TLineSeries(ser).LinePen.Color;
          TLineSeries(ser).Pointer.Style := POINTER_STYLES[TsLineSeries(ASeries).Symbol];
          TlineSeries(ser).Pointer.HorizSize := mmToPx(TsLineSeries(ASeries).SymbolWidth, ppi);
          TlineSeries(ser).Pointer.VertSize := mmToPx(TsLineSeries(ASeries).SymbolHeight, ppi);
        end;
      end;
    ctArea:
      begin
        ser := TAreaSeries.Create(FChart);
        UpdateChartBrush(ASeries.Fill, TAreaSeries(ser).AreaBrush);
        UpdateChartPen(ASeries.Line, TAreaSeries(ser).AreaContourPen);
        TAreaSeries(ser).AreaLinesPen.Style := psClear;
      end;
  end;
  src.SetTitleAddr(ASeries.TitleAddr);
  ser.Source := src;
  ser.Title := src.Title;
  ser.Transparency := round(ASeries.Fill.Transparency);
  FChart.AddSeries(ser);
end;

procedure TsWorkbookChartLink.ClearChart;
var
  i, j: Integer;
  ser: TChartSeries;
  src: TCustomChartSource;
begin
  if FChart = nil then
    exit;

  // Clear chart sources
  for i := 0 to FChart.SeriesCount-1 do
  begin
    if (FChart.Series[i] is TChartSeries) then
    begin
      ser :=  TChartSeries(FChart.Series[i]);
      src := ser.Source;
      if src is TsWorkbookChartSource then
        src.Free;
    end;
  end;

  // Clear the series
  FChart.ClearSeries;

  // Clear the axes
  for i := FChart.AxisList.Count-1 downto 0 do
  begin
    case FChart.AxisList[i].Alignment of
      calLeft, calBottom:
        FChart.AxisList[i].Title.Caption := '';
      calTop, calRight:
        FChart.AxisList.Delete(i);
    end;
    for j := FChart.AxisList[i].Minors.Count-1 downto 0 do
      FChart.AxisList[i].Minors.Delete(j);
  end;

  // Clear the title
  FChart.Title.Text.Clear;

  // Clear the footer
  FChart.Foot.Text.Clear;
end;

// Fix area series zero level not being clipped at chart's plotrect.
procedure TsWorkbookChartLink.FixAreaSeries(AWorkbookChart: TsChart);
var
  i: Integer;
  ser: TAreaSeries;
  ext: TDoubleRect;
begin
  if AWorkbookChart.GetChartType <> ctArea then
    exit;

  ext := FChart.LogicalExtent;
  for i := 0 to FChart.SeriesCount-1 do
    if FChart.Series[i] is TAreaSeries then
    begin
      ser := TAreaSeries(FChart.Series[i]);
      if ser.ZeroLevel < ext.a.y then
        ser.ZeroLevel := ext.a.y;
      if ser.ZeroLevel > ext.b.y then
        ser.ZeroLevel := ext.b.y;
      ser.UseZeroLevel := true;
    end;
end;

function TsWorkbookChartLink.GetWorkbookChart: TsChart;
begin
  if (FWorkbook <> nil) and (FWorkbookChartIndex > -1) then
    Result := FWorkbook.GetChartByIndex(FWorkbookChartIndex)
  else
    Result := nil;
end;

procedure TsWorkbookChartLink.ListenerNotification(AChangedItems: TsNotificationItems;
  AData: Pointer = nil);
begin
  // to be completed
end;

procedure TsWorkbookChartLink.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited Notification(AComponent, Operation);
  if (Operation = opRemove) then
  begin
    if (AComponent = FWorkbookSource) then
      SetWorkbookSource(nil)
    else
    if (AComponent = FChart) then
      SetChart(nil);
  end;
end;

procedure TsWorkbookChartLink.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

procedure TsWorkbookChartLink.SetChart(AValue: TChart);
begin
  if FChart = AValue then
    exit;
  FChart := AValue;
  UpdateChart;
end;

procedure TSWorkbookChartLink.SetWorkbookChartIndex(AValue: Integer);
begin
  if AValue = FWorkbookChartIndex then
    exit;
  FWorkbookChartIndex := AValue;
  UpdateChart;
end;

procedure TsWorkbookChartLink.SetWorkbookSource(AValue: TsWorkbookSource);
begin
  if AValue = FWorkbookSource then
    exit;
  if FWorkbookSource <> nil then
    FWorkbookSource.RemoveListener(self);
  FWorkbookSource := AValue;
  if FWorkbookSource <> nil then
  begin
    FWorkbookSource.AddListener(self);
    FWorkbook := FWorkbookSource.Workbook;
  end else
    FWorkbook := nil;
  ListenerNotification([lniWorkbook, lniWorksheet]);
  UpdateChart;
end;

procedure TsWorkbookChartLink.UpdateChart;
var
  ch: TsChart;
  i: Integer;
begin
  if (FChart = nil) then
    exit;
  if (FWorkbookSource = nil) or (FWorkbookChartIndex < 0) then
  begin
    ClearChart;
    exit;
  end;

  ch := GetWorkbookChart;
  UpdateChartBackground(ch);
  UpdateChartTitle(ch.Title, FChart.Title);
  UpdateChartTitle(ch.Subtitle, FChart.Foot);
  UpdateChartLegend(ch.Legend, FChart.Legend);
  UpdateChartAxis(ch.XAxis);
  UpdateChartAxis(ch.YAxis);
  UpdateChartAxis(ch.X2Axis);
  UpdateChartAxis(ch.Y2Axis);

  for i := 0 to ch.Series.Count-1 do
    AddSeries(ch.Series[i]);

  FChart.Prepare;
  UpdateChartAxisLabels(ch);
  UpdateBarSeries(ch);
  FixAreaSeries(ch);
end;

procedure TsWorkbookChartLink.UpdateChartAxis(AWorkbookAxis: TsChartAxis);
var
  align: TChartAxisAlignment;
  axis: TChartAxis;
  minorAxis: TChartMinorAxis;
begin
  if AWorkbookAxis = nil then
    exit;
  if AWorkbookAxis = AWorkbookAxis.Chart.XAxis then
    align := calBottom
  else if AWorkbookAxis = AWorkbookAxis.Chart.X2Axis then
    align := calTop
  else if AWorkbookAxis = AWorkbookAxis.Chart.YAxis then
    align := calLeft
  else if AWorkbookAxis = AWorkbookAxis.Chart.Y2Axis then
    align := calRight
  else
    raise Exception.Create('Unsupported axis alignment');
  axis := FChart.AxisList.GetAxisByAlign(align);

  if AWorkbookAxis.Visible and (axis = nil) then
  begin
    axis := FChart.AxisList.Add;
    axis.Alignment := align;
  end;

  if axis = nil then
    exit;

  // Entire axis visible?
  axis.Visible := AWorkbookAxis.Visible;

  // Axis title
  axis.Title.Caption := AWorkbookAxis.Title.Caption;
  axis.Title.Visible := true;
  Convert_sFont_to_Font(AWorkbookAxis.Title.Font, axis.Title.LabelFont);

  // Labels
  Convert_sFont_to_Font(AWorkbookAxis.LabelFont, axis.Marks.LabelFont);
//  if not AWorkbookAxis.CategoryRange.IsEmpty then   --- fix me
//    axis.Marks.Style := smsLabel;

  // Axis line
  UpdateChartPen(AWorkbookAxis.AxisLine, axis.AxisPen);
  axis.AxisPen.Visible := axis.AxisPen.Style <> psClear;

  // Major axis grid
  UpdateChartPen(AWorkbookAxis.MajorGridLines, axis.Grid);
  axis.Grid.Visible := axis.Grid.Style <> psClear;
  axis.TickLength := IfThen(catOutside in AWorkbookAxis.MajorTicks, 4, 0);
  axis.TickInnerLength := IfThen(catInside in AWorkbookAxis.MajorTicks, 4, 0);
  axis.TickColor := axis.Grid.Color;
  axis.TickWidth := axis.Grid.Width;

  // Minor axis grid
  if AWorkbookAxis.MinorGridLines.Style <> clsNoLine then
  begin
    minorAxis := axis.Minors.Add;
    UpdateChartPen(AWorkbookAxis.MinorGridLines, minorAxis.Grid);
    minorAxis.Grid.Visible := true;
    minorAxis.Intervals.Count := AWorkbookAxis.MinorCount;
    minorAxis.TickLength := IfThen(catOutside in AWorkbookAxis.MinorTicks, 2, 0);
    minorAxis.TickInnerLength := IfThen(catInside in AWorkbookAxis.MinorTicks, 2, 0);
    minorAxis.TickColor := minorAxis.Grid.Color;
    minorAxis.TickWidth := minorAxis.Grid.Width;
  end;

  // Inverted?
  axis.Inverted := AWorkbookAxis.Inverted;

  // Logarithmic?
  // to do....

  // Scaling
  axis.Range.UseMin := not AWorkbookAxis.AutomaticMin;
  axis.Range.UseMax := not AWorkbookAxis.AutomaticMax;
  axis.Range.Min := AWorkbookAxis.Min;
  axis.Range.Max := AWorkbookAxis.Max;
end;

procedure TsWorkbookChartLink.UpdateChartAxisLabels(AWorkbookChart: TsChart);
begin
  if (FChart.SeriesCount > 0) and
     (AWorkbookChart.GetChartType in [ctBar, ctLine, ctArea]) then
  begin
    FChart.BottomAxis.Marks.Source := TChartSeries(FChart.Series[0]).Source;
    if not AWorkbookChart.Series[0].LabelRange.IsEmpty then
      FChart.BottomAxis.Marks.Style := smsLabel
    else
      FChart.BottomAxis.Marks.Style := smsXValue;
  end;
end;

procedure TsWorkbookChartLink.UpdateChartBackground(AWorkbookChart: TsChart);
begin
  FChart.Color := Convert_sColor_to_Color(AWorkbookChart.Background.Color);
  FChart.BackColor := Convert_sColor_to_Color(AWorkbookChart.PlotArea.Background.Color);
  UpdateChartPen(AWorkbookChart.PlotArea.Border, FChart.Frame);
  FChart.Frame.Visible := AWorkbookChart.PlotArea.Border.Style <> clsNoLine;
end;

procedure TsWorkbookChartLink.UpdateBarSeries(AWorkbookChart: TsChart);
var
  i, n: Integer;
  ser: TBarSeries;
  barWidth, totalBarWidth: Integer;
begin
  if AWorkbookChart.GetChartType <> ctBar then
    exit;

  // Count the bar series
  n := 0;
  for i := 0 to AWorkbookChart.Series.Count-1 do
  begin
    if AWorkbookChart.Series[i].ChartType = ctBar then
      inc(n);
  end;

  // Iterate over bar series to put them side-by-side or to stack them
  totalBarWidth := 90;
  barWidth := round(totalBarWidth / n);
  for i := 0 to FChart.SeriesCount-1 do
    if FChart.Series[i] is TBarSeries then
    begin
      ser := TBarSeries(FChart.Series[i]);
      case AWorkbookChart.Stackmode of
        csmSideBySide:
          begin
            ser.BarWidthPercent := barWidth;
            ser.BarWidthStyle := bwPercentMin;
            ser.BarOffsetPercent := round((i - (n - 1)/2)*barWidth);
          end;
        csmStacked:
          ser.Stacked := true;
        csmStackedPercentage:
          begin
            ser.Stacked := true;
          end;
      end;
    end;
end;

procedure TsWorkbookChartLink.UpdateChartBrush(AWorkbookFill: TsChartFill;
  ABrush: TBrush);
begin
  if (AWorkbookFill <> nil) and (ABrush <> nil) then
  begin
    ABrush.Color := Convert_sColor_to_Color(AWorkbookFill.Color);
    if AWorkbookFill.Style = cfsNoFill then
      ABrush.Style := bsClear
    else
      ABrush.Style := bsSolid;
    // NOTE: TAChart will ignore gradient.
    // To be completed: hatched filles.
  end;
end;

procedure TsWorkbookChartLink.UpdateChartLegend(AWorkbookLegend: TsChartLegend;
  ALegend: TChartLegend);
const
  LEG_POS: array[TsChartLegendPosition] of TLegendAlignment = (
    laCenterRight,   // lpRight
    laTopCenter,     // lpTop
    laBottomCenter,  // lpBottom
    laCenterLeft     // lpLeft
  );
begin
  if (AWorkbookLegend <> nil) and (ALegend <> nil) then
  begin
    Convert_sFont_to_Font(AWorkbookLegend.Font, ALegend.Font);
    UpdateChartPen(AWorkbookLegend.Border, ALegend.Frame);
    UpdateChartBrush(AWorkbookLegend.Background, ALegend.BackgroundBrush);
    ALegend.Frame.Visible := (ALegend.Frame.Style <> psClear);
    ALegend.Alignment := LEG_POS[AWorkbookLegend.Position];
    ALegend.UseSidebar := not AWorkbookLegend.CanOverlapPlotArea;
    ALegend.Visible := AWorkbookLegend.Visible;
    ALegend.Inverted := true;
  end;
end;

procedure TsWorkbookChartLink.UpdateChartPen(AWorkbookLine: TsChartLine;
  APen: TPen);
begin
  if (AWorkbookLine <> nil) and (APen <> nil) then
  begin
    APen.Color := Convert_sColor_to_Color(AWorkbookLine.Color);
    APen.Width := mmToPx(AWorkbookLine.Width, GetParentForm(FChart).PixelsPerInch);
    case AWorkbookLine.Style of
      clsNoLine:
        APen.Style := psClear;
      clsSolid:
        APen.Style := psSolid;
      else  // to be fixed
        if (AWorkbookLine.Style in [clsDash, clsLongDash]) then
          APen.Style := psDash
        else
        if (AWorkbookLine.Style = clsDot) then
          APen.Style := psDot
        else
        if (AWorkbookLine.Style in [clsDashDot, clsLongDashDot]) then
          APen.Style := psDashDot
        else
        if (AWorkbookLine.Style in [clsLongDashDotDot]) then
          APen.Style := psDashDotDot
        else
          // to be fixed: create pattern as defined.
          APen.Style := psDash;
    end;
  end;
end;

{@@ Updates title and footer of the linked TAChart.
  NOTE: the workbook chart's subtitle is converted to TAChart's footer! }
procedure TsWorkbookChartLink.UpdateChartTitle(AWorkbookTitle: TsChartText;
  AChartTitle: TChartTitle);
begin
  if (AWorkbookTitle <> nil) and (AChartTitle <> nil) then
  begin
    AChartTitle.Text.Clear;
    AChartTitle.Text.Add(AWorkbookTitle.Caption);
    AChartTitle.Visible := AWorkbookTitle.Visible;
    AChartTitle.WordWrap := true;
    Convert_sFont_to_Font(AWorkbookTitle.Font, AChartTitle.Font);
    UpdateChartPen(AWorkbookTitle.Border, AChartTitle.Frame);
    UpdateChartBrush(AWorkbookTitle.Background, AChartTitle.Brush);
    AChartTitle.Font.Orientation := round(AWorkbookTitle.RotationAngle * 10);
    AChartTitle.Frame.Visible := (AChartTitle.Frame.Style <> psClear);
  end;
end;

end.
