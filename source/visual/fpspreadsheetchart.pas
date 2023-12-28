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
  // RTL/FCL
  Classes, Contnrs, SysUtils, Types, FPCanvas,
  // LCL
  LCLVersion, Forms, Controls, Graphics, GraphUtil, Dialogs,
  // TAChart
  TATypes, TATextElements, TAChartUtils, TADrawUtils, TALegend,
  TACustomSource, TASources, TACustomSeries, TAIntervalSources,
  TASeries, TARadialSeries, TAFitUtils, TAFuncSeries, TAMultiSeries,
  TATransformations, TAChartAxisUtils, TAChartAxis, TAStyles, TATools, TAGraph,
  // FPSpreadsheet
  fpsTypes, fpSpreadsheet, fpsUtils, fpsNumFormat, fpsChart, fpsStockSeries,
  // FPSpreadsheet Visual
  fpSpreadsheetCtrls, fpSpreadsheetGrid, fpsVisualUtils;

type

  {@@ Chart data source designed to work together with TChart from Lazarus
    to display the data.

    The data can be loaded from a TsWorksheetGrid Grid component or
    directly from a TsWorksheet FPSpreadsheet Worksheet }

  { TsWorkbookChartSource }

  TsXYLRange = (rngX, rngY, rngLabel, rngColor, rngXErrors, rngYErrors);

  TsWorkbookChartSource = class(TCustomChartSource, IsSpreadsheetControl)
  private
    FWorkbookSource: TsWorkbookSource;
    FWorksheets: array[TsXYLRange] of array of TsWorksheet;
    FRanges: array[TsXYLRange] of array of TsCellRangeArray;
    FRangeStr: array[TsXYLRange] of String;
    FPointsNumber: Cardinal;
    FTitleCol, FTitleRow: Cardinal;
    FTitleSheetName: String;
    FCyclicX: Boolean;
    FIntegerX: Boolean;       // only integers allowed for x values
    FDataPointColors: array of TsColor;
    function GetRange(AIndex: TsXYLRange): String;
    function GetTitle: String;
    function GetWorkbook: TsWorkbook;
    procedure GetXYItem(ARangeIndex:TsXYLRange; AListIndex,APointIndex: Integer;
      out ANumber: Double; out AText: String);
    procedure SetRange(AIndex: TsXYLRange; const AValue: String);
    procedure SetRangeFromChart(ARangeIndex: TsXYLRange; AListIndex: Integer; const ARange: TsChartRange);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
  protected
    FCurItem: TChartDataItem;
    FCurItemIndex: Integer;
    function BuildRangeStr(AIndex: TsXYLRange; AListSeparator: char = #0): String;
    procedure ClearRanges;
    function CountValues(AIndex: TsXYLRange): Integer;
    function GetCount: Integer; override;
    function GetItem(AIndex: Integer): PChartDataItem; override;
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;
    procedure Prepare; overload;
    procedure Prepare(AIndex: TsXYLRange); overload;
    procedure SetXCount(AValue: Cardinal); override;
    procedure SetYCount(AValue: Cardinal); override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function RangeIsEmpty(ARange: TsCellRange): Boolean;
    procedure Reset;
    procedure SetColorRange(ARange: TsChartRange);
    procedure SetLabelRange(ARange: TsChartRange);
    procedure SetXRange(XIndex: Integer; ARange: TsChartRange);
    procedure SetYRange(YIndex: Integer; ARange: TsChartRange);
    procedure SetTitleAddr(Addr: TsChartCellAddr);
    procedure SetXErrorBarRange(APosRange, ANegRange: TsChartRange);
    procedure SetYErrorBarRange(APosRange, ANegRange: TsChartRange);
    procedure UseDataPointColors(ASeries: TsChartSeries);
    property PointsNumber: Cardinal read FPointsNumber;
    property Workbook: TsWorkbook read GetWorkbook;
  public
    // Interface to TsWorkbookSource
    procedure ListenerNotification(AChangedItems: TsNotificationItems; AData: Pointer = nil);
    procedure RemoveWorkbookSource;
  published
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
    property ColorRange: String index rngColor read GetRange write SetRange;
    property CyclicX: Boolean read FCyclicX write FCyclicX default false;
    property IntegerX: Boolean read FIntegerX write FIntegerX default false;
    property LabelRange: String index rngLabel read GetRange write SetRange;
    property XErrorBarRange: String index rngXErrors read GetRange write SetRange;
    property YErrorBarRange: String index rngYErrors read GetRange write SetRange;
    property XRange: String index rngX read GetRange write SetRange;
    property YRange: String index rngY read GetRange write SetRange;
    property Title: String read GetTitle;

    property XErrorBarData;
    property YErrorBarData;
  end;

  {@@ Link between TAChart and the fpspreadsheet chart class }

  { TsWorkbookChartLink }

  TsWorkbookChartLink = class(TComponent, IsSpreadsheetControl)
  private
    FChart: TChart;
    FChartStyles: TChartStyles;
    FWorkbookSource: TsWorkbookSource;
    FWorkbook: TsWorkbook;
    FWorkbookChartIndex: Integer;
    FBrushBitmaps: TFPObjectList;
    FSavedAfterDraw: TChartDrawEvent;
    procedure SetChart(AValue: TChart);
    procedure SetWorkbookChartIndex(AValue: Integer);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);

    //procedure FitSeriesFitEquationText(ASeries: TFitSeries; AEquationText: IFitEquationText);

    procedure AfterDrawChartHandler(ASender: TChart; ADrawer: IChartDrawer);

  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

    function ActiveChartSeries(ASeries: TsChartSeries): TChartSeries;
    procedure AddSeries(ASeries: TsChartSeries);
    procedure FixAreaSeries({%H-}AWorkbookChart: TsChart);
    procedure FixBarSeries(AWorkbookChart: TsChart);
    procedure FixSource(AChartSeries: TBasicPointSeries);
    procedure ClearChart;
    procedure ConstructHatchPattern(AWorkbookChart: TsChart; AFill: TsChartFill; ABrush: TBrush);
    procedure ConstructHatchPatternSolid(AWorkbookChart: TsChart; AFill: TsChartFill; ABrush: TBrush);
    procedure ConstructSeriesMarks(AWorkbookSeries: TsChartSeries; AChartSeries: TChartSeries);
    function GetAutoscaleAxisTransform(AChartAxis: TChartAxis): TAutoScaleAxisTransform;
    function GetAxisTransform(AChartAxis: TChartAxis; AClass: TAxisTransformClass): TAxisTransform;
    function GetLogAxisTransform(AChartAxis: TChartAxis): TLogarithmAxisTransform;
    function GetWorkbookChart: TsChart;
    function IsLogarithmic(Axis: TChartAxis): Boolean;
    function IsSecondaryAxis(Axis: TsChartAxis): boolean;
    function IsStackable(ASeries: TsChartSeries): Boolean;

    procedure UpdateChartAxis(AWorkbookAxis: TsChartAxis);
    procedure UpdateChartAxisLabels(AWorkbookChart: TsChart);
    procedure UpdateChartBackground(AWorkbookChart: TsChart);
//    procedure UpdateBarSeries(AWorkbookChart: TsChart);
    procedure UpdateChartBrush(AWorkbookChart: TsChart; AWorkbookFill: TsChartFill; ABrush: TBrush);
    procedure UpdateChartErrorBars(AWorkbookSeries: TsChartSeries; ASeries: TBasicPointSeries);
    procedure UpdateChartLegend(AWorkbookLegend: TsChartLegend; ALegend: TChartLegend);
    procedure UpdateChartPen(AWorkbookChart: TsChart; AWorkbookLine: TsChartLine; APen: TPen);
    procedure UpdateChartSeriesMarks(AWorkbookSeries: TsChartSeries; AChartSeries: TChartSeries);
    procedure UpdateChartStyle(AWorkbookSeries: TsChartSeries; AStyleIndex: Integer);
    procedure UpdateChartTitle(AWorkbookTitle: TsChartText; AChartTitle: TChartTitle);

    procedure UpdateAreaSeries(AWorkbookSeries: TsAreaSeries; AChartSeries: TAreaSeries);
    procedure UpdateBarSeries(AWorkbookSeries: TsBarSeries; AChartSeries: TBarSeries);
    procedure UpdateBubbleSeries(AWorkbookSeries: TsBubbleSeries; AChartSeries: TBubbleSeries);
    procedure UpdateCustomLineSeries(AWorkbookSeries: TsCustomLineSeries; AChartSeries: TLineSeries);
    procedure UpdatePieSeries(AWorkbookSeries: TsPieSeries; AChartSeries: TPieSeries);
    procedure UpdatePolarSeries(AWorkbookSeries: TsRadarSeries; AChartSeries: TPolarSeries);
    procedure UpdateScatterSeries(AWorkbookSeries: TsScatterSeries; AChartSeries: TLineSeries);
    procedure UpdateStockSeries(AWorkbookSeries: TsStockSeries; AChartSeries: TStockSeries);

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

procedure Convert_sChartLine_to_Pen(AChart: TsChart; ALine: TsChartLine; APen: TPen);


implementation

uses
  Math;

type
  TBasicPointSeriesOpener = class(TBasicPointSeries);

  TsCustomLineSeriesOpener = class(TsCustomLineSeries);

  TErrorbarSeries = class(TBasicPointSeries)
  public
    property XErrorBars;
    property YErrorBars;
  end;

function mmToPx(mm: Double; ppi: Integer): Integer;
begin
  Result := round(mmToIn(mm * ppi));
end;

{ Constructs a PenStyle from the TsChartLine pattern style.
  Note: the conversion is only very rough... }
procedure Convert_sChartLine_to_Pen(AChart: TsChart; ALine: TsChartLine; APen: TPen);
var
  sLineStyle: TsChartLineStyle;

  function IsDot(ASegment: TsChartLineSegment): Boolean;
  var
    len: Integer;
  begin
    if sLineStyle.RelativeToLineWidth then
      Result := (ASegment.Length < 200)
    else
    begin
      len := mmToPx(ASegment.Length, ScreenPixelsPerInch);
      Result := len < 4;
    end;
  end;

var
  dot1, dot2: Boolean;
begin
  sLineStyle := AChart.GetLineStyle(ALine.Style);
  if sLineStyle.Distance = 0 then
    APen.Style := psSolid
  else
  if (sLinestyle.Segment1.Count = 0) and (sLineStyle.Segment2.Count = 0) then
    APen.Style := psClear
  else
  if (sLinestyle.Segment1.Count > 0) and (sLineStyle.Segment2.Count = 0) then
  begin
    if IsDot(sLineStyle.Segment1) then
      APen.Style := psDot
    else
      APen.Style := psDash;
  end else
  if (sLineStyle.Segment1.Count = 0) and (sLineStyle.Segment2.Count > 0) then
  begin
    if IsDot(sLineStyle.Segment2) then
      APen.Style := psDot
    else
      APen.Style := psDash;
  end else
  if (sLineStyle.Segment1.Count = 1) and (sLineStyle.Segment2.Count = 1) then
  begin
    dot1 := IsDot(sLineStyle.Segment1);
    dot2 := IsDot(sLineStyle.Segment2);
    if (dot1 and not dot2) or (not dot1 and dot2) then
      APen.Style := psDashDot
    else
    if dot1 and dot2 then
      APen.Style := psDot
    else
    if (not dot1) and (not dot2) then
      APen.Style := psDash;
  end else
    APen.Style := psDashDotDot
end;

{ Converts an fps format string (e.g. '0.000') to a format string usable in
  the Format() command (e.g. '%.3f') }
function Convert_NumFormatStr_to_FormatStr(ANumFormat: String): String;
var
  isPercent: Boolean = false;
  hasThSep: Boolean = false;
  varDecs: Boolean = false;
  expFmt: Boolean = false;
  p, i: Integer;
  fixedDecs: Integer = 0;
begin
  if ANumFormat = '' then
  begin
    Result := '%.9g';
    exit;
  end;

  i := 1;
  while i <= Length(ANumFormat) do
  begin
    case ANumFormat[i] of
      ',': hasThSep := true;
      '%': isPercent := true;
      '.': begin
             inc(i);
             while (i <= Length(ANumFormat)) do
             begin
               case ANumFormat[i] of
                 '0': inc(fixedDecs);
                 '#': begin
                        varDecs := true;
                        break;
                      end;
                 'e', 'E':
                      begin
                        expFmt := true;
                        break;
                      end;
               end;
               inc(i);
             end;
           end;
    end;
    inc(i);
  end;
  Result := '%.' + IntToStr(fixedDecs);
  if expFmt then
    Result := Result + 'e'
  else
  if varDecs then
    Result := Result + 'g'
  else
  if hasThSep then
    Result := Result + 'n'
  else
    Result := Result + 'f';
  if isPercent then
    Result := Result + '%%';
end;


{------------------------------------------------------------------------------}
{                             TsWorkbookChartSource                            }
{------------------------------------------------------------------------------}

constructor TsWorkbookChartSource.Create(AOwner: TComponent);
begin
  inherited;
  FCurItemIndex := -1;
  ClearRanges;
end;

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
  string unique. In case of x and y which can contain several range groups for
  XIndex/YIndex, all parts for the same XIndex/YIndex are enclosed in parenthesis.

  @@Example
  If there are two y value ranges in sheet1 A1:A10 and B1:B5;B7:B12 then the
  result will be '(Sheet1!A1:A10) (Sheet1!B1:B5;Sheet1!B7:B12)'
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.BuildRangeStr(AIndex: TsXYLRange;
  AListSeparator: Char = #0): String;
var
  L: TStrings;
  range: TsCellRange;
  rangeStr: String;
  totalStr: String;
  i, n: Integer;
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

    n := Length(FRanges[AIndex]);
    if (n = 0) then
      exit('');

    totalStr := '';
    for i := 0 to n-1 do
    begin
      L.Clear;
      for range in FRanges[AIndex, i] do
        L.Add(GetCellRangeString(range, rfAllRel, true));
      rangeStr := FWorksheets[AIndex, i].Name + SHEETSEPARATOR + L.DelimitedText;
      if n = 1 then
        totalStr := rangeStr
      else
      if totalStr = '' then
        totalStr := '(' + rangeStr + ')'
      else
        totalStr := totalStr + ' (' + rangeStr + ')';
    end;
    Result := totalStr;
  finally
    L.Free;
  end;
end;

procedure TsWorkbookChartSource.ClearRanges;
begin
  SetLength(FRanges[rngX], 1);            FRanges[rngX, 0 ] := nil;
  SetLength(FRanges[rngY], 1);            FRanges[rngY, 0] := nil;
  SetLength(FRanges[rngLabel], 1);        FRanges[rngLabel, 0] := nil;
  SetLength(FRanges[rngColor], 1);        FRanges[rngColor, 0] := nil;

  SetLength(FWorksheets[rngX], 1);        FWorksheets[rngX, 0] := nil;
  SetLength(FWorksheets[rngY], 1);        FWorksheets[rngY, 0] := nil;
  SetLength(FWorksheets[rngLabel], 1);    FWorksheets[rngLabel, 0] := nil;
  SetLength(FWorksheets[rngColor], 1);    FWorksheets[rngColor, 0] := nil;

  FRangeStr[rngX] := '';
  FRangeStr[rngY] := '';
  FRangeStr[rngLabel] := '';
  FRangeStr[rngColor] := '';
end;


{@@ ----------------------------------------------------------------------------
  Counts the number of x or y values contained in the x/y ranges

  @param   AIndex   Identifies whether values in the x or y ranges are counted.
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.CountValues(AIndex: TsXYLRange): Integer;
var
  range: TsCellRange;
  i, n: Integer;
begin
  Result := 0;
  for i := 0 to High(FRanges[AIndex]) do
  begin
    n := 0;
    for range in FRanges[AIndex, i] do
    begin
      if range.Col1 = range.Col2 then
        inc(n, range.Row2 - range.Row1 + 1)
      else
      if range.Row1 = range.Row2 then
        inc(n, range.Col2 - range.Col1 + 1)
      else
        raise Exception.Create('x/y ranges can only be 1 column wide or 1 row high.');
    end;
    Result := Max(Result, n);
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
const
  TWO_PI = pi * 2.0;
var
  dummyNumber: Double;
  dummyString: String;
  tmpLabel: String;
  i: Integer;
  value: Double;
begin
  if AIndex = FCurItemIndex then
  begin
    Result := @FCurItem;
    exit;
  end;

  for i := 0 to XCount-1 do
  begin
    if (FRanges[rngX, i] <> nil) then
    begin
      GetXYItem(rngX, i, AIndex, value, tmpLabel);
      if FIntegerX then
        value := trunc(value);
    end else
    if FCyclicX then
      value := AIndex / FPointsNumber * TWO_PI
    else
      value := AIndex;
    FCurItem.SetX(i, value);
  end;

  for i := 0 to YCount-1 do
  begin
    GetXYItem(rngY, i, AIndex, value, dummyString);
    FCurItem.SetY(i, value);
  end;

  if Length(FRanges[rngLabel]) > 0 then
  begin
    GetXYItem(rngLabel, 0, AIndex, dummyNumber, FCurItem.Text);
    if FCurItem.Text = '' then FCurItem.Text := tmpLabel;
  end;

  FCurItem.Color := clTAColor;  // = clDefault
  if AIndex <= High(FDataPointColors) then
    FCurItem.Color := FDataPointColors[AIndex];
  if FRanges[rngColor] <> nil then
  begin
    GetXYItem(rngColor, 0, AIndex, dummyNumber, dummyString);
    if not IsNaN(dummyNumber) then
      FCurItem.Color := round(dummyNumber);
  end;

  FCurItemIndex := AIndex;
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
end;

{@@ ----------------------------------------------------------------------------
  Helper method to prepare the information required for the series data point.

  @param  ARangeIndex  Identifies whether the method retrieves the x or y
                       coordinate, or the label text
  @param  AListIndex   Index of the x or y range group when XCount or YCount is > 1
  @param  APointIndex  Index of the data point for which the data are required
  @param  ANumber      (output) x or y coordinate of the data point
  @param  AText        Data point marks label text
  @param  AColor       Individual data point color
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.GetXYItem(ARangeIndex:TsXYLRange;
  AListIndex, APointIndex: Integer; out ANumber: Double; out AText: String);
var
  range: TsCellRange;
  idx: Integer;
  len: Integer;
  row, col: Cardinal;
  cell: PCell;
begin
  ANumber := NaN;
  AText := '';

  if (FRanges[ARangeIndex, AListIndex] = nil) then
    exit;
  if (FWorksheets[ARangeIndex] = nil) or (FWorksheets[ARangeIndex, AListIndex] = nil) then
    exit;

  cell := nil;
  idx := 0;

  for range in FRanges[ARangeIndex, AListIndex] do
  begin
    if RangeIsEmpty(range) then
      Continue;

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

  cell := FWorksheets[ARangeIndex, AListIndex].FindCell(row, col);

  if cell <> nil then
    case cell^.ContentType of
      cctUTF8String:
        begin
          ANumber := APointIndex;
          AText := FWorksheets[ARangeIndex, AListIndex].ReadAsText(cell);
        end;
      else
        ANumber := FWorksheets[ARangeIndex, AListIndex].ReadAsNumber(cell);
        AText := '';
    end;
end;

{@@ ----------------------------------------------------------------------------
  Notification message received from the WorkbookSource telling which
  spreadsheet item has changed.
  Responds to workbook changes by reading the worksheet names into the tabs,
  and to worksheet changes by selecting the tab corresponding to the selected
  worksheet.

 (@param  AChangedItems  Set of elements identifying whether workbook,
                         worksheet, cell content or cell formatting has changed)
 (@param  AData          Additional data, contains the worksheet for worksheet-related items)

  @see    TsNotificationItem
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.ListenerNotification(
  AChangedItems: TsNotificationItems; AData: Pointer = nil);
var
  ir, i, j: Integer;
  cell: PCell;
  ResetDone: Boolean;
  rng: TsXYLRange;
begin
  Unused(AData);

  // Workbook has been successfully loaded, all sheets are ready
  if (lniWorkbook in AChangedItems) then
  begin
    ClearRanges;
    Prepare;
  end;

  // Used worksheet has been renamed?
  if (lniWorksheetRename in AChangedItems) then
    for rng in TsXYLRange do
      for i := 0 to High(FWorksheets[rng]) do
        if TsWorksheet(AData) = FWorksheets[rng, i] then begin
          FRangeStr[rng] := BuildRangeStr(rng);
          Prepare(rng);
        end;

  // Used worksheet will be deleted?
  if (lniWorksheetRemoving in AChangedItems) then
  begin
    for rng in TsXYLRange do
      for i := 0 to High(FWorksheets[rng]) do
        if TsWorksheet(AData) = FWorksheets[rng, i] then
        begin
          for j := i+1 to High(FWorksheets[rng]) do
            FWorksheets[rng, j-1] := FWorksheets[rng, j];
          SetLength(FWorkSheets[rng], Length(FWorksheets[rng])-1);
          for j := i+1 to High(FRanges[rng]) do
            FRanges[rng, j-1] := FRanges[rng, j];
          SetLength(FRanges[rng], Length(FRanges[rng])-1);
        end;
    for rng in TsXYLRange do
    begin
      FRangeStr[rng] := BuildRangeStr(rng);
      Prepare(rng);
    end;
    Reset;
  end;

  // Cell changes: Enforce recalculation of axes if modified cell is within the
  // x or y range(s).
  if (lniCell in AChangedItems) and (Workbook <> nil) then
  begin
    cell := PCell(AData);
    if (cell <> nil) then begin
      ResetDone := false;
      for rng in TsXYLRange do
        for i := 0 to High(FRanges[rng]) do
          for ir:=0 to High(FRanges[rng, i]) do
          begin
            if FWorksheets[rng, i].CellInRange(cell^.Row, cell^.Col, FRanges[rng, i, ir]) then
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
  s: String;
  sa: TStringArray;
  ok: Boolean;
  i, j: Integer;
begin
  if (FWorkbookSource = nil) then
  begin
    FPointsNumber := 0;
    Reset;
    exit;
  end;

  s := FRangeStr[AIndex];
  if (s = '') then
  begin
    if AIndex = rngY then
    begin
      FPointsNumber := 0;
      Reset;
    end;
    exit;
  end;

  // Split range string into parts for the individual xindex and yindex parts.
  // Each part is enclosed by parenthesis.
  // Example for two y ranges:
  //   '(A1:A10) (B1:B5;B6:B11)' --> 1st y range is A1:A10, 2nd y range is B1:B5 and B6:B11
  if (s <> '') and (s[Length(s)] = ')') then
    Delete(s, Length(s), 1);
  for i := 1 to Length(s) do
    case s[i] of
      '(': s[i] := ' ';
      ')': s[i] := #1;
    end;
  sa := SplitStr(s, #1);
  ok := true;
  for i := 0 to High(sa) do
  begin
    sa[i] := Trim(sa[i]);
    if sa[i] = '' then
    begin
      ok := false;
      break;
    end;
  end;

  case AIndex of
    rngX: XCount := Max(1, Length(sa));
    rngY: YCount := Max(1, Length(sa));
    else ;
  end;

  // Extract range parameters and store them in FRanges
  SetLength(FRanges[AIndex], Length(sa));
  SetLength(FWorksheets[AIndex], Length(sa));
  for i := 0 to High(sa) do
  begin
    if Workbook.TryStrToCellRanges(sa[i], FWorksheets[AIndex, i], FRanges[AIndex, i])
    then begin
      for range in FRanges[AIndex, i] do
        if (range.Col1 <> range.Col2) and (range.Row1 <> range.Row2) then
          raise Exception.Create('x/y ranges can only be 1 column wide or 1 row high');
      FPointsNumber := Max(CountValues(rngX), CountValues(rngY));
      // If x and y ranges are of different size empty data points will be plotted.
      Reset;
    end else
    if (Workbook.GetWorksheetCount > 0) then begin
      if FWorksheets[AIndex, i] = nil then
        raise Exception.Create('Worksheet not found in ' + sa[i]);
    end;
  end;
  // Make sure to include worksheet name in RangeString.
  FRangeStr[AIndex] := BuildRangeStr(AIndex);
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the specified cell range is empty
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.RangeIsEmpty(ARange: TsCellRange): Boolean;
begin
  Result :=
    (ARange.Row1 = UNASSIGNED_ROW_COL_INDEX) and (ARange.Col1 = UNASSIGNED_ROW_COL_INDEX) and
    (ARange.Row2 = UNASSIGNED_ROW_COL_INDEX) and (ARange.Col2 = UNASSIGNED_ROW_COL_INDEX);
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
  SetRangeFromChart(rngColor, 0, ARange);
end;

procedure TsWorkbookChartSource.SetLabelRange(ARange: TsChartRange);
begin
  SetRangeFromChart(rngLabel, 0, ARange);
end;

{@@ ----------------------------------------------------------------------------
  Shared method to set the cell ranges for x, y, labels or colors directly from
  the chart ranges.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetRangeFromChart(ARangeIndex: TsXYLRange;
  AListIndex: Integer; const ARange: TsChartRange);
begin
  if ARange.Sheet1 <> ARange.Sheet2 then
    raise Exception.Create('A chart cell range can only be from a single worksheet.');

  // Auto-expand the FRanges amd FWorksheet arrays
  if AListIndex >= Length(FRanges[ARangeIndex]) then
    SetLength(FRanges[ARangeIndex], Length(FRanges[ARangeIndex]) + 1);
  if AListIndex >= Length(FWorksheets[ARangeIndex]) then
    SetLength(FWorksheets[ARangeIndex], Length(FWorksheets[ARangeIndex]) + 1);

  case ARangeIndex of
    rngX: XCount := Max(XCount, Length(FRanges[ARangeIndex]));
    rngY: YCount := Max(YCount, Length(FRanges[ARangeIndex]));
  end;

  SetLength(FRanges[ARangeIndex, AListIndex], 1);   // FIXME: Assuming here single-block range !!!
  FRanges[ARangeIndex, AListIndex, 0].Row1 := ARange.Row1;
  FRanges[ARangeIndex, AListIndex, 0].Col1 := ARange.Col1;
  FRanges[ARangeIndex, AListIndex, 0].Row2 := ARange.Row2;
  FRanges[ARangeIndex, AListIndex, 0].Col2 := ARange.Col2;
  FWorksheets[ARangeIndex, AListIndex] := FworkbookSource.Workbook.GetWorksheetByName(ARange.Sheet1);
  case ARangeIndex of
    rngX, rngY:
      FPointsNumber := Max(CountValues(rngX), CountValues(rngY));
  end;
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
                      If, in case of the x or y range, cell range strings are
                      put in parenthesis it is assumed that this indicates a
                      source with multiple x or y values.
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

procedure TsWorkbookChartSource.SetXErrorBarRange(APosRange, ANegRange: TsChartRange);
begin
  // TAChart supports error bars only for single-valued sources.
  SetRangeFromChart(rngXErrors, 0, APosRange);
  SetRangeFromChart(rngXErrors, 1, ANegRange)
end;

procedure TsWorkbookChartSource.SetXRange(XIndex: Integer; ARange: TsChartRange);
begin
  SetRangeFromChart(rngX, XIndex, ARange);
end;

procedure TsWorkbookChartSource.SetYErrorBarRange(APosRange, ANegRange: TsChartRange);
begin
  // TAChart supports error bars only for single-valued sources.
  SetRangeFromChart(rngYErrors, 0, APosRange);
  SetRangeFromChart(rngYErrors, 1, ANegRange);
end;

procedure TsWorkbookChartSource.SetYRange(YIndex: Integer; ARange: TsChartRange);
begin
  SetRangeFromChart(rngY, YIndex, ARange);
end;

{@@ ----------------------------------------------------------------------------
  Extracts the fill color from the DataPointStyle items of the series. All the
  other elements are ignored because TAChart does not support them.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.UseDataPointColors(ASeries: TsChartSeries);
var
  datapointStyle: TsChartDataPointStyle;
  i: Integer;
  c: TsColor;
  g: TsChartGradient;
begin
  if ASeries = nil then
  begin
    SetLength(FDataPointColors, 0);
    exit;
  end;

  SetLength(FDataPointColors, ASeries.DataPointStyles.Count);
  for i := 0 to High(FDataPointColors) do
  begin
    datapointStyle := ASeries.DatapointStyles[i];
    FDataPointColors[i] := clTAColor;
    if (dataPointStyle <> nil) and (dataPointStyle.Background <> nil) then
    begin
      if (datapointStyle.Background.Style in [cfsSolid, cfsSolidHatched]) then
        c := dataPointStyle.Background.Color
      else
      if (dataPointStyle.Background.Style = cfsGradient) then
      begin
        // TAChart does not support gradient fills. Let's use the start color
        // of the gradient for a solid fill.
        g := ASeries.Chart.Gradients[datapointStyle.Background.Gradient];
        c := g.StartColor;
      end else
        Continue;
      FDataPointColors[i] := Convert_sColor_to_Color(c);
    end;
  end;
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
  ListenerNotification([lniWorkbook, lniWorksheet]);
  Prepare;
end;

procedure TsWorkbookChartSource.SetXCount(AValue: Cardinal);
begin
  FXCount := AValue;
  SetLength(FCurItem.XList, XCount-1);
end;

{@@ ----------------------------------------------------------------------------
  Inherited ChartSource method telling the series how many y values are used.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetYCount(AValue: Cardinal);
begin
  {$IF LCL_FullVersion >= 3990000}
  inherited SetYCount(AValue);
  {$ELSE}
  FYCount := AValue;
  {$ENDIF}
  SetLength(FCurItem.YList, YCount-1);
end;


{------------------------------------------------------------------------------}
{                             TsWorkbookChartLink                              }
{------------------------------------------------------------------------------}

constructor TsWorkbookChartLink.Create(AOwner: TComponent);
begin
  inherited;
  FBrushBitmaps := TFPObjectList.Create;
  FChartStyles := TChartStyles.Create(self);
  FWorkbookChartIndex := -1;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the WorkbookChartLink.
  Removes itself from the WorkbookSource's listener list.
-------------------------------------------------------------------------------}
destructor TsWorkbookChartLink.Destroy;
begin
  if FWorkbookSource <> nil then FWorkbookSource.RemoveListener(self);
  FBrushBitmaps.Free;
  FChartStyles.Free;
  inherited;
end;

function TsWorkbookChartLink.ActiveChartSeries(ASeries: TsChartSeries): TChartSeries;
type
  TAxisType = (xAx, yAx);
const
  AXIS_ALIGNMENT: array[boolean, TsChartAxisLink, TAxisType] of TChartAxisAlignment = (
    ( (calBottom, calLeft), // not rotated - primary
      (calTop, calRight) ), // not rotated - secondary
    ( (calLeft, calBottom), // rotated - primary
      (calRight, calTop) )  // rotated - secondary
    );
var
  stackable: Boolean;
  firstSeries: TChartSeries;
  ch: TsChart;
  src: TsWorkbookChartSource;
  calcSrc: TCalculatedChartSource;
  style: TChartStyle;
  axAlign: TChartAxisAlignment;
begin
  if FChart.Series.Count > 0 then
    firstSeries := FChart.Series[0] as TChartSeries
  else
    firstSeries := nil;

  ch := ASeries.Chart;
  stackable := IsStackable(ASeries);

  if stackable and (firstSeries <> nil) then
  begin
    // A stackable series in TAChart must use multiple y values.
    Result := firstSeries;
    // For percent-stacking we need an additional chart source, a TCalculatedChartSource
    // which gets its data from the workbook chart source.
    if (firstSeries.Source is TCalculatedChartSource) then
    begin
      calcSrc := TCalculatedChartSource(firstSeries.Source);
      if (calcSrc.Origin is TsWorkbookChartSource) then
        src := TsWorkbookChartSource(calcSrc.Origin);
    end else
    // ... otherwise we use the workbook chart source directly.
    if (firstSeries.Source is TsWorkbookChartSource) then
    begin
      src := (firstSeries.Source as TsWorkbookChartSource);
      calcSrc := nil;
    end else
      raise Exception.Create('Unexpected chart source type.');

    src.SetYRange(src.YCount, ASeries.YRange);      // <--- This updates also the YCount
    src.FRangeStr[rngY] := src.BuildRangeStr(rngY);

    if Result is TBarSeries then
      TBarSeries(Result).Styles := FChartStyles
    else if Result is TLineSeries then
      TLineSeries(Result).Styles := FChartStyles
    else if Result is TAreaSeries then
      TAreaSeries(Result).Styles := FChartStyles;

    Result.Legend.Multiplicity := lmStyle;
    src.SetTitleAddr(ASeries.TitleAddr);

    // Trigger recalculation of YCount of the calculated chart source.
    if calcSrc <> nil then
    begin
      calcSrc.Origin := nil;
      calcSrc.Origin := src;
    end;
  end
  else
  begin
    // This is either for a non-stackable or the first stackable series.
    src := TsWorkbookChartSource.Create(self);
    src.WorkbookSource := FWorkbookSource;

    case ASeries.ChartType of
      ctBar:
        begin
          Result := TBarSeries.Create(FChart);
          src.IntegerX := true;
        end;
      ctLine, ctScatter:
        Result := TLineSeries.Create(FChart);
      ctArea:
        Result := TAreaSeries.Create(FChart);
      ctRadar, ctFilledRadar:
        Result := TPolarSeries.Create(FChart);
      ctBubble:
        begin
          Result := TBubbleSeries.Create(FChart);
          src.SetYRange(1, TsBubbleSeries(ASeries).BubbleRange);  // The radius is at YIndex 1
        end;
      ctPie:
        Result := TPieSeries.Create(FChart);
      ctStock:
        begin
          Result := TStockSeries.Create(FChart);
          src.YCount := 4;
          src.IntegerX := true;
          src.SetXRange(0, ASeries.Chart.XAxis.CategoryRange);
          src.SetYRange(0, TsStockSeries(ASeries).LowRange);   // 0=Low
          src.SetYRange(1, TsStockSeries(ASeries).OpenRange);  // 1=Open
          src.SetYRange(2, TsStockSeries(ASeries).CloseRange); // 2=Close (= Y)
          src.SetYRange(3, TsStockSeries(ASeries).HighRange);  // 3=High
        end;
      else
        exit(nil);
    end;

    // Get x and y ranges (except for OHLC which already has been handled)
    if not (Result is TStockSeries) then
    begin
      if not ASeries.XRange.IsEmpty then
        src.SetXRange(0, ASeries.XRange);
      if not ASeries.YRange.IsEmpty then
        src.SetYRange(0, ASeries.YRange);
    end;

    if not ASeries.LabelRange.IsEmpty then
      src.SetLabelRange(ASeries.LabelRange);
    if not ASeries.FillColorRange.IsEmpty then
      src.SetColorRange(ASeries.FillColorRange);
    src.SetTitleAddr(ASeries.TitleAddr);

    // Copy individual data point colors to the chart series.
    src.UseDataPointColors(ASeries);

    if stackable then begin
      calcSrc := TCalculatedChartSource.Create(self);
      calcSrc.Origin := src;
      Result.Source := calcSrc;
      src.Reset;
    end else
      Result.Source := src;
    Result.Title := src.Title;
  end;

  // Assign series index to axis for primary and secondary axes support
  axAlign := AXIS_ALIGNMENT[ch.RotatedAxes, ASeries.XAxis, xAx];
  Result.AxisIndexX := FChart.AxisList.GetAxisByAlign(axAlign).Index;
  axAlign := AXIS_ALIGNMENT[ch.RotatedAxes, ASeries.YAxis, yAx];
  Result.AxisIndexY := FChart.AxisList.GetAxisByAlign(axAlign).Index;
  {
  case ASeries.XAxis of
    alPrimary:
      Result.AxisIndexX := FChart.AxisList.GetAxisByAlign(calBottom).Index;
    alSecondary:
      Result.AxisIndexX := FChart.AxisList.GetAxisByAlign(calTop).Index;
  end;
  case ASeries.YAxis of
    alPrimary:
      Result.AxisIndexY := FChart.AxisList.GetAxisByAlign(calLeft).Index;
    alSecondary:
      Result.AxisIndexY := FChart.AxisList.GetAxisByAlign(calRight).Index;
  end;
  }

  if stackable then
  begin
    style := TChartStyle(FChartStyles.Styles.Add);
    style.Text := src.Title;
  end;
end;

procedure TsWorkbookChartLink.AddSeries(ASeries: TsChartSeries);
var
  ser: TChartSeries;
  axis: TsChartAxis;
begin
  ser := ActiveChartSeries(ASeries);
  if ser = nil then
  begin
    FWorkbook.AddErrorMsg('Series could not be loaded.');
    exit;
  end;

  ser.Transparency := round(ASeries.Fill.Transparency);
  axis := ASeries.Chart.YAxis;
  UpdateChartSeriesMarks(ASeries, ser);
  if IsStackable(ASeries) then
  begin
    UpdateChartStyle(ASeries, FChartStyles.Styles.Count-1);
    if ASeries.Chart.StackMode = csmStackedPercentage then
      FChart.LeftAxis.Marks.Format := Convert_NumFormatStr_to_FormatStr(axis.LabelFormatPercent)
    else
      FChart.LeftAxis.Marks.Format := Convert_NumFormatStr_to_FormatStr(axis.LabelFormat);
    FChart.Legend.Inverted := ASeries.Chart.StackMode <> csmSideBySide;
  end;

  FChart.AddSeries(ser);

  case ASeries.ChartType of
    ctArea:
      UpdateAreaSeries(TsAreaSeries(ASeries), TAreaSeries(ser));
    ctBar:
      UpdateBarSeries(TsBarSeries(ASeries), TBarSeries(ser));
    ctBubble:
      UpdateBubbleSeries(TsBubbleSeries(ASeries), TBubbleSeries(ser));
    ctLine:
      UpdateCustomLineSeries(TsLineSeries(ASeries), TLineSeries(ser));
    ctScatter:
      UpdateScatterSeries(TsScatterSeries(ASeries), TLineSeries(ser));
    ctStock:
      UpdateStockSeries(TsStockSeries(ASeries), TStockSeries(ser));
    ctPie, ctRing:
      UpdatePieSeries(TsPieSeries(ASeries), TPieSeries(ser));
    ctRadar, ctFilledRadar:
      UpdatePolarSeries(TsRadarSeries(ASeries), TPolarSeries(ser));
  end;
end;

procedure TsWorkbookChartLink.AfterDrawChartHandler(ASender: TChart;
  ADrawer: IChartDrawer);
begin
  if FSavedAfterDraw <> nil then
    FSavedAfterDraw(ASender, ADrawer);

  { TCanvasDrawer.SetBrushParams does not remove the Brush.Bitmap when then
    Brush.Style does not change. Since Brush.Style will be reset to bsSolid
    in the last statement of TChart.Draw this will be enforced by setting
    Brush.Style to bsClear here. }
  ADrawer.SetBrushParams(bsClear, clTAColor);
end;

procedure TsWorkbookChartLink.ClearChart;
var
  i, j: Integer;
  ser: TChartSeries;
  src, src1: TCustomChartSource;
begin
  // Clear the styles
  FChartStyles.Styles.Clear;

  if FChart = nil then
    exit;

  // Clear chart sources
  for i := 0 to FChart.SeriesCount-1 do
  begin
    if (FChart.Series[i] is TChartSeries) then
    begin
      ser :=  TChartSeries(FChart.Series[i]);
      src := ser.Source;
      if src is TCalculatedChartSource then
      begin
        src1 := TCalculatedChartSource(src).Origin;
        if src1 is TsWorkbookChartSource then
          src1.Free;
        src.Free;
      end else
      if src is TsWorkbookChartSource then
        src.Free;
    end;
  end;

  // Clear the series
  FChart.ClearSeries;

  // Clear the axes
  for i := FChart.AxisList.Count-1 downto 0 do
  begin
    if FChart.AxisList[i].Transformations <> nil then
      FChart.AxisList[i].Transformations.Free;

    if FChart.AxisList[i].Minors <> nil then
      for j := FChart.AxisList[i].Minors.Count-1 downto 0 do
        FChart.AxisList[i].Minors.Delete(j);

    if (FChart.AxisList[i].Marks.Source is TDateTimeIntervalChartSource) then
      FChart.AxisList[i].Marks.Source.Free;
    FChart.AxisList[i].Marks.Style := smsValue;

    case FChart.AxisList[i].Alignment of
      calLeft, calBottom:
        FChart.AxisList[i].Title.Caption := '';
      calTop, calRight:
        FChart.AxisList.Delete(i);
    end;
  end;

  // Clear the title
  FChart.Title.Text.Clear;

  // Clear the footer
  FChart.Foot.Text.Clear;
end;

{ Approximates the empty hatch patterns by the built-in TBrush styles. }
procedure TsWorkbookChartLink.ConstructHatchPattern(AWorkbookChart: TsChart;
  AFill: TsChartFill; ABrush: TBrush);
var
  hatch: TsChartHatch;
begin
  ABrush.Style := bsSolid;   // Fall-back style

  hatch := AWorkbookChart.Hatches[AFill.Hatch];
  case hatch.Style of
    chsSingle:
      if InRange(hatch.LineAngle mod 180, -22.5, 22.5) then  // horizontal "approximation"
        ABrush.Style := bsHorizontal
      else
      if InRange((hatch.LineAngle - 90) mod 180, -22.5, 22.5) then  // vertical
        ABrush.Style := bsVertical
      else
      if Inrange((hatch.LineAngle - 45) mod 180, -22.5, 22.5) then  // diagonal up
        ABrush.Style := bsBDiagonal
      else
      if InRange((hatch.LineAngle + 45) mod 180, -22.5, 22.5) then  // diagonal down
        ABrush.Style := bsFDiagonal;
    chsDouble,
    chsTriple:   // no triple hatches in LCL - fall-back to double hatch
      if InRange(hatch.LineAngle mod 180, -22.5, 22.5) then   // +++
        ABrush.Style := bsCross
      else
      if InRange((hatch.LineAngle - 45) mod 180, -22.5, 22.5) then // xxx
        ABrush.Style := bsDiagCross;
  end;
end;

{ Constructs a bitmap for the LCL brush. It is filled by AFill.Color and displays
  a hatch-pattern of hatch index AFill.Hatch. The bitmap is stored in the
  FBrushBitmaps list and assigned to the ABrush.Bitmap operating in fpImage
  style. }
procedure TsWorkbookChartLink.ConstructHatchPatternSolid(AWorkbookChart: TsChart;
  AFill: TsChartFill; ABrush: TBrush);
var
  hatch: TsChartHatch;
  d, ppi: Integer;
  png: TPortableNetworkGraphic;
  sa, ca: Double;
  bkCol: TColor;
  fgCol: TColor;

  procedure PrepareCanvas(w, h: Integer);
  begin
    png.SetSize(w, h);
    png.Canvas.Brush.Color := bkCol;
    png.Canvas.FillRect(0, 0, w, h);
    png.Canvas.Pen.Color := fgCol;
  end;

begin
  ABrush.Style := bsSolid;   // Fall-back style

  hatch := AWorkbookChart.Hatches[AFill.Hatch];
  ppi := GetParentForm(FChart).PixelsPerInch;
  d := mmToPx(hatch.LineDistance, ppi);              // line distance in px
  bkCol := Convert_sColor_to_Color(AFill.Color);     // background color
  fgCol := Convert_sColor_to_Color(hatch.LineColor); // foreground color

  png := TPortableNetworkGraphic.Create;

  case hatch.Style of
    chsSingle:
      begin
        // horizontal ---
        if hatch.LineAngle = 0 then
        begin
          PrepareCanvas(8, d);
          png.Canvas.Line(0, 0, png.Width, 0);
        end else
        // vertical  |||
        if hatch.LineAngle = 90 then
        begin
          PrepareCanvas(d, 0);
          png.Canvas.Line(0, 0, 0, png.Height);
        end else
        // any angle
        begin
          SinCos(DegToRad(hatch.LineAngle), sa, ca);
          PrepareCanvas(round(abs(d / sa)), round(abs(d / ca)));
          if sa/ca > 0 then  // sa/ca = tan
            png.Canvas.Line(0, png.Height-1, png.Width, -1)
          else
            png.Canvas.Line(0, 0, png.Width, png.Height);
        end;
        //png.SaveToFile('test.png');
      end;
    chsDouble, chsTriple:
      begin  // +++
        if InRange(hatch.LineAngle mod 180, -22.5, 22.5) then
        begin
          PrepareCanvas(d, d);
          png.Canvas.Line(0, d div 2, d, d div 2);
          png.Canvas.Line(d div 2, 0, d div 2, d);
          if hatch.Style = chsTriple then
            png.Canvas.Line(0, 0, d, d);
        end else
        // xxx
        if InRange((hatch.LineAngle-45) mod 180, -22.5, 22.5) then
        begin
          d := round(d * sqrt(2));
          PrepareCanvas(d, d);
          png.Canvas.Line(0, 0, d, d);
          png.Canvas.Line(0, d, d, 0);
          if hatch.Style = chsTriple then
            png.Canvas.Line(0, d div 2, d, d div 2);
        end;
      end;
  end;

  // Store the pattern image in the list...
  FBrushBitmaps.Add(png);
  // ... and assign the pattern to the brush
  ABrush.Style := bsImage;
  ABrush.Bitmap := png;
end;

procedure TsWorkbookChartLink.ConstructSeriesMarks(AWorkbookSeries: TsChartSeries;
  AChartSeries: TChartSeries);
var
  sep: String;
  textFmt: String;
  valueFmt: String;
  percentFmt: String;
begin
  if AWorkbookSeries.DataLabels = [cdlValue] then
    AChartSeries.Marks.Style := smsValue
  else if AWorkbookSeries.DataLabels = [cdlPercentage] then
    AChartSeries.Marks.Style := smsPercent
  else if AWorkbookSeries.DataLabels = [cdlCategory] then
    AChartSeries.Marks.Style := smsLabel
  else
  begin
    sep := AWorkbookSeries.LabelSeparator;
    valueFmt := '%0:.9g';
    percentFmt := '%1:.0f';
    textFmt := '%2:s';
    if (AWorkbookSeries.DataLabels * [cdlCategory, cdlValue, cdlPercentage] = [cdlCategory, cdlValue, cdlPercentage]) then
      AChartSeries.Marks.Format := textFmt + sep + valueFmt + sep + percentFmt
    else
    if AWorkbookSeries.DataLabels * [cdlValue, cdlPercentage] = [cdlValue, cdlPercentage] then
      AChartSeries.Marks.Format := valueFmt + sep + percentFmt
    else if AWorkbookSeries.DataLabels * [cdlCategory, cdlValue] = [cdlCategory, cdlValue] then
      AChartSeries.Marks.Format := textFmt + sep + valueFmt;
  end;
  AChartSeries.Marks.Alignment := taCenter;
end;

{
procedure TsWorkbookChartLink.FitSeriesFitEquationText(ASeries: TFitSeries;
  AEquationText: IFitEquationText);
begin
  if ASeries.ErrCode = fitOK then
  begin
    AEquationText.NumFormat('%.5f');
    AEquationText.TextFormat(tfHtml)
  end;
end;
 }

{@@ ----------------------------------------------------------------------------
  Adjusts the area series zero level which, otherwise, is not clipped at the
  chart's plotrect (in TAChart before v3.99)
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.FixAreaSeries(AWorkbookChart: TsChart);
{$IF LCL_FullVersion < 3990000}
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
{$ELSE}
begin
  //
end;
{$ENDIF}

{@@ ----------------------------------------------------------------------------
  Adjusts bar widths and offsets for side-by-side bar charts.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.FixBarSeries(AWorkbookChart: TsChart);
const
  TOTAL_BARWIDTH = 75;
var
  i, n: Integer;
  wBar: Integer;
  offs: Integer;
  ser: TBarSeries;
begin
  if AWorkbookChart.GetChartType <> ctBar then
    exit;

  // Count number of bar series
  n := 0;
  for i := 0 to FChart.SeriesCount - 1 do
    if FChart.Series[i] is TBarSeries then inc(n);

  // Calc bar width and adjust offset of each series within group
  wBar := TOTAL_BARWIDTH div n;
  offs := (wBar - TOTAL_BARWIDTH) div 2;
  for i := 0 to FChart.SeriesCount - 1 do
    if FChart.Series[i] is TBarSeries then
    begin
      ser := TBarSeries(FChart.Series[i]);
      ser.BarWidthPercent := wBar;
      ser.BarOffsetPercent := offs + wBar * i;
    end;
end;

procedure TsWorkbookChartLink.FixSource(AChartSeries: TBasicPointSeries);
var
  i, j, nx, ny: Integer;
  src: TsWorkbookChartSource;
  calcsrc: TCalculatedChartSource;
begin
  if AChartSeries.Source is TCalculatedChartSource then
  begin
    calcSrc := TCalculatedChartSource(AChartSeries.Source);
    if calcSrc.Origin is TsWorkbookChartSource then
      src := TsWorkbookChartSource(calcSrc.Origin)
    else
      exit;
  end else
  if AChartSeries.Source is TsWorkbookChartSource then
    src := TsWorkbookChartSource(AChartSeries.Source)
  else
    exit;

  { TAChart does supports error bars only for single-values chart sources. Only
    when FixSource is called it is known how many values are used by the source.
    If ranges have been added as rngXErrors or rgnYErrors display of error bars
    must be turned off in this case. Otherwise the main series would not be
    shown correctly. }
  if (src.XCount > 0) and (Length(src.FRanges[rngXErrors]) > 0) then
    src.XErrorBarData.Kind := ebkNone;

  if (src.YCount > 1) and (Length(src.FRanges[rngYErrors]) > 0) then
    src.YErrorbarData.Kind := ebkNone;

  { If we get error bar values from the worksheet the range information must
    be added to the rngX/rngY ranges because the Worksheet chartsource expects
    it to be there at the index specified in [X|Y]ErrorBarData.Index[Plus|Minus]. }
  if src.XErrorBarData.Kind = ebkChartSource then
  begin
    nx := src.XCount;
    SetLength(src.FRanges[rngX], nx + Length(src.FRanges[rngXErrors]));
    SetLength(src.FWorksheets[rngX], Length(src.FRanges[rngX]));
    for i := 0 to High(src.FRanges[rngXErrors]) do
    begin
      src.FRanges[rngX, nx + i] := src.FRanges[rngXErrors, i];
      src.FWorksheets[rngX, nx + i] := src.FWorksheets[rngXErrors, i];
      if odd(i) then
        src.XErrorBarData.IndexMinus := nx + i
      else
        src.XErrorBarData.IndexPlus := nx + i;
    end;
    src.XCount := Length(src.FRanges[rngX]);
  end;

  if src.YErrorBarData.Kind = ebkChartSource then
  begin
    ny := src.YCount;
    SetLength(src.FRanges[rngY], ny + Length(src.FRanges[rngYErrors]));
    SetLength(src.FWorksheets[rngY], Length(src.FRanges[rngY]));
    for i := 0 to High(src.FRanges[rngYErrors]) do
    begin
      src.FRanges[rngY, ny + i] := src.FRanges[rngYErrors, i];
      src.FWorksheets[rngY, ny + i] := src.FWorksheets[rngYErrors, i];
      if odd(i) then
        src.YErrorBarData.IndexMinus := ny + i
      else
        src.YErrorBarData.IndexPlus := ny + i;
    end;
    src.YCount := Length(src.Franges[rngY]);
  end;
end;

function TsWorkbookChartLink.GetAutoScaleAxisTransform(AChartAxis: TChartAxis): TAutoScaleAxisTransform;
begin
  Result := TAutoScaleAxisTransform(GetAxisTransform(AChartAxis, TAutoScaleAxisTransform));
end;

function TsWorkbookChartLink.GetAxisTransform(AChartAxis: TChartAxis;
  AClass: TAxisTransformClass): TAxisTransform;
var
  T: TAxisTransform;
begin
  if AChartAxis.Transformations <> nil then
    for T in AChartAxis.Transformations.List do
      if T is AClass then
      begin
        Result := T;
        exit;
      end;
  Result := nil;
end;

function TsWorkbookChartLink.GetLogAxisTransform(AChartAxis: TChartAxis): TLogarithmAxisTransform;
begin
  Result := TLogarithmAxisTransform(GetAxisTransform(AChartAxis, TLogarithmAxisTransform))
end;

function TsWorkbookChartLink.GetWorkbookChart: TsChart;
begin
  if (FWorkbook <> nil) and (FWorkbookChartIndex > -1) then
    Result := FWorkbook.GetChartByIndex(FWorkbookChartIndex)
  else
    Result := nil;
end;

function TsWorkbookChartLink.IsLogarithmic(Axis: TChartAxis): Boolean;
var
  T: TLogarithmAxisTransform;
begin
  T := GetLogAxisTransform(Axis);
  Result := (T <> nil) and T.Enabled;
end;

function TsWorkbookChartLink.IsSecondaryAxis(Axis: TsChartAxis): Boolean;
begin
  Result := (Axis = Axis.Chart.Y2Axis) or (Axis = Axis.Chart.X2Axis);
end;

{@@ ----------------------------------------------------------------------------
  Bar, line and area series can be stacked if they are assigned to the same axis.
-------------------------------------------------------------------------------}
function TsWorkbookChartLink.IsStackable(ASeries: TsChartSeries): Boolean;
var
  ch: TsChart;
  i, numSeries: Integer;
begin
  Result := (ASeries.ChartType in [ctBar, ctLine, ctArea]);
  if Result then
  begin
    ch := ASeries.Chart;
    numSeries := ch.Series.Count;
    if numSeries = 1 then
      exit;

    // Check whether all series are the same type and same y axis as ASeries.
    // NOTE: Not perfect yet since there might abe two stackable groups,
    //       one for the left and one for the right axis...
    for i := 0 to numSeries - 1 do
      if (ch.Series[i].ChartType <> ASeries.ChartType) or
         (ch.Series[i].YAxis <> ASeries.YAxis) then
      begin
        Result := false;
        exit;
      end;
  end;
end;

procedure TsWorkbookChartLink.ListenerNotification(AChangedItems: TsNotificationItems;
  AData: Pointer = nil);
begin
  Unused(AData);

  // Workbook has been successfully loaded, all sheets are ready
  if (lniWorkbook in AChangedItems) then
    ClearChart;
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
  if FChart <> nil then
  begin
    FSavedAfterDraw := FChart.OnAfterDraw;
    FChart.OnAfterDraw := @AfterDrawChartHandler;
  end else
    FSavedAfterDraw := nil;
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

procedure TsWorkbookChartLink.UpdateAreaSeries(AWorkbookSeries: TsAreaSeries;
  AChartSeries: TAreaSeries);
begin
  UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.Fill, AChartSeries.AreaBrush);
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.AreaContourPen);
  AChartSeries.AreaLinesPen.Style := psClear;
  AChartSeries.Stacked := AWorkbookSeries.Chart.StackMode <> csmSideBySide;
  AChartSeries.UseZeroLevel := true;
  if AChartSeries.Source is TCalculatedChartSource then
    TCalculatedChartSource(AChartSeries.Source).Percentage := (AWorkbookSeries.Chart.StackMode = csmStackedPercentage);
  UpdateChartErrorBars(AWorkbookSeries, AChartSeries);
end;

procedure TsWorkbookChartLink.UpdateBarSeries(AWorkbookSeries: TsBarSeries;
  AChartSeries: TBarSeries);
begin
  UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.Fill, AChartSeries.BarBrush);
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.BarPen);
  AChartSeries.BarWidthStyle := bwPercentMin;
  AChartSeries.Stacked := AWorkbookSeries.Chart.StackMode <> csmSideBySide;
  if AChartSeries.Source is TCalculatedChartSource then
    TCalculatedChartSource(AChartSeries.Source).Percentage := (AWorkbookSeries.Chart.StackMode = csmStackedPercentage);
  UpdateChartErrorBars(AWorkbookSeries, AChartSeries);
end;

procedure TsWorkbookChartlink.UpdateBubbleSeries(AWorkbookSeries: TsBubbleSeries;
  AChartSeries: TBubbleSeries);
begin
  UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.Fill, AChartSeries.BubbleBrush);
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.BubblePen);
  {$IF LCL_FullVersion >= 3990000}
  AChartSeries.BubbleRadiusUnits := bruPercentage;
  AChartSeries.ParentChart.ExpandPercentage := 10;
  {$IFEND}
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
  FChart.Proportional := false;
  FChart.ExpandPercentage := 0;

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

  for i := 0 to FChart.SeriesCount-1 do
    if FChart.Series[i] is TBasicPointSeries then
      FixSource(TBasicPointSeries(FChart.Series[i]));

  FChart.Prepare;
  UpdateChartAxisLabels(ch);
  FixAreaSeries(ch);
  FixBarSeries(ch);
end;

procedure TsWorkbookChartLink.UpdateChartAxis(AWorkbookAxis: TsChartAxis);
const
  ROTATED_ALIGNMENT: array[TChartAxisAlignment] of TChartAxisAlignment = (
  // calLeft,   calTop,  calRight, calBottom
     calBottom, calRight, calTop,  calLeft);
var
  align: TChartAxisAlignment;
  axis: TChartAxis;
  minorAxis: TChartMinorAxis;
  T: TAxisTransform;
  logTransf: TLogarithmAxisTransform;
begin
  if AWorkbookAxis = nil then
    exit;

  if AWorkbookAxis = AWorkbookAxis.Chart.XAxis then
    align := calBottom
  else if AWorkbookAxis = AWorkbookAxis.Chart.YAxis then
    align := calLeft
  else if AWorkbookAxis = AWorkbookAxis.Chart.Y2Axis then
    align := calRight
  else if AWorkbookAxis = AWorkbookAxis.Chart.X2Axis then
    align := calTop
  else
    raise Exception.Create('Unsupported axis alignment');

  if AWorkBookAxis.Chart.RotatedAxes then
    align := ROTATED_ALIGNMENT[align];

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

  // Usually not needed, but axis handling is simplified when there is
  // an axis transformations object at each axis with all transforms prepared.
  if (axis.Transformations = nil) then
  begin
    axis.Transformations := TChartAxisTransformations.Create(FChart);

    // Logarithmic
    T := TLogarithmAxisTransform.Create(axis.Transformations);
    T.Transformations := axis.Transformations;
    TLogarithmAxisTransform(T).Base := 10;
    TLogarithmAxisTransform(T).Enabled := AWorkbookAxis.Logarithmic;

    // Autoscale transformation for primary and secondary axes
    T := TAutoScaleAxisTransform.Create(axis.Transformations);
    T.Transformations := axis.Transformations;
    T.Enabled := AWorkbookAxis.Visible and AWorkbookAxis.GetOtherAxis.Visible;
  end;

  // Axis title
  axis.Title.Caption := AWorkbookAxis.Title.Caption;
  axis.Title.Visible := true;
  Convert_sFont_to_Font(AWorkbookAxis.Title.Font, axis.Title.LabelFont);
  axis.Title.LabelFont.Orientation := round(AWorkbookAxis.Title.RotationAngle * 10);

  // Labels
  Convert_sFont_to_Font(AWorkbookAxis.LabelFont, axis.Marks.LabelFont);
  axis.Marks.LabelFont.Orientation := round(AWorkbookAxis.LabelRotation * 10);
  if (AWorkbookAxis.LabelFormat <> '') and not IsDateTimeFormat(AWorkbookAxis.LabelFormat) then
    axis.Marks.Format := Convert_NumFormatStr_to_FormatStr(AWorkbookAxis.LabelFormat);

  // Axis line
  UpdateChartPen(AWorkbookAxis.Chart, AWorkbookAxis.AxisLine, axis.AxisPen);
  axis.AxisPen.Visible := axis.AxisPen.Style <> psClear;

  // Major axis grid
  UpdateChartPen(AWorkbookAxis.Chart, AWorkbookAxis.MajorGridLines, axis.Grid);
  axis.Grid.Visible := (axis.Grid.Style <> psClear) and not IsSecondaryAxis(AWorkbookAxis);
  axis.TickLength := IfThen(catOutside in AWorkbookAxis.MajorTicks, 4, 0);
  axis.TickInnerLength := IfThen(catInside in AWorkbookAxis.MajorTicks, 4, 0);
  axis.TickColor := axis.AxisPen.Color;
  {$IF LCL_FullVersion >= 3000000}
  axis.TickWidth := axis.AxisPen.Width;
  {$IFEND}

  // Minor axis grid
  if AWorkbookAxis.MinorGridLines.Style <> clsNoLine then
  begin
    minorAxis := axis.Minors.Add;
    UpdateChartPen(AWorkbookAxis.Chart, AWorkbookAxis.MinorGridLines, minorAxis.Grid);
    minorAxis.Grid.Visible := not IsSecondaryAxis(AWorkbookAxis);
    if AWorkbookAxis.Logarithmic then
      minorAxis.Intervals.Count := 9
    else
      minorAxis.Intervals.Count := AWorkbookAxis.MinorCount;
    minorAxis.TickLength := IfThen(catOutside in AWorkbookAxis.MinorTicks, 2, 0);
    minorAxis.TickInnerLength := IfThen(catInside in AWorkbookAxis.MinorTicks, 2, 0);
    minorAxis.TickColor := axis.AxisPen.Color;
    {$IF LCL_FullVersion >= 3000000}
    minorAxis.TickWidth := minorAxis.Grid.Width;
    {$IFEND}
  end;

  // Inverted?
  axis.Inverted := AWorkbookAxis.Inverted;

  // Scaling
  axis.Range.UseMin := not AWorkbookAxis.AutomaticMin;
  axis.Range.UseMax := not AWorkbookAxis.AutomaticMax;
  axis.Range.Min := AWorkbookAxis.Min;
  axis.Range.Max := AWorkbookAxis.Max;

  // Logarithmic
  logTransf := GetLogAxisTransform(axis);
  if logTransf <> nil then
    logTransf.Enabled := AWorkbookAxis.Logarithmic;
  if AWorkbookAxis.Logarithmic then
  begin
    axis.Intervals.Options := axis.Intervals.Options + [aipGraphCoords{$IF LCL_FullVersion >= 3990000}, aipInteger{$IFEND}];
    axis.Intervals.MaxLength := 150;
    axis.Intervals.MinLength := 30;
    axis.Intervals.Tolerance := 30;
  end else
  begin
    axis.Intervals.Options := axis.Intervals.Options - [aipGraphCoords{$IF LCL_FullVersion >= 3990000}, aipInteger{$IFEND}];
    axis.Intervals.MaxLength := 100;
    axis.Intervals.MinLength := 20;
    axis.Intervals.Tolerance := 0;
  end;
end;

procedure TsWorkbookChartLink.UpdateChartAxisLabels(AWorkbookChart: TsChart);
var
  axis: TChartAxis;
begin
  if FChart.SeriesCount = 0 then
    exit;

  if AWorkbookChart.RotatedAxes then
    axis := FChart.LeftAxis
  else
    axis := FChart.BottomAxis;
  axis.Marks.SourceExchangeXY := AWorkbookChart.RotatedAxes;

  case AWorkbookChart.GetChartType of
    ctScatter, ctBubble:
      begin
        axis.Marks.Source := nil;
        axis.Marks.Style := smsValue;
      end;
    ctBar, ctLine, ctArea:
      begin
        axis.Marks.Source := TChartSeries(FChart.Series[0]).Source;
        if not AWorkbookChart.Series[0].LabelRange.IsEmpty then
          axis.Marks.Style := smsLabel
        else
          axis.Marks.Style := smsXValue;
      end;
  end;

  // Date/time?
  if AWorkbookChart.XAxis.DateTime then
  begin
    axis.Marks.Source := TDateTimeIntervalChartsource.Create(FChart);
    axis.Marks.Style := smsLabel;
    with TDateTimeIntervalChartSource(axis.Marks.Source) do
    begin
      Params.MaxLength := 120;
      SuppressPrevUnit := false;
    end;
  end;
end;

procedure TsWorkbookChartLink.UpdateChartBackground(AWorkbookChart: TsChart);
begin
  FChart.Color := Convert_sColor_to_Color(AWorkbookChart.Background.Color);
  FChart.BackColor := Convert_sColor_to_Color(AWorkbookChart.PlotArea.Background.Color);
  UpdateChartPen(AWorkbookChart, AWorkbookChart.PlotArea.Border, FChart.Frame);
  FChart.Frame.Visible := AWorkbookChart.PlotArea.Border.Style <> clsNoLine;
end;

procedure TsWorkbookChartLink.UpdateChartBrush(AWorkbookChart: TsChart;
  AWorkbookFill: TsChartFill; ABrush: TBrush);
var
  img: TsChartImage;
  png: TCustomBitmap;
  w, h, ppi: Integer;
begin
  if (AWorkbookFill <> nil) and (ABrush <> nil) then
  begin
    ABrush.Color := Convert_sColor_to_Color(AWorkbookFill.Color);
    case AWorkbookFill.Style of
      cfsNoFill:
        ABrush.Style := bsClear;
      cfsSolid:
        ABrush.Style := bsSolid;
      cfsGradient:
        ABrush.Style := bsSolid;  // NOTE: TAChart cannot display gradients
      cfsHatched:
        ConstructHatchPattern(AWorkbookChart, AWorkbookFill, ABrush);
      cfsSolidHatched:
        ConstructHatchPatternSolid(AWorkbookChart, AWorkbookFill, ABrush);
      cfsImage:
        begin
          img := AWorkbookChart.Images[AWorkbookFill.Image];
          if img <> nil then
          begin
            ppi := GetParentForm(FChart).PixelsPerInch;
            w := mmToPx(img.Width, ppi);
            h := mmToPx(img.Height, ppi);
            png := TPortableNetworkGraphic.Create;
            png.Assign(img.Image);
            {$IF LCL_FullVersion >= 2020000}
            ScaleImg(png, w, h);
            {$IFEND}
            FBrushBitmaps.Add(png);
            ABrush.Bitmap := png;
          end else
            ABrush.Style := bsSolid;
        end;
    end;
  end;
end;

procedure TsWorkbookChartLink.UpdateChartErrorBars(AWorkbookSeries: TsChartSeries;
  ASeries: TBasicPointSeries);
const
  ERRORBAR_KINDS: array[TsChartErrorBarKind] of TChartErrorBarKind = (
    ebkConst, ebkPercent, ebkChartSource);

  function GetChartSource(ASource: TCustomChartSource): TsWorkbookChartSource;
  begin
    if ASource is TsWorkbookChartSource then
      Result := TsWorkbookChartSource(ASource)
    else if (ASource is TCalculatedChartSource) then
      Result := GetChartSource(TCalculatedChartSource(ASource).Origin)
    else
      Result := nil;
  end;

var
  series: TErrorBarSeries;
  source: TsWorkbookChartSource;
begin
  series := TErrorBarSeries(ASeries);
  source := GetChartSource(ASeries.Source);
  if source = nil then
    exit;

  // TAChart supports error bars only for single-values sources!
  if source.XCount = 1 then
  begin
    series.XErrorBars.Visible := AWorkbookSeries.XErrorBars.ShowPos or AWorkbookSeries.XErrorBars.ShowNeg;;
    UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.XErrorBars.Line, series.XErrorBars.Pen);
    source.XErrorBarData.Kind := ERRORBAR_KINDS[AWorkbookSeries.XErrorBars.Kind];
    source.XErrorBarData.ValuePlus := AWorkbookSeries.XErrorBars.ValuePos;
    source.XErrorBarData.ValueMinus := AWorkbookSeries.XErrorBars.ValueNeg;
    if (AWorkbookSeries.XErrorBars.Kind = cebkCellRange) then
      source.SetXErrorBarRange(AWorkbookSeries.XErrorBars.RangePos, AWorkbookSeries.XErrorBars.RangeNeg);
  end;

  if source.YCount = 1 then
  begin
    series.YErrorBars.Visible := AWorkbookSeries.YErrorBars.ShowPos or AWorkbookSeries.YErrorBars.ShowNeg;;
    UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.YErrorBars.Line, series.YErrorBars.Pen);
    source.YErrorBarData.Kind := ERRORBAR_KINDS[AWorkbookSeries.YErrorBars.Kind];
    source.YErrorBarData.ValuePlus := AWorkbookSeries.YErrorBars.ValuePos;
    source.YErrorBarData.ValueMinus := AWorkbookSeries.YErrorBars.ValueNeg;
    if (AWorkbookSeries.YErrorBars.Kind = cebkCellRange) then
      source.SetYErrorBarRange(AWorkbookSeries.YErrorBars.RangePos, AWorkbookSeries.YErrorBars.RangeNeg);
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
    UpdateChartPen(AWorkbookLegend.Chart, AWorkbookLegend.Border, ALegend.Frame);
    UpdateChartBrush(AWorkbookLegend.Chart, AWorkbookLegend.Background, ALegend.BackgroundBrush);
    ALegend.Frame.Visible := (ALegend.Frame.Style <> psClear);
    ALegend.Alignment := LEG_POS[AWorkbookLegend.Position];
    ALegend.UseSidebar := not AWorkbookLegend.CanOverlapPlotArea;
    ALegend.Visible := AWorkbookLegend.Visible;
    ALegend.TextFormat := tfHTML;
  end;
end;

procedure TsWorkbookChartLink.UpdateChartPen(AWorkbookChart: TsChart;
  AWorkbookLine: TsChartLine; APen: TPen);
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
          Convert_sChartLine_to_Pen(AWorkbookChart, AWorkbookLine, APen);
          // To do: not very precise - need to create custom patterns!
    end;
  end;
end;

procedure TsWorkbookChartLink.UpdateChartSeriesMarks(AWorkbookSeries: TsChartSeries;
  AChartSeries: TChartSeries);
begin
  ConstructSeriesMarks(AWorkbookSeries, AChartSeries);
  AChartSeries.Marks.LinkPen.Visible := false;

  AChartSeries.Marks.YIndex := -1;
  AChartSeries.Marks.Distance := 20;
  AChartSeries.Marks.Attachment := maDefault;
  Convert_sFont_to_Font(AWorkbookSeries.LabelFont, AChartSeries.Marks.LabelFont);

  if (AChartSeries is TBubbleSeries) then
    case AWorkbookSeries.LabelPosition of
      lpDefault, lpOutside:
        begin
          TBubbleSeries(AChartSeries).MarkPositions := lmpPositive;
          TBubbleSeries(AChartSeries).Marks.YIndex := 1;
          TBubbleSeries(AChartSeries).Marks.Distance := 5;
        end;
      lpInside:
        begin
          TBubbleSeries(AChartSeries).MarkPositions := lmpInside;
          TBubbleSeries(AChartSeries).Marks.YIndex := 1;
          TBubbleSeries(AChartSeries).Marks.Distance := 5;
        end;
      lpCenter:
        begin
          TBubbleSeries(AChartSeries).MarkPositions := lmpInside;
          TBubbleSeries(AChartSeries).Marks.YIndex := 0;  // 0 --> at data point
          TBubbleSeries(AChartSeries).Marks.Distance := 0;
          TBubbleSeries(AChartSeries).Marks.Attachment := maCenter;
        end;
    end
  else
  if (AChartSeries is TPieSeries) then
    case AWorkbookSeries.LabelPosition of
      lpInside:
        TPieSeries(AChartSeries).MarkPositions := pmpInside;
      {$IF LCL_FullVersion >= 2020000}
      lpCenter:
        TPieSeries(AChartSeries).MarkPositionCentered := true;
      {$IFEND}
      else
        TPieSeries(AChartSeries).MarkPositions := pmpAround;
    end
  else
  if (AChartSeries is TBasicPointSeries) then
    case AWorkbookSeries.LabelPosition of
      lpDefault:
        TBasicPointSeriesOpener(AChartSeries).MarkPositions := lmpOutside;
      lpOutside:
        TBasicPointSeriesOpener(AChartSeries).MarkPositions := lmpOutside;
      lpInside:
        TBasicPointSeriesOpener(AChartSeries).MarkPositions := lmpInside;
      lpCenter:
        begin
          TBasicPointSeriesOpener(AChartSeries).MarkPositions := lmpInside;
          TBasicPointSeriesOpener(AChartSeries).MarkPositionCentered := true;
        end;
    end;

  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.LabelBorder, AChartSeries.Marks.Frame);
  UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.LabelBackground, AChartSeries.Marks.LabelBrush);
end;

procedure TsWorkbookChartLink.UpdateChartStyle(AWorkbookSeries: TsChartSeries;
  AStyleIndex: Integer);
var
  style: TChartStyle;
begin
  style := TChartStyle(FChartStyles.Styles[AStyleIndex]);
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, style.Pen);
  UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.Fill, style.Brush);
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
    {$IF LCL_FullVersion >= 2020000}
    AChartTitle.WordWrap := true;
    {$IFEND}
    Convert_sFont_to_Font(AWorkbookTitle.Font, AChartTitle.Font);
    UpdateChartPen(AWorkbookTitle.Chart, AWorkbookTitle.Border, AChartTitle.Frame);
    UpdateChartBrush(AWorkbookTitle.Chart, AWorkbookTitle.Background, AChartTitle.Brush);
    AChartTitle.Font.Orientation := round(AWorkbookTitle.RotationAngle * 10);
    AChartTitle.Frame.Visible := (AChartTitle.Frame.Style <> psClear);
  end;
end;

procedure TsWorkbookChartLink.UpdateCustomLineSeries(AWorkbookSeries: TsCustomLineSeries;
  AChartSeries: TLineSeries);
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
  ppi: Integer;
  openedWorkbookSeries: TsCustomLineSeriesOpener absolute AWorkbookSeries;
begin
  ppi := GetParentForm(FChart).PixelsPerInch;

  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.LinePen);
  AChartSeries.ShowLines := AWorkbookSeries.Line.Style <> clsNoLine;
  AChartSeries.ShowPoints := openedWorkbookSeries.ShowSymbols;
  if AChartSeries.ShowPoints then
  begin
    UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.Fill, AChartSeries.Pointer.Brush);
    AChartSeries.Pointer.Pen.Color := AChartSeries.LinePen.Color;
    AChartSeries.Pointer.Style := POINTER_STYLES[openedWorkbookSeries.Symbol];
    AChartSeries.Pointer.HorizSize := mmToPx(openedWorkbookSeries.SymbolWidth, ppi);
    AChartSeries.Pointer.VertSize := mmToPx(openedWorkbookSeries.SymbolHeight, ppi);
  end;
  AChartSeries.Stacked := AWorkbookSeries.Chart.StackMode <> csmSideBySide;
  if AChartSeries.Source is TCalculatedChartSource then
    TCalculatedChartSource(AChartSeries.Source).Percentage := (AWorkbookSeries.Chart.StackMode = csmStackedPercentage);

  UpdateChartErrorBars(AWorkbookSeries, AChartSeries);
end;

procedure TsWorkbookChartLink.UpdatePieSeries(AWorkbookSeries: TsPieSeries;
  AChartSeries: TPieSeries);
begin
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.EdgePen);
  AChartSeries.Legend.Multiplicity := lmPoint;
  AChartSeries.Legend.Format := '%2:s';
  {$IF LCL_FullVersion >= 2020000}
  AChartSeries.StartAngle := AWorkbookSeries.StartAngle;
  if AWorkbookSeries is TsRingSeries then
    AChartSeries.InnerRadiusPercent := TsRingSeries(AWorkbookSeries).InnerRadiusPercent;
  {$IFEND}

  FChart.BottomAxis.Visible := false;
  FChart.LeftAxis.Visible := false;
  FChart.Legend.Inverted := false;
  FChart.Frame.Visible := false;
end;

procedure TsWorkbookChartLink.UpdatePolarSeries(AWorkbookSeries: TsRadarSeries;
  AChartSeries: TPolarSeries);
begin
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.LinePen);
  UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.Fill, AChartSeries.Brush);
  AChartSeries.Transparency := round(AWorkbookSeries.Fill.Transparency * 255);
  AChartSeries.CloseCircle := true;
  AChartSeries.Filled := (AWorkbookSeries.ChartType = ctFilledRadar);
  (AChartSeries.Source as TsWorkbookChartSource).CyclicX := true;

  FChart.LeftAxis.Minors.Clear;
  FChart.LeftAxis.Grid.Visible := false;
  FChart.BottomAxis.Minors.Clear;
  FChart.BottomAxis.Grid.Visible := false;
  FChart.Proportional := true;
end;

procedure TsWorkbookChartLink.UpdateScatterSeries(AWorkbookSeries: TsScatterSeries;
  AChartSeries: TLineSeries);
var
  ser: TFitSeries;
  s: String;
begin
  UpdateCustomLineSeries(AWorkbookSeries, AChartSeries);

  if AWorkbookSeries.Regression.RegressionType = rtNone then
    exit;

  // Create series and assign chartsource
  ser := TFitSeries.Create(FChart);
  ser.Source := AChartSeries.Source;

  // Fit equation
  case AWorkbookSeries.Regression.RegressionType of
    rtLinear: ser.FitEquation := feLinear;
    // rtLogarithmic: ser.FitEquation := feLogarithmic;   // to do: implement this!
    rtExponential: ser.FitEquation := feExp;
    rtPower: ser.FitEquation := fePower;
    rtPolynomial:
      begin
        ser.FitEquation := fePolynomial;
        ser.ParamCount := AWorkbookSeries.Regression.PolynomialDegree + 1;
      end;
  end;

  // Take care of y intercept
  if AWorkbookSeries.Regression.ForceYIntercept then
  begin
    str(AWorkbookSeries.Regression.YInterceptValue, s);
    ser.FixedParams := s;
  end;

  // style of regression line
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Regression.Line, ser.Pen);
  ser.AxisIndexX := AChartSeries.AxisIndexX;
  ser.AxisIndexY := AChartSeries.AxisIndexY;

  FChart.AddSeries(ser);

  // Legend text
  ser.Title := AWorkbookSeries.Regression.Title;

  {
  // Show fit curve in legend after series.
  ser.Legend.Order := AChartseries.Legend.Order + 1;
  }

  // Regression equation
  if AWorkbookSeries.Regression.DisplayEquation or AWorkbookSeries.Regression.DisplayRSquare then
  begin
    ser.ExecFit;
    s := '';
    if AWorkbookSeries.Regression.DisplayEquation then
      s := s + ser.EquationText.
        X(AWorkbookSeries.Regression.Equation.XName).
        Y(AWorkbookSeries.Regression.Equation.YName).
        NumFormat(Convert_NumFormatStr_to_FormatStr(AWorkbookSeries.Regression.Equation.NumberFormat)).
        DecimalSeparator('.').
        TextFormat(tfHtml).
        Get;
    if AWorkbookSeries.Regression.DisplayRSquare then
      s := s + LineEnding + 'R<sup>2</sup> = ' + FormatFloat('0.00', ser.FitStatistics.R2);
    if s <> '' then
      ser.Title := ser.Title + LineEnding + s;
//    ser.Legend.Format := '%0:s' + LineEnding + '%2:s';
  end;

  // Error bars
  UpdateChartErrorBars(AWorkbookSeries, AChartSeries);
end;

procedure TsWorkbookChartLink.UpdateStockSeries(AWorkbookSeries: TsStockSeries;
  AChartSeries: TStockSeries);
begin
  if AWorkbookSeries.CandleStick then
  begin
    AChartSeries.Mode := mCandleStick;
    UpdateChartBrush(AWorkbookseries.Chart, AWorkbookseries.CandleStickDownFill, AChartSeries.CandleStickDownBrush);
    UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.CandleStickUpFill, AChartSeries.CandleStickUpBrush);
    UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.CandleStickDownBorder, AChartSeries.CandleStickDownPen);
    UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.CandleStickUpBorder, AChartSeries.CandleStickUpPen);
  end else
  begin
    AChartSeries.Mode := mOHLC;
  end;
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.RangeLine, AChartSeries.LinePen);
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.RangeLine, AChartSeries.DownLinePen);
  AChartSeries.TickWidthStyle := twsPercentMin;
end;

end.
