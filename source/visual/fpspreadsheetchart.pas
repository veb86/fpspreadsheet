{ fpspreadsheetchart.pas }

{@@ ----------------------------------------------------------------------------
  Unit to interface the TAChart package from Lazarus to display the data
  with FPSpreadsheet to load data.

AUTHORS: Felipe Monteiro de Carvalho, Werner Pamler

LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
         distribution, for details about the license.
-------------------------------------------------------------------------------}

unit fpSpreadsheetChart;

{$mode objfpc}{$H+}
{$include ..\fps.inc}

interface

{$ifdef FPS_CHARTS}

uses
  lazloggerbase,

  // RTL/FCL
  Classes, Contnrs, SysUtils, Types, FPImage, FPCanvas,
  // LCL
  LCLVersion, Forms, Controls, IntfGraphics, Graphics, GraphUtil, Dialogs,
  // TAChart
  TATypes, TATextElements, TAChartUtils, TADrawUtils, TALegend,
  TACustomSource, TASources, TACustomSeries, TAIntervalSources,
  TASeries, TARadialSeries, TAFitUtils, TAFuncSeries, TAMultiSeries,
  TATransformations, TAChartAxisUtils, TAChartAxis, TAStyles, TATools, TAGraph,
  // FPSpreadsheet
  fpsTypes, fpSpreadsheet, fpsUtils, fpsNumFormat, fpsImages, fpsChart, fpsStockSeries,
  // FPSpreadsheet Visual
  fpSpreadsheetCtrls, fpSpreadsheetGrid, fpsVisualUtils;

type

  {@@ Enumeration to classify the cell ranges needed for various purposes
   @value rngX        Cell range contains the x values
   @value rngY        Cell range contains the y values
   @value rngLabel    Cell range contains axis or data point labels
   @value rngColor    Cell range contains individual data point colors
   @value rngXErrors  Cell range contains the errors of the x values
   @value rngYErrors  cell range contains the errors of the y values }
  TsXYLRange = (rngX, rngY, rngLabel, rngColor, rngXErrors, rngYErrors);

  {@@ Chart data source designed to pass the data values from the worksheet
    to the TAChart series for plotting.

    The data are loaded via a TsWorkbookSource directly from a TsWorksheet. }
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
    FPieSeriesMode: boolean;
    FPieOffsets: Array of Double;
    FStyles: TChartStyles;

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
    procedure CheckPieSeriesMode(ASeries: TsChartSeries);
    procedure Reset;
    procedure SetColorRange(ARange: TsChartRange);
    procedure SetLabelRange(ARange: TsChartRange);
    procedure SetXRange(XIndex: Integer; ARange: TsChartRange);
    procedure SetYRange(YIndex: Integer; ARange: TsChartRange);
    procedure SetTitleAddr(Addr: TsChartCellAddr);
    procedure SetXErrorBarRange(APosRange, ANegRange: TsChartRange);
    procedure SetYErrorBarRange(APosRange, ANegRange: TsChartRange);
    {@@ The read-only property PointsNumber returns the count of data points provided by the chart source. }
    property PointsNumber: Cardinal read FPointsNumber;
    {@@ Identifies the workbook represented by the assigned WorkbookSource (read-only) }
    property Workbook: TsWorkbook read GetWorkbook;
  public
    // Interface to TsWorkbookSource
    procedure ListenerNotification(AChangedItems: TsNotificationItems; AData: Pointer = nil);
    procedure RemoveWorkbookSource;
  public
    // Special methods to be called by TsWorkbookChartLink
    function RangeIsEmpty(ARange: TsCellRange): Boolean;
    procedure UseDatapointColors(ASeries: TsChartSeries);
    property Styles: TChartStyles read FStyles;
  published
    {@@ Identifies the workbook source from which the chart source gets the data to plot. }
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
    {@@ Cell range defining individual data point colors }
    property ColorRange: String index rngColor read GetRange write SetRange;
    {@@ Set to true by the polar series to rescale the x values to a full circle. }
    property CyclicX: Boolean read FCyclicX write FCyclicX default false;
    {@@ When true, only integers are allowed for x labels. }
    property IntegerX: Boolean read FIntegerX write FIntegerX default false;
    {@@ Cell range defining individual data point labels. }
    property LabelRange: String index rngLabel read GetRange write SetRange;
    {@@ Cell range defining the lenghts of the x error bars. }
    property XErrorBarRange: String index rngXErrors read GetRange write SetRange;
    {@@ Cell range defining the lengths of the y error bars. }
    property YErrorBarRange: String index rngYErrors read GetRange write SetRange;
    {@@ Cell range, given in Excel notation, for the x values of the data points }
    property XRange: String index rngX read GetRange write SetRange;
    {@@ Cell range, given in Excel notation, for the y values of the data points }
    property YRange: String index rngY read GetRange write SetRange;
    {@@ Returns the title of the series assigned to the data source (read-only). Uses the parameters given in SetTitleAddr(). }
    property Title: String read GetTitle;

    {@@ Defines parameters for the x error bars. }
    property XErrorBarData;
    {@@ Defines parameters for the y error bars. }
    property YErrorBarData;
  end;


  { TsWorkbookChartLink }

  {@@ Link between TAChart and the fpspreadsheet chart class }
  TsWorkbookChartLink = class(TComponent, IsSpreadsheetControl)
  private
    FChart: TChart;
    FChartStyles: TChartStyles;
    FWorkbookSource: TsWorkbookSource;
    FWorkbookChartIndex: Integer;
    FBrushBitmaps: TFPObjectList;
    FSavedAfterDraw: TChartDrawEvent;
    FLogLabelSource: TListChartSource;
    function GetWorkbook: TsWorkbook;
    procedure SetChart(AValue: TChart);
    procedure SetWorkbookChartIndex(AValue: Integer);
    procedure SetWorkbookSource(AValue: TsWorkbookSource);
    procedure AfterDrawChartHandler(ASender: TChart; ADrawer: IChartDrawer);

  protected
    procedure Notification(AComponent: TComponent; Operation: TOperation); override;

    procedure AddSeries(ASeries: TsChartSeries);
    function CreateChartSeries(ASeries: TsChartSeries): TChartSeries;
    procedure FixAreaSeries({%H-}AWorkbookChart: TsChart);
    procedure FixSource(AChartSeries: TBasicPointSeries);
    procedure ClearChart;
    procedure ConstructFillPattern(AWorkbookChart: TsChart; AFill: TsChartFill; ABrush: TBrush);
    procedure ConstructHatchPattern(AWorkbookChart: TsChart; AFill: TsChartFill; ABrush: TBrush);
    procedure ConstructImagePattern(AWorkbookChart: TsChart; AFill: TsChartFill; ABrush: TBrush);
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
    procedure UpdateChartBrush(AWorkbookChart: TsChart; AWorkbookFill: TsChartFill; ABrush: TBrush);
    procedure UpdateChartErrorBars(AWorkbookSeries: TsChartSeries; ASeries: TBasicPointSeries);
    procedure UpdateChartLegend(AWorkbookLegend: TsChartLegend; ALegend: TChartLegend);
    procedure UpdateChartPen(AWorkbookChart: TsChart; AWorkbookLine: TsChartLine; APen: TPen);
    procedure UpdateChartSeriesMarks(AWorkbookSeries: TsChartSeries; AChartSeries: TChartSeries);
    procedure UpdateChartSeriesTrendline(AWorkbookSeries: TsChartSeries; AChartSeries: TChartSeries);
    procedure UpdateChartStyle(AWorkbookSeries: TsChartSeries; AStyleIndex: Integer);
    procedure UpdateChartTitle(AWorkbookTitle: TsChartText; AChartTitle: TChartTitle);

    procedure UpdateAreaSeries(AWorkbookSeries: TsAreaSeries; AChartSeries: TAreaSeries);
    procedure UpdateBarSeries(AWorkbookSeries: TsBarSeries; AChartSeries: TBarSeries);
    procedure UpdateBubbleSeries(AWorkbookSeries: TsBubbleSeries; AChartSeries: TBubbleSeries);
    procedure UpdateCustomLineSeries(AWorkbookSeries: TsCustomLineSeries; AChartSeries: TBasicPointSeries);
    procedure UpdatePieSeries(AWorkbookSeries: TsPieSeries; AChartSeries: TPieSeries);
    procedure UpdatePolarSeries(AWorkbookSeries: TsRadarSeries; AChartSeries: TPolarSeries);
    procedure UpdateScatterSeries(AWorkbookSeries: TsScatterSeries; AChartSeries: TLineSeries);
    procedure UpdateStockSeries(AWorkbookSeries: TsStockSeries; AChartSeries: TStockSeries);

    procedure CreateChartStylesFromDatapoints(AWorkbookSeries: TsChartSeries; AChartStyles: TChartStyles);

  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;

    procedure UpdateChart;

    { Interfacing with WorkbookSource}
    procedure ListenerNotification(AChangedItems: TsNotificationItems; AData: Pointer = nil);
    procedure RemoveWorkbookSource;

    {@@ Identifies the workbook represented by the assigned WorkbookSource (read-only) }
    property Workbook: TsWorkbook read GetWorkbook;

  published
    {@@ Identifies the TAChart instance which is linked by the WorkbookChartlink to the workbook. }
    property Chart: TChart read FChart write SetChart;
    {@@ Determines the index of the workbook chart to be handled. }
    property WorkbookChartIndex: Integer read FWorkbookChartIndex write SetWorkbookChartIndex;
    {@@ Determines the workbook source which provides the workbook data. }
    property WorkbookSource: TsWorkbookSource read FWorkbookSource write SetWorkbookSource;
  end;

procedure Convert_sChartLine_to_Pen(AChart: TsChart; ALine: TsChartLine; APen: TPen);

{$endif}

implementation

{$ifdef FPS_CHARTS}

uses
  Math, fpsPatterns;

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
    psFullStar,
    psHorBar,
    psPoint
  );

type
  TOpenedBasicPointSeries = class(TBasicPointSeries);

  TsOpenedCustomLineSeries = class(TsCustomLineSeries);

  TsOpenedTrendlineSeries = class(TsChartSeries)
  public
    property Trendline;
  end;

  TOpenedErrorbarSeries = class(TBasicPointSeries)
  public
    property XErrorBars;
    property YErrorBars;
  end;


function mmToPx(mm: Double; ppi: Integer): Integer;
begin
  Result := round(mmToIn(mm * ppi));
end;

{@@ ----------------------------------------------------------------------------
  Constructs a GUI PenStyle from the TsRawLinePattern data.

  @param AChart  Chart in which this pen is used
  @param ALine   TsChartLine instance providing the parameters used in the workbook
  @param APen    GUI pen to which the ALine parameters are converted.

  @Note          The conversion is only very rough...
-------------------------------------------------------------------------------}
procedure Convert_sChartLine_to_Pen(AChart: TsChart; ALine: TsChartLine; APen: TPen);
var
  pattern: TsRawLinePattern;

  function IsDot(AElement: TsRawLinePatternElement): Boolean;
  var
    len: Integer;
  begin
    case pattern.LengthUnit of
      cluMillimeters:
        begin
          len := mmToPx(AElement.Length, ScreenPixelsPerInch);
          Result := len < 4;
        end;
      cluPercentage:
        Result := (AElement.Length <= 200)
    end;
  end;

var
  dot1, dot2: Boolean;
begin
  pattern := GetRawLinePattern(ALine.PatternIndex);
  if pattern.DistanceLength = 0 then
    APen.Style := psSolid
  else
  if (pattern.Element1.Count = 0) and (pattern.Element2.Count = 0) then
    APen.Style := psClear
  else
  if (pattern.Element1.Count > 0) and (pattern.Element2.Count = 0) then
  begin
    if IsDot(pattern.Element1) then
      APen.Style := psDot
    else
      APen.Style := psDash;
  end else
  if (pattern.Element1.Count = 0) and (pattern.Element2.Count > 0) then
  begin
    if IsDot(pattern.Element2) then
      APen.Style := psDot
    else
      APen.Style := psDash;
  end else
  if (pattern.Element1.Count = 1) and (pattern.Element2.Count = 1) then
  begin
    dot1 := IsDot(pattern.Element1);
    dot2 := IsDot(pattern.Element2);
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


{===============================================================================
                              TsWorkbookChartSource
===============================================================================}

{@@ ----------------------------------------------------------------------------
  Constructor of the TsWorkbookChartSource class
-------------------------------------------------------------------------------}
constructor TsWorkbookChartSource.Create(AOwner: TComponent);
begin
  inherited;
  FStyles := TChartStyles.Create(self);
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

  Example:
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

{@@ ----------------------------------------------------------------------------
  Empties all cell ranges used by the chart source.
-------------------------------------------------------------------------------}
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

  FStyles.Styles.Clear;
  SetLength(FDatapointColors, 0);
  SetLength(FPieOffsets, 0);
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

  @param  AIndex   Index of the data point in the series.
  @return(Pointer to a TChartDataItem record containing the x and y coordinates,
           the data point mark text, and the individual data point color.)
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

  { In PieSeriesMode, the x values are not taken from the worksheet, but
    have been made available in the separate array FPieOffsets. }
  if FPieSeriesMode then
  begin
    if (AIndex >= 0) and (AIndex < Length(FPieOffsets)) then
      FCurItem.SetX(0, FPieOffsets[AIndex])
    else
      FCurItem.SetX(0, 0.0);
  end
  else
    for i := 0 to XCount-1 do
    begin
      if (FRanges[rngX, i] <> nil) then
      begin
        GetXYItem(rngX, i, AIndex, value, tmpLabel);
        if FIntegerX then
          value := trunc(value);
      end else
        value := AIndex;
      // For polar series (which sets CyclicX to true) we rescale the x values
      // to a full circle. And the angle begins at the 90° position.
      if FCyclicX then
        value := value / FPointsNumber * TWO_PI + pi/2;
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

  @param   AIndex   Determines whether the methods deals with x, y values or value labels.
  @return(An Excel string containing workbookname and cell block(s) in A1
           notation. Multiple blocks are separated by the ListSeparator defined
           by the workbook's FormatSettings.)
-------------------------------------------------------------------------------}
function TsWorkbookChartSource.GetRange(AIndex: TsXYLRange): String;
begin
  Result := FRangeStr[AIndex];
end;

{@@ ----------------------------------------------------------------------------
  Getter method returning the string to be used as title of the associated series.
  The string is read from the cell identified by parameters of the SetTitleAddr
  method.
-------------------------------------------------------------------------------}
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

  @param  ARangeIndex  Identifies whether the method retrieves the x or y coordinate, or the label text
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
  begin
    case cell^.ContentType of
      cctUTF8String:
        begin
          ANumber := APointIndex;
//          AText := FWorksheets[ARangeIndex, AListIndex].ReadAsText(cell);
        end;
      cctDateTime:
        begin
          if not FWorksheets[ARangeIndex, AListIndex].ReadAsDateTime(cell, ANumber) then
            ANumber := APointIndex;
  //        AText := FWorksheets[ARangeIndex, AListIndex].ReadAsText(cell);
        end;
      else
        ANumber := FWorksheets[ARangeIndex, AListIndex].ReadAsNumber(cell);
//        AText := '';
    end;
    AText := FWorksheets[ARangeIndex, AListIndex].ReadAsText(cell);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Notification message received from the WorkbookSource telling which
  spreadsheet item has changed.

  @unorderedlist(
   @item(Prepares the chart when all sheets are completely loaded.)
   @item(Fixes cell ranges when a worksheet has been renamed or deleted.)
   @item(Enforces axis scale recalculation when a cell value has been changed.)
  )

 @param(AChangedItems  Set of elements identifying whether workbook,
                       worksheet, cell content or cell formatting has changed)
 @param AData          Additional data, contains the worksheet for worksheet-related items

 @seeAlso    TsNotificationItem
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
  WorkbookSource is being removed and sets the internal WorkbookSource variable
  to nil.
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

  @param  ARange  Cell range to be checked.
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
var
  sheet: TsWorksheet;
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
  sheet := FWorkbookSource.Workbook.GetWorksheetByName(ARange.Sheet1);
  if sheet = nil then sheet := FWorkbookSource.Workbook.ActiveWorksheet;
  FWorksheets[ARangeIndex, AListIndex] := sheet;
  case ARangeIndex of
    rngX, rngY:
      FPointsNumber := Max(CountValues(rngX), CountValues(rngY));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the cell range used for x or y data (or labels) in the chart
  If it does not contain the worksheet name the currently active worksheet of
  the WorkbookSource is assumed.

  @param(AIndex Distinguishes whether the method deals with x, y or label ranges.
  @param(AValue String in Excel syntax containing the cell range to be
                used for x or y, depending on AIndex. Can contain multiple
                cell blocks which must be separator by the ListSeparator
                character defined in the Workbook's FormatSettings.
                If, in case of the x or y range, cell range strings are
                put in parenthesis it is assumed that this indicates a
                source with multiple x or y values.)
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetRange(AIndex: TsXYLRange;
  const AValue: String);
begin
  FRangeStr[AIndex] := AValue;
  Prepare;
end;

{@@ ----------------------------------------------------------------------------
  Setter method for the address of the cell providing the string to be used for
  the series title.

  @param  Addr  A TsChartCellAddr record containing row and column index as well as sheet name of the cell
  @seeAlso Title
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetTitleAddr(Addr: TsChartCellAddr);
begin
  FTitleRow := Addr.Row;
  FTitleCol := Addr.Col;
  FTitleSheetName := UnquoteStr(Addr.GetSheetName);
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
  In case of a pie series, both xlsx and ods files do not provide the
  pie offsets as worksheet cell ranges but as attributes in the xml files.
  Therefore, we store these offsets separately in an array, FPieOffsets.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.CheckPieSeriesMode(ASeries: TsChartSeries);
var
  i, j: Integer;
  datapointStyle: TsChartDataPointStyle;
begin
  FPieSeriesMode := (ASeries is TsPieSeries);
  if FPieSeriesMode then
  begin
    SetLength(FPieOffsets, ASeries.Count);
    for i := 0 to ASeries.Count-1 do
    begin
      j := ASeries.DataPointStyles.IndexOfDataPoint(i);
      FPieOffsets[i] := 0;
      dataPointStyle := ASeries.DataPointStyles[j];
      if dataPointStyle <> nil then
        FPieOffsets[i] := dataPointStyle.PieOffset * 0.01;
    end;
  end else
    SetLength(FPieOffsets, 0);
end;

{@@ ----------------------------------------------------------------------------
  Extracts the fill color from the DataPointStyle items of the series. All the
  other elements are ignored because TAChart does not support them.

  But note: Some series types allow to use chartstyles for individual data point
  formatting. In this case this method is not executed.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.UseDataPointColors(ASeries: TsChartSeries);

  function ColorFromDatapointStyle(ADatapointStyle: TsChartDatapointStyle): TColor;
  var
    c: TsChartColor;
    g: TsChartGradient;
    fill: TsChartFill;
    coloredPatt: TsChartFillPattern;
  begin
    Result := clTAColor;
    if (ADatapointStyle <> nil) then
    begin
      if ADatapointStyle.Background <> nil then
        fill := ADataPointstyle.Background
      else
        fill := ASeries.Fill;
      case fill.Style of
        cfsSolidFill:
          c := fill.Color;
        cfsPattern:
          begin
            coloredPatt := ASeries.Chart.FillPatterns[fill.Pattern];
            c := coloredPatt.BgColor;
          end;
        cfsGradient:
          begin
            if (fill.Gradient = -1) or (ASeries.Chart.Gradients.Count = 0) then
              exit;
            g := ASeries.Chart.Gradients[fill.Gradient];
            c := g.StartColor;
          end;
      end;
      Result := Convert_sColor_to_Color(c.Color);
    end;
  end;

var
  datapointStyle: TsChartDataPointStyle;
  i, j: Integer;
begin
  if ASeries = nil then
  begin
    SetLength(FDataPointColors, 0);
    exit;
  end;

  SetLength(FDataPointColors, ASeries.Count);
  i := 0;
  for j := 0 to ASeries.DataPointStyles.Count-1 do
  begin
    datapointStyle := ASeries.DatapointStyles[j];
    while (datapointStyle <> nil) and (i < dataPointStyle.DataPointIndex) do
    begin
      FDataPointColors[i] := clTAColor;
      inc(i);
    end;
    FDataPointColors[i] := ColorFromDatapointStyle(datapointStyle);
    inc(i);
  end;
  while i <= High(FDataPointColors) do
  begin
    FDataPointColors[i] := clTAColor;
    inc(i);
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

{@@ ----------------------------------------------------------------------------
  Inherited ChartSource method telling the series how many x values are used
  per data point.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartSource.SetXCount(AValue: Cardinal);
begin
  FXCount := AValue;
  SetLength(FCurItem.XList, XCount-1);
end;

{@@ ----------------------------------------------------------------------------
  Inherited ChartSource method telling the series how many y values are used
  per data point.
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


{===============================================================================
                            TsWorkbookChartLink
===============================================================================}

{@@ ----------------------------------------------------------------------------
  Constructor of the TsWorkbookChartLink class
-------------------------------------------------------------------------------}
constructor TsWorkbookChartLink.Create(AOwner: TComponent);
begin
  inherited;
  FBrushBitmaps := TFPObjectList.Create;
  FChartStyles := TChartStyles.Create(self);
  FWorkbookChartIndex := -1;
  FLogLabelSource := TListChartSource.Create(Self);
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
  FLogLabelSource.Free;
  if Assigned(FChart) then
  begin
    FChart.OnAfterDraw := FSavedAfterDraw;
    FChart := nil;
  end;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Creates a TAChart series of the type which corresponds to the given
  workbook series, translates all the workbook series properties to it,
  and adds it to the GUI chart.

  @param   ASeries  Workbook series which is processed.
  @seeAlso CreateChartSeries
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.AddSeries(ASeries: TsChartSeries);
var
  ser: TChartSeries;
  axis: TsChartAxis;
begin
  ser := CreateChartSeries(ASeries);
  if (ser = nil) or (Workbook = nil) then
  begin
    Workbook.AddErrorMsg('Series could not be loaded.');
    exit;
  end;

  ser.Transparency := round(ASeries.Fill.Color.Transparency);
  axis := ASeries.Chart.YAxis;
  UpdateChartSeriesMarks(ASeries, ser);
  if IsStackable(ASeries) then
  begin
    UpdateChartStyle(ASeries, FChartStyles.Styles.Count-1);
    if ASeries.Chart.StackMode = csmStackedPercentage then
      FChart.LeftAxis.Marks.Format := Convert_NumFormatStr_to_FormatStr(axis.LabelFormatPercent)
    else
      FChart.LeftAxis.Marks.Format := Convert_NumFormatStr_to_FormatStr(axis.LabelFormat);
    FChart.Legend.Inverted := ASeries.Chart.StackMode <> csmDefault;
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

  //ser.Index := ASeries.Order;
end;

{@@ ----------------------------------------------------------------------------
  Handler for the GUI chart's OnAfterDraw event to reset the drawer's brush
  parameters.

  Explanation: TCanvasDrawer.SetBrushParams does not remove the Brush.Bitmap
    when the Brush.Style does not change. Since Brush.Style will be reset to
    bsSolid in the last statement of TChart.Draw this will be enforced
    here by setting Brush.Style to bsClear.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.AfterDrawChartHandler(ASender: TChart;
  ADrawer: IChartDrawer);
begin
  if FSavedAfterDraw <> nil then
    FSavedAfterDraw(ASender, ADrawer);
  ADrawer.SetBrushParams(bsClear, clTAColor);
end;

{@@ ----------------------------------------------------------------------------
  Removes and destroys all series (and attached chart sources) from the linked
  TAChart instance.
-------------------------------------------------------------------------------}
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

  // Restore default background
  FChart.BackColor := clWindow;
  FChart.Frame.Color := clDefault;
  FChart.Frame.Width := 1;
  FChart.Frame.Style := psSolid;
  FChart.Frame.Visible := true;

//  FChart.OnAfterDraw := FSavedAfterDraw;
end;

{@@ ----------------------------------------------------------------------------
  Approximates the empty hatch patterns by the built-in TBrush styles which are
  rendered without background by TAChart.

  @param  AWorkbookChart   Chart in the workbook which is processed here.
  @param  AFill            Fill parameters provided by the workbook chart.
  @param  ABrush           Graphics brush to be used by the user-interface chart.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.ConstructHatchPattern(AWorkbookChart: TsChart;
  AFill: TsChartFill; ABrush: TBrush);
var
  book: TsWorkbook;
  coloredPattern: TsChartFillPattern;
  rawPattern: TsRawFillPattern;
begin
  ABrush.Style := bsSolid;   // Fall-back style
  if AFill.Pattern = -1 then
    exit;

  book := TsWorkbook(AWorkbookChart.Workbook);
  coloredPattern := AWorkbookChart.FillPatterns[AFill.Pattern];
  rawPattern := GetRawFillPattern(coloredPattern.Index);
  ABrush.Color := Convert_sColor_to_Color(coloredpattern.FgColor.Color);
  if (rawPattern.LinePattern <> nil) and coloredPattern.IsClearPattern then
    case rawPattern.LinePattern.Multiplier of
      lfpmSingle:
        if InRange(FMod(rawPattern.LinePattern.Angle, 180.0), -22.5, 22.5) then  // horizontal "approximation"
          ABrush.Style := bsHorizontal
        else
        if InRange(FMod(rawPattern.LinePattern.Angle - 90, 180.0), -22.5, 22.5) then  // vertical
          ABrush.Style := bsVertical
        else
        if Inrange(FMod(rawPattern.LinePattern.Angle - 45, 180.0), -22.5, 22.5) then  // diagonal up
          ABrush.Style := bsBDiagonal
        else
        if InRange(FMod(rawPattern.linePattern.Angle + 45, 180.0), -22.5, 22.5) then  // diagonal down
          ABrush.Style := bsFDiagonal;
      lfpmDouble,
      lfpmTriple:   // no triple hatches in LCL - fall-back to double hatch
        if InRange(FMod(rawPattern.LinePattern.Angle, 180.0), -22.5, 22.5) then   // +++
          ABrush.Style := bsCross
        else
        if InRange(FMod(rawPattern.LinePattern.Angle - 45, 180.0), -22.5, 22.5) then // xxx
          ABrush.Style := bsDiagCross;
    end
  else
    ConstructFillPattern(AWorkbookChart, AFill, ABrush);
end;

{@@ ----------------------------------------------------------------------------
  Constructs a bitmap for the LCL brush. It is filled by AFill.Color and displays
  a hatch-pattern of hatch index AFill.Hatch. The bitmap is stored in the
  FBrushBitmaps list and assigned to the ABrush.Bitmap operating in fpImage
  style.

  @param  AWorkbookChart   Chart in the workbook which is processed here.
  @param  AFill            Fill parameters provided by the workbook chart.
  @param  ABrush           Graphics brush to be used by the user-interface chart.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.ConstructFillPattern(AWorkbookChart: TsChart;
  AFill: TsChartFill; ABrush: TBrush);
var
  rawPattern: TsRawFillPattern;
  coloredPattern: TsChartFillPattern;
  png: TPortableNetworkGraphic;
  img: TLazIntfImage;
//  chBkCol: TsChartColor;
  bkCol, fgCol: TFPColor;
  x, y: Integer;
  b: byte;
begin
  ABrush.Style := bsSolid;   // Solid fill as fall-back "pattern"

  coloredpattern := AWorkbookChart.FillPatterns[AFill.Pattern];
  rawPattern := GetRawFillPattern(coloredPattern.Index);

  // Pattern color
  fgCol := TColorToFPColor(Convert_sColor_to_Color(coloredPattern.FgColor.Color));

  // Background color
  bkCol := TColorToFPColor(Convert_sColor_to_Color(coloredPattern.BgColor.Color));

  png := TPortableNetworkGraphic.Create;
  png.SetSize(8, 8);
  img := png.CreateIntfImage;
  for y := 0 to img.Height-1 do
    for x := 0 to img.Width-1 do
    begin
      b := 1 shl x;
      if rawPattern.DotPattern[y] and b <> 0 then
        img.Colors[x, y] := fgCol
      else
        img.Colors[x, y] := bkCol;
    end;
  png.LoadFromIntfImage(img);
  img.Free;
  FBrushBitmaps.Add(png);

  // ... and assign the pattern to the brush
  ABrush.Style := bsImage;
  ABrush.Bitmap := png;
end;

{@@ ----------------------------------------------------------------------------
  Constructs a bitmap fill pattern for the GUI brush from the image fill pattern
  defined in the workbook chart.

  @param  AWorkbookChart  Workbook chart to which the pattern belongs
  @param  AFill           Fill parameters in the workbook chart from which the image pattern is to be converted
  @param  ABrush          Brush to which the bitmap will be assigned.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.ConstructImagePattern(AWorkbookChart: TsChart;
  AFill: TsChartFill; ABrush: TBrush);
var
  wBook: TsWorkbook;
  img: TsChartImage;
  obj: TsEmbeddedObj;
  pic: TPicture;
  png: TPortableNetworkGraphic;
  w, h, ppi: Integer;
begin
  wBook := TsWorkbook(AWorkbookChart.Workbook);

  if AFill.Image = -1 then
    exit;
  img := AWorkbookChart.Images[AFill.Image];
  if img.EmbeddedObjIndex = -1 then
    exit;

  obj := wBook.GetEmbeddedObj(img.EmbeddedObjIndex);
  pic := TPicture.Create;
  try
    obj.Stream.Position := 0;
    pic.LoadFromStream(obj.Stream);
    png := TPortableNetworkGraphic.Create;
    if (img.Width > 0) and (img.Height > 0) then
    begin
      png.PixelFormat := pf32Bit;
      ppi := GetParentForm(FChart).PixelsPerInch;
      png.SetSize(mmToPx(img.Width, ppi), mmToPx(img.Height, ppi));
      png.Canvas.StretchDraw(Rect(0, 0, png.Width, png.Height), pic.PNG);
    end else
      png.Assign(pic.PNG);
    FBrushBitmaps.Add(png);
    ABrush.Bitmap := png;
  finally
    pic.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Constructs the format strings for the series marks allowing: multiple items
  separated by the WorkbookSeries.LabelSeparator, formatting of numbers as
  specified by aWorkbookSeries.LabelFormat and AWorkbookseries.LabelFormatPercent.

  @note(There are some issues with TAChart:

  A 100%-stacked series gets its values from a CalculatedChartSource which
  already delivers the percentages, but does not give access to the original
  values. As a consequence the original y values cannot be displayed when
  cdlValue is in the series' DataLabels. We display the percentage as a
  fall-back solution when needed.

  Another issue is that TAChart has only a single format for multiple stack
  layers. We use the one defined by the last series of the stack.

  And: TAChart calculates percentages in non-100% series as percentage of the
  all-series-max rather than percentage of the individual stack.)
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.ConstructSeriesMarks(AWorkbookSeries: TsChartSeries;
  AChartSeries: TChartSeries);
var
  sep: String;
  percentMode: Boolean;
  valueFmt: String = '%.9g';
  percentFmt: String = '%.0f%%';
  textFmt: String = '%2:s';
  totalFmt: String = '';
begin
  percentMode := AWorkbookSeries.Chart.StackMode = csmStackedPercentage;

  // Number format
  if AWorkbookSeries.LabelFormat <> '' then
    valueFmt := Convert_NumFormatStr_to_FormatStr(AWorkbookSeries.LabelFormat);
  System.Delete(valueFmt, 1, 1);  // Delete the '%'; will be re-added later with  value selector.

  // Percent format
  if AWorkbookSeries.LabelFormatPercent <> '' then
    percentFmt := Convert_NumFormatStr_to_FormatStr(AWorkbookSeries.LabelFormatPercent);
  System.Delete(percentFmt, 1, 1);

  // Working around some restrictions of TAChart...
  if percentMode then
  begin
    if cdlValue in AWorkbookSeries.DataLabels then
      valueFmt := '%0:' + percentFmt
    else
      valueFmt := '%0:' + valueFmt;
    if cdlPercentage in AWorkbookSeries.DataLabels then
      percentFmt := '%0:' + percentFmt;
  end else
  begin
    valueFmt := '%0:' + valueFmt;
    percentFmt := '%1:' + percentFmt;
  end;

  if AWorkbookSeries.DataLabels = [cdlCategory] then
    AChartSeries.Marks.Style := smsLabel
  else
  begin
    sep := AWorkbookSeries.LabelSeparator;
    if cdlCategory in AWorkbookSeries.DataLabels then
      totalFmt := textFmt;
    if cdlValue in AWorkbookSeries.DataLabels then
    begin
      if totalFmt <> '' then
        totalFmt := totalFmt + sep + valuefmt
      else
        totalFmt := valueFmt;
    end;
    if cdlPercentage in AWorkbookSeries.DataLabels then
    begin
      if totalFmt <> '' then
        totalFmt := totalFmt + sep + percentFmt
      else
        totalFmt := percentFmt;
    end;
    AChartSeries.Marks.Format := totalFmt;
  end;
  AChartSeries.Marks.Alignment := taCenter;
  AChartSeries.Marks.LinkPen.Visible := cdlLeaderLines in AWorkbookSeries.DataLabels;
end;

{@@ ----------------------------------------------------------------------------
  Creates a TAChart series from the provided workbook chart series.

  Stackable series are created only once, subsequent calls for the stack levels
  just provide data in the multi-y values chart source. Moreover, the
  necessary calculated chart sources for percent presentation are created, and
  the chart styles for individual data point styles are added.
-------------------------------------------------------------------------------}
function TsWorkbookChartLink.CreateChartSeries(ASeries: TsChartSeries): TChartSeries;
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
  interpolation: TsChartInterpolation = ciLinear;
  calcSrc: TCalculatedChartSource;
  style: TChartStyle;
  axAlign: TChartAxisAlignment;
  ax: TChartAxis;
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

    if (ASeries is TsCustomLineSeries) then
    begin
      interpolation := TsOpenedCustomLineSeries(ASeries).Interpolation;
      // TAChart cannot stack spline series.
      if (interpolation in [ciBSpline, ciCubicSpline]) and (ch.StackMode <> csmDefault) then
        interpolation := ciLinear;
    end;

    case ASeries.ChartType of
      ctBar:
        begin
          Result := TBarSeries.Create(FChart);
          src.IntegerX := true;
        end;
      ctLine, ctScatter:
        case interpolation of
          ciLinear, ciStepStart, ciStepEnd, ciStepCenterX, ciStepCenterY:
            begin
              Result := TLineSeries.Create(FChart);
              case interpolation of
                ciLinear: TLineSeries(Result).LineType := ltFromPrevious;
                ciStepStart: TLineSeries(Result).LineType := ltStepXY;
                ciStepEnd: TLineSeries(Result).LineType := ltStepYX;
                {$IF LCL_FullVersion >= 3990000}
                ciStepCenterX: TLineSeries(Result).LineType := ltStepCenterXY;
                ciStepCenterY: TLineSeries(Result).LineType := ltStepCenterYX;
                {$ENDIF}
                else TLineSeries(Result).LineType := ltFromPrevious;
              end;
            end;
          ciCubicSpline:
            Result := TCubicSplineSeries.Create(FChart);
          ciBSpline:
            Result := TBSplineSeries.Create(FChart);
        end;
      ctArea:
        Result := TAreaSeries.Create(FChart);
      ctRadar, ctFilledRadar:
        Result := TPolarSeries.Create(FChart);
      ctBubble:
        begin
          Result := TBubbleSeries.Create(FChart);
          src.SetYRange(1, TsBubbleSeries(ASeries).BubbleRange);  // The radius is at YIndex 1
        end;
      ctPie, ctRing:
        Result := TPieSeries.Create(FChart);
      ctStock:
        begin
          Result := TStockSeries.Create(FChart);
          src.YCount := 4;
          src.IntegerX := true;
          src.SetLabelRange(TsStockSeries(ASeries).LabelRange);
          src.SetXRange(0, TsStockSeries(ASeries).LabelRange);
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

    // Send pie offsets to chart source...
    src.CheckPieSeriesMode(ASeries);
    // ... as well as datapoint styles/colors
    CreateChartStylesFromDatapoints(ASeries, src.Styles);
    src.UseDataPointColors(ASeries);
    {$if LCL_FullVersion >= 3990000}
    if (Result is TPieSeries) then
      TPieSeries(Result).Styles := src.Styles
    else if (Result is TBubbleSeries) then
      TBubbleSeries(Result).Styles := src.Styles
    else if (Result is TBarSeries) then
      TBarSeries(Result).Styles := src.Styles;
    {$ifend}

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
  ax := Chart.AxisList.GetAxisByAlign(axAlign);
  if ax <> nil then
    Result.AxisIndexX := ax.Index;

  axAlign := AXIS_ALIGNMENT[ch.RotatedAxes, ASeries.YAxis, yAx];
  ax := FChart.AxisList.GetAxisByAlign(axAlign);
  if ax <> nil then
    Result.AxisIndexY := ax.Index;

  if stackable then
  begin
    style := TChartStyle(FChartStyles.Styles.Add);
    style.Text := src.Title;
  end;

  //ax := Chart.AxisList.GetAxisByAlign(calRight);
end;

{@@ ----------------------------------------------------------------------------
  Adjusts the area series zero level which otherwise is not clipped at the
  chart's plotrect (in TAChart before v3.99)

  @param  AWorkbookChart  Chart in the workbook which is handled here.
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
  Fixes the chart source for error bars.

  TAChart does supports error bars only for single-values chart sources. Only
  when FixSource is called it is known how many values are used by the source.
  If ranges have been added as rngXErrors or rgnYErrors display of error bars
  must be turned off in this case. Otherwise the main series would not be
  shown correctly.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Returns the TAChart auto-scale axis transform assigned to the specified
  chart axis, or nil if this axis transform type is not available here.

  @param     AChartAxis   TAChart axis which is investigated
  @returns   TAutoScaleAxisTransform instance if attached to the axis, or nil otherwise
-------------------------------------------------------------------------------}
function TsWorkbookChartLink.GetAutoScaleAxisTransform(AChartAxis: TChartAxis): TAutoScaleAxisTransform;
begin
  Result := TAutoScaleAxisTransform(GetAxisTransform(AChartAxis, TAutoScaleAxisTransform));
end;

{@@ ----------------------------------------------------------------------------
  Returns the TAChart axis transform when an axis transform of the specified
  class has been assigned to the specified axis, or nil when no such axis
  transform is assigned to the axis.

  @param  AChartAxis       TAChart axis for which the axis transform is searched.
  @param  AClass           Class of the axis transform searched.
  @returns  TAxisTransform instance of the requested class if found, or nil otherwise
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Returns the TAChart logarithmic axis transform assigned to the specified
  chart axis, or nil if this axis transform type is not used here.

  @param    AChartAxis  TAChart axis for which the axis transform is searched
  @returns  TLogarithmAxisTransform instance if found, or nil otherwise
-------------------------------------------------------------------------------}
function TsWorkbookChartLink.GetLogAxisTransform(AChartAxis: TChartAxis): TLogarithmAxisTransform;
begin
  Result := TLogarithmAxisTransform(GetAxisTransform(AChartAxis, TLogarithmAxisTransform))
end;

{@@ ----------------------------------------------------------------------------
  Returns the workbook chart which is handled by the chart link

  @returns  Workbook chart handled by the Chartlink
-------------------------------------------------------------------------------}
function TsWorkbookChartLink.GetWorkbookChart: TsChart;
begin
  if (Workbook <> nil) and (FWorkbookChartIndex > -1) then
    Result := Workbook.GetChartByIndex(FWorkbookChartIndex)
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Returns the workbook to which the chart belongs.
-------------------------------------------------------------------------------}
function TsWorkbookChartlink.GetWorkbook: TsWorkbook;
begin
  if FWorkbookSource <> nil then
    Result := FWorkbookSource.Workbook
  else
    Result := nil;
end;

{@@ ----------------------------------------------------------------------------
  Returns @true if the specified TAChart axis has logarithmic divisions.

  @param    Axis  TChart axis which is queried
  @returns  @True if the axis has logarithmic divisions.
-------------------------------------------------------------------------------}
function TsWorkbookChartLink.IsLogarithmic(Axis: TChartAxis): Boolean;
var
  T: TLogarithmAxisTransform;
begin
  T := GetLogAxisTransform(Axis);
  Result := (T <> nil) and T.Enabled;
end;

{@@ ----------------------------------------------------------------------------
  Returns true if the specified axis is a secondary axis, i.e. a right or top
  axis.

  @param    Axis   Workbook chart axis which is queried
  @returns  @true if the axis is at the right or top of the chart.
-------------------------------------------------------------------------------}
function TsWorkbookChartLink.IsSecondaryAxis(Axis: TsChartAxis): Boolean;
begin
  Result := (Axis = Axis.Chart.Y2Axis) or (Axis = Axis.Chart.X2Axis);
end;

{@@ ----------------------------------------------------------------------------
  Returns true if the specified workbook chart series can be stacked.
  This is possible bar, line and area series if they are assigned to the same
  axis.

  @param    ASeries    Workbook chart series which is queried
  @returns  @true if the series type principally can contain stacked levels.
-------------------------------------------------------------------------------}
function TsWorkbookChartLink.IsStackable(ASeries: TsChartSeries): Boolean;
var
  ch: TsChart;
  ser: TsChartSeries;
  i, numSeries: Integer;
begin
  Result := (ASeries.ChartType in [ctBar, ctLine, ctArea]) and (ASeries.GroupIndex > -1);
  if Result then
  begin
    ch := ASeries.Chart;
    numSeries := ch.Series.Count;
    if numSeries = 1 then
    begin
      Result := false;
      exit;
    end;

    // Check whether all series are the same type and same y axis as ASeries.
    // NOTE: Not perfect yet since there might abe two stackable groups,
    //       one for the left and one for the right axis...
    for i := 0 to numSeries - 1 do
    begin
      ser := ch.Series[i];
      if (ser.ChartType <> ASeries.ChartType) or (ser.GroupIndex <> ASeries.GroupIndex) or
         (ser.YAxis <> ASeries.YAxis) then
      begin
        Result := false;
        exit;
      end;
      if (ser is TsLineSeries) and (TsLineSeries(ser).Interpolation in [ciCubicSpline, ciBSpline]) then
      begin
        Result := false;
        exit;
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
   Notification message received from the WorkbookSource telling which
  spreadsheet item has changed.

  @unorderedlist(
   @item(Updates the chart when the message is received that all work sheets are loaded.)
   @item(When another worksheet has been selected the first chart of this worksheet is
         picked for the TAChart.)
  )

 @param(AChangedItems  Set of elements identifying whether workbook,
                       worksheet, cell content or cell formatting has changed)
 @param AData          Additional data, contains the worksheet for worksheet-related items

 @seeAlso    TsNotificationItem
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.ListenerNotification(AChangedItems: TsNotificationItems;
  AData: Pointer = nil);
var
  charts: TsChartArray;
begin
  Unused(AData);

  // Workbook has been successfully loaded, all sheets are ready
  if (lniWorkbook in AChangedItems) then
    UpdateChart;

  // Another worksheet is selected --> Select the first chart of the worksheet
  if (lniWorksheet in AChangedItems) and (WorkbookSource <> nil) then
  begin
    charts := WorkbookSource.Worksheet.GetCharts;
    if Length(charts) > 0 then
      WorkbookChartIndex := WorkbookSource.Workbook.GetChartIndex(charts[0])
    else
      WorkbookChartIndex := -1;
    if Assigned(FChart) then
      FChart.Visible := WorkbookChartIndex > -1;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Standard component notification: The ChartLink is notified that the
  WorkbookSource or chart is being removed and sets the internal
  WorkbookSource or Chart variables to nil, respectively.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Removes the chartbook source from the chart link
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.RemoveWorkbookSource;
begin
  SetWorkbookSource(nil);
end;

{@@ ----------------------------------------------------------------------------
  Connects the specified TAChart to the chart link

  @param AValue   TAChart to be linked
-------------------------------------------------------------------------------}
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
    FWorkbookSource.AddListener(self);
  ListenerNotification([lniWorkbook, lniWorksheet]);
  UpdateChart;
end;

{@@ ----------------------------------------------------------------------------
  Updates the properties of the TAChart area series by those provided by the
  workbook area series.

  @param  AWorkbookSeries  Series in the workbook chart which provides the series parameters.
  @param  AChartseries     TAChart series which is updated by these parameters.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateAreaSeries(AWorkbookSeries: TsAreaSeries;
  AChartSeries: TAreaSeries);
begin
  UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.Fill, AChartSeries.AreaBrush);
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.AreaContourPen);
  AChartSeries.Transparency := round(AWorkbookSeries.Fill.Color.Transparency * 255);
  AChartSeries.AreaLinesPen.Style := psClear;
  AChartSeries.Stacked := AWorkbookSeries.Chart.StackMode <> csmDefault;
  AChartSeries.UseZeroLevel := true;
  if AChartSeries.Source is TCalculatedChartSource then
    TCalculatedChartSource(AChartSeries.Source).Percentage := (AWorkbookSeries.Chart.StackMode = csmStackedPercentage);

  // Trend line
  UpdateChartSeriesTrendline(AWorkbookSeries, AChartSeries);

  // Error bars
  UpdateChartErrorBars(AWorkbookSeries, AChartSeries);
end;

{@@ ----------------------------------------------------------------------------
  Updates the properties of the TAChart bar series by those provided by the
  workbook bar series.

  @param  AWorkbookSeries  Series in the workbook chart which provides the series parameters.
  @param  AChartseries     TAChart series which is updated by these parameters.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateBarSeries(AWorkbookSeries: TsBarSeries;
  AChartSeries: TBarSeries);

  function CalcBarWidthPercent: Integer;
  var
    ser: TsChartSeries;
    gapwidth: Integer;
    i, n: Integer;
  begin
    n := 1;
    if (AWorkbookSeries.Chart.GetChartType = ctBar) and (AWorkbookSeries.Chart.StackMode = csmDefault) then
      for i := 0 to AWorkbookSeries.Chart.Series.Count-1 do
      begin
        ser := AWorkbookSeries.Chart.Series[i];
        if (ser <> AWorkbookSeries) and (ser.GroupIndex = AWorkbookSeries.GroupIndex) then
          inc(n);
      end;
    gapWidth := AWorkbookSeries.Chart.BarGapWidthPercent;
    Result := round(100/(n + gapWidth/100) * n);
  end;

begin
  UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.Fill, AChartSeries.BarBrush);
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.BarPen);
  AChartSeries.Transparency := round(AWorkbookSeries.Fill.Color.Transparency * 255);
  AChartSeries.BarWidthPercent := CalcBarWidthPercent;
  AChartSeries.BarOffsetPercent := 0; // TAChart currently does not support offsets in multiple-y bar series.
  AChartSeries.BarWidthStyle := bwPercentMin;
  AChartSeries.Stacked := AWorkbookSeries.Chart.StackMode <> csmDefault;
  if AChartSeries.Source is TCalculatedChartSource then
    TCalculatedChartSource(AChartSeries.Source).Percentage := (AWorkbookSeries.Chart.StackMode = csmStackedPercentage);

  // Trend line
  UpdateChartSeriesTrendLine(AWorkbookSeries, AChartSeries);
end;

{@@ ----------------------------------------------------------------------------
  Updates the properties of the LCL bubble series by those provided by the
  workbook bubble series.

  @param  AWorkbookSeries  Series in the workbook's chart which provides the series parameters.
  @param  AChartseries     TAChart series which is updated by these parameters.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartlink.UpdateBubbleSeries(AWorkbookSeries: TsBubbleSeries;
  AChartSeries: TBubbleSeries);
begin
  UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.Fill, AChartSeries.BubbleBrush);
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.BubblePen);
  AChartSeries.Transparency := round(AWorkbookSeries.Fill.Color.Transparency * 255);
  AChartSeries.Legend.Multiplicity := lmPoint;

  {$IF LCL_FullVersion >= 3990000}
  case AWorkbookSeries.BubbleSizeMode of
    bsmRadius: AChartSeries.BubbleRadiusUnits := bruPercentageRadius;
    bsmArea: AChartSeries.BubbleRadiusUnits := bruPercentageArea;
  end;
  AChartSeries.ParentChart.ExpandPercentage := 10;
  {$IFEND}

  // Trend line
  UpdateChartSeriesTrendline(AWorkbookSeries, AChartSeries);
end;

{@@ ----------------------------------------------------------------------------
  Updates all chart properties with thoser provided by the linked workbook chart
  and its elements.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateChart;
var
  ch: TsChart;
  i: Integer;
begin
  ClearChart;

  if (FChart = nil) or (FWorkbookSource = nil) or (FWorkbookChartIndex < 0) then
    exit;

  FChart.Proportional := false;
  FChart.ExpandPercentage := 0;

  ch := GetWorkbookChart;
  if ch = nil then
    exit;

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
end;

{@@ ----------------------------------------------------------------------------
  Updates the TAChart axis parameters with those provided by the specified
  workbook chart axis.

  @param AWorkbookAxis   Axis in the workbook chart which provides the parameters for the associated TAChart axis.
-------------------------------------------------------------------------------}
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
  if (AWorkbookAxis = nil) or (not AWorkbookAxis.Visible) then
    exit;

  align := TChartAxisAlignment(ord(AWorkbookAxis.Alignment));
  if AWorkBookAxis.Chart.RotatedAxes then
    align := ROTATED_ALIGNMENT[align];

  axis := FChart.AxisList.GetAxisByAlign(align);

  if (axis = nil) then
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
    TLogarithmAxisTransform(T).Base := AWorkbookAxis.LogBase;
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
  axis.Title.LabelFont.Orientation := round(AWorkbookAxis.TitleRotationAngle * 10);

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
  minorAxis := axis.Minors.Add;
  UpdateChartPen(AWorkbookAxis.Chart, AWorkbookAxis.MinorGridLines, minorAxis.Grid);
  minorAxis.Grid.Visible := (axis.Grid.Style <> psClear) and not IsSecondaryAxis(AWorkbookAxis);
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
    // Next line is not needed because we provide log labels in a ListChartSource now.
//    axis.Intervals.Options := axis.Intervals.Options + [aipGraphCoords{$IF LCL_FullVersion >= 3990000}, aipInteger{$IFEND}];
    axis.Intervals.MaxLength := 150;
    axis.Intervals.MinLength := 30;
    axis.Intervals.Tolerance := 30;
  end else
  begin
//    axis.Intervals.Options := axis.Intervals.Options - [aipGraphCoords{$IF LCL_FullVersion >= 3990000}, aipInteger{$IFEND}];
    axis.Intervals.MaxLength := 100;
    axis.Intervals.MinLength := 20;
    axis.Intervals.Tolerance := 0;
  end;

  // Axis position
  case AWorkbookAxis.Position of
    capStart:
      begin
        if (axis.Alignment = calLeft) or (axis.Alignment = calBottom) then
          axis.Position := 0
        else
          axis.Position := 100;
        axis.PositionUnits := cuPercent;
      end;
    capEnd:
      begin
        if (axis.Alignment = calLeft) or (axis.Alignment = calBottom) then
          axis.Position := 100
        else
          axis.Position := 0;
        axis.PositionUnits := cuPercent;
        // To do: Move TAChart axis labels and title to the other side
      end;
    capValue:
      begin
        if AWorkbookAxis.GetRotatedAxis.Logarithmic then
          axis.Position := log10(AWorkbookAxis.PositionValue)
        else
          axis.Position := AWorkbookAxis.PositionValue;
        axis.PositionUnits := cuGraph;  // To do: cuAxis not yet implemented in TAChart...

        // Workaround for barseries issue: When Position = 0 and PositionUnits = cuGraph
        // the axis would be moved to the center of the 1st bar which looks
        // ugly --> revert to cuPercent
        if (axis.Position = 0) and (AWorkbookAxis.Chart.GetChartType = ctBar) then
          axis.PositionUnits := cuPercent;
      end;
  end;
end;

{@@-----------------------------------------------------------------------------
  Updates the TAChart axis label parameters with those provided by the linked
  workbook chart.

  @param  AWorkbookChart   Chart in the workbook from which the properties are taken
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateChartAxisLabels(AWorkbookChart: TsChart);

  procedure MayBeLogAxis(Axis: TChartAxis);
  begin
    if Axis <> nil then
    begin
      if IsLogarithmic(Axis) then
      begin
        Axis.Marks.Source := FLogLabelSource;
        Axis.Marks.Style := smsLabel;
        Axis.Marks.TextFormat := tfHtml;
      end else
      begin
        Axis.Marks.Source := nil;
        Axis.Marks.Style := smsValue;
      end;
    end;
  end;

  function FindFirstSeriesOfType(ASeriesClass: TSeriesClass): TBasicChartSeries;
  var
    i: Integer;
  begin
    for i := 0 to FChart.Series.Count-1 do
      if FChart.Series[i] is ASeriesClass then
      begin
        Result := FChart.Series[i] as ASeriesClass;
        exit;
      end;
    Result := nil;
  end;

var
  axis: TChartAxis;
  ser: TBasicChartSeries;
  i: Integer;
  value: Double;
begin
  if FChart.SeriesCount = 0 then
    exit;

  if (AWorkbookChart.GetChartType in [ctScatter, ctBubble]) then
  begin
    FLogLabelSource.Clear;
    value := 1E-20;
    i := -20;
    while value < 1E20 do
    begin
      FLogLabelSource.Add(value, value, Format('10<sup>%d</sup>', [i]));
      value := value * 10;
      inc(i);
    end;
    MayBeLogAxis(FChart.BottomAxis);
    MayBeLogAxis(FChart.LeftAxis);
    MayBeLogAxis(FChart.AxisList.GetAxisByAlign(calRight));
  end;

  if AWorkbookChart.RotatedAxes then
    axis := FChart.LeftAxis
  else
    axis := FChart.BottomAxis;
  {$IF LCL_FullVersion >= 2020000}
  axis.Marks.SourceExchangeXY := AWorkbookChart.RotatedAxes;
  {$IFEND}

  case AWorkbookChart.GetChartType of
    ctScatter, ctBubble:
      begin
        {
        if IsLogarithmic(axis) then
        begin
          axis.Marks.Source := FLogLabelSource;
          axis.Marks.Style := smsLabel;
        end else
        begin
          axis.Marks.Source := nil;
          axis.Marks.Style := smsValue;
        end;
        }
      end;
    ctBar, ctLine, ctArea:
      begin
        axis.Marks.Source := TChartSeries(FChart.Series[0]).Source;
        if not AWorkbookChart.Series[0].LabelRange.IsEmpty then
          axis.Marks.Style := smsLabel
        else
          axis.Marks.Style := smsXValue;
      end;
    ctStock:
      begin
        ser := FindFirstSeriesOfType(TStockSeries);
        if Assigned(ser) then
          FChart.BottomAxis.Marks.Source := TChartSeries(ser).Source;
        FChart.BottomAxis.Marks.Style := smsLabel;
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
      DateTimeFormat := AWorkbookChart.XAxis.LabelFormat;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Updates the properties of the TAChart background according to the
  background of the workbook chart.

  @param  AWorkbookChart   Workbook chart which provides the background parameters.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateChartBackground(AWorkbookChart: TsChart);
begin
  FChart.Color := Convert_sColor_to_Color(AWorkbookChart.Background.Color.Color);
  FChart.BackColor := Convert_sColor_to_Color(AWorkbookChart.PlotArea.Background.Color.Color);
  UpdateChartPen(AWorkbookChart, AWorkbookChart.PlotArea.Border, FChart.Frame);
  FChart.Frame.Visible := not AWorkbookChart.PlotArea.Border.IsHidden;
end;

{@@ ----------------------------------------------------------------------------
  Updates the properties of an LCL brush by those provided by the
  AWorkbookFill instance inside the workbook chart.

  @param  AWorkbookChart Workbook chart which contains the workbook fill instance.
  @param  AWorkbookFill  A TsChartFill instance providing the properties for the TAChart
  @param  ABrush         LCL brush used by the TAChart which is updated by the AWorkbookFill properties
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateChartBrush(AWorkbookChart: TsChart;
  AWorkbookFill: TsChartFill; ABrush: TBrush);
begin
  if (AWorkbookFill <> nil) and (ABrush <> nil) then
  begin
    ABrush.Color := Convert_sColor_to_Color(AWorkbookFill.Color.Color);
    case AWorkbookFill.Style of
      cfsNoFill:
        ABrush.Style := bsClear;
      cfsSolidFill:
        ABrush.Style := bsSolid;
      cfsGradient:
        ABrush.Style := bsSolid;  // NOTE: TAChart cannot display gradients
      cfsPattern:
        ConstructHatchPattern(AWorkbookChart, AWorkbookFill, ABrush);
      cfsImage:
        ConstructImagePattern(AWorkbookChart, AWorkbookFill, ABrush);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Updates the properties of the error bars of a TAChart series by those provided by
  a workbook's chart series.

  @param  AWorkbookSeries  Series in the workbook's chart which provides the error bar parameters.
  @param  ASeries          TAChart series in which the error bars are updated by these parameters.
-------------------------------------------------------------------------------}

procedure TsWorkbookChartLink.UpdateChartErrorBars(AWorkbookSeries: TsChartSeries;
  ASeries: TBasicPointSeries);
const
  EPS = 1E-16;
  ERRORBAR_KINDS: array[TsChartErrorBarKind] of TChartErrorBarKind = (
    ebkNone, ebkConst, ebkPercent, ebkChartSource);

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
  series: TOpenedErrorbarSeries;
  source: TsWorkbookChartSource;
begin
  series := TOpenedErrorbarSeries(ASeries);
  source := GetChartSource(ASeries.Source);
  if source = nil then
    exit;

  // TAChart supports error bars only for single-values sources!
  if source.XCount = 1 then
  begin
    series.XErrorBars.Visible := AWorkbookSeries.XErrorBars.ShowPos or AWorkbookSeries.XErrorBars.ShowNeg;
    UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.XErrorBars.Line, series.XErrorBars.Pen);
    source.XErrorBarData.Kind := ERRORBAR_KINDS[AWorkbookSeries.XErrorBars.Kind];
    source.XErrorBarData.ValuePlus := AWorkbookSeries.XErrorBars.ValuePos;
    source.XErrorBarData.ValueMinus := AWorkbookSeries.XErrorBars.ValueNeg;
    if not AWorkbookSeries.XErrorBars.ShowPos then
      source.XErrorBarData.ValuePlus := EPS;    // Note: 0 would mean "no error bar at all" for TAChart!
    if not AWorkbookSeries.XErrorBars.ShowNeg then
      source.XErrorBarData.ValueMinus := EPS;
    if (AWorkbookSeries.XErrorBars.Kind = cebkCellRange) then
      source.SetXErrorBarRange(AWorkbookSeries.XErrorBars.RangePos, AWorkbookSeries.XErrorBars.RangeNeg);
    if not AWorkbookSeries.XErrorBars.ShowEndCap then
      series.XErrorBars.Width := 0;
  end;

  if source.YCount = 1 then
  begin
    series.YErrorBars.Visible := AWorkbookSeries.YErrorBars.ShowPos or AWorkbookSeries.YErrorBars.ShowNeg;;
    UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.YErrorBars.Line, series.YErrorBars.Pen);
    source.YErrorBarData.Kind := ERRORBAR_KINDS[AWorkbookSeries.YErrorBars.Kind];
    source.YErrorBarData.ValuePlus := AWorkbookSeries.YErrorBars.ValuePos;
    source.YErrorBarData.ValueMinus := AWorkbookSeries.YErrorBars.ValueNeg;
    if not AWorkbookSeries.YErrorBars.ShowPos then
      source.YErrorBarData.ValuePlus := EPS;
    if not AWorkbookSeries.YErrorBars.ShowNeg then
      source.YErrorBarData.ValueMinus := EPS;
    if (AWorkbookSeries.YErrorBars.Kind = cebkCellRange) then
      source.SetYErrorBarRange(AWorkbookSeries.YErrorBars.RangePos, AWorkbookSeries.YErrorBars.RangeNeg);
    if not AWorkbookSeries.YErrorBars.ShowEndCap then
      series.YErrorBars.Width := 0;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Updates the properties of the TAChart chart legend with those provided by the
  workbook chart legend.

  @param  AWorkbookLegend  Legend in the workbook chart which provides the parameters of the TAChart legend.
  @param  ALegend          TAChart legend which is updated by these parameters.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateChartLegend(AWorkbookLegend: TsChartLegend;
  ALegend: TChartLegend);
const
  LEG_POS: array[TsChartLegendPosition] of TLegendAlignment = (
    laCenterRight,   // legRight
    laTopCenter,     // legTop
    laBottomCenter,  // legBottom
    laCenterLeft     // legLeft
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
    {$IF LCL_FullVersion >= 3990000}
    ALegend.ColumnCount := 0;
    {$IFEND}
  end;
end;

{@@ ----------------------------------------------------------------------------
  Updates the properties of an LCL pen by those provided by the
  AWorkbookLine instance inside the workbook chart.

  @param  AWorkbookChart Workbook chart which contains the line instance.
  @param  AWorkbookLine  A TsChartLine instance providing the properties for the TAChart
  @param  APen           LCL pen used by the TAhart which is updated by the AWorkbookLine properties
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateChartPen(AWorkbookChart: TsChart;
  AWorkbookLine: TsChartLine; APen: TPen);
begin
  if (AWorkbookLine <> nil) and (APen <> nil) then
  begin
    APen.Color := Convert_sColor_to_Color(AWorkbookLine.Color.Color);
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

{@@ ----------------------------------------------------------------------------
  Updates the properties of the marks of an LCL brush series by those provided
  by the given workbook chart series.

  @param  AWorkbookSeries Workbook series which contains the data point marks to be used.
  @param  AChartseries    TAChart series to which the marks of the workbook chart series will be applied.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateChartSeriesMarks(AWorkbookSeries: TsChartSeries;
  AChartSeries: TChartSeries);
begin
  ConstructSeriesMarks(AWorkbookSeries, AChartSeries);

  AChartSeries.Marks.YIndex := -1;
  AChartSeries.Marks.Attachment := maDefault;
  Convert_sFont_to_Font(AWorkbookSeries.LabelFont, AChartSeries.Marks.LabelFont);

  if (AChartSeries is TBubbleSeries) then
    case AWorkbookSeries.LabelPosition of
      lpDefault, lpOutside, lpAbove:
        begin
          TBubbleSeries(AChartSeries).MarkPositions := lmpPositive;
          TBubbleSeries(AChartSeries).Marks.YIndex := 1;
          TBubbleSeries(AChartSeries).Marks.Distance := 5;
        end;
      lpBelow:
        begin
          TBubbleSeries(AChartSeries).MarkPositions := lmpNegative;
          TBubbleSeries(AChartSeries).Marks.YIndex := 1;
          TBubbleSeries(AChartSeries).Marks.Distance := 5;
        end;
      lpInside, lpNearOrigin:
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
        TOpenedBasicPointSeries(AChartSeries).MarkPositions := lmpOutside;
      lpOutside:
        TOpenedBasicPointSeries(AChartSeries).MarkPositions := lmpOutside;
      lpInside, lpNearOrigin:
        TOpenedBasicPointSeries(AChartSeries).MarkPositions := lmpInside;
      lpCenter:
        begin
          TOpenedBasicPointSeries(AChartSeries).MarkPositions := lmpInside;
          TOpenedBasicPointSeries(AChartSeries).MarkPositionCentered := true;
        end;
      lpAbove:
        TOpenedBasicPointSeries(AChartSeries).MarkPositions := lmpPositive;
      lpBelow:
        TOpenedBasicPointSeries(AChartSeries).MarkPositions := lmpNegative;
    end;

  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.LabelBorder, AChartSeries.Marks.Frame);
  UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.LabelBackground, AChartSeries.Marks.LabelBrush);
  case AWorkbookSeries.DataLabelCalloutShape of
    lcsRectangle: AChartSeries.Marks.Shape := clsRectangle;
    lcsRoundRect: AChartSeries.Marks.Shape := clsRoundRect;
    lcsEllipse: AChartSeries.Marks.Shape := clsEllipse;
    lcsRectangleWedge: AChartSeries.Marks.Shape := clsRectangle;  // replacement
    lcsRoundRectWedge: AChartSeries.Marks.Shape := clsRoundRect;  // replacement
    lcsEllipseWedge: AChartSeries.Marks.Shape := clsEllipse;      // replacement
    else AChartSeries.Marks.Shape := clsRectangle;                // replacement
  end;

  AChartSeries.Marks.LinkPen.Visible := (cdlLeaderLines in AWorkbookSeries.DataLabels);
  AChartSeries.Marks.LinkPen.Color := AChartSeries.Marks.Frame.Color;
  if AChartSeries.Marks.LinkPen.Visible then
    AChartSeries.Marks.Distance := 16
  else
    AChartSeries.Marks.Distance := 5;
end;

{@@ ----------------------------------------------------------------------------
  Updates a TAChart trend line with the parameters of a workbook chart trendline.

  @param  AWorkbookSeries  Series in the workbook chart which provides the trendline parameters
  @param  AChartSeries     TAChart series to which a trendline is created as a TFitSeries
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateChartSeriesTrendline(AWorkbookSeries: TsChartSeries;
  AChartSeries: TChartSeries);
var
  trendlineSeries: TsOpenedTrendlineSeries;
  trendline: TsChartTrendline;
  ser: TFitSeries;
  s: String;
begin
  if not AWorkbookSeries.SupportsTrendline then
    exit;

  trendlineSeries := TsOpenedTrendlineSeries(AWorkbookSeries);
  trendline := trendlineSeries.Trendline;

  if trendline.TrendlineType = tltNone then
    exit;

  // Create series and assign chartsource
  ser := TFitSeries.Create(FChart);
  ser.Source := AChartSeries.Source;

  // Fit equation
  case trendline.TrendlineType of
    tltLinear: ser.FitEquation := feLinear;
    // rtLogarithmic: ser.FitEquation := feLogarithmic;   // to do: implement this!
    tltExponential: ser.FitEquation := feExp;
    tltPower: ser.FitEquation := fePower;
    tltPolynomial:
      begin
        ser.FitEquation := fePolynomial;
        ser.ParamCount := trendline.PolynomialDegree + 1;
      end;
  end;

  // Take care of y intercept
  if trendline.ForceYIntercept then
  begin
    str(trendline.YInterceptValue, s);
    ser.FixedParams := s;
  end;

  // style of trend line
  UpdateChartPen(AWorkbookSeries.Chart, trendline.Line, ser.Pen);
  ser.AxisIndexX := AChartSeries.AxisIndexX;
  ser.AxisIndexY := AChartSeries.AxisIndexY;

  FChart.AddSeries(ser);

  // Legend text
  ser.Title := trendline.Title;

  {
  // Show fit curve in legend after series.
  ser.Legend.Order := AChartseries.Legend.Order + 1;
  }

  // Trendline equation
  if trendline.DisplayEquation or trendline.DisplayRSquare then
  begin
    ser.ExecFit;
    s := '';
    if trendline.DisplayEquation then
      s := s + ser.EquationText.
        X(trendline.Equation.XName).
        Y(trendline.Equation.YName).
        NumFormat(Convert_NumFormatStr_to_FormatStr(trendline.Equation.NumberFormat)).
        DecimalSeparator('.').
        TextFormat(tfHtml).
        Get;
    if trendline.DisplayRSquare then
      s := s + LineEnding + 'R<sup>2</sup> = ' + FormatFloat('0.00', ser.FitStatistics.R2);
    if s <> '' then
      ser.Title := ser.Title + LineEnding + s;
//    ser.Legend.Format := '%0:s' + LineEnding + '%2:s';
  end;
end;

{@@ ----------------------------------------------------------------------------
  Updates the TAChart chartstyle having the specified index with parameters
  defined for a workbook chart series.

  @param  AWorkbookSeries  Workbook chart series from which the properties are applied to the ChartStyle
  @param  AStyleIndex      Index of the TChartStyle to which the properties are copied
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateChartStyle(AWorkbookSeries: TsChartSeries;
  AStyleIndex: Integer);
var
  style: TChartStyle;
  ch: TsChart;
  ppi: Integer;
begin
  ch := AWorkbookSeries.Chart;
  style := TChartStyle(FChartStyles.Styles[AStyleIndex]);
  UpdateChartPen(ch, AWorkbookSeries.Line, style.Pen);
  if (AWorkbookSeries is TsCustomLineSeries) then
  begin
    {$IF LCL_FullVersion >= 3990000}
    ppi := GetParentForm(FChart).PixelsPerInch;
    style.UsePointer := true;
    style.Pointer.Visible := true;
    style.Pointer.Style := POINTER_STYLES[TsOpenedCustomLineSeries(AWorkbookSeries).Symbol];
    style.Pointer.HorizSize := mmToPx(TsOpenedCustomLineSeries(AWorkbookSeries).SymbolWidth, ppi);
    style.Pointer.VertSize := mmToPx(TsOpenedCustomLineSeries(AWorkbookSeries).SymbolHeight, ppi);
    UpdateChartBrush(ch, TsOpenedCustomLineSeries(AWorkbookSeries).SymbolFill, style.Pointer.Brush);
    UpdateChartPen(ch, TsOpenedCustomLineSeries(AWorkbookSeries).SymbolBorder, style.Pointer.Pen);
    {$ELSE}
    UpdateChartBrush(ch, TsOpenedCustomLineSeries(AWorkbookSeries).SymbolFill, style.Brush);
    {$ENDIF}
  end
  else
    UpdateChartBrush(ch, AWorkbookSeries.Fill, style.Brush);
end;

{@@ ----------------------------------------------------------------------------
  Updates title and footer of the linked TAChart.

  @param  AWorkbookTitle  Title of the workbook chart which provides the parameters
  @param  AChartTitle     TAChart title to which these parameters are applied
  @note   The workbook chart's subtitle is converted to TAChart's footer!
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Updates a TAChart TBasicPointSeries with the parameters of a workbook TsCustomLineSeries.
  Both series types are ancestors of more complex types.

  @param  AWorkbookSeries  Workbook TsCustomLineSeries instance providing the parameters for the TAChart series.
  @param  AChartSeries     Basic TAChart point series to which these parameters are applied
-------------------------------------------------------------------------------}

procedure TsWorkbookChartLink.UpdateCustomLineSeries(AWorkbookSeries: TsCustomLineSeries;
  AChartSeries: TBasicPointSeries);
var
  ppi: Integer;
  openedWorkbookSeries: TsOpenedCustomLineSeries absolute AWorkbookSeries;
  lineSeries: TLineSeries absolute AChartSeries;
  cubicSplineSeries: TCubicSplineSeries absolute AChartSeries;
  bSplineSeries: TBSplineSeries absolute AChartSeries;
  seriesPointer: TSeriesPointer;
begin
  ppi := GetParentForm(FChart).PixelsPerInch;

  if AChartSeries is TLineSeries then
  begin
    UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, lineSeries.LinePen);
    lineSeries.ShowLines := not AWorkbookSeries.Line.IsHidden;
    seriesPointer := lineSeries.Pointer;
    lineSeries.Stacked := AWorkbookSeries.Chart.StackMode <> csmDefault;
    if lineSeries.Source is TCalculatedChartSource then
      TCalculatedChartSource(lineSeries.Source).Percentage := (AWorkbookSeries.Chart.StackMode = csmStackedPercentage);
  end
  else
  if AChartSeries is TCubicSplineSeries then
  begin
    UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, cubicSplineSeries.Pen);
    cubicSplineSeries.Pen.Visible := not AWorkbookSeries.Line.IsHidden;
    seriesPointer := cubicSplineSeries.Pointer;
  end
  else
  if AChartSeries is TBSplineSeries then
  begin
    UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, bSplineSeries.Pen);
    bSplineSeries.Pen.Visible := not AWorkbookSeries.Line.IsHidden;
    seriesPointer := bSplineSeries.Pointer;
  end;

  seriesPointer.Visible := openedWorkbookSeries.ShowSymbols;
  UpdateChartBrush(AWorkbookSeries.Chart, openedWorkbookSeries.SymbolFill, seriesPointer.Brush);
  UpdateChartPen(AWorkbookSeries.Chart, openedWorkbookSeries.SymbolBorder, seriesPointer.Pen);
  seriesPointer.Style := POINTER_STYLES[openedWorkbookSeries.Symbol];
  seriesPointer.HorizSize := mmToPx(openedWorkbookSeries.SymbolWidth / 2, ppi);
  seriesPointer.VertSize := mmToPx(openedWorkbookSeries.SymbolHeight / 2, ppi);

  // Error bars
  UpdateChartErrorBars(AWorkbookSeries, AChartSeries);

  // Trend line
  UpdateChartSeriesTrendline(AWorkbookSeries, AChartSeries);
end;

{@@ ----------------------------------------------------------------------------
  Updates a TAChart TPieSeries with the parameters of a workbook TsPieSeries

  @param  AWorkbookSeries  Workbook TsPieSeries instance providing the parameters for the TAChart series.
  @param  AChartSeries     TAChart series (type TPieSeries) to which these parameters are applied
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdatePieSeries(AWorkbookSeries: TsPieSeries;
  AChartSeries: TPieSeries);
begin
  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.EdgePen);
  AChartSeries.Legend.Multiplicity := lmPoint;
  AChartSeries.Legend.Format := '%2:s';
  {$IF LCL_FullVersion >= 2020000}
  AChartSeries.StartAngle := AWorkbookSeries.StartAngle;
  if AWorkbookSeries is TsPieSeries then
    AChartSeries.InnerRadiusPercent := TsPieSeries(AWorkbookSeries).InnerRadiusPercent;
  {$IFEND}
  AChartSeries.Exploded := true;

  FChart.BottomAxis.Visible := false;
  FChart.LeftAxis.Visible := false;
  FChart.Legend.Inverted := false;
  FChart.Frame.Visible := false;
end;

{@@ ----------------------------------------------------------------------------
  Updates a TAChart PolarSeries with the parameters of a workbook TsRadarSeries.

  @param  AWorkbookSeries  Workbook TsRadarSeries instance providing the parameters for the TAChart series.
  @param  AChartSeries     TAChart series (type TPolarSeries) to which these parameters are applied
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdatePolarSeries(AWorkbookSeries: TsRadarSeries;
  AChartSeries: TPolarSeries);
var
  ppi: Integer;
begin
  ppi := GetParentForm(FChart).PixelsPerInch;

  AChartSeries.Transparency := round(AWorkbookSeries.Fill.Color.Transparency * 255);
  AChartSeries.CloseCircle := true;
  AChartSeries.Filled := (AWorkbookSeries.ChartType = ctFilledRadar);
  AChartSeries.ShowPoints := AWorkbookSeries.ShowSymbols;
  (AChartSeries.Source as TsWorkbookChartSource).CyclicX := true;

  UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.Line, AChartSeries.LinePen);
  if AWorkbookSeries.ChartType = ctFilledRadar then
    UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.Fill, AChartSeries.Brush);
  if AWorkbookSeries.ShowSymbols then
  begin
    UpdateChartBrush(AWorkbookSeries.Chart, AWorkbookSeries.SymbolFill, AChartSeries.Pointer.Brush);
    UpdateChartPen(AWorkbookSeries.Chart, AWorkbookSeries.SymbolBorder, AChartSeries.Pointer.Pen);
    AChartSeries.Pointer.Style := POINTER_STYLES[AWorkbookSeries.Symbol];
    AChartSeries.Pointer.HorizSize := mmToPx(AWorkbookSeries.SymbolWidth / 2, ppi);
    AChartSeries.Pointer.VertSize := mmToPx(AWorkbookSeries.SymbolHeight / 2, ppi);
  end;

  FChart.LeftAxis.Minors.Clear;
  FChart.LeftAxis.Grid.Visible := false;
  FChart.BottomAxis.Minors.Clear;
  FChart.BottomAxis.Grid.Visible := false;
  FChart.Proportional := true;
end;

{@@ ----------------------------------------------------------------------------
  Applies the properties of workbook scatter series to a TAChart TLineSeries.

  @param  AWorkbookSeries  The scatter series in the workbook chart providing the formatting parameters
  @param  AChartSeries     The TAChart line series to which these parameters are applied.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.UpdateScatterSeries(AWorkbookSeries: TsScatterSeries;
  AChartSeries: TLineSeries);
begin
  UpdateCustomLineSeries(AWorkbookSeries, AChartSeries);
end;

{@@ ----------------------------------------------------------------------------
  Updates the properties of a TAChart StockSeries/HighLowCloseSeries by those
  of a workbook TsStockSeries

  @param  AWorkbookSeries  The stock series in the workbook chart providing the formatting parameters
  @param  AChartSeries     The TAChart series to which these parameters are applied.
-------------------------------------------------------------------------------}
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
  AChartSeries.TickWidth := AWorkbookSeries.TickWidthPercent div 2;

  // Trend line
  UpdateChartSeriesTrendline(AWorkbookSeries, AChartSeries);
end;

{@@ ----------------------------------------------------------------------------
  Creates the chart styles for a workbook series containing several y values.
  Chart styles are used in TAChart to format the subseries resulting from a
  multi-y valued chart source.

  @param  AWorkbookSeries  Workbook chart series providing the formatting parameters in its DataPointStyles property
  @param  AChartStyles     TAChart ChartStyles instance which gets populated according to the DataPointStyles of the workbook series.
-------------------------------------------------------------------------------}
procedure TsWorkbookChartLink.CreateChartStylesFromDatapoints(AWorkbookSeries: TsChartSeries;
  AChartStyles: TChartStyles);

  procedure FillAndLineToStyle(AFill: TsChartFill; ALine: TsChartLine; AStyle: TChartStyle);
  begin
    UpdateChartBrush(AWorkbookSeries.Chart, AFill, AStyle.Brush);
    UpdateChartPen(AWorkbookSeries.Chart, ALine, ASTyle.Pen);
  end;

var
  style: TChartStyle;
  datapointStyle: TsChartDatapointStyle;
  i, j: Integer;
  fill: TsChartFill;
  line: TsChartLine;
begin
  AChartStyles.Styles.Clear;

  if AWorkbookSeries.DataPointStyles.Count = 0 then
    exit;
  if not ((AWorkbookSeries is TsPieSeries) or (AWorkbookSeries is TsBubbleSeries)) then
    exit;
  if (AWorkbookSeries is TsBarSeries) and (AWorkbookSeries.Chart.Series.Count > 1) then
    exit;  // TAChart cannot handle datapoint styles for layered bar series

  j := 0;
  datapointStyle := AWorkbookSeries.DataPointStyles[0];
  for i := 0 to AWorkbookSeries.Count-1 do
  begin
    style := AChartStyles.Add;
    if (datapointStyle = nil) or (i < datapointStyle.DatapointIndex) then
      FillAndLineToStyle(AWorkbookSeries.Fill, AWorkbookSeries.Line, style)
    else
    begin
      if datapointStyle.Background <> nil then
        fill := datapointstyle.Background
      else
        fill := AWorkbookSeries.Fill;
      if datapointStyle.Border <> nil then
        line := datapointstyle.Border
      else
        line := AWorkbookSeries.Line;
      FillAndLineToStyle(fill, line, style);
      if j < AWorkbookSeries.DataPointStyles.Count-1 then
      begin
        inc(j);
        datapointStyle := AWorkbookSeries.DataPointStyles[j];
      end else
        datapointStyle := nil;
    end;
  end;
end;

{$ENDIF}

end.
