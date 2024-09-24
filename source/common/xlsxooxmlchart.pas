unit xlsxooxmlChart;

{$mode objfpc}{$H+}
{$include ..\fps.inc}

interface

{$ifdef FPS_CHARTS}

uses                                                  //LazLoggerBase,
  Classes, SysUtils, StrUtils, Contnrs, FPImage,
  {$ifdef FPS_PATCHED_ZIPPER}fpszipper,{$else}zipper,{$endif}
  laz2_xmlread, laz2_DOM,
  fpsTypes, fpSpreadsheet, fpsChart, fpsUtils, fpsNumFormat, fpsImages,
  fpsReaderWriter, fpsXMLCommon;

type
  { TsSpreadOOXMLChartReader }

  TsSpreadOOXMLChartReader = class(TsBasicSpreadChartReader)
  private
    FPointSeparatorSettings: TFormatSettings;
    FImages: TFPObjectList;
    FXAxisID, FYAxisID, FX2AxisID, FY2AxisID: DWord;
    FXAxisDelete, FYAxisDelete, FX2AxisDelete, FY2AxisDelete: Boolean;

    procedure ReadChartColor(ANode: TDOMNode; var AColor: TsChartColor);
    function ReadChartColorDef(ANode: TDOMNode; ADefault: TsChartColor): TsChartColor;
    procedure ReadChartFillAndLineProps(ANode: TDOMNode;
      AChart: TsChart; AFill: TsChartFill; ALine: TsChartLine);
    procedure ReadChartFontProps(ANode: TDOMNode; AFont: TsFont);
    procedure ReadChartGradientFillProps(ANode: TDOMNode; AChart: TsChart; AFill: TsChartFill);
    procedure ReadChartHatchFillProps(ANode: TDOMNode; AChart: TsChart; AFill: TsChartFill);
    procedure ReadChartImageFillProps(ANode: TDOMNode; AChart: TsChart; AFill: TsChartFill);
    procedure ReadChartLineProps(ANode: TDOMNode; AChart: TsChart; AChartLine: TsChartLine);
    procedure ReadChartTextProps(ANode: TDOMNode; AFont: TsFont; var AFontRotation: Single);
    procedure SetAxisDefaults(AWorkbookAxis: TsChartAxis);
    procedure SetDefaultSeriesColor(ASeries: TsChartSeries);
  protected
    procedure ReadChart(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartAreaSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartAxis(ANode: TDOMNode; AChart: TsChart; AChartAxis: TsChartAxis;
      var AxisID: DWord; var ADelete: Boolean);
    procedure ReadChartAxisScaling(ANode: TDOMNode; AChartAxis: TsChartAxis);
    function ReadChartAxisTickMarks(ANode: TDOMNode): TsChartAxisTicks;
    procedure ReadChartBarSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartBubbleSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartImages(AStream: TStream; AChart: TsChart; ARelsList: TFPList);
    procedure ReadChartLegend(ANode: TDOMNode; AChartLegend: TsChartLegend);
    procedure ReadChartLineSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartPieSeries(ANode: TDOMNode; AChart: TsChart; RingMode: Boolean);
    procedure ReadChartPlotArea(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartRadarSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartScatterSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartSeriesAxis(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartSeriesDataPointStyles(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartSeriesErrorBars(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartSeriesLabels(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartSeriesMarker(ANode: TDOMNode; ASeries: TsCustomLineSeries);
    procedure ReadChartSeriesProps(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartSeriesRange(ANode: TDOMNode; ARange: TsChartRange; var AFormat: String);
    procedure ReadChartSeriesTitle(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartSeriesTrendLine(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartStockSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartStockSeriesUpDownBars(ANode: TDOMNode; ASeries: TsStockSeries);
    procedure ReadChartTitle(ANode: TDOMNode; ATitle: TsChartText);
  public
    constructor Create(AReader: TsBasicSpreadReader); override;
    destructor Destroy; override;
    procedure ReadChartXML(AStream: TStream; AChart: TsChart; AChartXMLFile: String);
  end;

  { TsSpreadOOXMLChartWriter }

  TsSpreadOOXMLChartWriter = class(TsBasicSpreadChartWriter)
  private
    FSCharts: array of TStream;
    FSChartRels: array of TStream;
    FSChartStyles: array of TStream;
    FSChartColors: array of TStream;
    FPointSeparatorSettings: TFormatSettings;
    FAxisID: array[TsChartAxisAlignment] of DWord;
    FSeriesIndex: Integer;
    function GetChartColorXML(AIndent: Integer; ANodeName: String; AColor: TsChartColor): String;
    function GetChartColorXML(AColor: TsChartColor): String;
    function GetChartFillAndLineXML(AIndent: Integer; AChart: TsChart; AFill: TsChartFill; ALine: TsChartLine): String;
    function GetChartFillXML(AIndent: Integer; AChart: TsChart; AFill: TsChartFill): String;
    function GetChartFontXML(AIndent: Integer; AFont: TsFont; ANodeName: String): String;
    function GetChartLineXML(AIndent: Integer; AChart: TsChart; ALine: TsChartLine; OverrideOff: Boolean = false): String;
    function GetChartSeriesMarkerXML(AIndent: Integer; AChart: TsChart;
      AShowSymbols: Boolean; ASymbolKind: TsChartSeriesSymbol = cssRect;
      ASymbolWidth: Double = 3.0; ASymbolHeight: Double = 3.0;
      ASymbolFill: TsChartFill = nil; ASymbolBorder: TsChartLine = nil): String;

  protected
    // Called by the public functions
    procedure WriteChartColorsXML(AStream: TStream; AChartIndex: Integer);
    procedure WriteChartRelsXML(AStream: TStream; AChartIndex: Integer);
    procedure WriteChartStylesXML(AStream: TStream; AChartIndex: Integer);
    procedure WriteChartSpaceXML(AStream: TStream; AChartIndex: Integer);

    // Writing the main chart xml nodes
    procedure WriteChartNode(AStream: TStream; AIndent: Integer; AChartIndex: Integer);

    procedure WriteChartAxisNode(AStream: TStream; AIndent: Integer; Axis: TsChartAxis; ANodeName: String);
    procedure WriteChartAxisScaling(AStream: TStream; AIndent: Integer; Axis: TsChartAxis);
    procedure WriteChartAxisTitle(AStream: TStream; AIndent: Integer; Axis: TsChartAxis);
    procedure WriteChartLegendNode(AStream: TStream; AIndent: Integer; ALegend: TsChartLegend);
    procedure WriteChartPlotAreaNode(AStream: TStream; AIndent: Integer; AChart: TsChart);
    procedure WriteChartRange(AStream: TStream; AIndent: Integer; ARange: TsChartRange; ANodeName, ARefName: String; WriteCache: Boolean = false);
    procedure WriteChartTrendline(AStream: TStream; AIndent: Integer; ASeries: TsChartSeries);
    function WriteChartSeries(AStream: TStream; AIndent: Integer; AChart: TsChart; AxisLink: TsChartAxisLink; out xAxKind: String): Boolean;
    procedure WriteChartSeriesDatapointLabels(AStream: TStream; AIndent: Integer; ASeries: TsChartSeries);
    procedure WriteChartSeriesDatapointStyles(AStream: TStream; AIndent: Integer; ASeries: TsChartSeries);
    procedure WriteChartSeriesErrorBars(AStream: TStream; AIndent: Integer; ASeries: TsChartSeries; IsYError: Boolean);
    procedure WriteChartSeriesNode(AStream: TStream; AIndent: Integer; ASeries: TsChartSeries);
    procedure WriteChartSeriesTitle(AStream: TStream; AIndent: Integer; ASeries: TsChartSeries);
    procedure WriteChartTitleNode(AStream: TStream; AIndent: Integer; ATitle: TsChartText);

    // Writing the nodes of the series types
    procedure WriteAreaSeries(AStream: TStream; AIndent: Integer; ASeries: TsAreaSeries; ASeriesIndex, APosInAxisGroup: Integer);
    procedure WriteBarSeries(AStream: TStream; AIndent: Integer; ASeries: TsBarSeries; ASeriesIndex, APosInAxisGroup: Integer);
    procedure WriteBubbleSeries(AStream: TStream; AIndent: Integer; ASeries: TsBubbleSeries; APosInAxisGroup: Integer);
    procedure WriteLineSeries(AStream: TStream; AIndent: Integer; ASeries: TsLineSeries; ASeriesIndex, APosInAxisGroup: Integer);
    procedure WritePieSeries(AStream: TStream; AIndent: Integer; ASeries: TsPieSeries);
    procedure WriteRadarSeries(AStream: TStream; AIndent: Integer; ASeries: TsRadarSeries);
    procedure WriteScatterSeries(AStream: TStream; AIndent: Integer; ASeries: TsScatterSeries; APosInAxisGroup: Integer);
    procedure WriteStockSeries(AStream: TStream; AIndent: Integer; ASeries: TsStockSeries; APosInAxisGroup: Integer);
    procedure WriteStockSeriesNode(AStream: TStream; AIndent: Integer; ASeries: TsStockSeries; ASeriesIndex, OHLCPart: Integer; WriteCache: Boolean);

    procedure WriteChartLabels(AStream: TStream; AIndent: Integer; AFont: TsFont);
    procedure WriteChartText(AStream: TStream; AIndent: Integer; AText: TsChartText; ARotationAngle: Single);

    procedure WriteCellNumberValue(AStream: TStream; AIndent: Integer; AWorksheet: TsBasicWorksheet; ARow,ACol,AIndex: Cardinal);

  public
    constructor Create(AWriter: TsBasicSpreadWriter); override;
    destructor Destroy; override;

    // Public functions called by the main writer
    procedure AddChartsToZip(AZip: TZipper);
    procedure CreateStreams; override;
    procedure DestroyStreams; override;
    procedure ResetStreams; override;
    procedure WriteChartContentTypes(AStream: TStream);
    procedure WriteCharts; override;
  end;

{$ENDIF}

implementation

{$IFDEF FPS_CHARTS}

uses
  xlsxooxml;

const
  MIME_DRAWINGML_CHART        = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';
  MIME_DRAWINGML_CHART_STYLE  = 'application/vnd.ms-office.chartstyle+xml';
  MIME_DRAWINGML_CHART_COLORS = 'application/vnd.ms-office.chartcolorstyle+xml';

  SCHEMAS_RELS         = 'http://schemas.openxmlformats.org/package/2006/relationships';
  SCHEMAS_CHART_COLORS = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle';
  SCHEMAS_CHART_STYLE  = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle';

  OOXML_PATH_XL_CHARTS      = 'xl/charts/';
  OOXML_PATH_XL_CHARTS_RELS = 'xl/charts/_rels/';

  LE = LineEnding;

  PTS_MULTIPLIER = 12700;
  ANGLE_MULTIPLIER = 60000;
  PERCENT_MULTIPLIER = 1000;
  FACTOR_MULTIPLIER = 100000;

  DEFAULT_FONT_NAME = 'Liberation Sans';

  AX_POS: array[boolean, TsChartAxisAlignment] of string = (
    ('l', 't', 'r', 'b'),
    ('b', 'r', 't', 'l')
  );  //caaLeft, caaTop, caaRight, caaBottom

  FALSE_TRUE: Array[boolean] of Byte = (0, 1);

  LEGEND_POS: Array[TsChartLegendPosition] of string = ('r', 't', 'b', 'l');

  TRENDLINE_TYPES: Array[TsTrendlineType] of string = ('', 'linear', 'log', 'exp', 'power', 'poly');
    // 'movingAvg' and 'log' not supported, so far

  LABEL_POS: Array[TsChartLabelPosition] of string = ('', '', 'inEnd', 'ctr', '', 'inBase', 'inBase');
    // lpDefault, lpOutside, lpInside, lpCenter, lpAbove, lpBelow, lpNearOrigin

  DEFAULT_TEXT_DIR: array[boolean, TsChartAxisAlignment] of Integer = (
    (90, 0, 90, 0),  // not rotated for: caaLeft, caaTop, caaRight, caaBottom
    (0, 90, 0, 90)   // rotated for: caaLeft, caaTop, caaRight, caaBottom
  );

  XLSX_SERIES_COLORS: array[0..5] of DWord = (
    $D59B5B, $317DED, $A5A5A5, $00C0FF, $C47244, $47AD70
  );

  OHLC_OPEN = 0;
  OHLC_HIGH = 1;
  OHLC_LOW = 2;
  OHLC_CLOSE = 3;

{$INCLUDE xlsxooxmlchart_hatch.inc}

type
  TNamedStreamItem = class
    Name: String;
    Stream: TStream;
  end;

  TNamedStreamList = class(TFPObjectList)
  public
    function FindStreamByName(const AName: String): TStream;
  end;

function TNamedStreamList.FindStreamByName(const AName: String): TStream;
var
  i: Integer;
begin
  for i := 0 to Count-1 do
    if TNamedStreamItem(Items[i]).Name = AName then
    begin
      Result := TNamedStreamItem(Items[i]).Stream;
      exit;
    end;
  Result := nil;
end;

function PositiveAngle(Angle: Double): Double;
begin
  Result := Angle;
  while Result < 0 do
    Result := Result + 360.0;
end;

type
  TsOpenedCustomLineSeries = class(TsCustomLineSeries)
  public
    property Symbol;
    property SymbolBorder;
    property SymbolFill;
    property SymbolHeight;
    property SymbolWidth;
    property ShowLines;
    property ShowSymbols;
  end;

  TsOpenedTrendlineSeries = class(TsChartSeries)
  public
    property Trendline;
  end;


function CalcDefaultSeriesColor(AIndex: Integer): TsChartColor;
const
  LUM:  array[0..8] of Double = (1.0, 0.6, 0.8, 0.8, 0.6, 0.5, 0.7, 0.7, 0.5);
  OFFS: array[0..8] of Double = (0.0, 0.0, 0.2, 0.0, 0.4, 0.0, 0.3, 0.0, 0.5);
var
  c: TsColor;
  idx, modif: Integer;
begin
  idx := AIndex div Length(XLSX_SERIES_COLORS);
  modif := AIndex mod Length(XLSX_SERIES_COLORS);
  c := LumModOff(XLSX_SERIES_COLORS[idx], LUM[modif], OFFS[modif]);
  Result := ChartColor(c);
end;

function HTMLColorStr(AValue: TsColor): string;
var
  rgb: TRGBA absolute AValue;
begin
  Result := Format('%.2x%.2x%.2x', [rgb.r, rgb.g, rgb.b]);
end;


{ TsSpreadOOXMLChartReader }

constructor TsSpreadOOXMLChartReader.Create(AReader: TsBasicSpreadReader);
begin
  inherited Create(AReader);

  FPointSeparatorSettings := SysUtils.DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator:='.';

  FImages := TFPObjectList.Create;
end;

destructor TsSpreadOOXMLChartReader.Destroy;
begin
  FImages.Free;
  inherited;
end;

procedure TsSpreadOOXMLChartReader.ReadChart(ANode: TDOMNode; AChart: TsChart);
var
  nodeName: String;
begin
  // Defaults  (to be completed...)
  AChart.Legend.Visible := false;

  ANode := ANode.FirstChild;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:title':
        ReadChartTitle(ANode.FirstChild, AChart.Title);
      'c:plotArea':
        ReadChartPlotArea(ANode.FirstChild, AChart);
      'c:legend':
        ReadChartlegend(ANode.FirstChild, AChart.Legend);
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartAreaSeries(ANode: TDOMNode; AChart: TsChart);
var
  nodeName: String;
  s: String;
  ser: TsAreaSeries;
begin
  if ANode = nil then
    exit;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:ser':
        begin
          ser := TsAreaSeries.Create(AChart);
          SetDefaultSeriesColor(ser);
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:grouping':
        begin
          s := GetAttrValue(ANode, 'val');
          case s of
            'stacked': AChart.StackMode := csmStacked;
            'percentStacked': AChart.StackMode := csmStackedPercentage;
          end;
        end;
      'c:varyColors':
        ;
      'c:dLbls':
        ;
      'c:axId':
        ReadChartSeriesAxis(ANode, ser);
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartAxis(ANode: TDOMNode;
  AChart: TsChart; AChartAxis: TsChartAxis; var AxisID: DWord; var ADelete: Boolean);
var
  nodeName, s: String;
  n: LongInt;
  x: Single;
  node: TDOMNode;
begin
  if ANode = nil then
    exit;

  ADelete := false;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:axId':
        if (s <> '') and TryStrToInt(s, n) then
          AxisID := n;
      'c:delete':
        if s = '1' then
          ADelete := true;
      'c:axPos':
        case s of
          'l': AChartAxis.Alignment := caaLeft;
          'b': AChartAxis.Alignment := caaBottom;
          'r': AChartAxis.Alignment := caaRight;
          't': AChartAxis.Alignment := caaTop;
        end;
      'c:scaling':
        ReadChartAxisScaling(ANode.FirstChild, AChartAxis);
      'c:majorGridlines':
        begin
          node := ANode.FindNode('c:spPr');
          if Assigned(node) then
            ReadChartLineProps(node.FirstChild, AChart, AChartAxis.MajorGridLines);
        end;
      'c:minorGridlines':
        begin
          node := ANode.FindNode('c:spPr');
          if Assigned(node) then
            ReadChartLineProps(node.FirstChild, AChart, AChartAxis.MinorGridLines);
        end;
      'c:title':
        ReadChartTitle(ANode.FirstChild, AChartAxis.Title);
      'c:numFmt':
        begin
          s := GetAttrValue(ANode, 'formatCode');
          if (s = 'm/d/yyyy') or (s = 'mm/dd/yyyy') then
            AChartAxis.LabelFormat := FReader.Workbook.FormatSettings.ShortDateFormat
          else
          if IsDateTimeFormat(s) then
          begin
            AChartAxis.DateTime := true;
            AChartAxis.LabelFormatDateTime := s;
          end else
          if s = 'General' then
            AChartAxis.LabelFormat := ''
          else
            AChartAxis.LabelFormat := s;
        end;
      'c:majorTickMark':
        AChartAxis.MajorTicks := ReadChartAxisTickMarks(ANode);
      'c:minorTickMark':
        AChartAxis.MinorTicks := ReadChartAxisTickMarks(ANode);
      'c:tickLblPos':
        ;
      'c:spPr':
        ReadChartLineProps(ANode.FirstChild, AChart, AChartAxis.AxisLine);
      'c:majorUnit':
        if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        begin
          AChartAxis.AutomaticMajorInterval := false;
          AChartAxis.MajorInterval := x;
        end;
      'c:minorUnit':
        if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        begin
          AChartAxis.AutomaticMinorInterval := false;
          AChartAxis.MinorInterval := x;
        end;
      'c:txPr':  // Axis labels
        begin
          x := 0;
          ReadChartTextProps(ANode, AChartAxis.LabelFont, x);
          if x = 1000 then  // default rotation
            x := 0;
          AChartAxis.LabelRotation := x;
        end;
      'c:crossAx':
        ;
      'c:crosses':
        case s of
          'min': AChartAxis.Position := capStart;
          'max': AChartAxis.Position := capEnd;
          'autoZero': AChartAxis.Position := capStart;
        end;
      'c:crossesAt':
        if TryStrToFloat(s, x, FPointSeparatorSettings) then
        begin
          AChartAxis.Position := capValue;
          AChartAxis.PositionValue := x;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartAxisScaling(ANode: TDOMNode;
  AChartAxis: TsChartAxis);
var
  nodeName, s: String;
  node: TDOMNode;
  x: Double;
begin
  if ANode = nil then
    exit;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:orientation':
        AChartAxis.Inverted := (s = 'maxMin');
      'c:max':
        if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        begin
          AChartAxis.AutomaticMax := false;
          AChartAxis.Max := x;
        end;
      'c:min':
        if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        begin
          AChartAxis.AutomaticMin := false;
          AChartAxis.Min := x;
        end;
      'c:logBase':
        if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        begin
          AChartAxis.Logarithmic := true;
          AChartAxis.LogBase := x;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

function TsSpreadOOXMLChartReader.ReadChartAxisTickMarks(ANode: TDOMNode): TsChartAxisTicks;
var
  s: String;
begin
  s := GetAttrValue(ANode, 'val');
  case s of
    'none': Result := [];
    'in': Result := [catInside];
    'out': Result := [catOutside];
    'cross': Result := [catInside, catOutside];
    else Result := [];
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartBarSeries(ANode: TDOMNode; AChart: TsChart);
var
  nodeName: String;
  savedNode: TDOMNode;
  s: String;
  n: Double;
  ser: TsBarSeries = nil;
begin
  if ANode = nil then
    exit;
  savedNode := ANode;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:ser':
        begin
          ser := TsBarSeries.Create(AChart);
          SetDefaultSeriesColor(ser);
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:barDir':
        case s of
          'col': AChart.RotatedAxes := false;
          'bar': AChart.RotatedAxes := true;
        end;
      'c:grouping':
        case s of
          'stacked': AChart.StackMode := csmStacked;
          'percentStacked': AChart.StackMode := csmStackedPercentage;
        end;
      'c:varyColors':
        ;
      'c:dLbls':
        s := '';
      'c:gapWidth':
        ;  // see BarSeries
      'c:overlap':
        ;  // see BarSeries
      'c:axId':
        ReadChartSeriesAxis(ANode, ser);
    end;
    ANode := ANode.NextSibling;
  end;

  if ser = nil then
    exit;

  // Make sure that the series exists
  ANode := savedNode;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:gapWidth':
        if TryStrToFloat(s, n, FPointSeparatorSettings) then
          AChart.BarGapWidthPercent := round(n);
      'c:overlap':
        if TryStrToFloat(s, n, FPointSeparatorSettings) then
          AChart.BarOverlapPercent := round(n);
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a bubble series and reads its parameters.

  @@param   ANode    Child of a <c:bubbleChart> node.
  @@param   AChart   Chart into which the series will be inserted.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartBubbleSeries(ANode: TDOMNode;
  AChart: TsChart);
var
  nodeName: String;
  s: String;
  ser: TsBubbleSeries;
  mode: TsBubbleSizeMode = bsmArea;
  scale: Integer = 100;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:ser':
        begin
          ser := TsBubbleSeries.Create(AChart);
          SetDefaultSeriesColor(ser);
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:sizeRepresents':
        if s = 'w' then mode := bsmRadius;
      'c:bubbleScale':
        scale := StrToIntDef(s, 100);
      'c:showNegBubbles': ;
      'c:varyColors':  ;
      'c:dLbls':
        ;
      'c:axId':
        ReadChartSeriesAxis(ANode, ser);
    end;
    ANode := ANode.NextSibling;
  end;
  if Assigned(ser) then
  begin
    ser.BubbleSizeMode := mode;
    ser.BubbleScale := scale / 100;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartColor(ANode: TDOMNode;
  var AColor: TsChartColor);

  function ColorAlias(AColorName: String): String;
  const
    DARK_MODE: Boolean = false;
  begin
    case AColorName of
      'tx1': if DARK_MODE then Result := 'lt1' else Result := 'dk1';
      'tx2': if DARK_MODE then Result := 'lt2' else Result := 'dk2';
      else   Result := AColorName;
    end;
  end;

var
  nodeName, s: String;
  n: Integer;
  child: TDOMNode;
  themeRGB: TsColor;
  lumMod: Single = 1.0;
  lumOff: Single = 0.0;
begin
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'a:schemeClr':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') then
          begin
            themeRGB := TsSpreadOOXMLReader(Reader).GetThemeColor(ColorAlias(s));
            if themeRGB <> scNotDefined then
            begin
              AColor.Color := themeRGB;
              child := ANode.FirstChild;
              while Assigned(child) do
              begin
                nodeName := child.NodeName;
                s := GetAttrValue(child, 'val');
                case nodeName of
                  'a:tint':
                    if TryStrToInt(s, n) then
                      AColor.Color := TintedColor(AColor.Color, n/FACTOR_MULTIPLIER);
                  'a:lumMod':     // luminance modulated
                    if TryStrToInt(s, n) then
                      lumMod := n/FACTOR_MULTIPLIER;
                  'a:lumOff':
                    if TryStrToInt(s, n) then
                      lumOff := n/FACTOR_MULTIPLIER;
                  'a:alpha':
                    if TryStrToInt(s, n) then
                      AColor.Transparency := 1.0 - n / FACTOR_MULTIPLIER;
                end;
                child := child.NextSibling;
              end;
            end;
            AColor.Color := LumModOff(AColor.Color, lumMod, lumOff);
          end;
        end;
      'a:srgbClr':
        begin
          s := GetAttrValue(ANode, 'val');
          if s <> '' then
            AColor.Color := HTMLColorStrToColor(s);
          child := ANode.FirstChild;
          while Assigned(child) do
          begin
            nodeName := child.NodeName;
            s := GetAttrValue(child, 'val');
            case nodeName of
              'a:alpha':
                if TryStrToInt(s, n) then
                  AColor.Transparency := 1.0 - n / FACTOR_MULTIPLIER;
            end;
            child := child.NextSibling;
          end;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

function TsSpreadOOXMLChartReader.ReadChartColorDef(ANode: TDOMNode;
  ADefault: TsChartColor): TsChartColor;
begin
  Result := ADefault;
  ReadChartColor(ANode, Result);
end;

procedure TsSpreadOOXMLChartReader.ReadChartGradientFillProps(ANode: TDOMNode;
  AChart: TsChart; AFill: TsChartFill);
var
  nodeName, s: String;
  value: Double;
  color: TsChartColor;
  child: TDOMNode;
  gradient: TsChartGradient;
begin
  if ANode = nil then
    exit;

  AFill.Style := cfsGradient;
  gradient := TsChartGradient.Create;   // Do not destroy gradient, it will be added to the chart.
  ANode := ANode.FirstChild;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'a:gsLst':
        begin
          child := ANode.FirstChild;
          while Assigned(child) do
          begin
            nodeName := child.NodeName;
            if nodeName = 'a:gs' then
            begin
              s := GetAttrValue(child, 'pos');
              value := StrToFloatDef(s, 0.0, FPointSeparatorSettings) / FACTOR_MULTIPLIER;
              color := ChartColor(scWhite);
              ReadChartColor(child.FirstChild, color);
              gradient.AddStep(value, color);
            end;
            child := child.NextSibling;
          end;
        end;
      'a:lin':
        begin
          gradient.Style := cgsLinear;
          s := GetAttrValue(ANode, 'ang');
          if TryStrToFloat(s, value, FPointSeparatorSettings) then
            gradient.Angle := -value / ANGLE_MULTIPLIER;     // xlsx CW, fps CCW
        end;
      'a:path':
        begin
          s := GetAttrValue(ANode, 'path');
          case s of
            'rect': gradient.Style := cgsRectangular;
            'circle': gradient.Style := cgsRadial;
            'shape': gradient.Style := cgsShape;
          end;
          child := ANode.FindNode('a:fillToRect');
          s := GetAttrValue(ANode, 'l');
          if TryStrToFloat(s, value, FPointSeparatorSettings) then
            gradient.CenterX := value / FACTOR_MULTIPLIER
          else
            gradient.CenterX := 0.0;
          s := GetAttrValue(aNode, 't');
          if tryStrToFloat(s, value, FPointSeparatorSettings) then
            gradient.CenterY := value / FACTOR_MULTIPLIER
          else
            gradient.CenterY := 0.0;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
  AFill.Gradient := AChart.Gradients.AddGradient('', gradient);
end;

procedure TsSpreadOOXMLChartReader.ReadChartHatchFillProps(ANode: TDOMNode;
  AChart: TsChart; AFill: TsChartFill);
var
  nodeName: String;
  hatch: String;
  color: TsChartColor;
begin
  AFill.Style := cfsSolidHatched;
  hatch := GetAttrValue(ANode, 'prst');

  ANode := ANode.FirstChild;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'a:fgClr':
        color := ReadChartColorDef(ANode.FirstChild, ChartColor(scBlack));
      'a:bgClr':
        AFill.Color := ReadChartColorDef(ANode.FirstChild, ChartColor(scWhite));
    end;
    ANode := ANode.NextSibling;
  end;

  case hatch of
    'pct5':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY5_PATTERN);
    'pct10':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY10_PATTERN);
    'pct20':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY20_PATTERN);
    'pct25':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY25_PATTERN);
    'pct30':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY30_PATTERN);
    'pct40':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY40_PATTERN);
    'pct50':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY50_PATTERN);
    'pct60':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY60_PATTERN);
    'pct70':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY70_PATTERN);
    'pct75':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY75_PATTERN);
    'pct80':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY80_PATTERN);
    'pct90':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_GRAY90_PATTERN);
    'dashDnDiag':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_DASH_DNDIAG_PATTERN);
    'dashUpDiag':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_DASH_UPDIAG_PATTERN);
    'dashHorz':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_DASH_HORZ_PATTERN);
    'dashVert':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_DASH_VERT_PATTERN);
    'smConfetti':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_SMALL_CONFETTI_PATTERN);
    'lgConfetti':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_LARGE_CONFETTI_PATTERN);
    'zigZag':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_ZIGZAG_PATTERN);
    'wave':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_WAVE_PATTERN);
    'diagBrick':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_DIAG_BRICK_PATTERN);
    'horzBrick':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_HORZ_BRICK_PATTERN);
    'weave':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_WEAVE_PATTERN);
    'plaid':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_PLAID_PATTERN);
    'divot':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_DIVOT_PATTERN);
    'dotGrid':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_DOT_GRID_PATTERN);
    'dotDmnd':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_DOT_DIAMOND_PATTERN);
    'shingle':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_SHINGLE_PATTERN);
    'trellis':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_TRELLIS_PATTERN);
    'sphere':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_SPHERE_PATTERN);
    'smCheck':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_SMALL_CHECKERBOARD_PATTERN);
    'lgCheck':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_LARGE_CHECKBOARD_PATTERN);
    'solidDmnd':
      AFill.Hatch := AChart.Hatches.AddDotHatch(hatch, color, 8, 8, OOXML_SOLID_DIAMOND_PATTERN);

    // The following patterns are line patterns to simplify interfacing with ODS.
    'ltDnDiag':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 1.0, 0.1, -45);
    'ltUpDiag':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 1.0, 0.1, +45);
    'dkDnDiag':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 1.0, 0.5, -45);
    'dkUpDiag':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 1.0, 0.5, +45);
    'wdDnDiag':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 2.0, 0.7, -45);
    'wdUpDiag':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 2.0, 0.7, +45);
    'ltHorz':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 1.0, 0.1, 0);
    'ltVert':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 1.0, 0.1, 90);
    'narVert':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 0.6, 0.3, 90);
    'narHorz':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 0.6, 0.3, 0);
    'dkHorz':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 1.0, 0.7, 0);
    'dkVert':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsSingle, color, 1.0, 0.7, 90);
    'smGrid':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsDouble, color, 1.0, 0.1, 0);
    'lgGrid':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsDouble, color, 2.0, 0.1, 0);
    'openDmnd':
      AFill.Hatch := AChart.Hatches.AddLineHatch(hatch, chsDouble, color, 2.0, 0.1, 45);
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartFillAndLineProps(ANode: TDOMNode;
  AChart: TsChart; AFill: TsChartFill; ALine: TsChartLine);
var
  nodeName, s: String;
  child1, child2: TDOMNode;
  n: Integer;
  value: Double;
  alpha: Double;
  gradient: TsChartGradient;
  color: TsColor;
begin
  if ANode = nil then
    exit;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      // Solid fill
      'a:solidFill':
        begin
          AFill.Style := cfsSolid;
          ReadChartColor(ANode.FirstChild, AFill.Color);
        end;

      // Gradient fill
      'a:gradFill':
        ReadChartGradientFillProps(ANode, AChart, AFill);

      // Hatched fill
      'a:pattFill':
        ReadChartHatchFillProps(ANode, AChart, AFill);

      // Image fill
      'a:blipFill':
        ReadChartImageFillProps(ANode, AChart, AFill);

      // Line style
      'a:ln':
        ReadChartLineProps(ANode, AChart, ALine);

      // Drawing effects (not supported ATM)
      'a:effectLst':
        ;
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Extracts the properties of a font

  @param  ANode  This is a "a:defRPr" node
  @param  AFont  Font to which the parameters are applied
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartFontProps(ANode: TDOMNode;
  AFont: TsFont);
{  Example:
      <a:defRPr sz="1000" b="1" i="0" u="none" strike="noStrike"
         kern="1200" spc="-1" baseline="0">
        <a:solidFill>
          <a:schemeClr val="tx1"/>
        </a:solidFill>
        <a:latin typeface="Arial"/>
        <a:ea typeface="+mn-ea"/>
        <a:cs typeface="+mn-cs"/>
      </a:defRPr> }
var
  node: TDOMNode;
  nodeName, s: String;
  x: Double;
  n: Integer;
begin
  if ANode = nil then
    exit;

  nodeName := ANode.NodeName;
  if not ((nodeName = 'a:defRPr') or (nodeName = 'a:rPr')) then
    exit;

  // Font size
  s := GetAttrValue(ANode, 'sz');
  if (s <> '') and TryStrToInt(s, n) then
    AFont.Size := n/100;

  // Font styles
  if GetAttrValue(ANode, 'b') = '1' then
    AFont.Style := AFont.Style + [fssBold];

  if GetAttrValue(ANode, 'i') = '1' then
    AFont.Style := AFont.Style + [fssItalic];

  if GetAttrValue(ANode, 'u') = '1' then
    AFont.Style := AFont.Style + [fssUnderline];

  s := GetAttrValue(ANode, 'strike');
  if (s <> '') and (s <> 'noStrike') then
    AFont.Style := AFont.Style + [fssStrikeOut];

  node := ANode.FirstChild;
  while Assigned(node) do
  begin
    nodeName := node.NodeName;
    case nodeName of
      // Font color
      'a:solidFill':
        AFont.Color := ReadChartColorDef(node.FirstChild, ChartColor(scBlack)).Color;

      // font name
      'a:latin':
        begin
          s := GetAttrValue(node, 'typeface');
          if s <> '' then
            AFont.FontName := s;
        end;

      // not supported
      'a:ea': ;
      'a:cs': ;
    end;
    node := node.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartImageFillProps(ANode: TDOMNode;
  AChart: TsChart; AFill: TsChartFill);
var
  nodeName: String;
  relID: String = '';
  widthFactor: Double = 1.0;
  heightFactor: Double = 1.0;
  imgWidthInches, imgHeightInches: Double;
  img: TFPCustomImage;
  sImg: TsChartImage;
  stream: TStream;
begin
  if ANode = nil then
    exit;

  AFill.Style := cfsImage;
  ANode := ANode.FirstChild;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'a:blip':
        relID := GetAttrValue(ANode, 'r:embed');
      'a:tile':
        begin
          widthFactor := StrToFloatDef(GetAttrValue(ANode, 'cx'), 100000, FPointSeparatorSettings) / 100000;
          heightFactor := StrToFloatDef(GetAttrValue(ANode, 'cy'), 100000, FPointSeparatorSettings) / 100000;
        end;
    end;
    ANode := ANode.NextSibling;
  end;

  if relID <> '' then
  begin
    stream := TNamedStreamList(FImages).FindStreamByName(relID);
    if stream <> nil then
    begin
      stream.Position := 0;
      GetImageInfo(stream, imgWidthInches, imgHeightInches);
      stream.Position := 0;
      img := TFPMemoryImage.Create(0, 0); // will be destroyed by the chart's images list.
      img.LoadFromStream(stream);
      AFill.Image := AChart.Images.AddImage('', img);
      sImg := AChart.Images[AFill.Image];
      sImg.Width := InToMM(imgWidthInches) * widthFactor;
      sImg.Height := InToMM(imgHeightInches) * heightFactor;
    end;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartImages(AStream: TStream;
  AChart: TsChart; ARelsList: TFPList);
var
  i: Integer;
  rel: TXlsxRelationship;
  img: TFPCustomImage;
  imgFileName: string;
  namedStreamItem: TNamedStreamItem;
  unzip: TStreamUnzipper;
begin
  FImages.Clear;

  for i := 0 to ARelsList.Count-1 do
  begin
    rel := TXlsxRelationshipList(ARelsList)[i];
    if rel.Schema = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' then
    begin
      imgFileName := MakeXLPath(rel.Target);
      if imgFileName = '' then
        Continue;
      unzip := TStreamUnzipper.Create(AStream);
      try
        unzip.Examine;
        namedStreamItem := TNamedStreamItem.Create;
        namedStreamItem.Name := rel.RelID;
        namedStreamItem.Stream := TMemoryStream.Create;
        unzip.UnzipFile(imgFileName, namedStreamItem.Stream);
        namedStreamItem.Stream.Position := 0;
        FImages.Add(namedStreamItem);
      finally
        unzip.Free;
      end;
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the individual data point styles of a series.

  @param  ANode    First child of the <c:dPt> node
  @param  ASeries  Series to which these data points belong
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartSeriesDataPointStyles(ANode: TDOMNode;
  ASeries: TsChartSeries);
var
  nodename, s: String;
  fill: TsChartFill;
  line: TsChartLine;
  idx: Integer;
  explosion: Integer = 0;
begin
  if ANode = nil then
    exit;

  fill := TsChartFill.Create;
  line := TsChartLine.Create;
  try
    while Assigned(ANode) do
    begin
      nodeName := ANode.NodeName;
      s := GetAttrValue(ANode, 'val');
      case nodeName of
        'c:idx':
          if not TryStrToInt(s, idx) then  // This is an error condition!
            exit;
        'c:spPr':
          ReadChartFillAndLineProps(ANode.FirstChild, ASeries.Chart, fill, line);
        'c:explosion':
          explosion := StrToIntDef(s, 0);
      end;
      ANode := ANode.NextSibling;
    end;
    ASeries.DataPointStyles.AddFillAndLine(idx, fill, line, explosion);   // fill and line are copied here
  finally
    line.Free;
    fill.Free;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Extracts the legend properties

  @param  ANode         This is the "c:legend" node
  @param  AChartLegend  Legend to which the values are applied
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartLegend(ANode: TDOMNode;
  AChartLegend: TsChartLegend);
var
  nodeName, s: String;
  dummy: Single;
  lp: TsChartLegendPosition;
begin
  if ANode = nil then
    exit;

  AChartLegend.Visible := true;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      // Legend position with respect to the plot area
      'c:legendPos':
        begin
          s := GetAttrValue(ANode, 'val');
          for lp in TsChartLegendPosition do
            if s = LEGEND_POS[lp] then
            begin
              AChartLegend.Position := lp;
              break;
            end;
        end;

      // Formatting of individual legend items, not supported
      'c:legendEntry':
        ;

      // Overlap with plot area
      'c:overlay':
        begin
          s := GetAttrValue(ANode, 'val');
          AChartLegend.canOverlapPlotArea := (s = '1');
        end;

      // Background and border
      'c:spPr':
        ReadChartFillAndLineProps(ANode.FirstChild, AChartLegend.Chart, AChartLegend.Background, AChartLegend.Border);

      // Legend font
      'c:txPr':
        begin
          dummy := 0;
          ReadChartTextProps(ANode, AChartLegend.Font, dummy);
          //AChartLegend.RotationAngle := dummy;  // we do not support rotated text in legend
        end;
    end;

    ANode := ANode.NextSibling;
  end;
end;

{ Example:   (parent node is "c:spPr")
        <a:ln w="10800" cap="rnd">
          <a:solidFill>
            <a:srgbClr val="C0C0C0"/>
          </a:solidFill>
          <a:custDash>
            <a:ds d="300000" sp="150000"/>
          </a:custDash>
          <a:round/>
        </a:ln>  }
procedure TsSpreadOOXMLChartReader.ReadChartLineProps(ANode: TDOMNode;
  AChart: TsChart; AChartLine: TsChartLine);
var
  child, child2: TDOMNode;
  nodeName, s: String;
  w, d, sp: Int64;
  dMM, spMM: Double;
  noLine: Boolean;
begin
  if ANode = nil then
    exit;
  noLine := false;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'a:ln':
        begin
          s := GetAttrValue(ANode, 'w');
          if (s <> '') and TryStrToInt64(s, w) then
            AChartLine.Width := ptsToMM(w/PTS_MULTIPLIER);
          child := ANode.FirstChild;
          while Assigned(child) do
          begin
            nodeName := child.NodeName;
            case nodeName of
              'a:noFill':
                noLine := true;
              'a:solidFill':
                begin
                  AChartLine.Color := ReadChartColorDef(child.FirstChild, ChartColor(scBlack));
                  AChartLine.Style := clsSolid;
                end;
              'a:prstDash':
                begin
                  s := GetAttrValue(child, 'val');
                  case s of
                    'solid': AChartLine.Style := clsSolid;
                    'dot', 'sysDot': AChartLine.Style := clsDot;
                    'dash', 'sysDash': AChartLine.Style := clsDash;
                    'dashDot': AChartLine.Style := clsDashDot;
                    'lgDash': AChartLine.Style := clsLongDash;
                    'lgDashDot': AChartLine.Style := clsLongDashDot;
                    'lgDashDotDot': AChartLine.Style := clsLongDashDotDot;
                  end;
                end;
              'a:custDash':
                begin
                  child2 := child.FindNode('a:ds');
                  if Assigned(child2) then
                  begin
                    s := GetAttrValue(child2, 'd');
                    if TryStrToInt64(s, d) then
                    begin
                      s := GetAttrValue(child2, 'sp');
                      if TryStrToInt64(s, sp) then
                      begin
                        dMM := PtsToMM(d / PTS_MULTIPLIER);
                        spMM := PtsToMM(sp / PTS_MULTIPLIER);
                        AChartLine.Style := AChart.LineStyles.Add('', dMM, 1, 0, 0, (dMM+spMM), false);
                      end;
                    end;
                  end;
                end;
            end;
            child := child.NextSibling;
          end;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
  if noLine then
    AChartLine.Style := clsNoLine;
end;

{@@ ----------------------------------------------------------------------------
  Creates a line series and reads its properties
  In contrast to a scatter series, a line series has equidistance x values!

  @@param   ANode   First child of the <c:lineChart> node
  @@param   AChart  Chart to which the series will be attached
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartLineSeries(ANode: TDOMNode; AChart: TsChart);
var
  nodeName: String;
  s: String;
  n: LongInt;
  node: TDOMNode;
  ser: TsLineSeries = nil;
begin
  if ANode = nil then
    exit;
  node := ANode;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:ser':
        begin
          ser := TsLineSeries.Create(AChart);
          SetDefaultSeriesColor(ser);
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:grouping':
        case s of
          'stacked': AChart.StackMode := csmStacked;
          'percentStacked': AChart.StackMode := csmStackedPercentage;
        end;
      'c:varyColors':
        ;
      'c:dLbls':
        ;
      {
      'c:axId':
        ReadChartSeriesAxis(ANode, ser);
        }
    end;
    ANode := ANode.NextSibling;
  end;

  if ser = nil then
    exit;

  ANode := node;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:axId':
        ReadChartSeriesAxis(ANode, ser);
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a pie series and reads its properties

  @@param   ANode   First child of the <c:pieChart> node
  @@param   AChart  Chart to which the series will be attached
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartPieSeries(ANode: TDOMNode;
  AChart: TsChart; RingMode: Boolean);
var
  nodeName, s: String;
  ser: TsPieSeries;
  x: Double;
  ringRadius: Integer = 50;
  startAngle: Integer = 90;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:ser':
        begin
          ser := TsPieSeries.Create(AChart);
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:firstSliceAng':
        if TryStrToFloat(s, x, FPointSeparatorSettings) then
          startAngle := round(x) + 90;
      'c:holeSize':
        if RingMode then
          if TryStrToFloat(s, x, FPointSeparatorSettings) then
            ringRadius := round(x);
    end;
    ANode := ANode.NextSibling;
  end;

  if ser <> nil then
  begin
    TsPieSeries(ser).StartAngle := startAngle;
    if RingMode then
      TsPieSeries(ser).InnerRadiusPercent := ringRadius;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the properties of the series marker

  @@param  ANode    Points to the <c:marker> subnode of <c:ser> node, or to the first child of the <c:marker> node.
  @@param  ASeries  Instance of the TsCustomLineSeries created by ReadChartLineSeries
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartSeriesMarker(ANode: TDOMNode; ASeries: TsCustomLineSeries);
var
  nodeName, s: String;
  n: Integer;
begin
  if ANode = nil then
    exit;

  nodeName := ANode.NodeName;
  if nodeName = 'c:marker' then
    ANode := ANode.FirstChild;

  TsOpenedCustomLineSeries(ASeries).ShowSymbols := true;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:symbol':
        with TsOpenedCustomLineSeries(ASeries) do
          case s of
            'none': ShowSymbols := false;
            'square': Symbol := cssRect;
            'circle': Symbol := cssCircle;
            'diamond': Symbol := cssDiamond;
            'triangle': Symbol := cssTriangle;
            'star': Symbol := cssStar;
            'x': Symbol := cssX;
            'plus': Symbol := cssPlus;
            'dash': Symbol := cssDash;
            'dot': Symbol := cssDot;
            'picture': Symbol := cssAsterisk;
            // to do: read following blipFill node to determine the
            // bitmap to be used for symbol
          else
            Symbol := cssAsterisk;
        end;

      'c:size':
        if TryStrToInt(s, n) then
          with TsOpenedCustomLineSeries(ASeries) do
          begin
            SymbolWidth := PtsToMM(n);
            SymbolHeight := SymbolWidth;
          end;

      'c:spPr':
        with TsOpenedCustomLineSeries(ASeries) do
          ReadChartFillAndLineProps(ANode.FirstChild, Chart, SymbolFill, SymbolBorder);
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the plot area node and its subnodes.

  @@param   ANode   Is the first subnode of the <c:plotArea> node
  @@param   AChart  Chart to which the found property values are assigned
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartPlotArea(ANode: TDOMNode;
  AChart: TsChart);
var
  nodeName: String;
  workNode: TDOMNode;
  isScatterChart: Boolean = false;
  catAxCounter: Integer = 0;
  dateAxCounter: Integer = 0;
  valAxCounter: Integer = 0;
begin
  if ANode = nil then
    exit;

  // We need to know first whether the chart is a scatterchart because the way
  // how the axes are assigned is special here.
  isScatterChart := false;
  workNode := ANode;
  while Assigned(workNode) do
  begin
    nodeName := workNode.NodeName;
    if (nodeName = 'c:scatterChart') or (nodeName = 'c:bubbleChart') then
    begin
      isScatterChart := true;
      break;
    end;
    workNode := workNode.NextSibling;
  end;

  AChart.PlotArea.Background.Color := ChartColor(scWhite);
  AChart.PlotArea.Background.Style := cfsSolid;
  SetAxisDefaults(AChart.XAxis);
  SetAxisDefaults(AChart.YAxis);
  SetAxisDefaults(AChart.X2Axis);
  SetAxisDefaults(AChart.Y2Axis);

  // We need the axis IDs before creating the series.
  workNode := ANode;
  while Assigned(workNode) do
  begin
    nodeName := workNode.NodeName;
    case nodeName of
      'c:catAx':
        begin
          case catAxCounter of
            0: ReadChartAxis(workNode.FirstChild, AChart, AChart.XAxis, FXAxisID, FXAxisDelete);
            1: ReadChartAxis(workNode.FirstChild, AChart, AChart.X2Axis, FX2AxisID, FX2AxisDelete);
          end;
          inc(catAxCounter);
        end;
      'c:dateAx':
        begin
          case dateAxCounter of
            0: begin
                 ReadChartAxis(workNode.FirstChild, AChart, AChart.XAxis, FXAxisID, FXAxisDelete);
                 AChart.XAxis.DateTime := true;
                 if AChart.XAxis.LabelFormatDateTime = '' then
                   AChart.XAxis.LabelFormatDateTime := 'yyyy-mm';
               end;
            1: begin
                 ReadChartAxis(workNode.FirstChild, AChart, AChart.X2Axis, FX2AxisID, FX2AxisDelete);
                 AChart.X2Axis.DateTime := true;
                 if AChart.X2Axis.LabelFormatDateTime = '' then
                   AChart.X2Axis.LabelFormatDateTime := 'yyyy-mm';
               end;
          end;
          inc(dateAxCounter);
        end;
      'c:valAx':
        begin
          if isScatterChart then
          begin
            { Order of value axes in the <c:plotArea> node of a scatterchart:
                #1: primary x axis
                #2: primary y axis
                #3: secondary y axis
                #4: secondary x axis }
            case valAxCounter of
              0: ReadChartAxis(workNode.FirstChild, AChart, AChart.XAxis, FXAxisID, FXAxisDelete);
              1: ReadChartAxis(workNode.FirstChild, AChart, AChart.YAxis, FYAxisID, FYAxisDelete);
              2: ReadChartAxis(workNode.FirstChild, AChart, AChart.Y2Axis, FY2AxisID, FY2AxisDelete);
              3: ReadChartAxis(workNode.FirstChild, AChart, AChart.X2Axis, FX2AxisID, FX2AxisDelete);
            end;
          end else
          begin
            case valAxCounter of
              0: ReadChartAxis(workNode.FirstChild, AChart, AChart.YAxis, FYAxisID,  FYAxisDelete);
              1: ReadChartAxis(workNode.FirstChild, AChart, AChart.Y2Axis, FY2AxisID, FY2AxisDelete);
            end;
          end;
          inc(valAxCounter);
        end;
      'c:spPr':
        ReadChartFillAndLineProps(workNode.FirstChild, AChart, AChart.PlotArea.Background, AChart.PlotArea.Border);
    end;
    workNode := workNode.NextSibling;
  end;

  if FX2AxisDelete then
  begin
    // Force using only a single x axis in this case.
    FX2AxisID := FXAxisID;
    AChart.X2Axis.Visible := false;
  end;
  if FY2AxisDelete then
  begin
    FY2AxisID := FYAxisID;
    AChart.Y2Axis.Visible := false;
  end;

  // The RotatedAxes option handles all axis rotations by itself. Therefore we
  // must reset the Axis.Alignment back to the normal settings, otherwise
  // rotation will not be correct.
  if AChart.XAxis.Alignment = caaLeft then
  begin
    AChart.XAxis.Alignment := caaBottom;
    AChart.YAxis.Alignment := caaLeft;
    AChart.X2Axis.Alignment := caaTop;
    AChart.Y2Axis.Alignment := caaRight;
    AChart.RotatedAxes := true;        // Note: this rotates the axis titles!
  end;

  workNode := ANode;
  while Assigned(workNode) do
  begin
    nodeName := workNode.NodeName;
    case nodeName of
      'c:areaChart':
        ReadChartAreaSeries(workNode.FirstChild, AChart);
      'c:barChart':
        ReadChartBarSeries(workNode.FirstChild, AChart);
      'c:bubbleChart':
        ReadChartBubbleSeries(workNode.FirstChild, AChart);
      'c:lineChart':
        ReadChartLineSeries(workNode.FirstChild, AChart);
      'c:pieChart', 'c:doughnutChart':
        ReadChartPieSeries(workNode.FirstChild, AChart, nodeName = 'c:doughnutChart');
      'c:radarChart':
        ReadChartRadarSeries(workNode.FirstChild, AChart);
      'c:scatterChart':
        ReadChartScatterSeries(workNode.FirstChild, AChart);
      'c:stockChart':
        ReadChartStockSeries(workNode.FirstChild, AChart);
      'c:spPr':
        ReadChartFillAndLineProps(workNode.FirstChild, AChart, AChart.PlotArea.Background, AChart.PlotArea.Border);
    end;
    workNode := workNode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a radar series and reads its parameters

  @param  ANode   Child of a <c:radarChart> node.
  @param  AChart  Chart into which the series will be inserted.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartRadarSeries(ANode: TDOMNode; AChart: TsChart);
var
  node: TDOMNode;
  nodeName, s: String;
  ser: TsRadarSeries;
  filled: Boolean = false;
begin
  if ANode = nil then
    exit;

  // At first, we need the value of c:radarStyle because it determines the
  // series class to be created.
  node := ANode;
  while Assigned(node) do
  begin
    nodeName := node.NodeName;
    case nodeName of
      'c:radarStyle':
        begin
          s := GetAttrValue(node, 'val');
          filled := s = 'filled';
        end;
    end;
    node := node.NextSibling;
  end;

  // Search the series node. Then create the series and read its properties
  // from the subnodes.
  node := ANode;
  while Assigned(node) do
  begin
    nodeName := node.NodeName;
    case nodeName of
      'c:ser':
        begin
          if filled then
            ser := TsFilledRadarSeries.Create(AChart)
          else
            ser := TsRadarSeries.Create(AChart);
          SetDefaultSeriesColor(ser);
          ReadChartSeriesProps(node.FirstChild, ser);
        end;
    end;
    node := node.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Creates a scatter series and reads its parameters.
  A scatter series is a "line series" with irregularly spaced x values.

  @@param   ANode    Child of a <c:scatterChart> node.
  @@param   AChart   Chart into which the series will be inserted.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartScatterSeries(ANode: TDOMNode;
  AChart: TsChart);
var
  nodeName: String;
  s: String;
  ser: TsScatterSeries;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:ser':
        begin
          ser := TsScatterSeries.Create(AChart);
          SetDefaultSeriesColor(ser);
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:scatterStyle':
        begin
          {
          s := GetAttrValue(ANode, 'val');
          if (s = 'smoothMarker') then
            smooth := true;
            }
        end;
      'c:varyColors':
        ;
      'c:dLbls':
        ;
      'c:axId':
        ReadChartSeriesAxis(ANode, ser);
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads and assigns the axes used by the given series.

  @@param ANode    Node <c:axId>, a child of the <c:XXXXchart> node (XXXX = scatter, bar, line, etc)
  @@param ASeries  Series to which the axis is assigned.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartSeriesAxis(ANode: TDOMNode;
  ASeries: TsChartSeries);
var
  nodeName, s: String;
  n: Integer;
begin
  if ANode = nil then
    exit;

  nodeName := ANode.NodeName;
  if nodeName <> 'c:axId' then
    exit;

  s := GetAttrValue(ANode, 'val');
  if (s = '') or not TryStrToInt(s, n) then
    exit;

  if n = FXAxisID then
  begin
    ASeries.XAxis := calPrimary;
    ASeries.Chart.XAxis.Visible := not FXAxisDelete;
  end
  else if n = FYAxisID then
  begin
    ASeries.YAxis := calPrimary;
    ASeries.Chart.YAxis.Visible := not FYAxisDelete;
  end
  else if n = FX2AxisID then
  begin
    ASeries.XAxis := calSecondary;
    ASeries.Chart.X2Axis.Visible := not FX2AxisDelete;
  end
  else if n = FY2AxisID then
  begin
    ASeries.YAxis := calSecondary;
    ASeries.Chart.Y2Axis.Visible := not FY2AxisDelete;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the error bar parameters of a series.

  @@param ANode    Is the first child of the <c:errBars> subnode of <c:ser>.
  @@param ASeries  Series to which the error bars belong.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartSeriesErrorBars(ANode: TDOMNode;
  ASeries: TsChartSeries);
var
  workbook: TsWorkbook;
  nodeName, s: String;
  node: TDOMNode;
  val: Double;
  errorBars: TsChartErrorBars = nil;
  part: String = '';
  fmt: String = '';
begin
  if ANode = nil then
    exit;

  workbook := TsSpreadOOXMLReader(Reader).Workbook as TsWorkbook;

  // We must first find out whether the node is for x or y error bars and
  // whether it is for positive, negative or both error parts.
  node := ANode;
  while Assigned(node) do
  begin
    nodeName := node.NodeName;
    s := GetAttrValue(node, 'val');
    case nodeName of
      'c:errDir':
        begin
          case s of
            'x': errorBars := ASeries.XErrorBars;
            'y': errorBars := ASeries.YErrorBars;
          end;
        end;
      'c:errBarType':
        part := s;
    end;
    if (errorBars <> nil) and (part <> '') then
      break;
    node := node.NextSibling;
  end;

  errorBars.ShowPos := (part = 'both') or (part = 'plus');
  errorBars.ShowNeg := (part = 'both') or (part = 'minus');

  node := ANode;
  while Assigned(node) do
  begin
    nodeName := node.NodeName;
    s := GetAttrValue(node, 'val');
    case nodeName of
      'c:errValType':
        case s of
          'fixedVal':
            errorBars.Kind := cebkConstant;
          'percentage':
            errorBars.Kind := cebkPercentage;
          'cust':
            errorBars.Kind := cebkCellRange;
          'stdDev':
            begin
              errorBars.Visible := false;
              workbook.AddErrorMsg('Error bar kind "stdDev" not supported');
            end;
          'stdErr':
            begin
              errorBars.Visible := false;
              workbook.AddErrorMsg('Error bar kind "stdErr" not supported.');
            end;
        end;
      'c:val':
        if (s <> '') and TryStrToFloat(s, val, FPointSeparatorSettings) then
          case part of
            'both':
              begin
                errorBars.ValuePos := val;
                errorBars.ValueNeg := val;
              end;
            'plus':
              errorBars.ValuePos := val;
            'minus':
              errorBars.ValueNeg := val;
          end;
      'c:plus':
        ReadChartSeriesRange(node.FirstChild, errorBars.RangePos, fmt);
      'c:minus':
        ReadChartSeriesRange(node.FirstChild, errorBars.RangeNeg, fmt);
      'c:spPr':
        ReadChartLineProps(node.FirstChild, ASeries.Chart, errorBars.Line);
      'c:noEndCap':
        errorBars.ShowEndCap := (s <> '1');
    end;
    node := node.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the labels assigned to the series data points.

  @@param ANode    Is the first child of the <c:dLbls> subnode of <c:ser>.
  @@param ASeries  Series to which the labels belong.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartSeriesLabels(ANode: TDOMNode;
  ASeries: TsChartSeries);
var
  nodeName, s: String;
  child1, child2, child3: TDOMNode;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:numFmt':
        begin
          s := GetAttrValue(ANode, 'formatCode');
          if s <> '' then
            ASeries.LabelFormat := s;
        end;
      'c:spPr':
        ReadChartFillAndLineProps(ANode.FirstChild, ASeries.Chart, ASeries.LabelBackground, ASeries.LabelBorder);
      'c:txPr':
        begin
          child1 := ANode.FindNode('a:p');
          if Assigned(child1) then
          begin
            child2 := child1.FirstChild;
            while Assigned(child2) do
            begin
              nodeName := child2.NodeName;
              if nodeName = 'a:pPr' then
              begin
                child3 := child2.FindNode('a:defRPr');
                if Assigned(child3) then
                  ReadChartFontProps(child3, ASeries.LabelFont);
              end;
              child2 := child2.NextSibling;
            end;
          end;
        end;
      'c:dlblPos':
        case s of
          '': ASeries.LabelPosition := lpOutside;
          'ctr': ASeries.LabelPosition := lpCenter;
          'inBase': ASeries.LabelPosition := lpNearOrigin;
          'inEnd': ASeries.LabelPosition := lpInside;
        end;
      'c:showLegendKey':
        if (s = '1') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlSymbol];
      'c:showVal':
        if (s = '1') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlValue];
      'c:showCatName':
        if (s = '1') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlCategory];
      'c:showSerName':
        if (s = '1') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlSeriesName];
      'c:showPercent':
        if (s = '1') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlPercentage];
      'c:showBubbleSize':
        if (s = '1') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlValue];
      'c:showLeaderLines':
        if (s = '1') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlLeaderLines];
      'c:extLst':
        begin
          child1 := ANode.FirstChild;
          while Assigned(child1) do
          begin
            nodeName := child1.NodeName;
            if nodeName = 'c:ext' then
            begin
              child2 := child1.FirstChild;
              while Assigned(child2) do
              begin
                nodeName := child2.NodeName;
                if nodeName = 'c15:spPr' then
                begin
                  child3 := child2.FindNode('a:prstGeom');
                  if Assigned(child3) then
                  begin
                    s := GetAttrValue(child3, 'prst');
                    case s of
                      'rect': ASeries.DataLabelCalloutShape := lcsRectangle;
                      'roundRect': ASeries.DataLabelCalloutShape := lcsRoundRect;
                      'ellipse': ASeries.DataLabelCalloutShape := lcsEllipse;
                      'rightArrowCallout': ASeries.DataLabelCalloutShape := lcsRightArrow;
                      'downArrowCallout': ASeries.DataLabelCalloutShape := lcsDownArrow;
                      'leftArrowCallout': ASeries.DataLabelCalloutShape := lcsLeftArrow;
                      'upArrowCallout': ASeries.DataLabelCalloutShape := lcsUpArrow;
                      'wedgeRectCallout': ASeries.DataLabelCalloutShape := lcsRectangleWedge;
                      'wedgeRoundRectCallout': ASeries.DataLabelCalloutShape := lcsRoundRectWedge;
                      'wedgeEllipseCallout': ASeries.DataLabelCalloutShape := lcsEllipseWedge;
                      else ASeries.DataLabelCalloutShape := lcsRectangle;
                      {
                      'borderCallout1': ;
                      'borderCallout2': ;
                      'accentCallout1': ;
                      'accentCallout2': ;
                      }
                    end;
                  end;
                end;
                child2 := child2.NextSibling;
              end;
            end;
            child1 := child1.NextSibling;
          end;
        end;
      'c:separator':
        begin
          s := GetNodeValue(ANode);
          if (s = #10) or (s = #13#10) or (s = #13) then s := LineEnding;
          ASeries.LabelSeparator := s;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartSeriesProps(ANode: TDOMNode; ASeries: TsChartSeries);
var
  nodeName, fmt, s: String;
  n: Integer;
  idx: Integer;
  ax: TsChartAxis;
  smooth: Boolean = false;
begin
  if ANode = nil then
    exit;
  idx := 0;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:idx': ;
      'c:order':
        if TryStrToInt(s, n) then
          ASeries.Order := n;
      'c:tx':
        ReadChartSeriesTitle(ANode.FirstChild, ASeries);
      'c:cat':       // Category axis
        begin
          ax := ASeries.GetXAxis;
          ReadChartSeriesRange(ANode.FirstChild, ASeries.LabelRange, fmt);
          if ax.DateTime then ASeries.XRange.CopyFrom(ASeries.LabelRange);
          if IsDateTimeFormat(fmt) then
          begin
            if ax.LabelFormatDateTime = '' then
              ax.LabelFormatDateTime := fmt;
          end else
          begin
            if ax.LabelFormat = '' then
              ax.LabelFormat := fmt;
          end;
        end;
      'c:xVal':   // x value axis
        ReadChartSeriesRange(ANode.FirstChild, ASeries.XRange, fmt);
      'c:val',    // y value axis in categorized series
      'c:yVal':   // y value axis
        if ASeries.YRange.IsEmpty then  // TcStockSeries already has read the y range...
        begin
          ReadChartSeriesRange(ANode.FirstChild, ASeries.YRange, fmt);
          ASeries.LabelFormat := fmt;
        end;
      'c:bubbleSize':
        if ASeries is TsBubbleSeries then
          ReadChartSeriesRange(ANode.FirstChild, TsBubbleSeries(ASeries).BubbleRange, fmt);
      'c:bubble3D':
        ;
      'c:spPr':
        ReadChartFillAndLineProps(ANode.FirstChild, ASeries.Chart, ASeries.Fill, ASeries.Line);
      'c:marker':
        if ASeries is TsCustomLineSeries then
          ReadChartSeriesMarker(ANode.FirstChild, TsCustomLineSeries(ASeries));
      'c:dLbls':
        ReadChartSeriesLabels(ANode.FirstChild, ASeries);
      'c:dPt':
        ReadChartSeriesDataPointStyles(ANode.FirstChild, ASeries);
      'c:trendline':
        ReadChartSeriesTrendLine(ANode.FirstChild, ASeries);
      'c:errBars':
        ReadChartSeriesErrorBars(ANode.FirstChild, ASeries);
      'c:smooth':
        smooth := (s <> '0');
      'c:invertIfNegative':
        ;
      'c:extLst':
        ;
    end;
    ANode := ANode.NextSibling;
  end;

  if (ASeries is TsCustomLineSeries) and smooth then
    TsOpenedCustomLineSeries(ASeries).Interpolation := ciCubicSpline; //ciBSpline;
end;

{@@ ----------------------------------------------------------------------------
  Reads the cell range for a series.

  @@param  ANode   First child of a <c:val>, <c:yval> or <c:cat> node below <c:ser>.
  @@param  ARange  Cell range to which the range parameters will be assigned.
  @@param  AFormat Numberformat string
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartSeriesRange(ANode: TDOMNode;
  ARange: TsChartRange; var AFormat: String);
var
  node, child: TDomNode;
  nodeName, s: String;
  sheet1, sheet2: String;
  r1, c1, r2, c2: Cardinal;
  flags: TsRelFlags;
begin
  AFormat := '';
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    if (nodeName = 'c:strRef') or (nodeName = 'c:numRef') then
    begin
      node := ANode.FirstChild;
      while Assigned(node) do
      begin
        nodeName := node.NodeName;
        case nodeName of
          'c:f':
            begin
              s := GetNodeValue(node);
              if ParseCellRangeString(s, sheet1, sheet2, r1, c1, r2, c2, flags) then
              begin
                if sheet2 = '' then sheet2 := sheet1;
                ARange.Sheet1 := sheet1;
                ARange.Sheet2 := sheet2;
                ARange.Row1 := r1;
                ARange.Col1 := c1;
                ARange.Row2 := r2;
                ARange.Col2 := c2;
              end;
            end;
          'c:numCache':
            begin
              child := node.FindNode('c:formatCode');
              if Assigned(child) then
                AFormat := GetNodeValue(child);
            end;
        end;
        node := node.NextSibling;
      end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartSeriesTitle(ANode: TDOMNode; ASeries: TsChartSeries);
var
  nodeName, s: String;
  sheet: String;
  r, c: Cardinal;
begin
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    if (nodeName = 'c:strRef') then
    begin
      ANode := ANode.FindNode('c:f');
      if ANode <> nil then
      begin
        s := GetNodeValue(ANode);
        if ParseSheetCellString(s, sheet, r, c) then
        begin
          ASeries.SetTitleAddr(sheet, r, c);
          exit;
        end;
      end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the trend-line fitted to a series (which has SupportsRegression true).

  @@param ANode    Is the first child of the <c:trendline> subnode of <c:ser>.
  @@param ASeries  Series to which the fit was applied.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartSeriesTrendLine(ANode: TDOMNode;
  ASeries: TsChartSeries);
var
  nodeName, s: String;
  trendline: TsChartTrendline;
  child: TDOMNode;
  n: Integer;
  x: Double;
begin
  if ANode = nil then
    exit;
  if not ASeries.SupportsTrendline then
    exit;

  trendline := TsOpenedTrendlineSeries(ASeries).Trendline;

  while Assigned(ANode) do begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:name':
        trendline.Title := GetNodeValue(ANode);
      'c:spPr':
        ReadChartLineProps(ANode.FirstChild, ASeries.Chart, trendline.Line);
      'c:trendlineType':
        case s of
          'exp': trendline.TrendlineType := tltExponential;
          'linear': trendline.TrendlineType := tltLinear;
          'log': trendline.TrendlineType := tltNone;  // rtLog, but not supported.
          'movingAvg': trendline.TrendlineType := tltNone;  // rtMovingAvg, but not supported.
          'poly': trendline.TrendlineType := tltPolynomial;
          'power': trendline.TrendlineType := tltPower;
        end;
      'c:order':
        if (s <> '') and TryStrToInt(s, n) then
          trendline.PolynomialDegree := n;
      'c:period':
        if (s <> '') and TryStrToInt(s, n) then ;  // not supported
          // trendline.MovingAvgPeriod := n;
      'c:forward', 'c:backward':
        if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
          case nodeName of
            'c:forward': trendline.ExtrapolateForwardBy := x;
            'c:backward': trendline.ExtrapolateBackwardBy := x;
          end;
      'c:intercept':
        if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        begin
          trendline.YInterceptValue := x;
          trendline.ForceYIntercept := true;
        end;
      'c:dispRSqr':
        if s = '1' then
          trendline.DisplayRSquare := true;
      'c:dispEq':
        if s = '1' then
          trendline.DisplayEquation := true;
      'c:trendlineLbl':
        begin
          child := ANode.FirstChild;
          while child <> nil do
          begin
            nodeName := child.NodeName;
            case nodeName of
              'c:numFmt':
                begin
                  s := GetAttrValue(child, 'formatCode');
                  trendline.Equation.NumberFormat := s;
                end;
            end;
            child := child.NextSibling;
          end;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartStockSeries(ANode: TDOMNode;
  AChart: TsChart);
var
  ser: TsStockSeries;
  nodeName, fmt: String;
  sernode, child: TDOMNode;
begin
  if ANode = nil then
    exit;

  ser := TsStockSeries.Create(AChart);
  AChart.YAxis.AutomaticMin := false;
  AChart.Y2Axis.AutomaticMin := false;

  // Collecting the ranges which make up the stock series. Note that in Excel's
  // HLC series there are three ranges for high, low and close, while for the
  // OHLC series (candle stick) the first series is for "open". Therefore, we
  // iterate the siblings of ANode from the end.
  serNode := ANode.ParentNode.LastChild;
  while Assigned(serNode) do
  begin
    nodeName := serNode.NodeName;
    case nodeName of
      'c:ser':
        begin
          child := serNode.FirstChild;
          while Assigned(child) do
          begin
            nodeName := child.NodeName;
            case nodeName of
              {
              'c:cat':  // is read by ReadChartSeriesProps
                ReadChartSeriesRange(child.FirstChild, ser.LabelRange, fmt);
              }
              'c:val':
                if ser.CloseRange.IsEmpty then
                  ReadChartSeriesRange(child.FirstChild, ser.CloseRange, fmt)
                else if ser.LowRange.IsEmpty then
                  ReadChartSeriesRange(child.FirstChild, ser.LowRange, fmt)
                else if ser.HighRange.IsEmpty then
                  ReadChartSeriesRange(child.FirstChild, ser.HighRange, fmt)
                else if ser.OpenRange.IsEmpty then
                begin
                  ReadChartSeriesRange(child.FirstChild, ser.OpenRange, fmt);
                  ser.CandleStick := true;
                end;
            end;
            child := child.NextSibling;
          end;
        end;
    end;
    serNode := serNode.PreviousSibling;  // we must run backward
  end;

  serNode := ANode;
  while (serNode <> nil) do
  begin
    nodeName := serNode.NodeName;
    if nodeName = 'c:ser' then
      ReadChartSeriesProps(serNode.FirstChild, ser);
    serNode := serNode.NextSibling;
  end;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:hiLowLines':
        begin
          child := ANode.FindNode('c:spPr');
          if Assigned(child) then
            ReadChartLineProps(child.FirstChild, AChart, ser.Rangeline);
        end;
      'c:upDownBars':
        ReadChartStockSeriesUpDownBars(ANode.FirstChild, ser);
      'c:axId':
        ReadChartSeriesAxis(ANode, ser);
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the formatting of the stockseries candlesticks

  @@param   ANode     First child of <c:upDownBars>
  @@param   ASeries   Series to which the parameters will be applied
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartStockSeriesUpDownBars(ANode: TDOMNode;
  ASeries: TsStockSeries);
var
  nodeName, s: String;
  n: Double;
  child: TDOMNode;
begin
  if ANode = nil then
    exit;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:gapWidth':
        begin
          s := GetAttrValue(ANode, 'val');
          if TryStrToFloat(s, n, FPointSeparatorSettings) then
            ASeries.TickWidthPercent := round(100 / (1 + n/100));
        end;
      'c:upBars':
        begin
          child := ANode.FindNode('c:spPr');
          if Assigned(child) then
            ReadChartFillAndLineProps(child.FirstChild, ASeries.Chart, ASeries.CandleStickUpFill, ASeries.CandlestickUpBorder);
        end;
      'c:downBars':
        begin
          child := ANode.FindNode('c:spPr');
          if Assigned(child) then
            ReadChartFillAndLineProps(child.FirstChild, ASeries.Chart, ASeries.CandleStickDownFill, ASeries.CandlestickDownBorder);
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

{ Extracts the chart and axis titles, their formatting and their texts. }
procedure TsSpreadOOXMLChartReader.ReadChartTitle(ANode: TDOMNode; ATitle: TsChartText);
var
  nodeName, s, totalText: String;
  child, child2, child3, child4: TDOMNode;
  axis: TsChartAxis;
  chart: TsChart;
  n: Integer;
begin
  if ANode = nil then
    exit;

  chart := ATitle.Chart;
  if ATitle = chart.XAxis.Title then
    axis := chart.XAxis
  else if ATitle = chart.YAxis.Title then
    axis := chart.YAxis
  else if ATitle = chart.X2Axis.Title then
    axis := chart.X2Axis
  else if Atitle = chart.Y2Axis.Title then
    axis := chart.Y2Axis
  else
    axis := nil;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:tx':
        begin
          child := ANode.FindNode('c:rich');
          if Assigned(child) then
          begin
            child2 := child.FirstChild;
            while Assigned(child2) do
            begin
              nodeName := child2.NodeName;
              case nodeName of
                'a:bodyPr':
                  begin
                    s := GetAttrValue(child2, 'rot');
                    if (s <> '') and TryStrToInt(s, n) then
                    begin
                      if axis <> nil then
                      begin
                        if n = 1000 then
                          axis.DefaultTitleRotation := true
                        else
                        begin
                          if n = 0 then
                            ATitle.RotationAngle := 0
                          else
                            ATitle.RotationAngle := -n / ANGLE_MULTIPLIER;
                          axis.DefaultTitleRotation := false;
                        end;
                      end else
                      begin
                        if (n = 1000) or (n = 0) then
                          ATitle.RotationAngle := 0
                        else
                          Atitle.RotationAngle := -n/ANGLE_MULTIPLIER;
                      end;
                    end;
                  end;
                'a:lstStyle':
                  ;
                'a:p':
                  begin
                    totalText := '';
                    child3 := child2.FirstChild;
                    while Assigned(child3) do
                    begin
                      nodeName := child3.NodeName;
                      case NodeName of
                        'a:pPr':
                          begin
                            child4 := child3.FindNode('a:defRPr');
                            ReadChartFontProps(child4, ATitle.Font);
                          end;
                        'a:r':
                          begin
                            child4 := child3.FindNode('a:t');
                            totalText := totalText + GetNodeValue(child4);
                          end;
                      end;
                      child3 := child3.NextSibling;
                    end;
                    ATitle.Caption := totalText;
                  end;
              end;
              child2 := child2.NextSibling;
            end;
          end; // "rich" node
        end;  // "tx" node
      'c:overlay':
        ;
      'c:spPr':
        ReadChartFillAndLineProps(ANode.FirstChild, ATitle.Chart, ATitle.Background, ATitle.Border);
      'c:txPr':
      ;
    end;
    ANode := ANode.NextSibling;
  end;
end;

{ ANode is a "c:txPr" node }
procedure TsSpreadOOXMLChartReader.ReadChartTextProps(ANode: TDOMNode; AFont: TsFont;
  var AFontRotation: Single);
var
  nodeName, s: String;
  n: Integer;
  child1, child2: TDOMNode;
begin
  if (ANode = nil) or (ANode.NodeName <> 'c:txPr') then exit;

  ANode := ANode.FirstChild;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'a:bodyPr':
        { <a:bodyPr rot="-5400000" spcFirstLastPara="1" vertOverflow="ellipsis"
             vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/> }
        begin
          s := GetAttrValue(ANode, 'rot');
          if (s <> '') and TryStrToInt(s, n) then
            AFontRotation := -n / ANGLE_MULTIPLIER;
        end;
      'a:lstStyle':
        ;
      'a:p':
        begin
          child1 := ANode.FirstChild;
          if Assigned(child1) then
          begin
            child2 := child1.FindNode('a:defRPr');
            if Assigned(child2) then
              ReadChartFontProps(child2, AFont);
          end;
        end;
      'a:endParaRPr':
        ;
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the main xml file of a chart (and the associated rels file)

  @param   AStream        Stream of the xlsx file
  @param   AChart         Chart instance, already created, but empty
  @param   AChartXMLFile  Name of the xml file with the chart data, usually 'xl/charts/chart1.xml'
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartXML(AStream: TStream; AChart: TsChart;
  AChartXMLFile: String);
var
  lReader: TsSpreadOOXMLReader;
  xmlStream: TStream;
  doc: TXMLDocument = nil;
  node: TDOMNode;
  nodeName: String;
  relsFileName: String;
  relsList: TXlsxRelationshipList;
begin
  lReader := TsSpreadOOXMLReader(Reader);

  // Read the rels file of the chart. The items go into the FRelsList.
  relsFileName := ExtractFilePath(AChartXMLFile) + '_rels/' + ExtractFileName(AChartXMLFile) + '.rels';
  relsList := TXlsxRelationshipList.Create;
  try
    lReader.ReadRels(AStream, relsFileName, relsList);
    // Read the images mentioned in the rels file.
    ReadChartImages(AStream, AChart, relsList);
  finally
    relsList.Free;
  end;

  // Read the xml file of the chart
  xmlStream := lReader.CreateXMLStream;
  try
    if UnzipToStream(AStream, AChartXMLFile, xmlStream) then
    begin
      lReader.ReadXMLStream(doc, xmlStream);
      node := doc.DocumentElement.FirstChild;
      while Assigned(node) do
      begin
        nodeName := node.NodeName;
        case nodeName of
          'c:chart':
            ReadChart(node, AChart);
          'c:spPr':
            ReadChartFillAndLineProps(node.FirstChild, AChart, AChart.Background, AChart.Border);
        end;
        node := node.NextSibling;
      end;
      FreeAndNil(doc);
    end;
  finally
    xmlStream.Free;
  end;
end;

procedure TsSpreadOOXMLChartReader.SetAxisDefaults(AWorkbookAxis: TsChartAxis);
begin
  AWorkbookAxis.Title.Caption := '';
  AWorkbookAxis.DefaultTitleRotation := true;
  AWorkbookAxis.LabelRotation := 0;
  AWorkbookAxis.Visible := false;
  AWorkbookAxis.MajorGridLines.Style := clsNoLine;
  AWorkbookAxis.MinorGridLines.Style := clsNoLine;
end;

procedure TsSpreadOOXMLChartReader.SetDefaultSeriesColor(ASeries: TsChartSeries);
begin
  ASeries.Fill.Color := CalcDefaultSeriesColor(ASeries.Order);
  ASeries.Fill.Style := cfsSolid;
  ASeries.Line.Style := clsNoLine;
end;


{ TsSpreadOOXMLChartWriter }

constructor TsSpreadOOXMLChartWriter.Create(AWriter: TsBasicSpreadWriter);
begin
  inherited Create(AWriter);

  FPointSeparatorSettings := SysUtils.DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator:='.';
end;

destructor TsSpreadOOXMLChartWriter.Destroy;
begin
  inherited;
end;

procedure TsSpreadOOXMLChartWriter.AddChartsToZip(AZip: TZipper);
var
  i: Integer;
begin
  // Add chart relationships to zip
  for i := 0 to High(FSChartRels) do
  begin
    if (FSChartRels[i] = nil) or (FSChartRels[i].Size = 0) then Continue;
    FSChartRels[i].Position := 0;
    AZip.Entries.AddFileEntry(FSChartRels[i], OOXML_PATH_XL_CHARTS_RELS + Format('chart%d.xml.rels', [i+1]));
  end;

  // Add chart styles to zip
  for i:=0 to High(FSChartStyles) do
  begin
    if (FSChartStyles[i] = nil) or (FSChartStyles[i].Size = 0) then Continue;
    FSChartStyles[i].Position := 0;
    AZip.Entries.AddFileEntry(FSChartStyles[i], OOXML_PATH_XL_CHARTS + Format('style%d.xml', [i+1]));
  end;

  // Add chart colors to zip
  for i:=0 to High(FSChartColors) do
  begin
    if (FSChartColors[i] = nil) or (FSChartColors[i].Size = 0) then Continue;
    FSChartColors[i].Position := 0;
    AZip.Entries.AddFileEntry(FSChartColors[i], OOXML_PATH_XL_CHARTS + Format('colors%d.xml', [i+1]));
  end;

  // Add charts top zip
  for i:=0 to High(FSCharts) do
  begin
    if (FSCharts[i] = nil) or (FSCharts[i].Size = 0) then Continue;
    FSCharts[i].Position := 0;
    AZip.Entries.AddFileEntry(FSCharts[i], OOXML_PATH_XL_CHARTS + Format('chart%d.xml', [i+1]));
  end;
end;

procedure TsSpreadOOXMLChartWriter.CreateStreams;
var
  n, i: Integer;
  workbook: TsWorkbook;
begin
  workbook := TsWorkbook(Writer.Workbook);
  n := workbook.GetChartCount;
  SetLength(FSCharts, n);
  SetLength(FSChartRels, n);
  SetLength(FSChartStyles, n);
  SetLength(FSChartColors, n);

  for i := 0 to n - 1 do
  begin
    FSCharts[i] := CreateTempStream(workbook, Format('fpsCh%d', [i]));
    FSChartRels[i] := CreateTempStream(workbook, Format('fpsChRels%d', [i]));
    FSChartStyles[i] := CreateTempStream(workbook, Format('fpsChSty%d', [i]));
    FSChartColors[i] := CreateTempStream(workbook, Format('fpsChCol%d', [i]));
  end;
end;

procedure TsSpreadOOXMLChartWriter.DestroyStreams;
var
  stream: TStream;
begin
  for stream in FSCharts do DestroyTempStream(stream);
  SetLength(FSCharts, 0);

  for stream in FSChartRels do DestroyTempStream(stream);
  SetLength(FSChartRels, 0);

  for stream in FSChartStyles do DestroyTempStream(stream);
  SetLength(FSChartStyles, 0);

  for stream in FSChartColors do DestroyTempStream(stream);
  SetLength(FSChartColors, 0);
end;

procedure TsSpreadOOXMLChartWriter.ResetStreams;
var
  stream: TStream;
begin
  for stream in FSCharts do stream.Position := 0;
  for stream in FSChartRels do stream.Position := 0;
  for stream in FSChartStyles do stream.Position := 0;
  for stream in FSChartColors do stream.Position := 0;
end;

{@@ ----------------------------------------------------------------------------
  Writes a cell number value to the xml stream for the series' number cache

  @param  AStream    Stream for the chart
  @param  AIndent    Number of indentation spaces for better legibility
  @param  AWorksheet Worksheet to which the cell contains
  @param  ARow       Row index of the cell
  @param  ACol       Column index of the cell
  @param  AIndex     Index of the cell among the series data points
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteCellNumberValue(AStream: TStream;
  AIndent: Integer; AWorksheet: TsBasicWorksheet; ARow, ACol, AIndex: Cardinal);
var
  indent: String;
  value: Double;
begin
  indent := DupeString(' ', AIndent);
  value := TsWorksheet(AWorksheet).ReadAsNumber(ARow, ACol);
  AppendToStream(AStream, Format(
    indent + '<c:pt idx="%d">' + LE +
    indent + '  <c:v>%g</c:v>' + LE +
    indent + '</c:pt>' + LE,
    [ AIndex, value ], FPointSeparatorSettings
  ));
end;


{@@ ----------------------------------------------------------------------------
  Writes the xl/charts/colorsN.xml file where N is the number AChartIndex.

  So far, the code is just copied from a file writen by Excel.

  @param  AStream  Stream to which the xml text is written
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartColorsXML(AStream: TStream;
  AChartIndex: Integer);
begin
  AppendToStream(AStream,
    XML_Header);

  AppendToStream(AStream,
    '<?xml version="1.0" encoding="UTF-8"?>' + LE +

    '<cs:colorStyle ' + LE +
    '    xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle"' + LE +
    '    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' + LE +
    '    meth="cycle" id="10">' + LE +

    '  <a:schemeClr val="accent1"/>' + LE +
    '  <a:schemeClr val="accent2"/>' + LE +
    '  <a:schemeClr val="accent3"/>' + LE +
    '  <a:schemeClr val="accent4"/>' + LE +
    '  <a:schemeClr val="accent5"/>' + LE +
    '  <a:schemeClr val="accent6"/>' + LE +

    '  <cs:variation/>' + LE +
    '  <cs:variation>' + LE +
    '    <a:lumMod val="60000"/>' + LE +
    '  </cs:variation>' + LE +
    '  <cs:variation>' + LE +
    '    <a:lumMod val="80000"/>' + LE +
    '    <a:lumOff val="20000"/>' + LE +
    '  </cs:variation>' + LE +
    '  <cs:variation>' + LE +
    '    <a:lumMod val="80000"/>' + LE +
    '  </cs:variation>' + LE +
    '  <cs:variation>' + LE +
    '    <a:lumMod val="60000"/>' + LE +
    '    <a:lumOff val="40000"/>' + LE +
    '  </cs:variation>' + LE +
    '  <cs:variation>' + LE +
    '    <a:lumMod val="50000"/>'+ LE +
    '  </cs:variation>' + LE +
    '  <cs:variation>'+ LE +
    '    <a:lumMod val="70000"/>' + LE +
    '    <a:lumOff val="30000"/>' + LE +
    '  </cs:variation>' + LE +
    '  <cs:variation>' + LE +
    '    <a:lumMod val="70000"/>' + LE +
    '  </cs:variation>' + LE +
    '  <cs:variation>' + LE +
    '    <a:lumMod val="50000"/>' + LE +
    '    <a:lumOff val="50000"/>' + LE +
    '  </cs:variation>'+ LE +

    '</cs:colorStyle>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes the main chart node, below <c:chartSpace> in chartN.xml
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartNode(AStream: TStream;
  AIndent: Integer; AChartIndex: Integer);
var
  chart: TsChart;
  indent: String;
  savedRotatedAxes: Boolean;
begin
  indent := DupeString(' ', AIndent);
  chart := TsWorkbook(Writer.Workbook).GetChartByIndex(AChartIndex);

  // Only bar charts can have rotated axes in Excel
  savedRotatedAxes := chart.RotatedAxes;
  if (chart.GetChartType <> ctBar) and chart.RotatedAxes then
  begin
    FWriter.Workbook.AddErrorMsg('Axes can be rotated only in bar charts.');
    chart.RotatedAxes := false;
  end;

  AppendToStream(AStream,
    indent + '<c:chart>' + LE
  );

  WriteChartTitleNode(AStream, AIndent + 2, chart.Title);
  WriteChartPlotAreaNode(AStream, AIndent + 2, chart);
  WriteChartLegendNode(AStream, AIndent + 2, chart.Legend);

  AppendToStream(AStream,
    indent + '  <c:plotVisOnly val="1" />' + LE +
    indent + '</c:chart>' + LE
  );

  // Write chart background
  AppendToStream(AStream,
    indent + '<c:spPr>' + LE +
             GetChartFillAndLineXML(AIndent, chart, chart.Background, chart.Border) + LE +
    indent + '</c:spPr>' + LE
  );

  chart.RotatedAxes := savedRotatedAxes;
end;

{ Write the relationship file for the chart with the given index.
  The file defines which xml files contain the ChartStyles and Colors, as well
  as images needed by each chart. }
procedure TsSpreadOOXMLChartWriter.WriteChartRelsXML(AStream: TStream;
  AChartIndex: Integer);
begin
  AppendToStream(AStream,
    XML_HEADER);
  AppendToStream(AStream, Format(
    '<Relationships xmlns="%s">' + LE +
    '  <Relationship Id="rId1" Target="style%d.xml" Type="%s" />' + LE +
    '  <Relationship Id="rId2" Target="colors%d.xml" Type="%s" />' + LE +
    '</Relationships>' + LE, [
    SCHEMAS_RELS,
    AChartIndex + 1, SCHEMAS_CHART_STYLE,
    AChartIndex + 1, SCHEMAS_CHART_COLORS
  ]));
end;

{@@ ----------------------------------------------------------------------------
  Writes the xl/charts/stylesN.xml file where N is the number AChartIndex.

  So far, the code is just copied from a file written by Excel.

  @param  AStream   Stream to which the xml text is written.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartStylesXML(AStream: TStream;
  AChartIndex: Integer);
begin
  AppendToStream(AStream,
    XML_Header);

  AppendToStream(AStream,
    '<cs:chartStyle ' +
         'xmlns:cs="http://schemas.microsoft.com/office/drawing/2012/chartStyle" ' +
         'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" id="201">' + LE +
    '  <cs:axisTitle>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1">' + LE +
    '        <a:lumMod val="65000"/>' + LE +
    '        <a:lumOff val="35000"/>' + LE +
    '      </a:schemeClr>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:defRPr sz="1000" kern="1200"/>' + LE +
    '  </cs:axisTitle>' + LE +

    '  <cs:categoryAxis>'+ LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1">' + LE +
    '        <a:lumMod val="65000"/>' + LE +
    '        <a:lumOff val="35000"/>' + LE +
    '      </a:schemeClr>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="15000"/>' + LE +
    '            <a:lumOff val="85000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '    <cs:defRPr sz="900" kern="1200"/>' + LE +
    '  </cs:categoryAxis>' + LE +

    '  <cs:chartArea mods="allowNoFillOverride allowNoLineOverride">' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:solidFill>' + LE +
    '        <a:schemeClr val="bg1"/>' + LE +
    '      </a:solidFill>' + LE +
    '      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + LE+
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="15000"/>' + LE +
    '            <a:lumOff val="85000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '    <cs:defRPr sz="1000" kern="1200"/>' + LE +
    '  </cs:chartArea>' + LE +

    '  <cs:dataLabel>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1">' + LE +
    '        <a:lumMod val="75000"/>' + LE +
    '        <a:lumOff val="25000"/>' + LE +
    '      </a:schemeClr>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:defRPr sz="900" kern="1200"/>' + LE +
    '  </cs:dataLabel>' + LE +

    '  <cs:dataLabelCallout>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="dk1">' + LE +
    '        <a:lumMod val="65000"/>' + LE +
    '        <a:lumOff val="35000"/>' + LE +
    '      </a:schemeClr>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:solidFill>' + LE +
    '        <a:schemeClr val="lt1"/>' + LE +
    '      </a:solidFill>' + lE+
    '      <a:ln>' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="dk1">' + LE +
    '            <a:lumMod val="25000"/>' + LE +
    '            <a:lumOff val="75000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '    <cs:defRPr sz="900" kern="1200"/>' + LE +
    '    <cs:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="clip" ' + LE +
    '        horzOverflow="clip" vert="horz" wrap="square" lIns="36576" ' + LE +
    '        tIns="18288" rIns="36576" bIns="18288" anchor="ctr" anchorCtr="1">' + LE +
    '      <a:spAutoFit/>' + LE +
    '    </cs:bodyPr>' + LE +
    '  </cs:dataLabelCallout>' + LE +

    '  <cs:dataPoint>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="1">' + LE +
    '      <cs:styleClr val="auto"/>' + LE +
    '    </cs:fillRef>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '  </cs:dataPoint>' + LE +

    '  <cs:dataPoint3D>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="1">' + LE +
    '      <cs:styleClr val="auto"/>' + LE +
    '    </cs:fillRef>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '  </cs:dataPoint3D>' + LE +

    '  <cs:dataPointLine>' + LE +
    '    <cs:lnRef idx="0">' + LE +
    '      <cs:styleClr val="auto"/>' + LE +
    '    </cs:lnRef>' + LE +
    '    <cs:fillRef idx="1"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="28575" cap="rnd">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="phClr"/>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:dataPointLine>' + LE +

    '  <cs:dataPointMarker>' + LE +
    '    <cs:lnRef idx="0">' + LE +
    '      <cs:styleClr val="auto"/>' + LE +
    '    </cs:lnRef>' + LE +
    '    <cs:fillRef idx="1">' + LE +
    '      <cs:styleClr val="auto"/>' + LE +
    '    </cs:fillRef>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="9525">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="phClr"/>' + LE +
    '        </a:solidFill>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:dataPointMarker>' + LE +

    '  <cs:dataPointMarkerLayout symbol="circle" size="5"/>' + LE +

    '  <cs:dataPointWireframe>' + LE +
    '    <cs:lnRef idx="0">' + LE +
    '      <cs:styleClr val="auto"/>' + LE +
    '    </cs:lnRef>' + LE +
    '    <cs:fillRef idx="1"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="9525" cap="rnd">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="phClr"/>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:dataPointWireframe>' + LE +

    '  <cs:dataTable>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1">' + LE +
    '        <a:lumMod val="65000"/>' + LE +
    '        <a:lumOff val="35000"/>' + LE +
    '      </a:schemeClr>'  + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:noFill/>' + LE +
    '      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="15000"/>' + LE +
    '            <a:lumOff val="85000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>'  + lE +
    '    <cs:defRPr sz="900" kern="1200"/>' + LE +
    '  </cs:dataTable>' + LE +

    '  <cs:downBar>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="dk1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:solidFill>' + LE +
    '        <a:schemeClr val="dk1">' + LE +
    '          <a:lumMod val="65000"/>' + LE +
    '          <a:lumOff val="35000"/>' + LE +
    '        </a:schemeClr>' + LE +
    '      </a:solidFill>' + LE +
    '      <a:ln w="9525">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="65000"/>' + LE +
    '            <a:lumOff val="35000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:downBar>' + LE +

    '  <cs:dropLine>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="35000"/>' + LE +
    '            <a:lumOff val="65000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:dropLine>' + LE +

    '  <cs:errorBar>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="65000"/>' + LE +
    '            <a:lumOff val="35000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:errorBar>' + LE +

    '  <cs:floor>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:noFill/>' + LE +
    '      <a:ln>' + LE +
    '        <a:noFill/>' + LE+
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:floor>' + LE +

    '  <cs:gridlineMajor>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="15000"/>' + LE +
    '            <a:lumOff val="85000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:gridlineMajor>' + LE +

    '  <cs:gridlineMinor>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="5000"/>' + LE +
    '            <a:lumOff val="95000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:gridlineMinor>' + LE +

    '  <cs:hiLoLine>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="75000"/>' + LE +
    '            <a:lumOff val="25000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:hiLoLine>' + LE +

    '  <cs:leaderLine>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="35000"/>' + LE +
    '            <a:lumOff val="65000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:leaderLine>' + LE +

    '  <cs:legend>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1">'  + LE +
    '        <a:lumMod val="65000"/>' + LE +
    '        <a:lumOff val="35000"/>' + LE +
    '      </a:schemeClr>' + LE +
    '    </cs:fontRef>' + LE+
    '    <cs:defRPr sz="900" kern="1200"/>' + LE +
    '  </cs:legend>' + LE +

    '  <cs:plotArea mods="allowNoFillOverride allowNoLineOverride">' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '  </cs:plotArea>' + LE +

    '  <cs:plotArea3D mods="allowNoFillOverride allowNoLineOverride">' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">'  + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '  </cs:plotArea3D>' + LE +

    '  <cs:seriesAxis>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1">' + LE +
    '        <a:lumMod val="65000"/>' + LE +
    '        <a:lumOff val="35000"/>' + LE +
    '      </a:schemeClr>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:defRPr sz="900" kern="1200"/>' + LE +
    '  </cs:seriesAxis>' + LE +

    '  <cs:seriesLine>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '      <cs:fontRef idx="minor">' + LE +
    '        <a:schemeClr val="tx1"/>' + LE +
    '      </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE+
    '      <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="35000"/>' + LE +
    '            <a:lumOff val="65000"/>' + LE +
    '          </a:schemeClr>' + LE+
    '        </a:solidFill>' + LE +
    '        <a:round/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:seriesLine>' + LE +

    '  <cs:title>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1">' + LE +
    '        <a:lumMod val="65000"/>' + LE +
    '        <a:lumOff val="35000"/>' + LE +
    '      </a:schemeClr>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:defRPr sz="1400" b="0" kern="1200" spc="0" baseline="0"/>' + LE +
    '  </cs:title>' + LE +

    '  <cs:trendline>' + LE +
    '    <cs:lnRef idx="0">' + LE +
    '      <cs:styleClr val="auto"/>' + LE +
    '    </cs:lnRef>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:ln w="19050" cap="rnd">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="phClr"/>' + LE +
    '        </a:solidFill>' + LE +
    '        <a:prstDash val="sysDot"/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:trendline>' + LE +

    '  <cs:trendlineLabel>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1">' + LE +
    '        <a:lumMod val="65000"/>' + LE +
    '        <a:lumOff val="35000"/>' + LE +
    '      </a:schemeClr>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:defRPr sz="900" kern="1200"/>' + LE +
    '  </cs:trendlineLabel>' + LE +

    '  <cs:upBar>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>'  + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="dk1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:solidFill>' + LE +
    '        <a:schemeClr val="lt1"/>' + LE +
    '      </a:solidFill>' + LE +
    '      <a:ln w="9525">' + LE +
    '        <a:solidFill>' + LE +
    '          <a:schemeClr val="tx1">' + LE +
    '            <a:lumMod val="15000"/>' + LE +
    '            <a:lumOff val="85000"/>' + LE +
    '          </a:schemeClr>' + LE +
    '        </a:solidFill>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:upBar>' + LE +

    '  <cs:valueAxis>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE +
    '      <a:schemeClr val="tx1">' + LE +
    '        <a:lumMod val="65000"/>' + LE +
    '        <a:lumOff val="35000"/>' + LE +
    '      </a:schemeClr>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:defRPr sz="900" kern="1200"/>' + LE +
    '  </cs:valueAxis>' +

    '  <cs:wall>' + LE +
    '    <cs:lnRef idx="0"/>' + LE +
    '    <cs:fillRef idx="0"/>' + LE +
    '    <cs:effectRef idx="0"/>' + LE +
    '    <cs:fontRef idx="minor">' + LE+
    '      <a:schemeClr val="tx1"/>' + LE +
    '    </cs:fontRef>' + LE +
    '    <cs:spPr>' + LE +
    '      <a:noFill/>' + LE +
    '      <a:ln>' + LE +
    '        <a:noFill/>' + LE +
    '      </a:ln>' + LE +
    '    </cs:spPr>' + LE +
    '  </cs:wall>' + LE +

    '</cs:chartStyle>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes the root node of the file chartN.xml (where N is chart number)
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartSpaceXML(AStream: TStream;
  AChartIndex: Integer);
begin
  AppendToStream(AStream,
    '<?xml version="1.0" encoding="utf-8" standalone="yes"?>' + LE +
    '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ' + LE +
    '              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' + LE +
    '              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' + LE +
    '  <c:date1904 val="0"/>' + LE +      // to do: get correct value
    '  <c:roundedCorners val="0"/>' + LE

  );

  WriteChartNode(AStream, 2, AChartIndex);

  AppendToStream(AStream,
    '</c:chartSpace>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Creates the xml string representing a color for xlsx

  Example:
              <a:solidFill>
                <a:srgbClr val="FF0000">
                  <a:alpha val="60000"/>
                </a:srgbClr>
              </a:solidFill>

  @param  AIndent    Number of indentation spaces for better legibility
  @param  ANodeName  Name of the outermost xml node, in above example 'a:solidFill'
  @param  AColor     Color (including transparency) to be encoded
-------------------------------------------------------------------------------}
function TsSpreadOOXMLChartWriter.GetChartColorXML(AIndent: Integer;
  ANodeName: String; AColor: TsChartColor): String;
var
  indent: String;
  rgbStr: String;
begin
  indent := DupeString(' ', AIndent);
  rgbStr := GetChartColorXML(AColor);
  Result :=
    indent + '<' + ANodeName + '>' + LE +
    indent + '  ' + rgbStr +  LE +
    indent + '</' + ANodeName + '>';
end;

function TsSpreadOOXMLChartWriter.GetChartColorXML(AColor: TsChartColor): String;
var
  alpha: Integer;
begin
  if (AColor.Transparency > 0) then
  begin
    alpha := round((1.0 - AColor.Transparency) * FACTOR_MULTIPLIER);
    Result := Format('<a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr>',
      [HtmlColorStr(AColor.Color), alpha]
    );
  end else
    Result := Format('<a:srgbClr val="%s"/>',
      [ HtmlColorStr(AColor.Color) ]
    );
end;

{@@ ----------------------------------------------------------------------------
  Assembles the xml string for the children of a <c:spPr> node (fill and line style)
-------------------------------------------------------------------------------}
function TsSpreadOOXMLChartWriter.GetChartFillAndLineXML(AIndent: Integer;
  AChart: TsChart; AFill: TsChartFill; ALine: TsChartLine): String;
begin
  Result :=
    GetChartFillXML(AIndent, AChart, AFill) + LE +
    GetChartLineXML(AIndent, AChart, ALine);
  //  indent + '<a:effectLst/>';
end;

function TsSpreadOOXMLChartWriter.GetChartFillXML(AIndent: Integer;
  AChart: TsChart; AFill: TsChartFill): String;
const
  HATCH_NAMES: array[0..47] of string = (
    'pct5', 'pct10', 'pct20', 'pct25', 'pct30',                 // 0..4
    'pct40', 'pct50', 'pct60', 'pct70', 'pct75',                // 5..9
    'pct80', 'pct90', 'dashDnDiag', 'dashUpDiag', 'dashHorz',   // 10..14
    'dashVert', 'smConfetti', 'lgConfetti', 'zigZag', 'wave',   // 15..19
    'diagBrick', 'horzBrick', 'weave', 'plaid', 'divot',        // 20..24
    'dotGrid', 'dotDmnd', 'shingle', 'trellis', 'sphere',       // 25..29
    'smCheck', 'lgCheck', 'solidDmnd', 'ltDnDiag', 'ltUpDiag',  // 30..34
    'dkDnDiag', 'dkUpDiag', 'wdDnDiag', 'wdUpDiag', 'ltHorz',   // 35..39
    'ltVert', 'narVert', 'narHorz', 'dkHorz', 'dkVert',         // 40..44
    'smGrid', 'lgGrid', 'openDmnd'                              // 45..47
  );
var
  indent: String;
  hatch: TsChartHatch;
  gradient: TsChartGradient;
  step: TsChartGradientStep;
  gSteps: String = '';
  gStyle: String = '';
  lStr: String = '';
  tStr: String = '';
  rStr: String = '';
  bStr: String = '';
  i: Integer;
  presetIdx: Integer;
  alpha: Integer;
  rgbStr: String;
begin
  indent := DupeString(' ', AIndent);

  if (AFill = nil) or (AFill.Style = cfsNoFill) then
    Result := indent + '<a:noFill/>'
  else
    case AFill.Style of
      // Solid fills
      cfsSolid:
        Result := GetChartColorXML(AIndent + 2, 'a:solidFill', AFill.Color);

      // Gradient fills
      cfsGradient:
        begin
          gradient := AChart.Gradients[AFill.Gradient];
          gSteps := indent + '  <a:gsLst>' + LE;
          for i := 0 to gradient.NumSteps - 1 do
          begin
            step := gradient.Steps[i];
            gSteps := gSteps + Format(
              indent + '    <a:gs pos="%.0f">' + LE +
              indent + '      %s' + LE +
              indent + '    </a:gs>' + LE,
              [ step.Value * FACTOR_MULTIPLIER, GetChartColorXML(step.Color) ]
            );
          end;
          gSteps := gSteps + indent + '  </a:gsLst>' + LE;
          case gradient.Style of
            cgsLinear:
              gStyle := indent + Format('  <a:lin ang="%.0f" scaled="1"/>',
                [ PositiveAngle(-gradient.Angle) * ANGLE_MULTIPLIER ]   // xlsx gradient direction is CW, fps CCW
              );
            cgsAxial,
            cgsRadial,
            cgsElliptic,
            cgsSquare,
            cgsRectangular,
            cgsShape:
              begin
                case gradient.Style of
                  cgsRectangular, cgsAxial: gStyle := 'rect';
                  cgsElliptic, cgsRadial: gStyle := 'circle';
                  else gStyle := 'shape';
                end;
                if gradient.CenterX <> 0 then lStr := Format('l="%.0f" ', [gradient.CenterX * FACTOR_MULTIPLIER]);
                if gradient.CenterX <> 1.0 then rStr := Format('r="%.0f" ', [(1.0-gradient.CenterX) * FACTOR_MULTIPLIER]);
                if gradient.CenterY <> 0 then tStr := Format('t="%.0f" ', [gradient.CenterY * FACTOR_MULTIPLIER]);
                if gradient.CenterY <> 1.0 then bStr := Format('b="%.0f" ', [(1.0-gradient.CenterY) * FACTOR_MULTIPLIER]);
                gStyle := Format(
                  indent + '  <a:path path="%s">' + LE +
                  indent + '    <a:fillToRect %s%s%s%s/>' + LE +
                  indent + '  </a:path>' + LE,
                  [ gStyle, lStr, tStr, rStr, bStr ]
                );
              end;
          end;
          Result := indent + '<a:gradFill>' + LE +
                    gSteps +
                    gStyle +
                    indent + '</a:gradFill>' + LE;
        end;

      // Hatched and pattern fills
      cfsHatched, cfsSolidHatched:
        begin
          hatch := AChart.Hatches[AFill.Hatch];
          presetIdx := -1;
          for i := 0 to High(HATCH_NAMES) do
            if hatch.Name = HATCH_NAMES[i] then
            begin
              presetIdx := i;
              break;
            end;
          if presetIdx = -1 then
            case Lowercase(hatch.Name) of
              'crossed': presetIdx := 47;   // openDmnd
              'forward': presetIdx := 34;   // ltUpDiag
              'backward': presetIdx := 33;  // ltDnDiag
            end;
          if presetIdx > -1 then
            Result :=
              indent + '<a:pattFill prst="' + HATCH_NAMES[presetIdx] + '">' + LE +
                       GetChartColorXML(AIndent + 2, 'a:fgClr', hatch.PatternColor) + LE +
                       GetChartColorXML(AIndent + 2, 'a:bgClr', AFill.Color) + LE +
              indent + '</a:pattFill>'
          else
            // unknown pattern - use a solid fill
            Result :=
              indent + GetChartColorXML(AIndent + 2, 'a:solidFill', AFill.Color);
        end;
      else
        Result := indent + '<a:noFill/>';
    end;
end;

{@@ ----------------------------------------------------------------------------
  Assembles the xml string for a font to be used in a chart

  @param   AIndent  Number of indentation spaces, for better legibility
  @param   AFont    Font to be processed
  @param   ANode    String for the node in which the result is used. Either '<a:defRPr>', or '<a:rPr>'
-------------------------------------------------------------------------------}
function TsSpreadOOXMLChartWriter.GetChartFontXML(AIndent: Integer;
  AFont: TsFont; ANodeName: String): String;
var
  indent: String;
  fontname: String;
  bold: String;
  italic: String;
  strike: String;
  underline: String;
begin
  indent := DupeString('  ', AIndent);

  fontName := IfThen(AFont.FontName <> '', AFont.FontName, DEFAULT_FONTNAME);
  bold := IfThen(fssBold in AFont.Style, 'b="1" ', 'b="0" ');
  italic := IfThen(fssItalic in AFont.Style, 'i="1" ', '');
  strike := IfThen(fssStrikeOut in AFont.Style, 'strike="sngStrike" ', 'strike="noStrike" ');  // no support for double-strike...
  underline := IfThen(fssUnderline in AFont.Style, 'u="sng" ', '');  // no support for double-underline

  Result := Format(
    indent + '<%0:s sz="%d" spc="-1" %s%s%s%s>' + LE +
    indent + '  <a:solidFill>' + LE +
    indent + '    <a:srgbClr val="%s"/>' + LE +
    indent + '  </a:solidFill>' + LE +
    indent + '  <a:latin typeface="%s"/>' + LE +
    indent + '</%0:s>',
    [
      ANodeName,
      round(AFont.Size * 100),
      bold, italic, strike, underline,
      HTMLColorStr(AFont.Color),
      fontName
    ]
  );
end;

{@@ ----------------------------------------------------------------------------
  Creates an xml string for a line style node <a:ln>. Must be inserted into a
  <c:spPr> node.

  @param  AIndent  Number of indentation spaces for better legibility
  @param  AChart   Chart to which this line belongs
  @param  ALine    Line instance for which the string is created
  @param  OverrideOff  If true, an empty line string is created no matter which parameters are in ALine
-------------------------------------------------------------------------------}
function TsSpreadOOXMLChartWriter.GetChartLineXML(AIndent: Integer;
  AChart: TsChart; ALine: TsChartline; OverrideOff: Boolean = false): String;
var
  indent: String;
  noLine: Boolean;
  lineStyle: TsChartLineStyle;
  w: Double;
  len1: Double;
  len2: Double;
  space: Double;
begin
  indent := DupeString(' ', AIndent);

  if (ALine = nil) or (ALine.Style = clsNoLine) or OverrideOff then
    Result := indent + '<a:ln>' + LE +
              indent + '  <a:noFill/>' + LE +
              indent + '</a:ln>'
  else
  begin
    Result := Format(
      indent + '<a:ln w="%.0f">' + LE +
               GetChartColorXML(AIndent + 2, 'a:solidFill', ALine.Color) + LE,
      [ mmToPts(ALine.Width) ]
    );
    if ALine.Style <> clsSolid then
    begin
      lineStyle := AChart.LineStyles[ALine.Style];
      if lineStyle.RelativeToLineWidth then
      begin
        w := ALine.Width;
        if w < 1 then w := 1.0;
        len1 := w * lineStyle.Segment1.Length * 0.01;
        len2 := w * lineStyle.Segment2.Length * 0.01;
        space := w * lineStyle.Distance * 0.01;
      end else
      begin
        len1 := lineStyle.Segment1.Length;
        len2 := lineStyle.Segment2.Length;
        space := lineStyle.Distance;
      end;
      Result := Result + Format(
        indent + '  <a:custDash>' + LE +
        indent + '    <a:ds d="%.0f" sp="%.0f"/>' + LE +
        indent + '  </a:custDash>' + LE,
        [ mmToPts(len1) * PTS_MULTIPLIER, mmToPts(space) * PTS_MULTIPLIER ]
      );
      // To do: how to handle multiple segments?
    end;
    Result := Result + indent + '</a:ln>';
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the properties of the given area series to the <c:plotArea> node of
  file chartN.xml
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteAreaSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsAreaSeries; ASeriesIndex, APosInAxisGroup: Integer);
const
  GROUPING: Array[TsChartStackMode] of string = ('standard', 'stacked', 'percentStacked');
var
  indent: String;
  chart: TsChart;
  xAxis: TsChartAxis;
  isFirstOfGroup: Boolean;
  isLastOfGroup: Boolean;
  prevSeriesGroupIndex: Integer = -1;
  nextSeriesGroupIndex: Integer = -1;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  if (ASeriesIndex > 0) and (chart.Series[ASeriesIndex-1].YAxis = ASeries.YAxis) then
    prevSeriesGroupIndex := chart.Series[ASeriesIndex-1].GroupIndex;
  if (ASeriesIndex < chart.Series.Count-1) and (chart.Series[ASeriesIndex+1].YAxis = ASeries.YAxis) then
    nextSeriesGroupIndex := chart.Series[ASeriesIndex+1].GroupIndex;

  isFirstOfGroup := APosInAxisGroup and 1 = 1;
  isLastOfGroup := APosInAxisgroup and 2 = 2;

  if ((ASeries.GroupIndex > -1) and (prevSeriesGroupIndex = ASeries.GroupIndex)) then
    isFirstOfGroup := false;
  if ((ASeries.GroupIndex > -1) and (nextSeriesGroupIndex = ASeries.GroupIndex)) then
    isLastOfGroup := false;

  if isFirstOfGroup then
    AppendToStream(AStream, Format(
      indent + '<c:areaChart>' + LE +
      indent + '  <c:grouping val="%s"/>' + LE,
      [ GROUPING[chart.StackMode] ]
    ));

  WriteChartSeriesNode(AStream, AIndent + 2, ASeries);

  if isLastOfGroup then
  begin
    if ASeries.YAxis = calPrimary then
      xAxis := chart.XAxis
    else
      xAxis := chart.X2Axis;

    AppendToStream(AStream, Format(
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '</c:areaChart>' + LE,
      [
        FAxisID[xAxis.Alignment],  // <c:axId>
        FAxisID[ASeries.GetYAxis.Alignment]   // <c:axId>
      ]
    ));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the properties of the given bar series to the <c:plotArea> node of
  file chartN.xml
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteBarSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsBarSeries; ASeriesIndex, APosInAxisGroup: Integer);
const
  BAR_DIR: array[boolean] of string = ('col', 'bar');
  GROUPING: array[TsChartStackMode] of string = ('clustered', 'stacked', 'percentStacked');
var
  indent: String;
  chart: TsChart;
  xAxis: TsChartAxis;
  gapWidth: Integer = 0;
  overlap: Integer = 999;
  isFirstOfGroup: Boolean;
  isLastOfGroup: Boolean;
  prevSeriesGroupIndex: Integer = -1;
  nextSeriesGroupIndex: Integer = -1;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  if (ASeriesIndex > 0) and (chart.Series[ASeriesIndex-1].YAxis = ASeries.YAxis) then
    prevSeriesGroupIndex := chart.Series[ASeriesIndex-1].GroupIndex;
  if (ASeriesIndex < chart.Series.Count-1) and (chart.Series[ASeriesIndex+1].YAxis = ASeries.YAxis) then
    nextSeriesGroupIndex := chart.Series[ASeriesIndex+1].GroupIndex;

  isFirstOfGroup := APosInAxisGroup and 1 = 1;
  isLastOfGroup := APosInAxisgroup and 2 = 2;

  if ((ASeries.GroupIndex > -1) and (prevSeriesGroupIndex = ASeries.GroupIndex)) then
    isFirstOfGroup := false;
  if ((ASeries.GroupIndex > -1) and (nextSeriesGroupIndex = ASeries.GroupIndex)) then
    isLastOfGroup := false;

  if (ASeries.GroupIndex > -1) and (chart.StackMode <> csmDefault) then
    overlap := 100;

  if isFirstOfGroup then
    AppendToStream(AStream, Format(
      indent + '<c:barChart>' + LE +
      indent + '  <c:barDir val="%s"/>' + LE +
      indent + '  <c:varyColors val="0"/>' + LE +
      indent + '  <c:grouping val="%s"/>' + LE,
      [ BAR_DIR[chart.RotatedAxes], GROUPING[chart.StackMode] ]
    ));

  WriteChartSeriesNode(AStream, AIndent + 2, ASeries);

  if isLastOfGroup then
  begin
    if overlap = 999 then
      overlap := chart.BarOverlapPercent;
    gapWidth := chart.BarGapWidthPercent;

    if ASeries.YAxis = calPrimary then
      xAxis := chart.XAxis
    else
      xAxis := chart.X2Axis;

    AppendToStream(AStream, Format(
      indent + '  <c:gapWidth val="%d"/>' + LE +
      indent + '  <c:overlap val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '</c:barChart>' + LE,
      [
        gapWidth,                                // <c:gapWidth>
        overlap,                                 // <c:overlap>
        FAxisID[xAxis.Alignment],  // <c:axId>
        FAxisID[ASeries.GetYAxis.Alignment]   // <c:axId>
      ]
    ));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the properties of the given bubble series to the <c:plotArea> node
  of the chart.

  @param  AStream       Stream to be written, it becomes the chartN.xml file
  @param  AIndent       Count of indentation spaces, for better legibility
  @param  ASeries       Bubble series to be processed
  @param  APosInAxisGroup Bit 1 - first series on primary or secondary axis, bit 2 - last series on primary or secondary axis
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteBubbleSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsBubbleSeries; APosInAxisGroup: Integer);
var
  indent: String;
  chart: TsChart;
  xAxis: TsChartAxis;
  isFirstOfGroup: Boolean;
  isLastOfGroup: Boolean;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  isFirstOfGroup := (APosInAxisGroup and 1 = 1);
  isLastOfGroup := (APosInAxisGroup and 2 = 2);

  if isFirstOfGroup then
    AppendToStream(AStream,
      indent + '<c:bubbleChart>' + LE +
      indent + '  <c:varyColors val="0"/>' + LE
    );

  WriteChartSeriesNode(AStream, AIndent + 2, ASeries);

  if isLastOfGroup then
  begin
    if ASeries.YAxis = calPrimary then
      xAxis := chart.XAxis
    else
      xAxis := chart.X2Axis;

    if ASeries.BubbleSizeMode = bsmRadius then
      AppendToStream(AStream,
        indent + '  <c:sizeRepresents val="w"/>' + LE
      );

    AppendToStream(AStream, Format(
      indent + '  <c:bubbleScale val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '</c:bubbleChart>' + LE,
      [
        round(ASeries.BubbleScale*100),          // <c:bubbleScale>
        FAxisID[xAxis.Alignment],  // <c:axId>
        FAxisID[ASeries.GetYAxis.Alignment]   // <c:axId>
      ]
    ));
  end;

  // Note:  <c:showNegBubbles> not supported
end;

{@@ ----------------------------------------------------------------------------
  Writes the properties of the given chart axis to the chartN.xml file under
  the <c:plotArea> node

  Depending on AxisKind, the node is either <c:catAx> or <c:valAx>.

  @param  AStream     Stream of the chartN.xml file
  @param  AIndent     Count of indentation spaces to increase readability
  @param  Axis        Chart axis processed
  @param  ANodeName   'catAx' when Axis is a category axis, otherwise 'valAx'
//  @param  IsSecondary  is true when the axis is used as a secondary axis.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartAxisNode(AStream: TStream;
  AIndent: Integer; Axis: TsChartAxis; ANodeName: String);

  function GetTickMarkStr(ATicks: TsChartAxisTicks): String;
  begin
    if ATicks = [] then
      Result := 'none'
    else if ATicks = [catInside] then
      Result := 'in'
    else if ATicks = [catOutside] then
      Result := 'out'
    else
      Result := 'cross';
  end;

  function GetGridLineStr(AIndent: Integer; ANodeName: String; ALine: TsChartLine): String;
  var
    indent: String;
  begin
    if ALine.Style <> clsNoLine then
    begin
      indent := DupeString(' ', AIndent);
      Result := Format(
        indent + '<%0:s>' + LE +
        indent + '  <c:spPr>' + LE +
                      '%s' + LE +
        indent + '  </c:spPr>' + LE +
        indent + '</%0:s>' + LE,
        [ ANodeName, GetChartLineXML(AIndent + 4, Axis.Chart, ALine) ]
      )
    end else
      Result := '';
  end;

var
  indent: String;
  chart: TsChart;
  axID: DWord;
  crosses: String;
  delete: Integer = 0;
  axAlign: TsChartAxisAlignment;
  rotAxID: DWord;
  fmt: String;
begin
  indent := DupeString(' ', AIndent);
  chart := Axis.Chart;

  axID := FAxisID[Axis.Alignment];
  rotAxID := FAxisID[Axis.GetRotatedAxis.Alignment];
  if Axis = chart.X2Axis then
    delete := 1;

  AppendToStream(AStream, Format(
    indent + '<%s>' + LE +
    indent + '  <c:axId val="%d"/>' + LE,
    [ ANodeName, axID ]
  ));

  WriteChartAxisScaling(AStream, AIndent + 2, Axis);

  if Axis.Alignment = caaTop then axAlign := caaBottom else axAlign := Axis.Alignment;
  AppendToStream(AStream, Format(
    indent + '  <c:delete val="%d"/>' + LE +
    indent + '  <c:axPos val="%s" />' + LE,
    [ delete,
      AX_POS[Axis.Chart.RotatedAxes, axAlign]
    ]
    // axis rotation seems to be respected by Excel only for bar series.
  ));

  // Grid lines
  if Axis <> chart.Y2Axis then
    AppendToStream(AStream,
      GetGridLineStr(AIndent + 2, 'c:majorGridlines', Axis.MajorGridLines) +
      GetGridLineStr(AIndent + 2, 'c:minorGridlines', Axis.MinorGridLines)
    );

  // Axis title
  WriteChartAxisTitle(AStream, AIndent + 2, Axis);

  // Axis labels
  if Axis.ShowLabels then
    WriteChartLabels(AStream, AIndent + 2, Axis.LabelFont);

  // Axis position
  if delete = 0 then
  begin
    case Axis.Position of
      capStart:
        crosses := '  <c:crosses val="min"/>';
      capEnd:
        crosses := '  <c:crosses val="max"/>';
      capValue:
        crosses := Format('  <c:crossesAt val="%g"/>', [Axis.PositionValue], FPointSeparatorSettings);
      else
        raise Exception.Create('Unsupported value of Axis.Position');
      // not used here: "autoZero"
    end;
    if crosses <> '' then crosses := indent + crosses + LE;
  end;

  if Axis.DateTime then
    fmt := 'mm/dd/yyyy'
  else
    fmt := 'General';

  AppendToStream(AStream, Format(
    indent + '  <c:numFmt formatCode="%s" sourceLinked="1"/>' + LE +
    indent + '  <c:majorTickMark val="%s"/>' + LE +
    indent + '  <c:minorTickMark val="%s"/>' + LE +
    indent + '  <c:tickLblPos val="nextTo"/>' + LE +
    indent + '  <c:crossAx val="%d"/>' + LE +
                crosses +
//    indent + '  <c:auto val="1"/>' + LE +
    indent + '</%s>' + LE,
    [
      fmt,
      GetTickMarkStr(Axis.MajorTicks),  // <c:majorTickMark>
      GetTickMarkStr(Axis.MinorTicks),  // <c:minorTickMark>
      rotAxID,                          // <c:crossAx>
      ANodeName                         // </c:catAx> or </c:valAx>
    ]
  ));
end;

procedure TsSpreadOOXMLChartWriter.WriteChartAxisScaling(AStream: TStream;
  AIndent: Integer; Axis: TsChartAxis);
const
  INVERTED: array[boolean] of String = ('minMax', 'maxMin');
var
  indent: String;
  intv: Double;
  logStr: String = '';
  maxStr: String = '';
  minStr: String = '';
  orientationStr: String;
begin
  indent := DupeString(' ', AIndent);

  if not Axis.AutomaticMax then
    maxStr := indent + Format('  <c:max val="%g"/>', [Axis.Max], FPointSeparatorSettings) + LE;

  if not Axis.AutomaticMin then
    minStr := indent + Format('  <c:min val="%g"/>', [Axis.Min], FPointSeparatorSettings) + LE;

  if Axis.Logarithmic then
    logStr := indent + Format('  <c:logBase val="%g"/>', [Axis.LogBase], FPointSeparatorSettings) + LE;

  orientationStr := indent + Format('  <c:orientation val="%s"/>', [ INVERTED[Axis.Inverted] ]) + LE;

  AppendToStream(AStream,
    indent + '<c:scaling>' + LE +
                maxStr +
                minStr +
                logStr +
                orientationStr +
    indent + '</c:scaling>' + LE
  );

  // The following nodes are outside the <c:scaling node> !
  if not Axis.AutomaticMajorInterval then
    AppendToStream(AStream, Format(
      indent + '<c:majorUnit val="%g"/>', [Axis.MajorInterval], fPointSeparatorSettings) + LE
    );

  if not Axis.AutomaticMinorInterval then
    AppendToStream(AStream, Format(
      indent + '<c:minorUnit val="%g"/>', [Axis.MinorInterval], FPointSeparatorSettings) + LE
    );
end;

procedure TsSpreadOOXMLChartWriter.WriteChartAxisTitle(AStream: TStream;
  AIndent: Integer; Axis: TsChartAxis);
var
  indent: String;
  chart: TsChart;
begin
  if not Axis.Title.Visible or (Axis.Title.Caption = '') then
    exit;

  indent := DupeString(' ', AIndent);
  chart := Axis.Chart;

  AppendToStream(AStream,
    indent + '<c:title>' + LE
  );

  WriteChartText(AStream, AIndent + 4, Axis.Title, Axis.TitleRotationAngle);

  AppendToStream(AStream,
    indent + '  <c:overlay val="0"/>' + LE +
    indent + '  <c:spPr>' + LE +
                  GetChartFillAndLineXML(AIndent + 6, Axis.Chart, Axis.Title.Background, Axis.Title.Border) + LE +
    indent + '  </c:spPr>' + LE
  );

  AppendToStream(AStream,
    indent + '</c:title>' + LE
  );
end;


{@@ ----------------------------------------------------------------------------
  Writes the chart-related entries to the [Content_Types].xml file

  @param  AStream   Stream holding the other content types
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartContentTypes(AStream: TStream);
var
  i, j, n: Integer;
  workbook: TsWorkbook;
  sheet: TsWorksheet;
begin
  workbook := TsWorkbook(Writer.Workbook);
  n := 1;
  for i:=0 to workbook.GetWorksheetCount-1 do
  begin
    sheet := workbook.GetWorksheetByIndex(i);
    for j:=0 to sheet.GetChartCount-1 do
    begin
      AppendToStream(AStream, Format(
        '<Override PartName="/xl/charts/chart%d.xml" ContentType="%s" />' + LE,
          [n, MIME_DRAWINGML_CHART]));
      AppendToStream(AStream, Format(
        '<Override PartName="/xl/charts/style%d.xml" ContentType="%s" />' + LE,
          [n, MIME_DRAWINGML_CHART_STYLE]));
      AppendToStream(AStream, Format(
        '<Override PartName="/xl/charts/colors%d.xml" ContentType="%s" />' + LE,
          [n, MIME_DRAWINGML_CHART_COLORS]));
      inc(n);
    end;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a <c:txPr> node for chart or axis labels

  @param  AStream  Stream to be written to
  @param  AIndent  Number of indentation spaced, for better legibility
  @param  AFont    Font to be used by the labels
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartLabels(AStream: TStream;
  AIndent: Integer; AFont: TsFont);
var
  indent: String;
begin
  indent := DupeString(' ', AIndent);

  AppendToStream(AStream,
    indent + '<c:txPr>' + LE +
    indent + '  <a:bodyPr/>' + LE +
    indent + '  <a:lstStyle/>' + LE +
    indent + '  <a:p>' + LE +
    indent + '    <a:pPr>' + LE +
                    GetChartFontXML(AIndent + 6, AFont, 'a:defRPr') + LE +
    indent + '    </a:pPr>' + LE +
    indent + '  </a:p>' + LE +
    indent + '</c:txPr>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes the <c:legend> node of a chart.

  @param  AStream  Stream containing the chartN.xml file
  @param  AIndent  Count of indentation spaces, for better legibility
  @param  ALegend  Chart legend which is processed here
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartLegendNode(AStream: TStream;
  AIndent: Integer; ALegend: TsChartLegend);
var
  indent: String;
  indent2: String;
  overlay: String = '';
  legendPos: String = '';
  formatStr: String = '';
  fontStr: String = '';
begin
  if not ALegend.Visible then
    exit;

  indent := DupeString(' ', AIndent);
  indent2 := indent + '  ';

  // Legend position
  legendPos := indent2 + Format('<c:legendPos val="%s"/>', [LEGEND_POS[ALegend.Position]]) + LE;

  // Inside/outside plot area?
  overlay := indent2 + Format('<c:overlay val="%d"/>', [FALSE_TRUE[ALegend.CanOverlapPlotArea]]) + LE;

  // Background and border formatting
  formatStr :=
    indent2 + '<c:spPr>' + LE +
    GetChartFillAndLineXML(AIndent + 4, ALegend.Chart, ALegend.Background, ALegend.Border) + LE +
    indent2 + '</c:spPr>' + LE;

  // Font of text items
  fontStr :=
    indent2 + '<c:txPr>' + LE +
    indent2 + '  <a:bodyPr/>' + LE +
    indent2 + '  <a:lstStyle/>' + LE +
    indent2 + '  <a:p>' + LE +
    indent2 + '    <a:pPr>' + LE +
    GetChartFontXML(AIndent + 6, ALegend.Font, 'a:defRPr') + LE +
    indent2 + '    </a:pPr>' + LE +
    indent2 + '  </a:p>' + LE +
    indent2 + '</c:txPr>' + LE;

  // Write out
  AppendToStream(AStream,
    indent + '<c:legend>' + LE +
    legendPos +
    overlay +
    formatStr +
    fontStr +
    indent + '</c:legend>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Assembles the child nodes of the <c:marker> node of a scatter series as a
  string
-------------------------------------------------------------------------------}
function TsSpreadOOXMLChartWriter.GetChartSeriesMarkerXML(AIndent: Integer;
  AChart: TsChart; AShowSymbols: Boolean; ASymbolKind: TsChartSeriesSymbol = cssRect;
  ASymbolWidth: Double = 3.0; ASymbolHeight: Double = 3.0;
  ASymbolFill: TsChartFill = nil; ASymbolBorder: TsChartLine = nil): String;
var
  indent: String;
  markerStr: String;
  symbolSizePts: Integer;
begin
  indent := DupeString(' ', AIndent);

  if not AShowSymbols then
  begin
    Result := indent + '<c:symbol val="none"/>';
    exit;
  end;

  case ASymbolKind of
    cssRect: markerStr := 'square';
    cssDiamond: markerStr := 'diamond';
    cssTriangle: markerStr := 'triangle';
    cssTriangleDown: markerStr := 'triangle';  // !!!!
    cssTriangleLeft: markerStr := 'triangle';  // !!!!
    cssTriangleRight: markerStr := 'triangle';  // !!!!
    cssCircle: markerStr := 'circle';
    cssStar: markerStr := 'star';
    cssX: markerstr := 'x';
    cssPlus: markerStr := '+';
    cssAsterisk: markerStr := 'star';  // !!!
    cssDash: markerStr := 'dash';
    cssDot: markerStr := 'dot';
    else markerStr := 'star';  // !!!
  end;                  // The symbols marked by !!! are not available in Excel

  symbolSizePts := round(mmToPts((ASymbolWidth + ASymbolHeight)/2));
  if symbolSizePts > 72 then symbolSizePts := 72;

  Result := Format(
    indent + '<c:symbol val="%s"/>' + LE +
    indent + '<c:size val="%d"/>' + LE +
    indent + '<c:spPr>' + LE +
               GetChartFillAndLineXML(AIndent + 2, AChart, ASymbolFill, ASymbolBorder) + LE +
    indent + '</c:spPr>',
    [ markerStr, symbolSizePts ]
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes the <c:plotArea> node. It contains the series and axes as subnodes.

  @param  AStream   Stream for the chartN.xml file
  @param  AIndent   Count of indentation spaces, for better legibility
  @param  AChart    Chart which is being processed here
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartPlotAreaNode(AStream: TStream;
  AIndent: Integer; AChart: TsChart);
var
  indent: String;
  i: Integer;
  xAxKind, yAxKind, x2AxKind: String;
  hasSecondaryAxis: Boolean = false;
begin
  indent := DupeString(' ', AIndent);

  FAxisID[caaBottom] := Random(MaxInt);
  FAxisID[caaLeft] := Random(MaxInt);
  FAxisID[caaRight] := Random(MaxInt);
  FAxisID[caaTop] := Random(MaxInt);
  FSeriesIndex := 0;

  AppendToStream(AStream,
    indent + '<c:plotArea>' + LE
  );

  // Write series attached to primary y axis
  WriteChartSeries(AStream, AIndent + 2, AChart, calPrimary, xAxKind);

  // Write series attached to secondary y axis
  hasSecondaryAxis := WriteChartSeries(AStream, AIndent + 2, AChart, calSecondary, x2AxKind);

  // Write the x and y axes. No axes for pie series and related.
  if not (AChart.GetChartType in [ctPie, ctRing]) then
  begin
    yAxKind := 'c:valAx';
    WriteChartAxisNode(AStream, AIndent, AChart.XAxis, xAxKind);
    WriteChartAxisNode(AStream, AIndent, AChart.YAxis, yAxKind);

    // Write the secondary axes
    if hasSecondaryAxis then begin
      WriteChartAxisNode(AStream, AIndent, AChart.Y2Axis, yAxKind);
      x2AxKind := xAxKind;
      WriteChartAxisNode(AStream, AIndent, AChart.X2Axis, x2Axkind);
    end;
  end;

  // Write the plot area background
  AppendToStream(AStream,
    indent + '  <c:spPr>' + LE +
             GetChartFillAndLineXML(AIndent + 4, AChart, AChart.PlotArea.Background, AChart.PlotArea.Border) + LE +
    indent + '  </c:spPr>' + LE
  );

  AppendToStream(AStream,
    indent + '</c:plotArea>' + LE
  );
end;

{ ------------------------------------------------------------------------------
  Writes a cell range to the xml stream

  @param  AStream   Stream to which the range is written (chartX.xml)
  @param  AIndent   Number of indentation spaces, for better legibility
  @param  ARange    Reference to the cell range to be written
  @param  ANodeName Name to be used for the node into which the data are embedded
  @param  ARefName  Identification of the data type in the range, needed by Excel
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartRange(AStream: TStream;
  AIndent: Integer; ARange: TsChartRange; ANodeName, ARefName: String;
  WriteCache: Boolean = false);
var
  indent: String;
  rangeStr: String;
  r, c, idx: Integer;
  chart: TsChart;
  workbook: TsWorkbook;
  sheet: TsWorksheet;
begin
  indent := DupeString(' ', AIndent);
  chart := ARange.Chart;
  if ARange.Sheet1 <> '' then
  begin
    workbook := TsWorkbook(Writer.Workbook);
    sheet := workbook.GetWorksheetByName(ARange.Sheet1);
  end else
    sheet := TsWorksheet(chart.Worksheet);

  if ARange.Sheet1 = ARange.Sheet2 then
    rangeStr := ARange.GetSheet1Name + '!' + GetCellRangeString(ARange.Row1, ARange.Col1, ARange.Row2, ARange.Col2, [])
  else
    rangeStr := GetCellRangeString(ARange.GetSheet1Name, ARange.GetSheet2Name, ARange.Row1, ARange.Col1, ARange.Row2, ARange.Col2, []);

  AppendToStream(AStream, Format(
    indent + '<%s>' + LE +
    indent + '  <%s>' + LE +
    indent + '    <c:f>%s</c:f>' + LE,
    [ ANodeName, ARefName, rangeStr ]
  ));

  if WriteCache then
  begin
    // Number cache
    if (ARange.GetSheet1Name = ARange.GetSheet2Name) and (ARefName = 'c:numRef') then
    begin
      AppendToStream(AStream, Format(
        indent + '    <c:numCache>' + LE +
        indent + '      <c:ptCount val="%d"/>' + LE,
        [ ARange.NumCells ]
      ));
      idx := 0;
      // Column range
      if (ARange.Col1 = ARange.Col2) then
      begin
        for r := ARange.Row1 to ARange.Row2 do
        begin
          WriteCellNumberValue(AStream, AIndent + 6, sheet, r, ARange.Col1, idx);
          inc(idx);
        end
      end else
      // Row range
      if (ARange.Row1 = ARange.Row2) then
      begin
        for c := ARange.Col1 to ARange.Col2 do
        begin
          WriteCellNumberValue(AStream, AIndent, sheet, ARange.Row1, c, idx);
          inc(idx);
        end
      end;
      AppendToStream(AStream,
        indent + '    </c:numCache>' + LE
      );
    end;
  end;

  AppendToStream(AStream, Format(
    indent + '  </%s>' + LE +
    indent + '</%s>' + LE,
    [ ARefName,  ANodeName ]
  ));
end;

procedure TsSpreadOOXMLChartWriter.WriteCharts;
var
  i: Integer;
begin
  for i := 0 to TsWorkbook(Writer.Workbook).GetChartCount - 1 do
  begin
  //  WriteChartRelsXML(FSChartRels[i], i);
  //  WriteChartStylesXML(FSChartStyles[i], i);
  //  WriteChartColorsXML(FSChartColors[i], i);
    WriteChartSpaceXML(FSCharts[i], i);
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the nodes for all series which are attached to the same y axis.

  @param  AStream  Stream of the chartN.xml file
  @param  AIndent  Number of indentation spaces, for better legibility
  @param  AChart   Chart from which the series are taken
  @param  AxisLink Type of the y axis (primary or secondary) which selects the series handled in this run.
  @param  xAxKind  Returns the type name of the x axis: <c:catAx> for a category, <c:valAx> for a value axis

  @returns When no series were attached to the specified axis the function returns false
-------------------------------------------------------------------------------}
function TsSpreadOOXMLChartWriter.WriteChartSeries(AStream: TStream;
  AIndent: Integer; AChart: TsChart; AxisLink: TsChartAxisLink;
  out xAxKind: string): Boolean;
var
  i, j, n: Integer;
  ser: TsChartSeries;
  axisGroup: Array of Integer = nil;
  posInGroup: Integer;
begin
  Result := false;

  // Collect all series attached to the same y axis, depending on PrimaryAxis parameter.
  SetLength(axisGroup, AChart.Series.Count);
  n := 0;
  for i := 0 to AChart.Series.Count-1 do
  begin
    ser := TsChartSeries(AChart.Series[i]);
    if (AxisLink = calPrimary) and (ser.YAxis = calPrimary) then
    begin
      axisGroup[n] := i;
      inc(n);
    end;
    if (AxisLink = calSecondary) and (ser.YAxis = calSecondary) then
    begin
      axisGroup[n] := i;
      inc(n);
    end;
  end;
  SetLength(axisGroup, n);

  if n = 0 then
    exit;

  xAxKind := 'c:catAx';

  for i := 0 to High(axisGroup) do
  begin
    j := axisGroup[i];
    ser := TsChartSeries(AChart.Series[j]);

    posInGroup := 0;
    if i = 0 then
      posInGroup := posInGroup or 1;
    if i = High(axisGroup) then
      posInGroup := posInGroup or 2;

    case ser.ChartType of
      ctArea:
        WriteAreaSeries(AStream, AIndent + 2, TsAreaSeries(ser), j, posInGroup);
      ctBar:
        WriteBarSeries(AStream, AIndent + 2, TsBarSeries(ser), j, posInGroup);
      ctBubble:
        begin
          WriteBubbleSeries(AStream, AIndent + 2, TsBubbleSeries(ser), posInGroup);
          xAxKind := 'c:valAx';
        end;
      ctLine:
        WriteLineSeries(AStream, AIndent + 2, TsLineSeries(ser), j, posInGroup);
      ctPie, ctRing:
        WritePieSeries(AStream, AIndent + 2, TsPieSeries(ser));
      ctRadar, ctFilledRadar:
        WriteRadarSeries(AStream, AIndent + 2, TsRadarSeries(ser));
      ctScatter:
        begin
          WriteScatterSeries(AStream, AIndent + 2, TsScatterSeries(ser), posInGroup);
          xAxKind := 'c:valAx';
        end;
      ctStock:
        begin
          WriteStockSeries(AStream, AIndent + 2, TsStockSeries(ser), posInGroup);
          xAxKind := 'c:dateAx';
          if TsStockSeries(ser).CandleStick then inc(FSeriesIndex, 3) else inc(FSeriesIndex, 2);
          // Together with the following inc(FSeriesIndex) we increment FSeriesIndex by 4 or 3 here.
        end;
    end;
    inc(FSeriesIndex);
  end;
  Result := true;
end;


procedure TsSpreadOOXMLChartWriter.WritePieSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsPieSeries);
var
  indent: String;
  chart: TsChart;
  nodeName: String;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  if ASeries.InnerRadiusPercent > 0 then
    nodeName := 'c:doughnutChart'
  else
    nodeName := 'c:pieChart';

  AppendToStream(AStream,
    indent + '<' + nodeName + '>' + LE +
    indent + '  <c:varyColors val="1"/>' + LE
  );

  WriteChartSeriesNode(AStream, AIndent + 4, ASeries);

  if ASeries.InnerRadiusPercent > 0 then
    AppendToStream(AStream, Format(
      indent + '<c:holeSize val="%d"/>' + LE, [ASeries.InnerRadiusPercent ]
    ));

  AppendToStream(AStream, Format(
    indent + '<c:firstSliceAng val="%d"/>' + LE,
    [ (90 - ASeries.StartAngle) mod 360 ]
  ));

  AppendToStream(AStream,
    indent + '</' + nodeName + '>' + LE
  );
end;

procedure TsSpreadOOXMLChartWriter.WriteRadarSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsRadarSeries);
var
  indent: String;
  chart: TsChart;
  xAxis: TsChartAxis;
  radarStyle: String;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  if ASeries.ChartType = ctFilledRadar then
    radarStyle := 'filled'
  else
    radarStyle := 'marker';

  if ASeries.YAxis = calPrimary then
    xAxis := chart.XAxis
  else
    xAxis := chart.X2Axis;

  AppendToStream(AStream,
    indent + '<c:radarChart>' + LE +
    indent + '  <c:radarStyle val="' + radarStyle + '"/>' + LE
  );

  WriteChartSeriesNode(AStream, AIndent + 4, ASeries);

  AppendToStream(AStream, Format(
    indent + '  <c:axId val="%d"/>' + LE +
    indent + '  <c:axId val="%d"/>' + LE +
    indent + '</c:radarChart>' + LE,
    [
      FAxisID[xAxis.Alignment],             // <c:axId>
      FAxisID[ASeries.GetYAxis.Alignment]   // <c:axId>
    ]
  ));
end;

procedure TsSpreadOOXMLChartWriter.WriteScatterSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsScatterSeries; APosInAxisGroup: Integer);
var
  indent: String;
  chart: TsChart;
  xAxis: TsChartAxis;
  scatterStyleStr: String;
  isFirstOfGroup: Boolean = true;
  isLastOfGroup: Boolean = true;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  isFirstOfGroup := (APosInAxisGroup and 1 = 1);
  isLastOfGroup := (APosInAxisGroup and 2 = 2);

  case chart.Interpolation of
    ciLinear:
      scatterStyleStr := 'lineMarker';
    ciCubicSpline, ciBSpline:
      scatterStyleStr := 'smoothMarker';
    else
      //ciStepStart, ciStepEnd, ciCenterX, ciCenterY
      scatterStyleStr := 'lineMarker';      // better than nothing...
  end;

  if isFirstOfGroup then
    AppendToStream(AStream,
      indent + '<c:scatterChart>' + LE +
      indent + '  <c:varyColors val="0"/>' + LE +
      indent + '  <c:scatterStyle val="' + scatterStyleStr + '"/>' + LE
    );

  WriteChartSeriesNode(AStream, AIndent + 4, ASeries);

  if isLastOfGroup then
  begin
    if ASeries.YAxis = calPrimary then
      xAxis := chart.XAxis
    else
      xAxis := chart.X2Axis;

    AppendToStream(AStream, Format(
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '</c:scatterChart>' + LE,
      [
        FAxisID[xAxis.Alignment],  // <c:axId>
        FAxisID[ASeries.GetYAxis.Alignment]   // <c:axId>
      ]
    ));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes a <c:stockChart> node for a stock series.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteStockSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsStockSeries; APosInAxisGroup: Integer);
var
  indent: String;
  chart: TsChart;
  xAxis: TsChartAxis;
  isfirstOfGroup: Boolean;
  isLastOfGroup: Boolean;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  isFirstOfGroup := (APosInAxisGroup and 1 = 1);
  isLastOfGroup := (APosInAxisGroup and 2 = 2);

  if isFirstOfGroup then
    AppendToStream(AStream,
      indent + '<c:stockChart>' + LE
    );

  // Writes the ranges used by the open, high, low and close series of the chart.
  // Experiments show that at least the close series must be written with cache.
  // Otherwise the vertical range bar (from high to low) is not drawn by excel.
  if ASeries.CandleStick then
  begin
    WriteStockSeriesNode(AStream, AIndent + 2, ASeries, FSeriesIndex, OHLC_OPEN, false);
    WriteStockSeriesNode(AStream, AIndent + 2, ASeries, FSeriesIndex+1, OHLC_HIGH, false);
    WriteStockSeriesNode(AStream, AIndent + 2, ASeries, FSeriesIndex+2, OHLC_LOW, false);
    WriteStockSeriesNode(AStream, AIndent + 2, ASeries, FseriesIndex+3, OHLC_CLOSE, true);
  end else
  begin
    WriteStockSeriesNode(AStream, AIndent + 2, ASeries, FSeriesIndex, OHLC_LOW, false);
    WriteStockSeriesNode(AStream, AIndent + 2, ASeries, FSeriesIndex+1, OHLC_HIGH, false);
    WriteStockSeriesNode(AStream, AIndent + 2, ASeries, FSeriesIndex+2, OHLC_CLOSE, true);
  end;

  if isLastOfGroup then
  begin
    AppendToStream(AStream,
      indent + '  <c:hiLowLines>' + LE +
      indent + '    <c:spPr>' + LE +
      GetChartLineXML(AIndent + 6, chart, ASeries.RangeLine) + LE +
      indent + '    </c:spPr>' + LE +
      indent + '  </c:hiLowLines>' + LE
    );

    if ASeries.CandleStick then
      AppendToStream(AStream,
        indent + '  <c:upDownBars>' + LE +
        indent + '    <c:gapWidth val="150"/>' + LE +
        indent + '    <c:upBars>' + LE +
        indent + '      <c:spPr>' + LE +
        GetChartFillAndLineXML(AIndent + 6, chart, ASeries.CandleStickUpFill, ASeries.CandleStickUpBorder) + LE +
        indent + '      </c:spPr>' + LE +
        indent + '    </c:upBars>' + LE +
        indent + '    <c:downBars>' + LE +
        indent + '      <c:spPr>' + LE +
        GetChartFillAndLineXML(AIndent + 6, chart, ASeries.CandleStickDownFill, ASeries.CandleStickDownBorder) + LE +
        indent + '      </c:spPr>' + LE +
        indent + '    </c:downBars>' + LE +
        indent + '  </c:upDownBars>' + LE
      );

    if ASeries.YAxis = calPrimary then
      xAxis := chart.XAxis
    else
      xAxis := chart.X2Axis;

    AppendToStream(AStream, Format(
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '</c:stockChart>' + LE,
      [
        FAxisID[xAxis.Alignment],  // <c:axId>
        FAxisID[ASeries.GetYAxis.Alignment]   // <c:axId>
      ]
    ));
  end;
end;

{ OHLCLPart: 0=open, 1=high, 2=low, 3=close }
procedure TsSpreadOOXMLChartWriter.WriteStockSeriesNode(AStream: TStream;
  AIndent: Integer; ASeries: TsStockSeries; ASeriesIndex, OHLCPart: Integer;
  WriteCache: Boolean);
var
  indent: String;
  chart: TsChart;
  markerStr: String;
  xRng, yRng: TsChartRange;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  AppendToStream(AStream, Format(
    indent + '<c:ser>' + LE +
    indent + '  <c:idx val="%d"/>' + LE +
    indent + '  <c:order val="%d"/>' + LE,
    [ ASeriesIndex, ASeriesIndex]
  ));

  // No line connecting the data points
  AppendToStream(AStream,
    indent + '  <c:spPr>' + LE +
    indent + '    <a:ln>' + LE +
    indent + '      <a:noFill/>' + LE +
    indent + '    </a:ln>' + LE +
    indent + '  </c:spPr>' + LE
  );

  // Marker
  if ASeries.CandleStick or (OHLCPart <> OHLC_CLOSE) then
    markerStr := GetChartSeriesMarkerXML(AIndent + 4, chart, false) // no marker
  else
    markerStr := GetChartSeriesMarkerXML(AIndent + 4, chart, true, cssDot, 10, 10, nil, ASeries.Line);

  AppendToStream(AStream,
    indent + '  <c:marker>' + LE +
                 markerStr + LE +
    indent + '  </c:marker>' + LE
  );

  // x range
  xRng := ASeries.XRange;
  if xRng.IsEmpty then
    xRng := ASeries.LabelRange;
  if xRng.IsEmpty then
    xRng := chart.CategoryLabelRange;
  WriteChartRange(AStream, AIndent + 2, xRng, 'c:cat', 'c:numRef');

  // y range
  case OHLCPart of
    OHLC_OPEN: yRng := ASeries.OpenRange;
    OHLC_HIGH: yRng := ASeries.HighRange;
    OHLC_LOW: yRng := ASeries.LowRange;
    OHLC_CLOSE: yRng := ASeries.CloseRange;
  end;
  WriteChartRange(AStream, AIndent + 2, yRng, 'c:val', 'c:numRef', WriteCache);

  AppendToStream(AStream,
    indent + '  <c:smooth val="0"/>' + LE+
    indent + '</c:ser>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes the <c:trendline> node for the specified chart series if a trendline
  is activated.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartTrendline(AStream: TStream;
  AIndent: Integer; ASeries: TsChartSeries);
var
  indent: String;
  trendline: TsChartTrendline;
  nameStr: String = '';
  orderStr: String = '';
  interceptStr: String = '';
  backwardStr: String = '';
  forwardStr: String = '';
begin
  trendline := TsOpenedTrendlineSeries(ASeries).Trendline;
  if trendline.TrendlineType = tltNone then
    exit;

  indent := DupeString(' ', AIndent);

  if trendline.Title <> '' then
    nameStr := Format(
      indent + '  <c:name>%s</c:name>' + LE, [trendline.Title]);

  if trendline.TrendlineType = tltPolynomial then
    orderStr := Format(
      indent + '  <c:order val="%d"/>' + LE, [trendline.PolynomialDegree]);

  if trendline.ForceYIntercept then
    interceptStr := Format(
      indent + '  <c:intercept val="%g"/>' + LE, [trendline.YInterceptValue], FPointSeparatorSettings);

  if trendline.ExtrapolateForwardBy <> 0 then
    forwardStr := Format(
      indent + '  <c:forward val="%g"/>' + LE, [trendline.ExtrapolateForwardBy], FPointSeparatorSettings);

  if trendline.ExtrapolateBackwardBy <> 0 then
    backwardStr := Format(
      indent + '  <c:backward val="%g"/>' + LE, [trendline.ExtrapolateBackwardBy], FPointSeparatorSettings);

  AppendToStream(AStream, Format(
      indent + '<c:trendline>' + LE +
                  nameStr +
      indent + '  <c:spPr>' + LE +
                   GetChartLineXML(AIndent + 4, ASeries.Chart, trendline.Line) + LE +
      indent + '  </c:spPr>' + LE +
      indent + '  <c:trendlineType val="%s"/>' + LE +
                  orderStr +
                  interceptStr +
                  forwardStr +
                  backwardStr +
      indent + '  <c:dispRSqr val="%d"/>' + LE +
      indent + '  <c:dispEq val="%d"/>' + LE +
      indent + '</c:trendline>' + LE,
      [ TRENDLINE_TYPES[trendline.TrendlineType],
        FALSE_TRUE[trendline.DisplayRSquare],
        FALSE_TRUE[trendline.DisplayEquation]
      ]
  ));
end;

{@@ ----------------------------------------------------------------------------
  Write series data point labels if requested. The corresponding node is
  <c:dLbls> underneath <c:ser>.

  @param  AStream      Stream of the chartN.xml file
  @param  AIndent      Number of indentation spaces, for better legibility
  @param  ASeries      Series to which the labels are attached
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartSeriesDatapointLabels(AStream: TStream;
  AIndent: Integer; ASeries: TsChartSeries);
var
  indent: String;
  fillAndLineFmt: String = '';
  labelPos: String = '';
  labelItems: String = '';
  separator: String = '';
  numFmt: String = '';
begin
  if ASeries.DataLabels = [] then
    exit;

  indent := DupeString(' ', AIndent);

  if (ASeries.LabelFormat <> '') then
    numFmt := indent + '  <c:numFmt formatCode="' + ASeries.LabelFormat + '" sourceLinked="0"/>' + LE;

  fillAndLineFmt := indent + '  <c:spPr>' + LE +
    GetChartFillAndLineXML(AIndent + 4, ASeries.Chart, ASeries.LabelBackground, ASeries.LabelBorder) + LE +
    indent + '</c:spPr>' + LE;

  labelPos := LABEL_POS[ASeries.LabelPosition];
  if labelPos <> '' then
    labelPos := indent + '  <c:dLblPos val="' + labelPos + '"/>' + LE;

  labelItems := Format(
    indent + '  <c:showLegendKey val="%d"/>' + LE +
    indent + '  <c:showVal val="%d"/>' + LE +
    indent + '  <c:showCatName val="%d"/>' + LE +
    indent + '  <c:showSerName val="%d"/>' + LE +
    indent + '  <c:showPercent val="%d"/>' + LE +
    indent + '  <c:showBubbleSize val="%d"/>' + LE +
    indent + '  <c:showLeaderLines val="%d"/>' + LE,
    [ FALSE_TRUE[cdlSymbol in ASeries.DataLabels],
      FALSE_TRUE[cdlValue in ASeries.DataLabels],
      FALSE_TRUE[cdlCategory in ASeries.DataLabels],
      FALSE_TRUE[cdlSeriesName in ASeries.DataLabels],
      FALSE_TRUE[cdlPercentage in ASeries.DataLabels],
      FALSE_TRUE[false],  // bubble size -- to do...
      FALSE_TRUE[cdlLeaderLines in ASeries.DataLabels]
    ]
  );

  separator := trim(ASeries.LabelSeparator);
  case ASeries.LabelSeparator of
    '\n', #10, #13, #13#10:
      separator := FPS_LINE_ENDING;  // Excel wants #10
    ' ':
      separator := '';    // space is default and not stored in the xml.
    else
      separator := ASeries.LabelSeparator;
  end;
  if separator <> '' then
    separator := indent + '  <c:separator>' + separator + '</c:separator>' + LE;

  AppendToStream(AStream,
    indent + '<c:dLbls>' + LE +
                numFmt +
                fillAndLineFmt +
                labelPos +
                labelItems +
                separator +
    indent + '</c:dLbls>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Write individual data point formatting to the chartN.xml stream. The information
  is written to <c:dPt> nodes underneath <c:ser>, one <c:dPt> node per data point.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartSeriesDataPointStyles(AStream: TStream;
  AIndent: Integer; ASeries: TsChartSeries);
var
  indent: String;
  i: Integer;
  dps: TsChartDatapointStyle;
  explosionStr: String;
begin
  indent := DupeString(' ', AIndent);
  for i := 0 to ASeries.DataPointStyles.Count-1 do
  begin
    dps := ASeries.DataPointStyles[i];
    explosionStr := '';
    if dps <> nil then
    begin
      if dps.PieOffset > 0 then
        explosionStr := Format('<c:explosion val="%d"/>', [dps.PieOffset]);
      AppendToStream(AStream,
        indent + '<c:dPt>' + LE +
        indent + '  <c:idx val="' + IntToStr(dps.DataPointIndex) + '"/>' + LE +
                    explosionStr +
        indent + '  <c:spPr>' + LE +
                      GetChartFillAndLineXML(AIndent + 4, ASeries.Chart, dps.Background, dps.Border) + LE +
        indent + '  </c:spPr>' + LE +
        indent + '</c:dPt>' + LE
      );
    end;
  end;
end;

procedure TsSpreadOOXMLChartWriter.WriteChartSeriesErrorBars(AStream: TStream;
  AIndent: Integer; ASeries: TsChartSeries; IsYError: Boolean);
var
  indent: String;
  errBars: TsChartErrorBars;
  errDir: String[1];
  errBarType: String;
  noEndCap: String[1];
  valType: String;
  value: Double;
begin
  if IsYError then
  begin
    errBars := ASeries.YErrorBars;
    errDir := 'y';
  end else
  begin
    errBars := ASeries.XErrorBars;
    errDir := 'x';
  end;

  case errbars.Kind of
    cebkNone:
      exit;
    cebkConstant:
      valType := 'fixedVal';
    cebkPercentage:
      valType := 'percentage';
    cebkCellRange:
      valType := 'cust';
    else
      Writer.Workbook.AddErrorMsg(Format('Unsupported %s error bar kind', [errDir]));
      exit;
  end;

  if errBars.ShowPos and errBars.ShowNeg then
    errBarType := 'both'
  else if errBars.ShowPos then
    errBarType := 'plus'
  else if errBars.ShowNeg then
    errBarType := 'minus'
  else
    exit;

  if errBars.ShowEndCap then
    noEndCap := '0'
  else
    noEndCap := '1';

  indent := DupeString(' ', AIndent);

  AppendToStream(AStream, Format(
    indent + '<c:errBars>' + LE +
    indent + '  <c:errDir val="%s"/>' + LE +
    indent + '  <c:errBarType val="%s"/>' + LE +
    indent + '  <c:errValType val="%s"/>' + LE +
    indent + '  <c:noEndCap val="%s"/>' + LE,
    [ errDir, errBarType, valType, noEndCap ]
  ));

  if errbars.Kind = cebkCellRange then
  begin
    if errBars.ShowPos then
      WriteChartRange(AStream, AIndent + 2, errBars.RangePos, 'c:plus', 'c:numRef');
    if errBars.ShowNeg then
      WriteChartRange(AStream, AIndent + 2, errBars.RangeNeg, 'c:minus', 'c:numRef');
  end else
  begin
    if errBars.ShowPos and errBars.ShowNeg then
      value := errBars.ValuePos
    else if errBars.ShowPos then
      value := errBars.ValuePos
    else if errBars.ShowNeg then
      value := errBars.ValueNeg;
    AppendToStream(AStream, Format(
      indent + '  <c:val val="%g"/>' + LE,
      [ value ], FPointSeparatorSettings ));
  end;

  AppendToStream(AStream,
    indent + '  <c:spPr>' + LE +
    GetChartFillAndLineXML(AIndent + 4, ASeries.Chart, nil, errBars.Line) + LE +
    indent + '  </c:spPr>' + LE
  );

  AppendToStream(AStream,
    indent + '</c:errBars>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes the <c:ser> node for the specified chart series
  Is called by all series types.

  @param  AStream       Stream of the chartN.xml file
  @param  AIndent       Number of indentation spaces, for better legibility
  @param  ASeries       Series to be written
  @param  ASeriesIndex  Index of ther series to be written
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartSeriesNode(AStream: TStream;
  AIndent: Integer; ASeries: TsChartSeries);
var
  indent: string;
  chart: TsChart;
  xRng, yRng: TsChartRange;
  forceNoLine: Boolean;
  xValName, yValName, xRefName, yRefName: String;
  lser: TsOpenedCustomLineSeries;
  smoothVal: Integer = 0;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  AppendToStream(AStream,
    indent + '<c:ser>' + LE
  );

  AppendToStream(AStream, Format(
    indent + '  <c:idx val="%d"/>' + LE +
    indent + '  <c:order val="%d"/>' + LE,
    [
      FSeriesIndex,   // <c:idx>
      FSeriesIndex    // <c:order>
    ]
  ));

  // Series title
  WriteChartSeriesTitle(AStream, AIndent + 2, ASeries);

  // Individual data point formats
  WriteChartSeriesDatapointStyles(AStream, AIndent + 2, ASeries);

  // Bubble series
  if (ASeries is TsBubbleSeries) then
  begin
    AppendToStream(AStream,
      indent + '  <c:spPr>' + LE +
      GetChartFillAndLineXML(AIndent + 4, chart, ASeries.Fill, ASeries.Line) + LE +
      indent + '  </c:spPr>' + LE
    );
  end else
  // Line & scatter & radar series: symbol markers
  if (ASeries is TsCustomLineSeries) then
  begin
    lSer := TsOpenedCustomLineSeries(ASeries);

    if (ASeries.ChartType = ctFilledRadar) then
      AppendToStream(AStream,
        indent + '  <c:spPr>' + LE +
        GetChartFillAndLineXML(AIndent + 4, chart, ASeries.Fill, ASeries.Line) + LE +
        indent + '  </c:spPr>' + LE
      )
    else
    begin
      forceNoLine := not lSer.ShowLines;
      AppendToStream(AStream,
        indent + '  <c:spPr>' + LE +
        GetChartLineXML(AIndent + 4, chart, ASeries.Line, forceNoLine) + LE +
        indent + '  </c:spPr>' + LE
      );
      if lSer.Interpolation in [ciCubicSpline, ciBSpline] then
        smoothVal := 1;
    end;
    AppendToStream(AStream,
      indent + '  <c:marker>' + LE +
      GetChartSeriesMarkerXML(AIndent + 4, chart, lser.ShowSymbols, lser.Symbol, lSer.SymbolWidth, lSer.SymbolWidth, lSer.SymbolFill, lSer.SymbolBorder) + LE +
      indent + '  </c:marker>' + LE
    );
  end else
    // Series main formatting
    AppendToStream(AStream,
      indent + '  <c:spPr>' + LE +
      GetChartFillAndLineXML(AIndent + 4, chart, ASeries.Fill, ASeries.Line) + LE +
      indent + '  </c:spPr>' + LE
    );

  // Error bars
  if (ASeries.ChartType in [ctArea, ctBar, ctLine, ctScatter]) then
  begin
    WriteChartSeriesErrorBars(AStream, AIndent + 2, ASeries, false);
    WriteChartSeriesErrorBars(AStream, AIndent + 2, ASeries, true);
  end;

  // Trend line
  if ASeries.SupportsTrendline then
    WriteChartTrendline(AStream, AIndent + 2, ASeries);

  // Data point labels
  WriteChartSeriesDatapointLabels(AStream, AIndent + 2, ASeries);

  // Cell ranges
  if (ASeries is TsScatterSeries) or (ASeries is TsBubbleSeries) then
  begin
    xRng := ASeries.XRange;
    xValName := 'c:xVal';
    yValName := 'c:yVal';
    xRefName := 'c:numRef';
  end else
  begin
    xRng := ASeries.XRange;
    if xRng.IsEmpty then
      xRng := ASeries.LabelRange;
    if xRng.IsEmpty then
      xRng := chart.CategoryLabelRange;
    xValName := 'c:cat';
    yValName := 'c:val';
    xRefName := 'c:strRef';
  end;
  yRng := ASeries.YRange;
  yRefName := 'c:numRef';

  // x range
  WriteChartRange(AStream, AIndent + 2, xRng, xValName, xRefName);

  // y range
  WriteChartRange(AStream, AIndent + 2, yRng, yValName, yRefName);

  // Bubble series: Bubble size range
  if (ASeries is TsBubbleSeries) then
    WriteChartRange(AStream, AIndent, TsBubbleSeries(ASeries).BubbleRange, 'c:bubbleSize', 'c:numRef');

  // Line series: Interpolation
  if ASeries is TsCustomLineSeries then
    AppendToStream(AStream,
      indent + '  <c:smooth val="' + IntToStr(smoothVal) + '"/>' + LE
    );

  AppendToStream(AStream,
    indent + '</c:ser>' + LE
  );
end;

procedure TsSpreadOOXMLChartWriter.WriteChartSeriesTitle(AStream: TStream;
  AIndent: Integer; ASeries: TsChartSeries);
var
  indent: String;
  cellAddr: String;
begin
  with ASeries.TitleAddr do
  begin
    if not IsUsed then
      exit;

    cellAddr := Format('%s!%s', [GetSheetName, GetCellString(Row, Col, []) ]);
  end;

  indent := DupeString(' ', AIndent);

  AppendToStream(AStream,
    indent + '<c:tx>' + LE +
    indent + '  <c:strRef>' + LE +
    indent + '    <c:f>' + cellAddr + '</c:f>' + LE +
    indent + '  </c:strRef>' + LE +
    indent + '</c:tx>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes a <c:tx> node containing either the chart title or axis title.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartText(AStream: TStream;
  AIndent: Integer; AText: TsChartText; ARotationAngle: Single);
var
  indent: String;
  rotStr: String;
begin
  if not AText.Visible then
    exit;

  str(-ARotationAngle * ANGLE_MULTIPLIER:0:0, rotStr);

  indent := DupeString(' ', AIndent);
  AppendToStream(AStream,
    indent + '<c:tx>' + LE +
    indent + '  <c:rich>' + LE +
    indent + '    <a:bodyPr rot="' + rotStr + '"/>' + LE +
    indent + '    <a:p>' + LE +
    indent + '      <a:pPr>' + LE +
                      GetChartFontXML(AIndent + 8, AText.Font, 'a:defRPr') + LE +
    indent + '      </a:pPr>' + LE +
    indent + '      <a:r>' + LE +
                      GetChartFontXML(AIndent + 8, AText.Font, 'a:rPr') + LE +
    indent + '        <a:t>' + AText.Caption + '</a:t>' + LE +
    indent + '      </a:r>' + LE +
    indent + '    </a:p>' + LE +
    indent + '  </c:rich>' + LE +
    indent + '</c:tx>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes the <c:title> node defining the chart's title

  @param  AStream   Stream to receive the data
  @param  AIndent   Count of indentation spaces, fr better legibility
  @param  ATitle    Title of the chart which is being processed here
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartTitleNode(AStream: TStream;
  AIndent: Integer; ATitle: TsChartText);
var
  indent: String;
begin
  if not ATitle.Visible or (ATitle.Caption = '') then
    exit;

  indent := DupeString(' ', AIndent);
  AppendToStream(AStream,
    indent + '<c:title>' + LE
  );

  WriteChartText(AStream, AIndent + 2, ATitle, ATitle.RotationAngle);

  AppendToStream(AStream,
    indent + '  <c:overlay val="0"/>' + LE +
                GetChartFillAndLineXML(AIndent + 2, ATitle.Chart, ATitle.Background, ATitle.Border) + LE
  );

  AppendToStream(AStream,
    indent + '</c:title>' + LE
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes the properties of the given line series to the <c:plotArea> node of
  file chartN.xml
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteLineSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsLineSeries; ASeriesIndex, APosInAxisGroup: Integer);
const
  GROUPING: Array[TsChartStackMode] of string = ('standard', 'stacked', 'percentStacked');
var
  indent: String;
  chart: TsChart;
  xAxis: TsChartAxis;
  isFirstOfGroup: Boolean;
  isLastOfGroup: Boolean;
  prevSeriesGroupIndex: Integer = -1;
  nextSeriesGroupIndex: Integer = -1;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  if (ASeriesIndex > 0) and (chart.Series[ASeriesIndex-1].YAxis = ASeries.YAxis) then
    prevSeriesGroupIndex := chart.Series[ASeriesIndex-1].GroupIndex;
  if (ASeriesIndex < chart.Series.Count-1) and (chart.Series[ASeriesIndex+1].YAxis = ASeries.YAxis) then
    nextSeriesGroupIndex := chart.Series[ASeriesIndex+1].GroupIndex;

  isFirstOfGroup := APosInAxisGroup and 1 = 1;
  isLastOfGroup := APosInAxisgroup and 2 = 2;

  if ((ASeries.GroupIndex > -1) and (prevSeriesGroupIndex = ASeries.GroupIndex)) then
    isFirstOfGroup := false;
  if ((ASeries.GroupIndex > -1) and (nextSeriesGroupIndex = ASeries.GroupIndex)) then
    isLastOfGroup := false;

  if isFirstOfGroup then
    AppendToStream(AStream, Format(
      indent + '<c:lineChart>' + LE +
      indent + '  <c:grouping val="%s"/>' + LE,
      [ GROUPING[chart.StackMode] ]
    ));

  WriteChartSeriesNode(AStream, AIndent + 2, ASeries);

  if isLastOfGroup then
  begin
    if ASeries.YAxis = calPrimary then
      xAxis := chart.XAxis
    else
      xAxis := chart.X2Axis;

    AppendToStream(AStream, Format(
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '</c:lineChart>' + LE,
      [
        FAxisID[xAxis.Alignment],  // <c:axId>
        FAxisID[ASeries.GetYAxis.Alignment]   // <c:axId>
      ]
    ));
  end;
end;

{$ENDIF}

end.

