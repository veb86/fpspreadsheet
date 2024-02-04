unit xlsxooxmlChart;

{$mode objfpc}{$H+}
{$include ..\fps.inc}

interface

{$ifdef FPS_CHARTS}

uses
  Classes, SysUtils, StrUtils, Contnrs, FPImage, fgl,
  {$ifdef FPS_PATCHED_ZIPPER}fpszipper,{$else}zipper,{$endif}
  laz2_xmlread, laz2_DOM,
  fpsTypes, fpSpreadsheet, fpsChart, fpsUtils, fpsNumFormat, fpsImages,
  fpsReaderWriter, fpsXMLCommon;

type
  { TsSpreadOOXMLChartReader }

  TsSpreadOOXMLChartReader = class(TsBasicSpreadChartReader)
  private
    FPointSeparatorSettings: TFormatSettings;
    FColors: specialize TFPGMap<string, TsColor>;
    FImages: TFPObjectList;
    FXAxisID, FYAxisID, FX2AxisID, FY2AxisID: DWord;
    FXAxisDelete, FYAxisDelete, FX2AxisDelete, FY2AxisDelete: Boolean;

    function ReadChartColor(ANode: TDOMNode; ADefault: TsColor): TsColor;
    procedure ReadChartColor(ANode: TDOMNode; var AColor: TsColor; var Alpha: Double);
    procedure ReadChartFillAndLineProps(ANode: TDOMNode;
      AChart: TsChart; AFill: TsChartFill; ALine: TsChartLine);
    procedure ReadChartFontProps(ANode: TDOMNode; AFont: TsFont);
    procedure ReadChartGradientFillProps(ANode: TDOMNode; AChart: TsChart; AFill: TsChartFill);
    procedure ReadChartHatchFillProps(ANode: TDOMNode; AChart: TsChart; AFill: TsChartFill);
    procedure ReadChartImageFillProps(ANode: TDOMNode; AChart: TsChart; AFill: TsChartFill);
    procedure ReadChartLineProps(ANode: TDOMNode; AChart: TsChart; AChartLine: TsChartLine);
    procedure ReadChartTextProps(ANode: TDOMNode; AFont: TsFont; var AFontRotation: Single);
    procedure SetAxisDefaults(AWorkbookAxis: TsChartAxis);
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
    procedure ReadChartSeriesRange(ANode: TDOMNode; ARange: TsChartRange);
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
    function GetChartFillAndLineXML(AIndent: Integer; AChart: TsChart; AFill: TsChartFill; ALine: TsChartLine): String;
    function GetChartFillXML(AIndent: Integer; AChart: TsChart; AFill: TsChartFill): String;
    function GetChartFontXML(AIndent: Integer; AFont: TsFont; ANodeName: String): String;
    function GetChartLineXML(AIndent: Integer; AChart: TsChart; ALine: TsChartLine; OverrideOff: Boolean = false): String;
    function GetChartRangeXML(AIndent: Integer; ARange: TsChartRange; ARefKind: String): String;
    function GetChartSeriesMarkerXML(AIndent: Integer; ASeries: TsCustomLineSeries): String;

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
    procedure WriteChartRegression(AStream: TStream; AIndent: Integer; ASeries: TsChartSeries);
    procedure WriteChartSeriesNode(AStream: TStream; AIndent: Integer; ASeries: TsChartSeries; ASeriesIndex: Integer);
    procedure WriteChartSeriesTitle(AStream: TStream; AIndent: Integer; ASeries: TsChartSeries);
    procedure WriteChartTitleNode(AStream: TStream; AIndent: Integer; ATitle: TsChartText);

    // Writing the nodes of the series types
    procedure WriteAreaSeries(AStream: TStream; AIndent: Integer; ASeries: TsAreaSeries; ASeriesIndex: Integer);
    procedure WriteBarSeries(AStream: TStream; AIndent: Integer; ASeries: TsBarSeries; ASeriesIndex: Integer);
    procedure WriteBubbleSeries(AStream: TStream; AIndent: Integer; ASeries: TsBubbleSeries; ASeriesIndex: Integer);
    procedure WritePieSeries(AStream: TStream; AIndent: Integer; ASeries: TsPieSeries; ASeriesIndex: Integer);
    procedure WriteScatterSeries(AStream: TStream; AIndent: Integer; ASeries: TsScatterSeries; ASeriesIndex: Integer);

    procedure WriteChartLabels(AStream: TStream; AIndent: Integer; AFont: TsFont);
    procedure WriteChartText(AStream: TStream; AIndent: Integer; AText: TsChartText);

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

  AX_POS: array[TsChartAxisAlignment] of string = ('l', 't', 'r', 'b');
  FALSE_TRUE: Array[boolean] of String = ('0', '1');
  LEGEND_POS: Array[TsChartLegendPosition] of string = ('r', 't', 'b', 'l');
  TRENDLINE_TYPES: Array[TsTrendlineType] of string = ('', 'linear', 'log', 'exp', 'power', 'poly');
    // 'movingAvg' and 'log' not supported, so far


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

  // The following color values are directly copied from xlsx files written by Excel.
  // In the long term, they should be read from xl/theme/theme1.xml.
  FColors := specialize TFPGMap<string, TsColor>.Create;
  FColors.Add('dk1', scBlack);
  FColors.Add('lt1', scWhite);
  FColors.Add('dk2', FlipColorBytes($44546A));
  FColors.Add('lt2', FlipColorBytes($E7E6E6));
  FColors.Add('accent1', FlipColorBytes($4472C4));
  FColors.Add('accent2', FlipColorBytes($ED7D31));
  FColors.Add('accent3', FlipColorBytes($A5A5A5));
  FColors.Add('accent4', FlipColorBytes($FFC000));
  FColors.Add('accent5', FlipColorBytes($5B9BD5));
  FColors.Add('accent6', FlipColorBytes($70AD47));

  FImages := TFPObjectList.Create;
end;

destructor TsSpreadOOXMLChartReader.Destroy;
begin
  FImages.Free;
  FColors.Free;
  inherited;
end;

procedure TsSpreadOOXMLChartReader.ReadChart(ANode: TDOMNode; AChart: TsChart);
var
  nodeName: String;
begin
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
          if s = 'm/d/yyyy' then
            AChartAxis.LabelFormat := FReader.Workbook.FormatSettings.ShortDateFormat
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
          if x = 1000 then x := 0;  // not sure, but maybe 1000 means: default
          AChartAxis.LabelRotation := x;
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
        {
        if TryStrToFloat(s, n, FPointSeparatorSettings) then
          ser.BarWidthPercent := round(100 / (1 + n/100));
          }
      'c:overlap':
        if TryStrToFloat(s, n, FPointSeparatorSettings) then
          AChart.BarOverlapPercent := round(n);
      {
        if TryStrToFloat(s, n, FPointSeparatorSettings) then
          ser.BarOffsetPercent := round(n);
      }
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
  smooth: Boolean;
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
          ser.BubbleSizeMode := bsmArea;  // Excel always plots the area of the bubbles
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:bubbleScale': ;
      'c:showNegBubbles': ;
      'c:varyColors':  ;
      'c:dLbls':
        ;
      'c:axId':
        ReadChartSeriesAxis(ANode, ser);
    end;
    ANode := ANode.NextSibling;
  end;
end;

function TsSpreadOOXMLChartReader.ReadChartColor(ANode: TDOMNode;
  ADefault: TsColor): TsColor;
var
  alpha: Double;
begin
  Result := ADefault;
  ReadChartColor(ANode, Result, alpha);
end;

procedure TsSpreadOOXMLChartReader.ReadChartColor(ANode: TDOMNode;
  var AColor: TsColor; var Alpha: Double);

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
  idx: Integer;
  n: Integer;
  child: TDOMNode;
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
            idx := FColors.IndexOf(ColorAlias(s));
            if idx > -1 then
            begin
              AColor := FColors.Data[idx];
              child := ANode.FirstChild;
              while Assigned(child) do
              begin
                nodeName := child.NodeName;
                s := GetAttrValue(child, 'val');
                case nodeName of
                  'a:tint':
                    if TryStrToInt(s, n) then
                      AColor := TintedColor(AColor, n/FACTOR_MULTIPLIER);
                  'a:lumMod':     // luminance modulated
                    if TryStrToInt(s, n) then
                      AColor := LumModColor(AColor, n/FACTOR_MULTIPLIER);
                  'a:lumOff':
                    if TryStrToInt(s, n) then
                      AColor := LumOffsetColor(AColor, n/FACTOR_MULTIPLIER);
                  'a:alpha':
                    if TryStrToInt(s, n) then
                      Alpha := n / 100000;
                end;
                child := child.NextSibling;
              end;
            end;
          end;
        end;
      'a:srgbClr':
        begin
          s := GetAttrValue(ANode, 'val');
          if s <> '' then
            AColor := HTMLColorStrToColor(s);
          child := ANode.FirstChild;
          while Assigned(child) do
          begin
            nodeName := child.NodeName;
            s := GetAttrValue(child, 'val');
            case nodeName of
              'a:alpha':
                if TryStrToInt(s, n) then
                  Alpha := n / FACTOR_MULTIPLIER;
            end;
            child := child.NextSibling;
          end;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartGradientFillProps(ANode: TDOMNode;
  AChart: TsChart; AFill: TsChartFill);
var
  nodeName, s: String;
  value, alpha: Double;
  color: TsColor;
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
              color := scWhite;
              alpha := 1.0;
              ReadChartColor(child.FirstChild, color, alpha);
              gradient.AddStep(value, color, 1.0 - alpha, 1.0);
            end;
            child := child.NextSibling;
          end;
        end;
      'a:lin':
        begin
          gradient.Style := cgsLinear;
          s := GetAttrValue(ANode, 'ang');
          if TryStrToFloat(s, value, FPointSeparatorSettings) then
            gradient.Angle := value / ANGLE_MULTIPLIER;
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
  color: TsColor;
begin
  AFill.Style := cfsSolidHatched;
  hatch := GetAttrValue(ANode, 'prst');

  ANode := ANode.FirstChild;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'a:fgClr':
        color := ReadChartColor(ANode.FirstChild, scBlack);
      'a:bgClr':
        AFill.Color := ReadChartColor(ANode.FirstChild, scWhite);
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

  alpha := 1.0;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      // Solid fill
      'a:solidFill':
        begin
          AFill.Style := cfsSolid;
          ReadChartColor(ANode.FirstChild, AFill.Color, alpha);
          AFill.Transparency := 1.0 - alpha;
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

      // Drawing effects
      'a:effectLst': ;
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
        AFont.Color := ReadChartColor(node.FirstChild, scBlack);

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
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:spPr':
        begin
          fill := TsChartFill.Create;  // will be destroyed by the chart!
          line := TsChartLine.Create;
          ReadChartFillAndLineProps(ANode.FirstChild, ASeries.Chart, fill, line);
          ASeries.DataPointStyles.AddFillAndLine(fill, line);
        end;
      'c:explosion':
        ; // in case of pie series: movement of individual sector away from center
    end;
    ANode := ANode.NextSibling;
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
                  AChartLine.Color := ReadChartColor(child.FirstChild, scBlack);
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
    case nodeName of
      'c:ser':
        begin
          ser := TsLineSeries.Create(AChart);
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
      'c:gapWidth':
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
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:ser':
        begin
          if RingMode then
            ser := TsRingSeries.Create(AChart)
          else
            ser := TsPieSeries.Create(AChart);
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:firstSliceAng':
        begin
          s := GetAttrValue(ANode, 'val');
          if TryStrToFloat(s, x, FPointSeparatorSettings) then
            ser.StartAngle := round(x) + 90;
        end;
      'c:holeSize':
        if RingMode then
        begin
          s := GetAttrValue(ANode, 'val');
          if TryStrToFloat(s, x, FPointSeparatorSettings) then
            TsRingSeries(ser).InnerRadiusPercent := round(x);
        end;
    end;
    ANode := ANode.NextSibling;
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
            SymbolWidth := PtsToMM(n div 2);
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
               end;
            1: begin
                 ReadChartAxis(workNode.FirstChild, AChart, AChart.X2Axis, FX2AxisID, FX2AxisDelete);
                 AChart.X2Axis.DateTime := true;
               end;
          end;
          inc(dateAxCounter);
          if (dateAxCounter > 1) and (AChart.X2Axis.Alignment = AChart.XAxis.Alignment) and FX2AxisDelete then
          begin
            // Force using only a single x axis in this case.
            FX2AxisID := FXAxisID;
            AChart.X2Axis.Visible := false;
          end;
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
            if (AChart.X2Axis.Alignment = AChart.XAxis.Alignment) and FX2AxisDelete then
            begin
              // Force using only a single x axis in this case.
              FX2AxisID := FXAxisID;
              AChart.X2Axis.Visible := false;
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
    end;
    workNode := workNode.NextSibling;
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
  nodeName: String;
  ser: TsRadarSeries;
  radarStyle: String = '';
begin
  if ANode = nil then
    exit;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:radarStyle':
        radarStyle := GetAttrValue(ANode, 'val');
      'c:ser':
        begin
          ser := TsRadarSeries.Create(AChart);
          ReadChartSeriesProps(ANode.FirstChild, ser);
          if radarStyle <> 'filled' then
            ser.Fill.Style := cfsNoFill;
        end;
    end;
    ANode := ANode.NextSibling;
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
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:scatterStyle':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s = 'smoothMarker') then
            AChart.Interpolation := ciCubicSpline;
        end;
      'c:varyColors':  ;
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
        ReadChartSeriesRange(node.FirstChild, errorBars.RangePos);
      'c:minus':
        ReadChartSeriesRange(node.FirstChild, errorBars.RangeNeg);
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
        ;
      'c:showLegendKey':
        if (s <> '') and (s <> '0') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlSymbol];
      'c:showVal':
        if (s <> '') and (s <> '0') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlValue];
      'c:showCatName':
        if (s <> '') and (s <> '0') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlCategory];
      'c:showSerName':
        if (s <> '') and (s <> '0') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlSeriesName];
      'c:showPercent':
        if (s <> '') and (s <> '0') then
          ASeries.DataLabels := ASeries.DataLabels + [cdlPercentage];
      'c:showBubbleSize':
        if (s <> '') and (s <> '0') and (ASeries is TsBubbleSeries) then
          ASeries.DataLabels := ASeries.DataLabels + [cdlValue];
      'c:showLeaderLines':
        ;
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
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartSeriesProps(ANode: TDOMNode; ASeries: TsChartSeries);
var
  nodeName, s: String;
  n: Integer;
begin
  if ANode = nil then
    exit;
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
      'c:cat':
        ReadChartSeriesRange(ANode.FirstChild, ASeries.LabelRange);
      'c:xVal':
        ReadChartSeriesRange(ANode.FirstChild, ASeries.XRange);
      'c:val', 'c:yVal':
        if ASeries.YRange.IsEmpty then  // TcStockSeries already has read the y range...
          ReadChartSeriesRange(ANode.FirstChild, ASeries.YRange);
      'c:bubbleSize':
        if ASeries is TsBubbleSeries then
          ReadChartSeriesRange(ANode.FirstChild, TsBubbleSeries(ASeries).BubbleRange);
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
      'c:invertIfNegative':
        ;
      'c:extLst':
        ;
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the cell range for a series.

  @@param  ANode   First child of a <c:val>, <c:yval> or <c:cat> node below <c:ser>.
  @@param  ARange  Cell range to which the range parameters will be assigned.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartSeriesRange(ANode: TDOMNode; ARange: TsChartRange);
var
  nodeName, s: String;
  sheet1, sheet2: String;
  r1, c1, r2, c2: Cardinal;
  flags: TsRelFlags;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    if (nodeName = 'c:strRef') or (nodeName = 'c:numRef') then
    begin
      ANode := ANode.FindNode('c:f');
      if ANode <> nil then
      begin
        s := GetNodeValue(ANode);
        if ParseCellRangeString(s, sheet1, sheet2, r1, c1, r2, c2, flags) then
        begin
          if sheet2 = '' then sheet2 := sheet1;
          ARange.Sheet1 := sheet1;
          ARange.Sheet2 := sheet2;
          ARange.Row1 := r1;
          ARange.Col1 := c1;
          ARange.Row2 := R2;
          ARange.Col2 := C2;
          exit;
        end;
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
  nodeName: String;
  sernode, child: TDOMNode;
  rangeLine: TsChartLine = nil;
begin
  if ANode = nil then
    exit;

  ser := TsStockSeries.Create(AChart);

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
                ReadChartSeriesRange(child.FirstChild, ser.LabelRange);
                }
              'c:val':
                if ser.CloseRange.IsEmpty then
                  ReadChartSeriesRange(child.FirstChild, ser.CloseRange)
                else if ser.LowRange.IsEmpty then
                  ReadChartSeriesRange(child.FirstChild, ser.LowRange)
                else if ser.HighRange.IsEmpty then
                  ReadChartSeriesRange(child.FirstChild, ser.HighRange)
                else if ser.OpenRange.IsEmpty then
                begin
                  ReadChartSeriesRange(child.FirstChild, ser.OpenRange);
                  ser.CandleStick := true;
                end;
            end;
            child := child.NextSibling;
          end;
        end;
    end;
    serNode := serNode.PreviousSibling;  // we must run backward
  end;

  ReadChartSeriesProps(ANode.FirstChild, ser);

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:hiLowLines':
        begin
          child := ANode.FirstChild;
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
    child := ANode.FirstChild;
    case nodeName of
      'c:gapWidth':
        begin
          s := GetAttrValue(ANode, 'val');
          if TryStrToFloat(s, n, FPointSeparatorSettings) then
            ASeries.TickWidthPercent := round(100 / (1 + n/100));
        end;
      'c:upBars':
        if Assigned(child) then
          ReadChartFillAndLineProps(child.FirstChild, ASeries.Chart, ASeries.CandleStickUpFill, ASeries.CandlestickUpBorder);
      'c:downBars':
        if Assigned(child) then
          ReadChartFillAndLineProps(child.FirstChild, ASeries.Chart, ASeries.CandleStickDownFill, ASeries.CandlestickDownBorder);
    end;
    ANode := ANode.NextSibling;
  end;
end;

{ Extracts the chart and axis titles, their formatting and their texts. }
procedure TsSpreadOOXMLChartReader.ReadChartTitle(ANode: TDOMNode; ATitle: TsChartText);
var
  nodeName, s, totalText: String;
  child, child2, child3, child4: TDOMNode;
  n: Integer;
begin
  if ANode = nil then
    exit;
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
                    s := GetAttrValue(ANode, 'rot');
                    if (s <> '') and TryStrToInt(s, n) then
                      ATitle.RotationAngle := -n / ANGLE_MULTIPLIER;
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
  AWorkbookAxis.LabelRotation := 0;
  AWorkbookAxis.Visible := false;
  AWorkbookAxis.MajorGridLines.Style := clsNoLine;
  AWorkbookAxis.MinorGridLines.Style := clsNoLine;
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
begin
  indent := DupeString(' ', AIndent);
  chart := TsWorkbook(Writer.Workbook).GetChartByIndex(AChartIndex);

  AppendToStream(AStream,
    indent + '<c:chart>' + LE
  );

  WriteChartTitleNode(AStream, AIndent + 2, chart.Title);
  WriteChartPlotAreaNode(AStream, AIndent + 2, chart);
  WriteChartLegendNode(AStream, AIndent + 2, chart.Legend);

  AppendToStream(AStream,
    indent  + '  <c:plotVisOnly val="1" />' + LE +
    indent + '</c:chart>' + LE
  );
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
  i: Integer;
  presetIdx: Integer;
begin
  indent := DupeString(' ', AIndent);

  if (AFill = nil) or (AFill.Style = cfsNoFill) then
    Result := indent + '<a:noFill/>'
  else
    case AFill.Style of
      cfsSolid:
        begin
          Result := Format(
            indent + '<a:solidFill>' + LE +
            indent + '  <a:srgbClr val="%s"/>' + LE +
            indent + '</a:solidFill>',
            [ HtmlColorStr(AFill.Color) ]
          );
        end;
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
            Result := Format(
              indent + '<a:pattFill prst="%s">' + LE +
              indent + '  <a:fgClr>' + LE +
              indent + '    <a:srgbClr val="%s"/>' + LE +
              indent + '  </a:fgClr>' + LE +
              indent + '  <a:bgClr>' + LE +
              indent + '    <a:srgbClr val="%s"/>' + LE +
              indent + '  </a:bgClr>' + LE +
              indent + '</a:pattFill>',
              [ HATCH_NAMES[presetIdx], HtmlColorStr(hatch.PatternColor), HtmlColorStr(AFill.Color) ]
            )
          else
            // unknown pattern - use a solid fill
            Result := Format(
              indent + '<a:solidFill>' + LE +
              indent + '  <a:srgbClr val="%s"/>' + LE +
              indent + '</a:solidFill>',
              [ HtmlColorStr(AFill.Color) ]
            );
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

  if (ALine <> nil) and (ALine.Style <> clsNoLine) and not OverrideOff then
  begin
    Result := Format(
      indent + '<a:ln w="%.0f">' + LE +
      indent + '  <a:solidFill>' + LE +
      indent + '    <a:srgbClr val="%s"/>' + LE +
      indent + '  </a:solidFill>' + LE,
      [ mmToPts(ALine.Width) * PTS_MULTIPLIER, HtmlColorStr(ALine.Color) ]
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
  end else
    Result := indent + '<a:ln>' + LE +
              indent + '  <a:noFill/>' + LE +
              indent + '</a:ln>';
end;

{@@ ----------------------------------------------------------------------------
  Creates an xml string for a cell range to be used in ARange

  @param  AIndent   Number of intentation spaces, for better legibility
  @param  ARange    Range containing the worksheet names and cell references
  @param  ARefKind  Determines whether the range contains strings ('c:strRef') or numbers ('c:numRef'). Must be a valid Excel nodename (including "c:")
-------------------------------------------------------------------------------}
function TsSpreadOOXMLChartWriter.GetChartRangeXML(AIndent: Integer;
  ARange: TsChartRange; ARefKind: String): String;
var
  indent: String;
begin
  indent := DupeString(' ', AIndent);
  if ARange.Sheet1 = ARange.Sheet2 then
    Result := ARange.GetSheet1Name + '!' + GetCellRangeString(ARange.Row1, ARange.Col1, ARange.Row2, ARange.Col2, [])
  else
    Result := GetCellRangeString(ARange.GetSheet1Name, ARange.GetSheet2Name, ARange.Row1, ARange.Col1, ARange.Row2, ARange.Col2, []);

  Result := Format(
    indent + '<%0:s> ' + LE +
    indent + '  <c:f>' + Result + '</c:f>' + LE +
    indent + '</%0:s>',
    [ ARefKind ]
  );
end;

{@@ ----------------------------------------------------------------------------
  Writes the properties of the given area series to the <c:plotArea> node of
  file chartN.xml
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteAreaSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsAreaSeries; ASeriesIndex: Integer);
const
  GROUPING: Array[TsChartStackMode] of string = ('standard', 'stacked', 'percentStacked');
var
  indent: String;
  chart: TsChart;
  isFirstOfGroup: Boolean = true;
  isLastOfGroup: Boolean = true;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  if ASeries.GroupIndex > -1 then
  begin
    if (ASeriesIndex > 0) and (chart.Series[ASeriesIndex-1].GroupIndex = ASeries.GroupIndex) then
      isfirstOfGroup := false;
    if (ASeriesIndex < chart.Series.Count-1) and (chart.Series[ASeriesIndex+1].GroupIndex = ASeries.GroupIndex) then
      isLastOfGroup := false;
  end;

  if isFirstOfGroup then
    AppendToStream(AStream, Format(
      indent + '<c:areaChart>' + LE +
      indent + '  <c:grouping val="%s"/>' + LE,
      [ GROUPING[chart.StackMode] ]
    ));

  WriteChartSeriesNode(AStream, AIndent + 2, ASeries, ASeriesIndex);

  if isLastOfGroup then
  begin
    AppendToStream(AStream, Format(
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '</c:areaChart>' + LE,
      [
        FAxisID[ASeries.Chart.XAxis.Alignment],  // <c:axId>
        FAxisID[ASeries.Chart.YAxis.Alignment]   // <c:axId>
      ]
    ));
  end;
end;

{@@ ----------------------------------------------------------------------------
  Writes the properties of the given bar series to the <c:plotArea> node of
  file chartN.xml
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteBarSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsBarSeries; ASeriesIndex: Integer);
const
  GROUPING: Array[TsChartStackMode] of string = ('clustered', 'stacked', 'percentStacked');
var
  indent: String;
  chart: TsChart;
  gapWidth: Integer = 0;
  overlap: Integer = 999;
  isFirstOfGroup: Boolean = true;
  isLastOfGroup: Boolean = true;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  if ASeries.GroupIndex > -1 then
  begin
    if (ASeriesIndex > 0) and (chart.Series[ASeriesIndex-1].GroupIndex = ASeries.GroupIndex) then
      isfirstOfGroup := false;
    if (ASeriesIndex < chart.Series.Count-1) and (chart.Series[ASeriesIndex+1].GroupIndex = ASeries.GroupIndex) then
      isLastOfGroup := false;
    if chart.StackMode <> csmSideBySide then
      overlap := 100
  end;

  if isFirstOfGroup then
    AppendToStream(AStream, Format(
      indent + '<c:barChart>' + LE +
      indent + '  <c:barDir val="col"/>' + LE +
      indent + '  <c:varyColors val="0"/>' + LE +
      indent + '  <c:grouping val="%s"/>' + LE,
      [ GROUPING[chart.StackMode] ]
    ));

  WriteChartSeriesNode(AStream, AIndent + 2, ASeries, ASeriesIndex);

  if isLastOfGroup then
  begin
    if overlap = 999 then
      overlap := chart.BarOverlapPercent;
    gapWidth := chart.BarGapWidthPercent;
    AppendToStream(AStream, Format(
      indent + '  <c:gapWidth val="%d"/>' + LE +
      indent + '  <c:overlap val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '  <c:axId val="%d"/>' + LE +
      indent + '</c:barChart>' + LE,
      [
        gapWidth,                                // <c:gapWidth>
        overlap,                                 // <c:overlap>
        FAxisID[ASeries.Chart.XAxis.Alignment],  // <c:axId>
        FAxisID[ASeries.Chart.YAxis.Alignment]   // <c:axId>
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
  @param  ASeriesIndex  Index of the series in the chart's series list
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteBubbleSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsBubbleSeries; ASeriesIndex: Integer);
var
  indent: String;
  chart: TsChart;
  diameterMode: String;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  AppendToStream(AStream,
    indent + '<c:bubbleChart>' + LE +
    indent + '  <c:varyColors val="0"/>' + LE
  );

  WriteChartSeriesNode(AStream, AIndent + 2, ASeries, ASeriesIndex);

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
      FAxisID[ASeries.Chart.XAxis.Alignment],  // <c:axId>
      FAxisID[ASeries.Chart.YAxis.Alignment]   // <c:axId>
    ]
  ));
  // Note:  <c:showNegBubbles> not supported
end;

{@@ ----------------------------------------------------------------------------
  Writes the properties of the given chart axis to the chartN.xml file under
  the <c:plotArea> node

  Depending on AxisKind, the node is either <c:catAx> or <c:valAx>.

  @param  AStream   Stream of the chartN.xml file
  @param  AIndent   Count of indentation spaces to increase readability
  @param  Axis      Chart axis processed
  @param  AxisKind  'catAx' when Axis is a category axis, otherwise 'valAx'
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
  axID: DWord;
  rotAxID: DWord;
  crosses: String = 'autoZero';
  logarithmic: String = '';
begin
  indent := DupeString(' ', AIndent);

  axID := FAxisID[Axis.Alignment];
  rotAxID := FAxisID[Axis.GetRotatedAxis.Alignment];
  if (Axis = Axis.Chart.YAxis) and (Axis.Chart.GetChartType in [ctBar]) then
    crosses := 'min';

  AppendToStream(AStream, Format(
    indent + '<%s>' + LE +
    indent + '  <c:axId val="%d"/>' + LE,
    [ ANodeName, axID ]
  ));

  WriteChartAxisScaling(AStream, AIndent + 2, Axis);

  AppendToStream(AStream, Format(
    indent + '  <c:delete val="0"/>' + LE +
    indent + '  <c:axPos val="%s" />' + LE,
    [ AX_POS[Axis.Alignment] ]
  ));

  // Grid lines
  AppendToStream(AStream,
    GetGridLineStr(AIndent + 2, 'c:majorGridlines', Axis.MajorGridLines) +
    GetGridLineStr(AIndent + 2, 'c:minorGridlines', Axis.MinorGridLines)
  );

  // Axis title
  WriteChartAxisTitle(AStream, AIndent + 2, Axis);

  // Axis labels
  if Axis.ShowLabels then
    WriteChartLabels(AStream, AIndent + 2, Axis.LabelFont);

  AppendToStream(AStream, Format(
    indent + '  <c:numFmt formatCode="General" sourceLinked="1"/>' + LE +
    indent + '  <c:majorTickMark val="%s"/>' + LE +
    indent + '  <c:minorTickMark val="%s"/>' + LE +
    indent + '  <c:tickLblPos val="nextTo"/>' + LE +
    indent + '  <c:crossAx val="%d" />' + LE +
    indent + '  <c:crosses val="%s"/>' + LE +
//    indent + '  <c:auto val="1"/>' + LE +
    indent + '</%s>' + LE,
    [
      GetTickMarkStr(Axis.MajorTicks),  // <c:majorTickMark>
      GetTickMarkStr(Axis.MinorTicks),  // <c:minorTickMark>
      rotAxID,                          // <c:crossAx>
      crosses,                          // <c:crosses>
      ANodeName                         // </c:catAx> or </c:valAx>
    ]
  ));
end;

procedure TsSpreadOOXMLChartWriter.WriteChartAxisScaling(AStream: TStream;
  AIndent: Integer; Axis: TsChartAxis);
var
  indent: String;
  intv: Double;
  logStr: String = '';
  maxStr: String = '';
  minStr: String = '';
begin
  indent := DupeString(' ', AIndent);

  if not Axis.AutomaticMax then
    maxStr := indent + Format('  <c:max val="%g"/>', [Axis.Max], FPointSeparatorSettings) + LE;

  if not Axis.AutomaticMin then
    minStr := indent + Format('  <c:min val="%g"/>', [Axis.Min], FPointSeparatorSettings) + LE;

  if Axis.Logarithmic then
    logStr := indent + Format('  <c:logBase val="%g"/>', [Axis.LogBase], FPointSeparatorSettings) + LE;

  AppendToStream(AStream,
    indent + '<c:scaling>' + LE +
                maxStr +
                minStr +
                logStr +
    indent + '  <c:orientation val="minMax"/>' + LE +
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
begin
  if not Axis.Title.Visible or (Axis.Title.Caption = '') then
    exit;

  indent := DupeString(' ', AIndent);

  AppendToStream(AStream,
    indent + '<c:title>' + LE
  );

  WriteChartText(AStream, AIndent + 4, Axis.Title);

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
  overlay := indent2 + Format('<c:overlay val="%s"/>', [FALSE_TRUE[ALegend.CanOverlapPlotArea]]) + LE;

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
  Assembles the <c:marker> node of a scatter series as a string
-------------------------------------------------------------------------------}
function TsSpreadOOXMLChartWriter.GetChartSeriesMarkerXML(AIndent: Integer;
  ASeries: TsCustomLineSeries): String;
var
  indent: String;
  markerStr: String;
  chart: TsChart;
  ser: TsOpenedCustomLineSeries;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;
  ser := TsOpenedCustomLineSeries(ASeries);

  if ser.ShowSymbols then
    case ser.Symbol of
      cssRect: markerStr := 'square';
      cssDiamond: markerStr := 'diamong';
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
    end
  else
    markerStr := 'none';

  Result := Format(
    indent + '<c:symbol val="%s"/>' + LE +
    indent + '<c:size val="%.0f"/>' + LE +
    indent + '<c:spPr>' + LE +
               GetChartFillAndLineXML(AIndent + 2, chart, ser.SymbolFill, ser.SymbolBorder) + LE +
    indent + '</c:spPr>',
    [ markerStr, mmToPts(ser.SymbolWidth + ser.SymbolHeight) ]
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
  ser: TsChartSeries;
  xAxKind, yAxKind: String;
begin
  indent := DupeString(' ', AIndent);
  FAxisID[caaBottom] := Random(MaxInt);
  FAxisID[caaLeft] := Random(MaxInt);

  AppendToStream(AStream,
    indent + '<c:plotArea>' + LE
  );

  xAxKind := 'c:catAx';
  yAxKind := 'c:valAx';

  for i := 0 to AChart.Series.Count-1 do
  begin
    ser := TsChartSeries(AChart.Series[i]);
    case ser.ChartType of
      ctArea:
        WriteAreaSeries(AStream, AIndent + 2, TsAreaSeries(ser), i);
      ctBar:
        WriteBarSeries(AStream, AIndent + 2, TsBarSeries(ser), i);
      ctBubble:
        begin
          WriteBubbleSeries(AStream, AIndent + 2, TsBubbleSeries(ser), i);
          xAxKind := 'c:valAx';
        end;
      ctPie:
        WritePieSeries(AStream, AIndent + 2, TsPieSeries(ser), i);
      ctScatter:
        begin
          WriteScatterSeries(AStream, AIndent + 2, TsScatterSeries(ser), i);
          xAxKind := 'c:valAx';
        end;
    end;
  end;

  if not (ser is TsPieSeries) then
  begin
    WriteChartAxisNode(AStream, AIndent, AChart.XAxis, xAxKind);
    WriteChartAxisNode(AStream, AIndent, AChart.YAxis, yAxKind);
  end;

  AppendToStream(AStream,
    indent + '</c:plotArea>' + LE
  );
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

procedure TsSpreadOOXMLChartWriter.WritePieSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsPieSeries; ASeriesIndex: Integer);
var
  indent: String;
  chart: TsChart;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  AppendToStream(AStream,
    indent + '<c:pieChart>' + LE +
    indent + '  <c:varyColors val="1"/>' + LE
  );

  WriteChartSeriesNode(AStream, AIndent + 4, ASeries, ASeriesIndex);

  AppendToStream(AStream, Format(
    indent + '<c:firstSliceAng val="%d"/>' + LE,
    [ (90 - ASeries.StartAngle) mod 360 ]
  ));
  AppendToStream(AStream,
    indent + '</c:pieChart>' + LE
  );
end;

procedure TsSpreadOOXMLChartWriter.WriteScatterSeries(AStream: TStream;
  AIndent: Integer; ASeries: TsScatterSeries; ASeriesIndex: Integer);
var
  indent: String;
  chart: TsChart;
begin
  indent := DupeString(' ', AIndent);
  chart := ASeries.Chart;

  AppendToStream(AStream,
    indent + '<c:scatterChart>' + LE +
    indent + '  <c:varyColors val="0"/>' + LE +
    indent + '  <c:scatterStyle val="lineMarker"/>' + LE
  );

  WriteChartSeriesNode(AStream, AIndent + 4, ASeries, ASeriesIndex);

  AppendToStream(AStream, Format(
    indent + '  <c:axId val="%d"/>' + LE +
    indent + '  <c:axId val="%d"/>' + LE +
    indent + '</c:scatterChart>' + LE,
    [
      FAxisID[chart.XAxis.Alignment],  // <c:axId>
      FAxisID[chart.YAxis.Alignment]   // <c:axId>
    ]
  ));
end;

{@@ ----------------------------------------------------------------------------
  Writes the <c:trendline> node for the specified chart series if a trendline
  is activated.
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartWriter.WriteChartRegression(AStream: TStream;
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
      indent + '  <c:dispRSqr val="%s"/>' + LE +
      indent + '  <c:dispEq val="%s"/>' + LE +
      indent + '</c:trendline>' + LE,
      [ TRENDLINE_TYPES[trendline.TrendlineType],
        FALSE_TRUE[trendline.DisplayRSquare],
        FALSE_TRUE[trendline.DisplayEquation]
      ]
  ));
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
  AIndent: Integer; ASeries: TsChartSeries; ASeriesIndex: integer);
var
  indent: string;
  chart: TsChart;
  xRng, yRng: TsChartRange;
  forceNoLine: Boolean;
  xValName, yValName, xRefName, yRefName: String;
  explosionStr: String = '';
  dps: TsChartDataPointStyle;
  i: Integer;
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
      ASeriesIndex,   // <c:idx>
      ASeriesIndex    // <c:order>
    ]
  ));

  // Series title
  WriteChartSeriesTitle(AStream, AIndent + 2, ASeries);

  // Individual data point formats
  if ASeries.DataPointStyles.Count > 0 then
    for i := 0 to ASeries.DataPointStyles.Count-1 do
    begin
      dps := ASeries.DataPointStyles[i];
      AppendToStream(AStream,
        indent + '  <c:dPt>' + LE +
        indent + '    <c:idx val="' + IntToStr(i) + '"/>' + LE +
        explosionStr +                        // to do: read explosion value from worksheet!
        indent + '    <c:spPr>' + LE +
        GetChartFillAndLineXML(AIndent + 6, chart, dps.Background, dps.Border) + LE +
        indent + '   </c:spPr>' + LE +
        indent + '  </c:dPt>' + LE
      );
    end;

  // Line & scatter series: symbol markers
  if (ASeries is TsCustomLineSeries) then
  begin
    forceNoLine := not TsOpenedCustomLineSeries(ASeries).ShowLines;
    AppendToStream(AStream,
      indent + '  <c:spPr>' + LE +
                    GetChartLineXML(AIndent, chart, ASeries.Line, forceNoLine) + LE +
      indent + '  </c:spPr>' + LE
    );
    if TsOpenedCustomLineSeries(ASeries).ShowSymbols then
      AppendToStream(AStream,
        indent + '  <c:marker>' + LE +
                      GetChartSeriesMarkerXML(AIndent + 4, TsOpenedCustomLineSeries(ASeries)) + LE +
        indent + '  </c:marker>' + LE
      );
  end else
    // Series main formatting
    AppendToStream(AStream,
      indent + '  <c:spPr>' + LE +
                    GetChartFillAndLineXML(AIndent + 4, chart, ASeries.Fill, ASeries.Line) + LE +
      indent + '  </c:spPr>' + LE
    );

  // Regression
  if ASeries.SupportsTrendline then
    WriteChartRegression(AStream, AIndent + 2, ASeries);

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
  AppendToStream(AStream, Format(
    indent + '  <%0:s>' + LE +
                 GetChartRangeXML(AIndent + 4, xRng, xRefName) + LE +
    indent + '  </%0:s>' + LE,
    [ xValName ]
  ));
  // y range
  AppendToStream(AStream, Format(
    indent + '  <%0:s>' + LE +
                 GetChartRangeXML(AIndent + 4, yRng, yRefName) + LE +
    indent + '  </%0:s>' + LE,
    [ yValName ]
  ));

  if (ASeries is TsBubbleSeries) then
  begin
    // Bubble size range
    AppendToStream(AStream,
      indent +  '<c:bubbleSize>' + LE +
      indent +   GetChartRangeXML(AIndent + 4, TsBubbleSeries(ASeries).BubbleRange, 'c:numRef') + LE +
      indent +  '</c:bubbleSize>' + LE
    );
  end;

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
  AIndent: Integer; AText: TsChartText);
var
  indent: String;
  rotStr: String;
begin
  if not AText.Visible then
    exit;

  str(-AText.RotationAngle * ANGLE_MULTIPLIER:0:0, rotStr);

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

  WriteChartText(AStream, AIndent + 2, ATitle);

  AppendToStream(AStream,
    indent + '  <c:overlay val="0"/>' + LE +
                GetChartFillAndLineXML(AIndent + 2, ATitle.Chart, ATitle.Background, ATitle.Border) + LE
  );

  AppendToStream(AStream,
    indent + '</c:title>' + LE
  );
end;

{$ENDIF}

end.

