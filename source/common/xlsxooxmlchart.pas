unit xlsxooxmlChart;

{$mode objfpc}{$H+}
{$include ..\fps.inc}

interface

{$ifdef FPS_CHARTS}

uses
  Classes, SysUtils, StrUtils, Contnrs, FPImage, fgl,
  {$ifdef FPS_PATCHED_ZIPPER}fpszipper,{$else}zipper,{$endif}
  laz2_xmlread, laz2_DOM,
  fpsTypes, fpSpreadsheet, fpsChart, fpsUtils, fpsNumFormat,
  fpsReaderWriter, fpsXMLCommon;

type
  { TsSpreadOOXMLChartReader }

  TsSpreadOOXMLChartReader = class(TsBasicSpreadChartReader)
  private
    FPointSeparatorSettings: TFormatSettings;
    FColors: specialize TFPGMap<string, TsColor>;
    FXAxisID, FYAxisID, FX2AxisID, FY2AxisID: DWord;
    FXAxisDelete, FYAxisDelete, FX2AxisDelete, FY2AxisDelete: Boolean;

    function ReadChartColor(ANode: TDOMNode; ADefault: TsColor): TsColor;
    procedure ReadChartColor(ANode: TDOMNode; var AColor: TsColor; var Alpha: Double);
    procedure ReadChartFillAndLineProps(ANode: TDOMNode;
      AChart: TsChart; AFill: TsChartFill; ALine: TsChartLine);
    procedure ReadChartFontProps(ANode: TDOMNode; AFont: TsFont);
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
    procedure ReadChartSeriesDataPointStyles(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartLegend(ANode: TDOMNode; AChartLegend: TsChartLegend);
    procedure ReadChartLineSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartPieSeries(ANode: TDOMNode; AChart: TsChart; RingMode: Boolean);
    procedure ReadChartPlotArea(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartRadarSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartScatterSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartSeriesAxis(ANode: TDOMNode; ASeries: TsChartSeries);
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
    procedure ReadChartXML(AStream: TStream; AChart: TsChart; AChartXML: String);

  end;

  TsSpreadOOXMLChartWriter = class(TsBasicSpreadChartWriter)
  private
    FPointSeparatorSettings: TFormatSettings;

  protected

  public
    constructor Create(AWriter: TsBasicSpreadWriter); override;
    destructor Destroy; override;

  end;

{$ENDIF}

implementation

{$IFDEF FPS_CHARTS}

uses
  xlsxooxml;

const
  PTS_MULTIPLIER = 12700;
  ANGLE_MULTIPLIER = 60000;

{$INCLUDE xlsxooxmlchart_hatch.inc}

type
  TsOpenCustomLineSeries = class(TsCustomLineSeries)
  public
    property Symbol;
    property SymbolBorder;
    property SymbolFill;
    property SymbolHeight;
    property SymbolWidth;
    property ShowLines;
    property ShowSymbols;
  end;

  TsOpenRegressionSeries = class(TsChartSeries)
  public
    property Regression;
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
end;

destructor TsSpreadOOXMLChartReader.Destroy;
begin
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
    case nodeName of
      'c:ser':
        begin
          ser := TsBarSeries.Create(AChart);
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:barDir':
        begin
          s := GetAttrValue(ANode, 'val');
          case s of
            'col': AChart.RotatedAxes := false;
            'bar': AChart.RotatedAxes := true;
          end;
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
        s := '';
      'c:gapWidth':
        ;
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
          ser.BarWidthPercent := round(100 / (1 + n/100));
      'c:overlap':
        if TryStrToFloat(s, n, FPointSeparatorSettings) then
          ser.BarOffsetPercent := round(n);
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
  if ANode = nil then
    exit;

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
                    AColor := TintedColor(AColor, n/100000);
                'a:lumMod':     // luminance modulated
                  if TryStrToInt(s, n) then
                    AColor := LumModColor(AColor, n/100000);
                'a:lumOff':
                  if TryStrToInt(s, n) then
                    AColor := LumOffsetColor(AColor, n/100000);
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
                Alpha := n / 100000;
          end;
          child := child.NextSibling;
        end;
      end;
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
  hatch: string;
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
        begin
          AFill.Style := cfsGradient;
          gradient := TsChartGradient.Create;   // Do not destroy gradient, it will be added to the chart.
          child1 := ANode.FirstChild;
          while Assigned(child1) do
          begin
            nodeName := child1.NodeName;
            case nodeName of
              'a:gsLst':
                begin
                  child2 := child1.FirstChild;
                  while Assigned(child2) do
                  begin
                    nodeName := child2.NodeName;
                    if nodeName = 'a:gs' then
                    begin
                      s := GetAttrValue(child2, 'pos');
                      value := StrToIntDef(s, 0) / 100000;
                      color := scWhite;
                      alpha := 1.0;
                      ReadChartColor(child2.FirstChild, color, alpha);
                      gradient.AddStep(value, color, 1.0 - alpha, 1.0);
                    end;
                    child2 := child2.NextSibling;
                  end;
                end;
              'a:lin':
                begin
                  gradient.Style := cgsLinear;
                  s := GetAttrValue(child1, 'ang');
                  if TryStrToInt(s, n) then
                    gradient.Angle := n / ANGLE_MULTIPLIER;
                end;
            end;
            child1 := child1.NextSibling;
          end;
          AFill.Gradient := AChart.Gradients.AddGradient('', gradient);
        end;

      // Hatched fill
      'a:pattFill':
        begin
          AFill.Style := cfsSolidHatched;
          hatch := GetAttrValue(ANode, 'prst');
          child1 := ANode.FirstChild;
          while Assigned(child1) do
          begin
            nodeName := child1.NodeName;
            case nodeName of
              'a:fgClr':
                color := ReadChartColor(child1.FirstChild, scBlack);
              'a:bgClr':
                AFill.Color := ReadChartColor(child1.FirstChild, scWhite);
            end;
            child1 := child1.NextSibling;
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

{@@ ----------------------------------------------------------------------------
  Reads the individual data point styles of a series.

  @param  ANode    First child of the <c:dPt> node
  @param  ASeries  Series to which these data points belong
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartreader.ReadChartSeriesDataPointStyles(ANode: TDOMNode;
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
          case s of
            't': AChartLegend.Position := lpTop;
            'b': AChartLegend.Position := lpBottom;
            'l': AChartLegend.Position := lpLeft;
            'r': AChartLegend.Position := lpRight;
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

  TsOpenCustomLineSeries(ASeries).ShowSymbols := true;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:symbol':
        with TsOpenCustomLineSeries(ASeries) do
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
          with TsOpenCustomlineSeries(ASeries) do
          begin
            SymbolWidth := PtsToMM(n div 2);
            SymbolHeight := SymbolWidth;
          end;

      'c:spPr':
        with TsOpenCustomLineSeries(ASeries) do
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
  smooth: Boolean;
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
          smooth := (s = 'smoothMarker');    // to do: use it to create a spline series when true.
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
  child, child2, child3: TDOMNode;
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
          child := ANode.FindNode('a:p');
          if Assigned(child) then
          begin
            child2 := child.FirstChild;
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
        ;
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
        if (s <> '') and TryStrToInt(s, n) then
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
end;

procedure TsSpreadOOXMLChartReader.ReadChartSeriesTitle(ANode: TDOMNode; ASeries: TsChartSeries);
var
  nodeName, s: String;
  sheet: String;
  r, c: Cardinal;
begin
  if ANode = nil then
    exit;
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
  regression: TsChartRegression;
  child: TDOMNode;
  n: Integer;
  x: Double;
begin
  if ANode = nil then
    exit;
  if not ASeries.SupportsRegression then
    exit;

  regression := TsOpenRegressionSeries(ASeries).Regression;

  while Assigned(ANode) do begin
    nodeName := ANode.NodeName;
    s := GetAttrValue(ANode, 'val');
    case nodeName of
      'c:name':
        regression.Title := GetNodeValue(ANode);
      'c:spPr':
        ReadChartLineProps(ANode.FirstChild, ASeries.Chart, regression.Line);
      'c:trendlineType':
        case s of
          'exp': regression.RegressionType := rtExponential;
          'linear': regression.RegressionType := rtLinear;
          'log': regression.RegressionType := rtNone;  // rtLog, but not supported.
          'movingAvg': regression.RegressionType := rtNone;  // rtMovingAvg, but not supported.
          'poly': regression.RegressionType := rtPolynomial;
          'power': regression.RegressionType := rtPower;
        end;
      'c:order':
        if (s <> '') and TryStrToInt(s, n) then
          regression.PolynomialDegree := n;
      'c:period':
        if (s <> '') and TryStrToInt(s, n) then ;  // not supported
          // regression.MovingAvgPeriod := n;
      'c:forward', 'c:backward':
        if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
          case nodeName of
            'c:forward': regression.ExtrapolateForwardBy := x;
            'c:backward': regression.ExtrapolateBackwardBy := x;
          end;
      'c:intercept':
        if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
        begin
          regression.YInterceptValue := x;
          regression.ForceYIntercept := true;
        end;
      'c:dispRSqr':
        if s = '1' then
          regression.DisplayRSquare := true;
      'c:dispEq':
        if s = '1' then
          regression.DisplayEquation := true;
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
                  regression.Equation.NumberFormat := s;
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

procedure TsSpreadOOXMLChartReader.ReadChartXML(AStream: TStream; AChart: TsChart;
  AChartXML: String);
var
  lReader: TsSpreadOOXMLReader;
  xmlStream: TStream;
  doc: TXMLDocument = nil;
  node: TDOMNode;
  nodeName: String;
begin
  lReader := TsSpreadOOXMLReader(Reader);

  xmlStream := lReader.CreateXMLStream;
  try
    if UnzipToStream(AStream, AChartXML, xmlStream) then
    begin
      lReader.ReadXMLStream(doc, xmlStream);
      node := doc.DocumentElement.FirstChild; //FindNode('c:chart');
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

{$ENDIF}

end.

