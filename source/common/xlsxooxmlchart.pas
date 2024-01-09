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

    function ReadChartColor(ANode: TDOMNode; ADefault: TsColor): TsColor;
    procedure ReadChartFillAndLineProps(ANode: TDOMNode;
      AChart: TsChart; AFill: TsChartFill; ALine: TsChartLine);
    procedure ReadChartFontProps(ANode: TDOMNode; AFont: TsFont);
    procedure ReadChartLineProps(ANode: TDOMNode; AChart: TsChart; AChartLine: TsChartLine);
    procedure ReadChartTextProps(ANode: TDOMNode; AFont: TsFont; var AFontRotation: Single);
  protected
    procedure ReadChart(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartAreaSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartAxis(ANode: TDOMNode; AChart: TsChart; AChartAxis: TsChartAxis);
    procedure ReadChartAxisScaling(ANode: TDOMNode; AChartAxis: TsChartAxis);
    function ReadChartAxisTickMarks(ANode: TDOMNode): TsChartAxisTicks;
    procedure ReadChartBarSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartLegend(ANode: TDOMNode; AChartLegend: TsChartLegend);
    procedure ReadChartLineSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartPlotArea(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartScatterSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartSeriesLabels(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartSeriesMarker(ANode: TDOMNode; ASeries: TsCustomLineSeries);
    procedure ReadChartSeriesProps(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartSeriesRange(ANode: TDOMNode; ARange: TsChartRange);
    procedure ReadChartSeriesTitle(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartSeriesTrendLine(ANode: TDOMNode; ASeries: TsChartSeries);
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

  // The following color values are directly taken from xlsx files written by Excel.
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
      'c:ser':
        begin
          ser := TsAreaSeries.Create(AChart);
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:dLbls':
        ;
      'c:axId':
        ;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartAxis(ANode: TDOMNode;
  AChart: TsChart; AChartAxis: TsChartAxis);
var
  nodeName, s: String;
  n: Integer;
  x: Single;
  node: TDOMNode;
begin
  if ANode = nil then
    exit;

  // Defaults
  AChartAxis.Title.Caption := '';
  AChartAxis.LabelRotation := 0;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:axId':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') and TryStrToInt(s, n) then
            AChartAxis.ID := n;
        end;
      'c:axPos':
        ;
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
        ;
      'c:majorTickMark':
        AChartAxis.MajorTicks := ReadChartAxisTickMarks(ANode);
      'c:minorTickMark':
        AChartAxis.MinorTicks := ReadChartAxisTickMarks(ANode);
      'c:tickLblPos':
        ;
      'c:spPr':
        ReadChartLineProps(ANode.FirstChild, AChart, AChartAxis.AxisLine);
      'c:majorUnit':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
          begin
            AChartAxis.AutomaticMajorInterval := false;
            AChartAxis.MajorInterval := x;
          end;
        end;
      'c:minorUnit':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
          begin
            AChartAxis.AutomaticMinorInterval := false;
            AChartAxis.MinorInterval := x;
          end;
        end;
      'c:txPr':  // Axis labels
        begin
          x := 0;
          ReadChartTextProps(ANode, AChartAxis.LabelFont, x);
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
    case nodeName of
      'c:orientation':
        begin
          s := GetAttrValue(ANode, 'val');
          AChartAxis.Inverted := (s = 'maxMin');
        end;
      'c:max':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
          begin
            AChartAxis.AutomaticMax := false;
            AChartAxis.Max := x;
          end;
        end;
      'c:min':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
          begin
            AChartAxis.AutomaticMin := false;
            AChartAxis.Min := x;
          end;
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
  s: String;
  ser: TsBarSeries;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
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
      'c:ser':
        begin
          ser := TsBarSeries.Create(AChart);
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:dLbls':
        s := '';
      'c:gapWidth':
        ;
      'c:axId':
        ;
    end;
    ANode := ANode.NextSibling;
  end;
end;

function TsSpreadOOXMLChartReader.ReadChartColor(ANode: TDOMNode;
  ADefault: TsColor): TsColor;

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
  Result := ADefault;
  if ANode = nil then
    exit;

  nodeName := ANode.NodeName;
  case nodeName of
    'a:schemeClr':
      begin
        s := GetAttrValue(ANode, 'val');
        if (s <> '') then
        begin
          idx := FColors.IndexOf(s);
          if idx > -1 then
          begin
            Result := FColors.Data[idx];
            child := ANode.FirstChild;
            while Assigned(child) do
            begin
              nodeName := child.NodeName;
              case nodeName of
                'a:tint':
                  begin
                    s := GetAttrValue(child, 'val');
                    if (s <> '') and TryStrToInt(s, n) then
                      Result := TintedColor(Result, n/100000);
                  end;
                'a:lumMod':     // luminance modulated
                  begin
                    s := GetAttrValue(child, 'val');
                    if (s <> '') and TryStrToInt(s, n) then
                      Result := LumModColor(Result, n/100000);
                  end;
                'a:lumOff':
                  begin
                    s := GetAttrValue(child, 'val');
                    if (s <> '') and TryStrToInt(s, n) then
                      Result := LumOffsetColor(Result, n/100000);
                  end;
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
          Result := HTMLColorStrToColor(s);
      end;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartFillAndLineProps(ANode: TDOMNode;
  AChart: TsChart; AFill: TsChartFill; ALine: TsChartLine);
var
  nodeName, s: String;
begin
  if ANode = nil then
    exit;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'a:solidFill':
        begin
          AFill.Style := cfsSolid;
          AFill.Color := ReadChartColor(ANode.FirstChild, scWhite);
        end;
      'a:ln':
        ReadChartLineProps(ANode, AChart, ALine);
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

procedure TsSpreadOOXMLChartReader.ReadChartLineSeries(ANode: TDOMNode; AChart: TsChart);
var
  nodeName: String;
  s: String;
  ser: TsLineSeries;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
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
      'c:ser':
        begin
          ser := TsLineSeries.Create(AChart);
          ReadChartSeriesMarker(ANode.FirstChild, TsLineSeries(ser));
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:dLbls':
        ;
      'c:gapWidth':
        ;
      'c:axId':
        ;
    end;
    ANode := ANode.NextSibling;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Reads the properties of a line series

  @@param  ANode    Points to the <c:marker> subnode of <c:ser> node
  @@param  ASeries  Instance of the TsLineSeries created by ReadChartLineSeries
-------------------------------------------------------------------------------}
procedure TsSpreadOOXMLChartReader.ReadChartSeriesMarker(ANode: TDOMNode; ASeries: TsCustomLineSeries);
var
  nodeName, s: String;
  child: TDOMNode;
  n: Integer;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:marker':
        begin
          TsOpenCustomLineseries(ASeries).ShowSymbols := true;
          child := ANode.FirstChild;
          while Assigned(child) do
          begin
            nodeName := child.NodeName;
            case nodeName of
              'c:symbol':
                begin
                  s := GetAttrValue(child, 'val');
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
                        Symbol := cssAsterisk
                    end;
                end;

              'c:size':
                begin
                  s := GetAttrValue(child, 'val');
                  if (s <> '') and TryStrToInt(s, n) then
                    with TsOpenCustomlineSeries(ASeries) do
                    begin
                      SymbolWidth := PtsToMM(n div 2);
                      SymbolHeight := SymbolWidth;
                    end;
                end;

              'c:spPr':
                with TsOpenCustomLineSeries(ASeries) do
                  ReadChartFillAndLineProps(child.FirstChild, Chart, SymbolFill, SymbolBorder);
            end;
            child := child.NextSibling;
          end;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartPlotArea(ANode: TDOMNode; AChart: TsChart);
var
  nodeName: String;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:areaChart':
        ReadChartAreaSeries(ANode.FirstChild, AChart);
      'c:barChart':
        ReadChartBarSeries(ANode.FirstChild, AChart);
      'c:lineChart':
        ReadChartLineSeries(ANode.FirstChild, AChart);
      'c:scatterChart':
        ReadChartScatterSeries(ANode.FirstChild, AChart);
      'c:catAx':
        ReadChartAxis(ANode.FirstChild, AChart, AChart.XAxis);
      'c:valAx':
        ReadChartAxis(ANode.FirstChild, AChart, AChart.YAxis);
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartScatterSeries(ANode: TDOMNode; AChart: TsChart);
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
      'c:grouping':    ;
      'c:varyColors':  ;
      'c:ser':
        begin
          ser := TsScatterSeries.Create(AChart);
          ReadChartSeriesMarker(ANode.FirstChild, TsScatterSeries(ser));
          ReadChartSeriesProps(ANode.FirstChild, ser);
        end;
      'c:dLbls':
        ;
      'c:gapWidth':
        ;
      'c:axId':
        ;
    end;
    ANode := ANode.NextSibling;
  end;
end;

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
      'c:showLegendKey':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') and (s <> '0') then
            ASeries.DataLabels := ASeries.DataLabels + [cdlSymbol];
        end;
      'c:showVal':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') and (s <> '0') then
            ASeries.DataLabels := ASeries.DataLabels + [cdlValue];
        end;
      'c:showCatName':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') and (s <> '0') then
            ASeries.DataLabels := ASeries.DataLabels + [cdlCategory];
        end;
      'c:showSerName':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') and (s <> '0') then
            ASeries.DataLabels := ASeries.DataLabels + [cdlSeriesName];
        end;
      'c:showPercent':
        begin
          s := GetAttrValue(ANode, 'val');
          if (s <> '') and (s <> '0') then
            ASeries.DataLabels := ASeries.DataLabels + [cdlPercentage];
        end;
      'c:showBubbleSize':
        ;
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
  nodeName: String;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:idx': ;
      'c:order': ;
      'c:tx':
        ReadChartSeriesTitle(ANode.FirstChild, ASeries);
      'c:cat':
        ReadChartSeriesRange(ANode.FirstChild, ASeries.LabelRange);
      'c:xVal':
        ReadChartSeriesRange(ANode.FirstChild, ASeries.XRange);
      'c:val', 'c:yVal':
        ReadChartSeriesRange(ANode.FirstChild, ASeries.YRange);
      'c:spPr':
        ReadChartFillAndLineProps(ANode.FirstChild, ASeries.Chart, ASeries.Fill, ASeries.Line);
      'c:dLbls':
        ReadChartSeriesLabels(ANode.Firstchild, ASeries);
      'c:trendline':
        ReadChartSeriesTrendLine(ANode.FirstChild, ASeries);
      'c:invertIfNegative':
        ;
      'c:extLst':
        ;
    end;
    ANode := ANode.NextSibling;
  end;
end;

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
begin
  lReader := TsSpreadOOXMLReader(Reader);

  xmlStream := lReader.CreateXMLStream;
  try
    if UnzipToStream(AStream, AChartXML, xmlStream) then
    begin
      lReader.ReadXMLStream(doc, xmlStream);
      ReadChart(doc.DocumentElement.FindNode('c:chart'), AChart);
      FreeAndNil(doc);
    end;
  finally
    xmlStream.Free;
  end;
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

