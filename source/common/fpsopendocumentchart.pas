unit fpsOpenDocumentChart;

{$mode objfpc}{$H+}
{$include ..\fps.inc}

interface

{$IFDEF FPS_CHARTS}

uses
  Classes, SysUtils, StrUtils, Contnrs, FPImage, Math,
 {$IFDEF FPS_PATCHED_ZIPPER}
  fpszipper,
 {$ELSE}
  zipper,
 {$ENDIF}
  laz2_xmlread, laz2_DOM,
  fpsTypes, fpSpreadsheet, fpsChart, fpsUtils, fpsNumFormat,
  fpsReaderWriter, fpsXMLCommon;

type

  { TsSpreadOpenDocChartReader }

  TsSpreadOpenDocChartReader = class(TsBasicSpreadChartReader)
  private
    FChartFiles: TStrings;
    FPointSeparatorSettings: TFormatSettings;
    FNumberFormatList: TStrings;
    FPieSeriesStartAngle: Integer;
    FStreamList: TFPObjectList;
    FChartType: TsChartType;
    FStockSeries: TsStockSeries;
    function FindStyleNode(AStyleNodes: TDOMNode; AStyleName: String): TDOMNode;
    function GetChartFillProps(ANode: TDOMNode; AChart: TsChart; AFill: TsChartFill): Boolean;
    function GetChartLineProps(ANode: TDOMNode; AChart: TsChart; ALine: TsChartLine): Boolean;
    procedure GetChartTextProps(ANode: TDOMNode; AFont: TsFont);

    procedure ReadChartAxisGrid(ANode, AStyleNode: TDOMNode; AChart: TsChart; Axis: TsChartAxis);
    procedure ReadChartAxisProps(ANode, AStyleNode: TDOMNode; AChart: TsChart);
    procedure ReadChartAxisStyle(AStyleNode: TDOMNode; AChart: TsChart; Axis: TsChartAxis);
    procedure ReadChartBackgroundProps(ANode, AStyleNode: TDOMNode; AChart: TsChart; AElement: TsChartFillElement);
    procedure ReadChartBackgroundStyle(AStyleNode: TDOMNode; AChart: TsChart; AElement: TsChartFillElement);
    procedure ReadChartCellAddr(ANode: TDOMNode; ANodeName: String; ACellAddr: TsChartCellAddr);
    procedure ReadChartCellRange(ANode: TDOMNode; ANodeName: String; ARange: TsChartRange);
    procedure ReadChartProps(AChartNode, AStyleNode: TDOMNode; AChart: TsChart);
    procedure ReadChartPlotAreaProps(ANode, AStyleNode: TDOMNode; AChart: TsChart);
    procedure ReadChartPlotAreaStyle(AStyleNode: TDOMNode; AChart: TsChart);
    procedure ReadChartLegendProps(ANode, AStyleNode: TDOMNode; AChart: TsChart);
    procedure ReadChartLegendStyle(AStyleNode: TDOMNode; AChart: TsChart);
    procedure ReadChartRegressionEquationStyle(AStyleNode: TDOMNode; AChart: TsChart; ASeries: TsChartSeries);
    procedure ReadChartRegressionProps(ANode, AStyleNode: TDOMNode; AChart: TsChart; ASeries: TsChartSeries);
    procedure ReadChartRegressionStyle(AStyleNode: TDOMNode; AChart: TsChart; ASeries: TsChartSeries);
    procedure ReadChartSeriesDataPointStyle(AStyleNode: TDOMNode; AChart: TsChart;
      ASeries: TsChartSeries; var AFill: TsChartFill; var ALine: TsChartLine; var APieOffset: Integer);
    procedure ReadChartSeriesErrorBarProps(ANode, AStyleNode: TDOMNode; AChart: TsChart;
      ASeries: TsChartSeries);
    procedure ReadChartSeriesErrorBarStyle(AStyleNode: TDOMNode; AChart: TsChart;
      AErrorBars: TsChartErrorBars);
    procedure ReadChartSeriesProps(ANode, AStyleNode: TDOMNode; AChart: TsChart);
    procedure ReadChartSeriesStyle(AStyleNode: TDOMNode; AChart: TsChart; ASeries: TsChartSeries);
    procedure ReadChartStockSeriesStyle(AStyleNode: TDOMNode; AChart: TsChart;
      ASeries: TsStockSeries; ANodeName: String);
    procedure ReadChartTitleProps(ANode, AStyleNode: TDOMNode; AChart: TsChart; ATitle: TsChartText);
    procedure ReadChartTitleStyle(AStyleNode: TDOMNode; AChart: TsChart; ATitle: TsChartText);

    procedure ReadObjectFillImages(ANode: TDOMNode; AChart: TsChart; ARoot: String);
    procedure ReadObjectGradientStyles(ANode: TDOMNode; AChart: TsChart);
    procedure ReadObjectHatchStyles(ANode: TDOMNode; AChart: TsChart);
    procedure ReadObjectLineStyles(ANode: TDOMNode; AChart: TsChart);
  protected
    procedure ReadChartFiles(AStream: TStream; AFileList: String);
    procedure ReadChart(AChartNode, AStyleNode: TDOMNode; AChart: TsChart);
    procedure ReadObjectStyles(ANode: TDOMNode; AChart: TsChart; ARoot: String);
    procedure ReadPictureFile(AStream: TStream; AFileName: String);
  public
    constructor Create(AReader: TsBasicSpreadReader); override;
    destructor Destroy; override;
    procedure AddChartFiles(AFileList: String);
    procedure ReadCharts(AStream: TStream);
  end;

  TsSpreadOpenDocChartWriter = class(TsBasicSpreadChartWriter)
  private
    FSCharts: array of TStream;
    FSObjectStyles: array of TStream;
    FNumberFormatList: TStrings;
    FPointSeparatorSettings: TFormatSettings;
    function GetChartAxisStyleAsXML(Axis: TsChartAxis; AIndent, AStyleID: Integer): String;
    function GetChartBackgroundStyleAsXML(AChart: TsChart;
      AFill: TsChartFill; ABorder: TsChartLine; AIndent: Integer; AStyleID: Integer): String;
    function GetChartCaptionStyleAsXML(AChart: TsChart;
      ACaptionKind, AIndent, AStyleID: Integer): String;
    function GetChartErrorBarStyleAsXML(AChart: TsChart;
      AErrorBar: TsChartErrorBars; AIndent, AStyleID: Integer): String;
    function GetChartFillStyleGraphicPropsAsXML(AChart: TsChart;
      AFill: TsChartFill): String;
    function GetChartLegendStyleAsXML(AChart: TsChart;
      AIndent, AStyleID: Integer): String;
    function GetChartLineStyleAsXML(AChart: TsChart;
      ALine: TsChartLine; AIndent, AStyleID: Integer): String;
    function GetChartLineStyleGraphicPropsAsXML(AChart: TsChart;
      ALine: TsChartLine; ForceNoLine: Boolean = false): String;
    function GetChartPlotAreaStyleAsXML(AChart: TsChart;
      AIndent, AStyleID: Integer): String;
    function GetChartRegressionEquationStyleAsXML(AChart: TsChart;
      AEquation: TsTrendlineEquation; AIndent, AStyleID: Integer): String;
    function GetChartRegressionStyleAsXML(AChart: TsChart;
      ASeriesIndex, AIndent, AStyleID: Integer): String;
    function GetChartSeriesDataPointStyleAsXML(AChart: TsChart;
      ASeriesIndex, ADataPointStyleIndex, AIndent, AStyleID: Integer): String;
    function GetChartSeriesStyleAsXML(AChart: TsChart;
      ASeriesIndex, AIndent, AStyleID: integer): String;
    function GetChartStockSeriesStyleAsXML(AChart: TsChart;
      ASeries: TsStockSeries; AKind: Integer; AIndent, AStyleID: Integer): String;

    procedure CheckAxis(AChart: TsChart; Axis: TsChartAxis);
    function GetNumberFormatID(ANumFormat: String): String;
    procedure ListAllNumberFormats(AChart: TsChart);

  protected
    // Object X/styles.xml
    procedure WriteObjectStyles(AStream: TStream; AChart: TsChart);
    procedure WriteObjectGradientStyles(AStream: TStream; AChart: TsChart; AIndent: Integer);
    procedure WriteObjectHatchStyles(AStream: TStream; AChart: TsChart; AIndent: Integer);
    procedure WriteObjectLineStyles(AStream: TStream; AChart: TsChart; AIndent: Integer);

    // Object X/content.xml
    procedure WriteChart(AStream: TStream; AChart: TsChart);
    procedure WriteChartAxis(AChartStream, AStyleStream: TStream;
      AChartIndent, AStyleIndent: Integer; Axis: TsChartAxis; var AStyleID: Integer);
    procedure WriteChartBackground(AChartStream, AStyleStream: TStream;
      AChartIndent, AStyleIndent: Integer; AChart: TsChart; var AStyleID: Integer);
    procedure WriteChartLegend(AChartStream, AStyleStream: TStream;
      AChartIndent, AStyleIndent: Integer; AChart: TsChart; var AStyleID: Integer);
    procedure WriteChartNumberStyles(AStream: TStream;
      AIndent: Integer; {%H-}AChart: TsChart);
    procedure WriteChartPlotArea(AChartStream, AStyleStream: TStream;
      AChartIndent, AStyleIndent: Integer; AChart: TsChart; var AStyleID: Integer);
    procedure WriteChartSeries(AChartStream, AStyleStream: TStream;
      AChartIndent, AStyleIndent: Integer; AChart: TsChart; ASeriesIndex: Integer;
      var AStyleID: Integer);
    procedure WriteChartStockSeries(AChartStream, AStyleStream: TStream;
      AChartIndent, AStyleIndent: Integer; AChart: TsChart; ASeriesIndex: Integer;
      var AStyleID: Integer);
//    procedure WriteChartTable(AStream: TStream; AChart: TsChart; AIndent: Integer);
    procedure WriteChartTitle(AChartStream, AStyleStream: TStream;
      AChartIndent, AStyleIndent: Integer; AChart: TsChart; IsSubtitle: Boolean;
      var AStyleID: Integer);

  public
    constructor Create(AWriter: TsBasicSpreadWriter); override;
    destructor Destroy; override;

    procedure AddChartsToZip(AZip: TZipper);
    procedure AddToMetaInfManifest(AStream: TStream);

    procedure CreateStreams; override;
    procedure DestroyStreams; override;
    procedure ResetStreams; override;
    procedure WriteCharts; override;
  end;

{$ENDIF}

implementation

{$IFDEF FPS_CHARTS}

uses
  fpsOpenDocument;

type
  TAxisKind = 3..6;

  TsOpenedCustomLineSeries = class(TsCustomLineSeries)
  public
    property Interpolation;
  end;

  TsOpenedTrendlineSeries = class(TsChartSeries)
  public
    property Trendline;
  end;

const
  OPENDOC_PATH_METAINF_MANIFEST = 'META-INF/manifest.xml';
  OPENDOC_PATH_CHART_CONTENT    = 'Object %d/content.xml';
  OPENDOC_PATH_CHART_STYLES     = 'Object %d/styles.xml';

  DEFAULT_FONT_NAME = 'Liberation Sans';

  CHART_TYPE_NAMES: array[TsChartType] of string = (
    '', 'bar', 'line', 'area', 'barLine', 'scatter', 'bubble',
    'radar', 'filled-radar', 'circle', 'ring', 'stock'
  );

  SYMBOL_NAMES: array[TsChartSeriesSymbol] of String = (
    'square', 'diamond', 'arrow-up', 'arrow-down', 'arrow-left',
    'arrow-right', 'circle', 'star', 'x', 'plus', 'asterisk',
    'horizontal-bar', ''
  );  // unsupported: bow-tie, hourglass, vertical-bar

  GRADIENT_STYLES: array[TsChartGradientStyle] of string = (
    'linear', 'axial', 'radial', 'ellipsoid', 'square', 'rectangular', 'radial'
  );

  HATCH_STYLES: array[TsChartHatchStyle] of string = (
    '', 'single', 'double', 'triple'
  );

  LABEL_POSITION: array[TsChartLabelPosition] of string = (
    '', 'outside', 'inside', 'center', 'top', 'bottom', 'near-origin');

  LEGEND_POSITION: array[TsChartLegendPosition] of string = (
    'end', 'top', 'bottom', 'start'
  );

  AXIS_ID: array[TAxisKind] of string = ('x', 'y', 'x', 'y');
  AXIS_LEVEL: array[TAxisKind] of string = ('primary', 'primary', 'secondary', 'secondary');

  TRENDLINE_TYPE: array [TsTrendlineType] of string = (
    '', 'linear', 'logarithmic', 'exponential', 'power', 'polynomial');

  FALSE_TRUE: array[boolean] of string = ('false', 'true');

  LE = LineEnding;

// Replaces all non-letters/numbers by their hex ASCII value surrounded by '_'
function ASCIIName(AName: String): String;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(AName) do
    if AName[i] in ['a'..'z', 'A'..'Z', '0'..'9'] then
      Result := Result + AName[i]
    else
      Result := Result + Format('_%.2x_', [ord(AName[i])]);
end;

// Reverts the replacement done by ASCIIName.
function UnASCIIName(AName: String): String;
var
  i: Integer;
  s: String;
  decoding: Boolean;
begin
  Result := '';
  decoding := false;
  for i := 1 to Length(AName) do
  begin
    if AName[i] = '_' then
    begin
      if decoding then
        Result := Result + char(StrToInt('$'+s))
      else
        s := '';
      decoding := not decoding;
    end else
    if decoding then
      s := s + AName[i]
    else
      Result := Result + AName[i];
  end;
end;

{ Extracts the length from an ods length string, e.g. "3.5cm" or "300%". In the
  former case AValue become 35 (in millimeters), in the latter case AValue is
  300 and Relative becomes true }
function EvalLengthStr(AText: String; out AValue: Double; out Relative: Boolean): Boolean;
var
  i: Integer;
  res: Integer;
  units: String;
begin
  Result := false;

  if AText = '' then
    exit;

  units := '';
  for i := Length(AText) downto 1 do
    if AText[i] in ['%', 'm', 'c', 'p', 't', 'i', 'n'] then
    begin
      units := AText[i] + units;
      Delete(AText, i, 1);
    end;
  Val(AText, AValue, res);
  Result := (res = 0);
  if res = 0 then
  begin
    Relative := false;
    case units of
      '%': Relative := true;
      'mm': ;
      'cm': AValue := AValue * 10;
      'pt': AValue := PtsToMM(AValue);
      'in': AValue := InToMM(AValue);
      else  Result := false;
    end;
  end;
end;

function ModifyColor(AColor: TsChartColor; AIntensity: double): TsChartColor;
begin
  Result.Color := LumModOff(AColor.Color, AIntensity, 0.0);
end;

procedure SwapColors(var AColor1, AColor2: TsChartColor);
var
  tmp: TsChartColor;
begin
  tmp := AColor1;
  AColor1 := AColor2;
  AColor2 := tmp;
end;


{------------------------------------------------------------------------------}
{                        internal number formats                               }
{------------------------------------------------------------------------------}

type
  TsChartNumberFormatList = class(TStringList)
  public
    constructor Create;
    function Add(const ANumFormat: String): Integer; override;
    function FindFormatByName(const AName: String): String;
    function IndexOfFormat(ANumFormat: String): Integer;
    function IndexOfName(const AName: String): Integer; override;
  end;

constructor TsChartNumberFormatList.Create;
begin
  inherited;
  NameValueSeparator := ':';
  Add('');  // default number format
end;

// Adds a new format, but make sure to avoid duplicates.
// The format list stores the item internally as name:value pair with
// name = 'N'+index and value = ANumFormat
function TsChartNumberFormatList.Add(const ANumFormat: String): Integer;
begin
  if (ANumFormat = '') and (Count > 0) then
    Result := 0
  else
  begin
    Result := IndexOfFormat(ANumFormat);
    if Result = -1 then
      Result := inherited Add(ANumFormat); //(Format('N%d:%s', [Count, ANumFormat]));
  end;
end;

{ The reader adds formats in the form "name:format" where "name" is the
  identifier used in the style definition, e.g. "N0". }
function TsChartNumberFormatList.FindFormatByName(const AName: String): String;
var
  idx: Integer;
begin
  Result := '';
  idx := IndexOfName(AName);
  if idx <> -1 then
  begin
    Result := ValueFromIndex[idx];
    if Result = 'General' then
      Result := '';
  end;
end;

function TsChartNumberFormatList.IndexOfFormat(ANumFormat: String): Integer;
var
  i: Integer;
  fmt: String;
begin
  ANumFormat := lowercase(ANumFormat);
  for i := 0 to Count-1 do
  begin
    fmt := Lowercase(ValueFromIndex[i]);
    if fmt = ANumFormat then
    begin
      Result := i;
      exit;
    end;
  end;
  Result := -1;
end;

function TsChartNumberFormatList.IndexOfName(const AName: String): Integer;
begin
  Result := inherited IndexOfName(lowercase(AName));
end;

{------------------------------------------------------------------------------}
{                         internal picture storage                             }
{------------------------------------------------------------------------------}
type
  TStreamItem = class
    Name: String;
    Stream: TStream;
    destructor Destroy; override;
  end;

destructor TStreamItem.Destroy;
begin
  Stream.Free;
  inherited;
end;

type
  TStreamList = class(TFPObjectList)
  public
    function FindByName(AName: String): TStream;
  end;

function TStreamList.FindByName(AName: String): TStream;
var
  i: Integer;
begin
  for i := 0 to Count-1 do
    if TStreamItem(Items[i]).Name = AName then
    begin
      Result := TStreamItem(Items[i]).Stream;
      exit;
    end;
  Result := nil;
end;


{------------------------------------------------------------------------------}
{                        TsSpreadOpenDocChartReader                            }
{------------------------------------------------------------------------------}

constructor TsSpreadOpenDocChartReader.Create(AReader: TsBasicSpreadReader);
begin
  inherited Create(AReader);

  FPointSeparatorSettings := SysUtils.DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator:='.';

  FChartFiles := TStringList.Create;
  FNumberFormatList := TsChartNumberFormatList.Create;
  FStreamList := TStreamList.Create;

  FPieSeriesStartAngle := 999;
end;

destructor TsSpreadOpenDocChartReader.Destroy;
begin
  FStreamList.Free;
  FNumberFormatList.Free;
  FChartFiles.Free;
  inherited;
end;

{ Searches in the child nodes of AStyleNode for the style:style node with
  the attributes style:family = chart and style:name = AStyleName. }
function TsSpreadOpenDocChartReader.FindStyleNode(AStyleNodes: TDOMNode;
  AStyleName: String): TDOMNode;
var
  nodeName: String;
  sn, sf: String;
begin
  Result := AStyleNodes.FirstChild;
  while (Result <> nil) do
  begin
    nodeName := Result.NodeName;
    if nodeName = 'style:style' then
    begin
      sn := GetAttrValue(Result, 'style:name');
      sf := GetAttrValue(Result, 'style:family');
      if (sf = 'chart') and (sn = AStyleName) then
        exit;
    end;
    Result := Result.NextSibling;
  end;
  Result := nil;
end;

{ AFiles contains a sorted, comma-separated list of all files
  belonging to each chart. }
procedure TsSpreadOpenDocChartReader.AddChartFiles(AFileList: String);
begin
  FChartFiles.Add(AFileList);
end;

{@@ ----------------------------------------------------------------------------
  Reads the fill style properties from the specified node. Returns FALSE, if
  the node contains no fill-specific attributes.
-------------------------------------------------------------------------------}
function TsSpreadOpenDocChartReader.GetChartFillProps(ANode: TDOMNode;
  AChart: TsChart; AFill: TsChartFill): Boolean;
var
  {%H-}nodeName: String;
  sFill: String;
  sOpac: String;
  sc: String;
  sn: String;
  opacity: Double;
  img: TsChartImage;
  value: Double;
  rel: Boolean;
begin
  nodeName := ANode.NodeName;

  sFill := GetAttrValue(ANode, 'draw:fill');
  case sFill of
    'none':
      AFill.Style := cfsNoFill;
    '', 'solid':
      begin
        AFill.Style := cfsSolid;
        sc := GetAttrValue(ANode, 'draw:fill-color');
        if sc <> '' then
          AFill.Color := ChartColor(HTMLColorStrToColor(sc));
      end;
    'gradient':
      begin
        AFill.Style := cfsGradient;
        sn := GetAttrValue(ANode, 'draw:fill-gradient-name');
        if sn <> '' then
          AFill.Gradient := AChart.Gradients.IndexOfName(UnASCIIName(sn));
      end;
    'hatch':
      begin
        sc := GetAttrValue(ANode, 'draw:fill-hatch-solid');
        if sc = 'true' then
          AFill.Style := cfsSolidHatched
        else
          AFill.Style := cfsHatched;
        sn := GetAttrValue(ANode, 'draw:fill-hatch-name');
        if sn <> '' then
          AFill.Hatch := AChart.Hatches.IndexOfName(UnASCIIName(sn));
        sc := GetAttrValue(ANode, 'draw:fill-color');
        if sc <> '' then
          AFill.Color := ChartColor(HTMLColorStrToColor(sc));
      end;
    'bitmap':
      begin
        sn := GetAttrValue(ANode, 'draw:fill-image-name');
        if sn <> '' then
        begin
          AFill.Style := cfsImage;
          AFill.Image := AChart.Images.IndexOfName(UnASCIIName(sn));
          img := AChart.Images[AFill.Image];
          sc := GetAttrValue(ANode, 'draw:fill-image-width');
          if (sc <> '') and EvalLengthStr(sc, value, rel) then
            img.Width := value
          else
            img.Width := -1;
          sc := GetAttrValue(ANode, 'draw:fill-image-height');
          if (sc <> '') and EvalLengthStr(sc, value, rel) then
            img.Height := value
          else
            img.Height := -1;
        end else
          AFill.Style := cfsSolid;
      end;
  end;

  sOpac := GetAttrValue(ANode, 'draw:opacity');
  if (sOpac <> '') and TryPercentStrToFloat(sOpac, opacity) then
    AFill.Color.Transparency := 1.0 - opacity;

  Result := (sFill <> '') or (sc <> '') or (sn <> '') or (sOpac <> '');
end;

{ ------------------------------------------------------------------------------
  Reads the line formatting properties from the specified node.
  Returns FALSE, if there are no line-related attributes.
-------------------------------------------------------------------------------}
function TsSpreadOpenDocChartReader.GetChartLineProps(ANode: TDOMNode;
  AChart: TsChart; ALine: TsChartLine): Boolean;
var
  {%H-}nodeName: String;
  s: String;
  sn: String;
  sc: String;
  sw: String;
  so: String;
  value: Double;
  rel: Boolean;
begin
  nodeName := ANode.NodeName;

  s := GetAttrValue(ANode, 'draw:stroke');
  case s of
    'none':
      ALine.Style := clsNoLine;
    'solid':
      ALine.Style := clsSolid;
    'dash':
      begin
        sn := GetAttrValue(ANode, 'draw:stroke-dash');
        if sn <> '' then
          ALine.Style := AChart.LineStyles.IndexOfName(UnASCIIName(sn));
      end;
  end;

  sc := GetAttrValue(ANode, 'svg:stroke-color');
  if sc = '' then
    sc := GetAttrValue(ANode, 'draw:stroke-color');
  if sc <> '' then
    ALine.Color := ChartColor(HTMLColorStrToColor(sc));

  sw := GetAttrValue(ANode, 'svg:stroke-width');
  if sw = '' then
    sw := GetAttrValue(ANode, 'draw:stroke-width');
  if (sw <> '') and EvalLengthStr(sw, value, rel) then
    ALine.Width := value;

  so := GetAttrValue(ANode, 'draw:stroke-opacity');
  if (so <> '') and TryPercentStrToFloat(so, value) then
    ALine.Color.Transparency := 1.0 - value*0.01;

  Result := (s <> '') or (sc <> '') or (sw <> '') or (so <> '');
end;

procedure TsSpreadOpenDocChartReader.GetChartTextProps(ANode: TDOMNode;
  AFont: TsFont);
begin
  TsSpreadOpenDocReader(Reader).ReadFont(ANode, AFont);
  if AFont.FontName = '' then
    AFont.FontName := DEFAULT_FONT_NAME;
end;

procedure TsSpreadOpenDocChartReader.ReadChart(AChartNode, AStyleNode: TDOMNode;
  AChart: TsChart);
var
  nodeName: String;
  officeChartNode: TDOMNode;
  chartChartNode: TDOMNode;
  chartElementNode: TDOMNode;
begin
  // Default values
  AChart.Legend.Visible := false;

  nodeName := AStyleNode.NodeName;
  if nodeName = 'office:automatic-styles' then
    TsSpreadOpenDocReader(Reader).ReadNumFormats(AStyleNode, FNumberFormatList);

  nodeName := AChartNode.NodeName;
  officeChartNode := AChartNode.FirstChild;
  while officeChartNode <> nil do
  begin
    nodeName := officeChartNode.NodeName;
    if nodeName = 'office:chart' then
    begin
      chartChartNode := officeChartNode.FirstChild;
      while chartChartNode <> nil do
      begin
        nodeName := chartChartNode.NodeName;
        if nodeName = 'chart:chart' then
        begin
          ReadChartProps(chartChartNode, AStyleNode, AChart);
          chartElementNode := chartChartNode.FirstChild;
          while (chartElementNode <> nil) do
          begin
            nodeName := chartElementNode.NodeName;
            case nodeName of
              'chart:plot-area': ReadChartPlotAreaProps(chartElementNode, AStyleNode, AChart);
              'chart:legend': ReadChartLegendProps(chartElementNode, AStyleNode, AChart);
              'chart:title': ReadChartTitleProps(chartElementNode, AStyleNode, AChart, AChart.Title);
              'chart:subtitle': ReadChartTitleProps(chartElementNode, AStyleNode, AChart, AChart.Subtitle);
            end;
            chartElementNode := chartElementNode.NextSibling;
          end;
        end;
        chartChartNode := chartChartNode.NextSibling;
      end;
    end;
    officeChartNode := officeChartNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartAxisGrid(ANode, AStyleNode: TDOMNode;
  AChart: TsChart; Axis: TsChartAxis);
var
  nodeName: String;
  s: String;
  styleNode, subNode: TDOMNode;
  grid: TsChartLine;
begin
  nodeName := ANode.NodeName;

  s := GetAttrValue(ANode, 'chart:class');
  case s of
    'major': grid := Axis.MajorGridLines;
    'minor': grid := Axis.MinorGridLines;
    else exit;
  end;

  // Set defaults
  Axis.MajorTicks := [catOutside];
  grid.Style := clsSolid;
  grid.Color := ChartColor($c0c0c0);

  s := GetAttrValue(ANode, 'chart:style-name');
  styleNode := FindStyleNode(AStyleNode, s);
  if styleNode <> nil then
  begin
    subnode := styleNode.FirstChild;
    while (subNode <> nil) do
    begin
      nodeName := subNode.NodeName;
      if nodeName = 'style:graphic-properties' then
        GetChartLineProps(subNode, AChart, grid);
      subNode := subNode.NextSibling;
    end;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartAxisProps(ANode, AStyleNode: TDOMNode;
  AChart: TsChart);
var
  s, nodeName: String;
  styleNode, subNode: TDOMNode;
  axis: TsChartAxis;
begin
  s := GetAttrValue(ANode, 'chart:name');
  case s of
    'primary-x': axis := AChart.XAxis;
    'primary-y': axis := AChart.YAxis;
    'secondary-x': axis := AChart.X2Axis;
    'secondary-y': axis := AChart.Y2Axis;
    else raise Exception.Create('Unknown chart axis.');
  end;

  // Default values
  axis.Visible := true;  // The presence of this node makes the axis visible.
  axis.Title.Caption := '';
  axis.MajorGridLines.Style := clsNoLine;
  axis.MinorGridLines.Style := clsNoLine;
  axis.MajorTicks := [catOutside];
  axis.MinorTicks := [catOutside];

  s := GetAttrValue(ANode, 'chart:style-name');
  styleNode := FindStyleNode(AStyleNode, s);
  ReadChartAxisStyle(styleNode, AChart, axis);

  subNode := ANode.FirstChild;
  while subNode <> nil do
  begin
    nodeName := subNode.NodeName;
    case nodeName of
      'chart:title':
        ReadChartTitleProps(subNode, AStyleNode, AChart, axis.Title);
      'chart:categories':
        ReadChartCellRange(subNode, 'table:cell-range-address', axis.CategoryRange);
      'chart:grid':
        ReadChartAxisGrid(subNode, AStyleNode, AChart, axis);
      'chartooo:date-scale':
        axis.DateTime := true;
    end;
    subNode := subNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartAxisStyle(AStyleNode: TDOMNode;
  AChart: TsChart; Axis: TsChartAxis);
var
  nodeName: String;
  s: String;
  value: Double;
  n: Integer;
  ticks: TsChartAxisTicks = [];
begin
  nodeName := AStyleNode.NodeName;

  s := GetAttrValue(AStyleNode, 'style:data-style-name');
  if s <> '' then
    s := TsChartNumberFormatList(FNumberFormatList).FindFormatByName(s);
  if IsDateTimeFormat(s) then
  begin
    Axis.DateTime := true;
    Axis.LabelFormatDateTime := s;
  end else
  if (AChart.StackMode = csmStackedPercentage) and ((Axis = AChart.YAxis) or (Axis = AChart.Y2Axis)) then
    Axis.LabelFormatPercent := s
  else
    Axis.LabelFormat := s;

  AStyleNode := AStyleNode.FirstChild;
  while AStyleNode <> nil do
  begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:text-properties':
        GetChartTextProps(AStyleNode, Axis.LabelFont);
      'style:graphic-properties':
        GetChartLineProps(AStyleNode, AChart, Axis.AxisLine);
      'style:chart-properties':
        begin
          s := GetAttrValue(AStyleNode, 'chart:display-label');
          if s = 'true' then
            Axis.ShowLabels := true;

          s := GetAttrValue(AStyleNode, 'chart:logarithmic');
          if s = 'true' then
            Axis.Logarithmic := true;

          s := GetAttrValue(AStyleNode, 'chart:minimum');
          if (s <> '') and TryStrToFloat(s, value, FPointSeparatorSettings) then
          begin
            Axis.Min := value;
            Axis.AutomaticMin := false;
          end else
            Axis.AutomaticMin := true;

          s := GetAttrValue(AStyleNode, 'chart:maximum');
          if (s <> '') and TryStrToFloat(s, value, FPointSeparatorSettings) then
          begin
            Axis.Max := value;
            Axis.AutomaticMax := false;
          end else
            Axis.AutomaticMax := true;

          s := GetAttrValue(AStyleNode, 'chart:interval-major');
          if (s <> '') and TryStrToFloat(s, value, FPointSeparatorSettings) then
            Axis.MajorInterval := value;

          s := GetAttrValue(AStyleNode, 'chart:interval-minor-divisor');
          if (s <> '') and TryStrToInt(s, n) then
            Axis.MinorCount := n;

          s := GetAttrValue(AStyleNode, 'chart:axis-position');
          case s of
            'start':
              Axis.Position := capStart;
            'end':
              Axis.Position := capEnd;
            else
              if TryStrToFloat(s, value, FPointSeparatorSettings) then
              begin
                Axis.Position := capValue;
                Axis.PositionValue := value;
              end;
          end;

          ticks := [];  // To do: check defaults...
          s := GetAttrValue(AStyleNode, 'chart:tick-marks-major-inner');
          if s = 'true' then ticks := ticks + [catInside];
          s := GetAttrValue(AStyleNode, 'chart:tick-marks-major-outer');
          if s = 'true' then ticks := ticks + [catOutside];
          Axis.MajorTicks := ticks;

          ticks := [];  // To do: check defaults...
          s := GetAttrValue(AStyleNode, 'chart:tick-marks-minor-inner');
          if s = 'true' then ticks := ticks + [catInside];
          s := GetAttrValue(AStyleNode, 'chart:tick-marks-minor-outer');
          if s = 'true' then ticks := ticks + [catOutside];
          Axis.MinorTicks := ticks;

          s := GetAttrValue(AStyleNode, 'chart:reverse-direction');
          if s = 'true' then Axis.Inverted := true;

          s := GetAttrValue(AStyleNode, 'style:rotation-angle');
          if (s <> '') and TryStrToFloat(s, value, FPointSeparatorSettings) then
            Axis.LabelRotation := Round(value);

          s := GetAttrValue(AStyleNode, 'chart:gap-width');  // why did they put this here ???
          if TryStrToFloat(s, value, FPointSeparatorSettings) then
            AChart.BarGapWidthPercent := round(value);

          s := GetAttrValue(AStyleNode, 'chart:overlap');    // why did they put this here ???
          if TryStrToFloat(s, value, FPointSeparatorSettings) then
            AChart.BarOverlapPercent := round(value);
        end;
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartBackgroundProps(ANode, AStyleNode: TDOMNode;
  AChart: TsChart; AElement: TsChartFillElement);
var
  s: String;
  styleNode: TDOMNode;
begin
  s := GetAttrValue(ANode, 'chart:style-name');
  styleNode := FindStyleNode(AStyleNode, s);
  ReadChartBackgroundStyle(styleNode, AChart, AElement);
end;

procedure TsSpreadOpenDocChartReader.ReadChartBackgroundStyle(AStyleNode: TDOMNode;
  AChart: TsChart; AElement: TsChartFillElement);
var
  nodeName: String;
begin
  AElement.Border.Style := clsNoLine;

  nodeName := AStyleNode.NodeName;
  AStyleNode := AStyleNode.FirstChild;
  while AStyleNode <> nil do begin
    nodeName := AStyleNode.NodeName;
    if nodeName = 'style:graphic-properties' then
    begin
      GetChartLineProps(AStyleNode, AChart, AElement.Border);
      GetChartFillProps(AStyleNode, AChart, AElement.Background);
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartCellAddr(ANode: TDOMNode;
  ANodeName: String; ACellAddr: TsChartCellAddr);
var
  s: String;
  sh1, sh2: String;
  r1, c1, r2, c2: Cardinal;
  relFlags: TsRelFlags;
begin
  s := GetAttrValue(ANode, ANodeName);
  if (s <> '') and TryStrToCellRange_ODS(s, sh1, sh2, r1, c1, r2, c2, relFlags) then
  begin
    ACellAddr.Sheet := sh1;
    ACellAddr.Row := r1;
    ACellAddr.Col := c1;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartCellRange(ANode: TDOMNode;
  ANodeName: String; ARange: TsChartRange);
var
  s: String;
  sh1, sh2: String;
  r1, c1, r2, c2: Cardinal;
  relFlags: TsRelFlags;
begin
  s := GetAttrValue(ANode, ANodeName);
  if (s <> '') and TryStrToCellRange_ODS(s, sh1, sh2, r1, c1, r2, c2, relFlags) then
  begin
    ARange.Sheet1 := sh1;
    if (sh2 = '') and (ARange.Sheet1 <> '') then
      ARange.Sheet2 := ARange.Sheet1
    else
      ARange.Sheet2 := sh2;
    ARange.Row1 := r1;
    ARange.Col1 := c1;
    ARange.Row2 := r2;
    ARange.Col2 := c2;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartFiles(AStream: TStream;
  AFileList: String);
var
  sa: TStringArray;
  i, p: Integer;
  root, fn: String;
  contentFile: String = '';
  stylesFile: String = '';
  XMLStream: TStream;
  doc: TXMLDocument = nil;
  chart: TsChart;
  ok: Boolean;
  lReader: TsSpreadOpenDocReader;
begin
  lReader := TsSpreadOpenDocReader(Reader);

  sa := SplitStr(AFileList, ',');
  for i := 0 to High(sa) do
  begin
    fn := ExtractFileName(sa[i]);
    if fn = 'content.xml' then
      contentFile := sa[i]
    else if fn = 'styles.xml' then
      stylesFile := sa[i]
    else if pos('/Pictures/', sa[i]) > 0 then
      ReadPictureFile(AStream, sa[i]);
  end;

  for i := 0 to TsWorkbook(Reader.Workbook).GetChartCount-1 do
  begin
    chart := TsWorkbook(Reader.Workbook).GetChartByIndex(i);
    if pos(chart.Name, contentFile) = 1 then
      break;
    chart := nil;
  end;

  // Chart not found
  if chart = nil then
    raise Exception.Create('Chart in "' + contentfile + '" not found.');

  // Read the Object/styles.xml file
  if stylesFile <> '' then
  begin
    XMLStream := lReader.CreateXMLStream;
    try
      ok := UnzipToStream(AStream, stylesFile, XMLStream);
      if ok then
      begin
        lReader.ReadXMLStream(doc, XMLStream);
        if not Assigned(doc) then
          ok := false;
      end;
    finally
      XMLStream.Free;
    end;
    if not ok then
      raise Exception.Create('ODS chart reader: error reading styles file "' + stylesFile + '"');

    p := pos('/', stylesFile);
    root := copy(stylesFile, 1, p);
    ReadObjectStyles(doc.DocumentElement.FindNode('office:styles'), chart, root);
    FreeAndNil(doc);
  end;

  // Read the Object/content.xml file
  XMLStream := lReader.CreateXMLStream;
  try
    ok := UnzipToStream(AStream, contentFile, XMLStream);
    if ok then
    begin
      lReader.ReadXMLStream(doc, XMLStream);
      if not Assigned(doc) then
        ok := false;
    end;
  finally
    XMLStream.Free;
  end;

  if not ok then
    raise Exception.Create('ODS chart reader: error reading content file ' + contentFile);

  ReadChart(
    doc.DocumentElement.FindNode('office:body'),
    doc.DocumentElement.FindNode('office:automatic-styles'),
    chart
  );

  FreeAndNil(doc);
end;

procedure TsSpreadOpenDocChartReader.ReadChartProps(AChartNode, AStyleNode: TDOMNode;
  AChart: TsChart);
var
  ct: TsChartType;
  s: String;
  styleName: String;
  styleNode: TDOMNode;
begin
  s := GetAttrValue(AChartNode, 'chart:class');
  if s <> '' then
  begin
    Delete(s, 1, Pos(':', s));  // remove "chart:"
    for ct in TsChartType do
      if CHART_TYPE_NAMES[ct] = s then
      begin
        FChartType := ct;
        if FChartType = ctStock then
        begin
          FStockSeries := TsStockSeries.Create(AChart);
          FStockSeries.Fill.Style := cfsSolid;
          FStockSeries.Fill.Color := ChartColor(scWhite);
          FStockSeries.Line.Style := clsSolid;
          FStockSeries.Line.Color := ChartColor(scBlack);
          FStockSeries.RangeLine.Style := clsSolid;
          FStockSeries.RangeLine.Color := ChartColor(scBlack);
          FStockSeries.CandleStickDownFill.Style := cfsSolid;
          FStockSeries.CandleStickDownFill.Color := ChartColor(scBlack);
        end;
      end;
  end;

  styleName := GetAttrValue(AChartNode, 'chart:style-name');
  if styleName <> '' then
  begin
    styleNode := FindStyleNode(AStyleNode, styleName);
    ReadChartBackgroundStyle(styleNode, AChart, AChart);
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartPlotAreaProps(ANode, AStyleNode: TDOMNode;
  AChart: TsChart);
var
  nodeName: String;
  styleName: String;
  styleNode: TDOMNode;
begin
  styleName := GetAttrValue(ANode, 'chart:style-name');
  styleNode := FindStyleNode(AStyleNode, styleName);
  ReadChartPlotAreaStyle(styleNode, AChart);

  // Defaults
  AChart.XAxis.Visible := false;
  AChart.YAxis.Visible := false;
  AChart.X2Axis.Visible := false;
  AChart.Y2Axis.Visible := false;
  AChart.XAxis.DefaultTitleRotation := true;
  AChart.YAxis.DefaultTitleRotation := true;
  AChart.X2Axis.DefaultTitleRotation := true;
  AChart.Y2Axis.DefaultTitleRotation := true;
  AChart.PlotArea.Border.Style := clsNoLine;
  AChart.Floor.Border.Style := clsNoLine;

  ANode := ANode.FirstChild;
  while ANode <> nil do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'chart:axis':
        ReadChartAxisProps(ANode, AStyleNode, AChart);
      'chart:series':
        ReadChartSeriesProps(ANode, AStyleNode, AChart);
      'chart:wall':
        ReadChartBackgroundProps(ANode, AStyleNode, AChart, AChart.PlotArea);
      'chart:floor':
        ReadChartBackgroundProps(ANode, AStyleNode, AChart, AChart.Floor);
      'chart:stock-gain-marker',
      'chart:stock-loss-marker',
      'chart:stock-range-line':
        begin
          styleName := GetAttrValue(ANode, 'chart:style-name');
          if (styleName <> '') and (FStockSeries <> nil) then
          begin
            styleNode := FindStyleNode(AStyleNode, styleName);
            ReadChartStockSeriesStyle(styleNode, AChart, FStockSeries, nodeName);
          end;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartPlotAreaStyle(AStyleNode: TDOMNode; AChart: TsChart);
var
  nodeName, s: String;
begin
  AStyleNode := AStyleNode.FirstChild;
  while AStyleNode <> nil do
  begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:chart-properties':
        begin
          // Stacked
          s := GetAttrValue(AStyleNode, 'chart:stacked');
          if s = 'true' then
            AChart.StackMode := csmStacked;
          // Stacked as percentage
          s := GetAttrValue(AStyleNode, 'chart:percentage');
          if s = 'true' then
            AChart.StackMode := csmStackedPercentage;
          // Line series interpolation
          s := GetAttrValue(AStyleNode, 'chart:interpolation');
          case s of
            'cubic-spline': AChart.Interpolation := ciCubicSpline;
            'b-spline': AChart.Interpolation := ciBSpline;
            'step-start': AChart.Interpolation := ciStepStart;
            'step-end': AChart.Interpolation := ciStepEnd;
            'step-center-x': AChart.Interpolation := ciStepCenterX;
            'step-center-y': AChart.Interpolation := ciStepCenterY;
            else AChart.Interpolation := ciLinear;
          end;
          // Horizontal bars
          s := GetAttrValue(AStyleNode, 'chart:vertical');
          if s = 'true' then
            AChart.RotatedAxes := true;
          // Pie series start angle
          s := GetAttrValue(AStyleNode, 'chart:angle-offset');
          if s <> '' then
            FPieSeriesStartAngle := StrToInt(s);
          // Stockseries candlestick mode
          s := GetAttrValue(AStyleNode, 'chart:japanese-candle-stick');
          if (s <> '') and (FStockSeries <> nil) then
            FStockSeries.CandleStick := true;
        end;
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartLegendProps(ANode, AStyleNode: TDOMNode;
  AChart: TsChart);
var
  styleName: String;
  styleNode: TDOMNode;
  s: String;
  lp: TsChartLegendPosition;
  value: Double;
  rel: Boolean;
begin
  styleName := GetAttrValue(ANode, 'chart:style-name');
  styleNode := FindStyleNode(AStyleNode, styleName);
  ReadChartLegendStyle(styleNode, AChart);

  s := GetAttrValue(ANode, 'chart:legend-position');
  if s <> '' then
    for lp in TsChartLegendPosition do
      if s = LEGEND_POSITION[lp] then
      begin
        AChart.Legend.Visible := true;
        AChart.Legend.Position := lp;
        break;
      end;

  s := GetAttrValue(ANode, 'svg:x');
  if (s <> '') and EvalLengthStr(s, value, rel) then
    if not rel then
      AChart.Legend.PosX := value;

  s := GetAttrValue(ANode, 'svg:y');
  if (s <> '') and EvalLengthStr(s, value, rel) then
    if not rel then
      AChart.Legend.PosY := value;

  s := GetAttrValue(ANode, 'loext:overlay');
  AChart.Legend.CanOverlapPlotArea := (s = 'true');
end;

procedure TsSpreadOpenDocChartReader.ReadChartLegendStyle(AStyleNode: TDOMNode;
  AChart: TsChart);
var
  nodeName: String;
begin
  nodeName := AStyleNode.NodeName;
  AStyleNode := AStyleNode.FirstChild;
  while AStyleNode <> nil do begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:graphic-properties':
        begin
          GetChartLineProps(AStyleNode, AChart, AChart.Legend.Border);
          GetChartFillProps(AStyleNode, AChart, AChart.Legend.Background);
        end;
      'style:text-properties':
          GetChartTextProps(AStyleNode, AChart.Legend.Font);
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartRegressionEquationStyle(AStyleNode: TDOMNode;
  AChart: TsChart; ASeries: TsChartSeries);
var
  series: TsCustomScatterSeries;
  trendline: TsChartTrendline;
  odsReader: TsSpreadOpenDocReader;
  s, nodeName: String;
begin
  if not (ASeries is TsScatterSeries) then
    exit;

  series := TsCustomScatterSeries(ASeries);
  trendline := series.Trendline;
  odsReader := TsSpreadOpenDocReader(Reader);

  nodeName := AStyleNode.NodeName;
  s := GetAttrValue(AStyleNode, 'style:data-style-name');
  if s <> '' then
    s := TsChartNumberFormatList(FNumberFormatList).FindFormatByName(s);
  trendline.Equation.NumberFormat := s;

  AStyleNode := AStyleNode.FirstChild;
  while Assigned(AStyleNode) do
  begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:graphic-properties':
        begin
          GetChartLineProps(AStyleNode, AChart, trendline.Equation.Border);
          GetChartFillProps(AStyleNode, AChart, trendline.Equation.Fill);
        end;
      'style:text-properties':
        GetChartTextProps(AStyleNode, trendline.Equation.Font);
      'style:chart-properties':
        begin
          s := GetAttrValue(AStyleNode, 'loext:regression-x-name');
          if s <> '' then
            trendline.Equation.XName := s;

          s := GetAttrValue(AStyleNode, 'loext:regression-y-name');
          if s <> '' then
            trendline.Equation.YName := s;
        end;
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartRegressionProps(ANode, AStyleNode: TDOMNode;
  AChart: TsChart; ASeries: TsChartSeries);
var
  series: TsCustomScatterSeries;
  trendline: TsChartTrendline;
  s, nodeName: String;
  styleNode: TDOMNode;
  subNode: TDOMNode;
begin
  if not (ASeries is TsCustomScatterSeries) then
    exit;

  series := TsCustomScatterSeries(ASeries);
  trendline := series.Trendline;

  s := GetAttrValue(ANode, 'chart:style-name');
  styleNode := FindStyleNode(AStyleNode, s);
  ReadChartRegressionStyle(styleNode, AChart, ASeries);

  subNode := ANode.FirstChild;
  while Assigned(subNode) do
  begin
    nodeName := subNode.NodeName;
    if nodeName = 'chart:equation' then
    begin
      s := GetAttrValue(subNode, 'chart:display-equation');
      trendline.DisplayEquation := (s = 'true');

      s := GetAttrValue(subNode, 'chart:display-r-square');
      trendline.DisplayRSquare := (s = 'true');

      s := GetAttrValue(subNode, 'chart:style-name');
      styleNode := FindStyleNode(AStyleNode, s);
      ReadChartRegressionEquationStyle(styleNode, AChart, ASeries);
    end;
    subNode := subNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartRegressionStyle(AStyleNode: TDOMNode;
  AChart: TsChart; ASeries: TsChartSeries);
var
  s, nodeName: String;
  trendline: TsChartTrendline;
  rt: TsTrendlineType;
  value: Double;
  intValue: Integer;
begin
  if not ASeries.SupportsTrendline then
    exit;

  trendline := TsOpenedTrendlineSeries(ASeries).Trendline;

  AStyleNode := AStyleNode.FirstChild;
  while Assigned(AStyleNode) do
  begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:graphic-properties':
        GetChartLineProps(AStyleNode, AChart, trendline.Line);
      'style:chart-properties':
        begin
          s := GetAttrValue(AStyleNode, 'chart:regression-name');
          trendline.Title := s;

          s := GetAttrValue(AStyleNode, 'chart:regression-type');
          for rt in TsTrendlineType do
            if (s <> '') and (TRENDLINE_TYPE[rt] = s) then
            begin
              trendline.TrendlineType := rt;
              break;
            end;

          s := GetAttrValue(AStyleNode, 'chart:regression-max-degree');
          if TryStrToInt(s, intValue) then
            trendline.PolynomialDegree := intValue;

          s := GetAttrValue(AStyleNode, 'chart:regression-extrapolate-forward');
          if TryStrToFloat(s, value, FPointSeparatorSettings) then
            trendline.ExtrapolateForwardBy := value
          else
            trendline.ExtrapolateForwardBy := 0.0;

          s := GetAttrValue(AStyleNode, 'chart:regression-extrapolate-backward');
          if TryStrToFloat(s, value, FPointSeparatorSettings) then
            trendline.ExtrapolateBackwardBy := value
          else
            trendline.ExtrapolateBackwardBy := 0.0;

          s := GetAttrValue(AStyleNode, 'chart:regression-force-intercept');
          trendline.ForceYIntercept := (s = 'true');

          s := GetAttrValue(AStyleNode, 'chart:regression-intercept-value');
          if TryStrToFloat(s, value, FPointSeparatorSettings) then
            trendline.YInterceptValue := value;
        end;
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartSeriesDataPointStyle(AStyleNode: TDOMNode;
  AChart: TsChart; ASeries: TsChartSeries; var AFill: TsChartFill; var ALine: TsChartLine;
  var APieOffset: Integer);
var
  nodeName, s: string;
  value: Double;
begin
  AFill := nil;
  ALine := nil;
  APieOffset := 0;

  nodeName := AStyleNode.NodeName;
  AStyleNode := AStyleNode.FirstChild;
  while AStyleNode <> nil do
  begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:graphic-properties':
        begin
          AFill := TsChartFill.Create;
          if not GetChartFillProps(AStyleNode, AChart, AFill) then FreeAndNil(AFill);
          ALine := TsChartLine.Create;
          if not GetChartLineProps(AStyleNode, AChart, ALine) then FreeAndNil(ALine);
        end;
      'style:chart-properties':
        if ASeries is TsPieSeries then
        begin
          s := GetAttrValue(AStyleNode, 'chart:pie-offset');
          if TryStrToFloat(s, value, FPointSeparatorSettings) then
            APieOffset := round(value);
        end;
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartSeriesErrorBarProps(
  ANode, AStyleNode: TDOMNode; AChart: TsChart; ASeries: TsChartSeries);
var
  s: String;
  styleNode: TDOMNode;
  errorBars: TsChartErrorBars;
begin
  s := GetAttrValue(ANode, 'chart:dimension');
  case s of
    'x': errorBars := ASeries.XErrorBars;
    'y': errorBars := ASeries.YErrorBars;
    else exit;
  end;

  s := GetAttrValue(ANode, 'chart:style-name');
  if s <> '' then
  begin
    styleNode := FindStyleNode(AStyleNode, s);
    ReadChartSeriesErrorBarStyle(styleNode, AChart, errorBars);
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartSeriesErrorBarStyle(
  AStyleNode: TDOMNode; AChart: TsChart; AErrorBars: TsChartErrorBars);
var
  nodeName, s: String;
  x: Double;
begin
  AStyleNode := AStyleNode.FirstChild;
  while AStyleNode <> nil do
  begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:chart-properties':
        begin
          s := GetAttrValue(AStyleNode, 'chart:error-category');
          case s of
            'constant': AErrorBars.Kind := cebkConstant;
            'cell-range': AErrorBars.Kind := cebkCellRange;
            'percentage': AErrorBars.Kind := cebkPercentage;
            else
              exit;
            // To do: support the statistical categories 'standard-error',
            //        'variance', 'standard-deviation', 'error-margin'
          end;

          s := GetAttrValue(AStyleNode, 'chart:error-lower-limit');
          if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
            AErrorBars.ValueNeg := x;

          s := GetAttrValue(AStyleNode, 'chart:error-upper-limit');
          if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
            AErrorBars.ValuePos := x;

          s := GetAttrValue(AStyleNode, 'chart:error-lower-indicator');
          AErrorBars.ShowNeg := (s = 'true');

          s := GetAttrValue(AStyleNode, 'chart:error-upper-indicator');
          AErrorBars.ShowPos := (s = 'true');

          s := GetAttrValue(AStyleNode, 'chart:error-percentage');
          if (s <> '') and TryStrToFloat(s, x, FPointSeparatorSettings) then
          begin
            AErrorBars.ValueNeg := x;
            AErrorBars.ValuePos := x;
          end;

          ReadChartCellRange(AStyleNode, 'chart:error-lower-range', AErrorBars.RangeNeg);
          ReadChartCellRange(AStyleNode, 'chart:error-upper-range', AErrorBars.RangePos);
        end;
      'style:graphic-properties':
        GetChartLineProps(AStyleNode, AChart, AErrorBars.Line);
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartSeriesProps(ANode, AStyleNode: TDOMNode;
  AChart: TsChart);
var
  s, nodeName: String;
  series: TsChartSeries;
  fill: TsChartFill;
  line: TsChartLine;
  subNode: TDOMNode;
  styleNode: TDOMNode;
  xyCounter: Integer;
  i, n, pieOffset, ptIndex: Integer;
begin
  s := GetAttrValue(ANode, 'chart:class');
  if (FChartType = ctStock) and (s = '') then
    series := FStockSeries
  else
    case s of
      'chart:area':
        series := TsAreaSeries.Create(AChart);
      'chart:bar':
        series := TsBarSeries.Create(AChart);
      'chart:bubble':
        series := TsBubbleSeries.Create(AChart);
      'chart:circle':
        begin
          series := TsPieSeries.Create(AChart);
          if FChartType = ctRing then
            TsPieSeries(series).InnerRadiusPercent := 50;
        end;
      'chart:line':
        begin
          series := TsLineSeries.Create(AChart);
          TsLineSeries(series).Interpolation := AChart.Interpolation;
        end;
      'chart:radar':
        series := TsRadarSeries.Create(AChart);
        // Note: In ods, line and symbol colors are equal!
      'chart:filled-radar':
        series := TsFilledRadarSeries.Create(AChart);
      'chart:scatter':
        series := TsScatterSeries.Create(AChart);
      // 'chart:stock': --- has already been created
      else
        raise Exception.Create('Unknown/unsupported series type.');
    end;

  ReadChartCellAddr(ANode, 'chart:label-cell-address', series.TitleAddr);
  if (series is TsStockSeries) then
  begin
    // The file contains the range in the order Open-Low-High-Close
    if FStockSeries.OpenRange.IsEmpty and FStockSeries.CandleStick then
      ReadChartCellRange(ANode, 'chart:values-cell-range-address', FStockSeries.OpenRange)
    else
    if FStockSeries.LowRange.IsEmpty then
      ReadChartCellRange(ANode, 'chart:values-cell-range-address', FStockSeries.LowRange)
    else
    if FStockSeries.HighRange.IsEmpty then
      ReadChartCellRange(ANode, 'chart:values-cell-range-address', FStockSeries.HighRange)
    else
    if FStockSeries.CloseRange.IsEmpty then
      ReadChartCellRange(ANode, 'chart:values-cell-range-address', FStockSeries.CloseRange);
  end
  else
  if (series is TsBubbleSeries) then
  begin
    TsBubbleSeries(series).BubbleSizeMode := bsmArea;
    ReadChartCellRange(ANode, 'chart:values-cell-range-address', TsBubbleSeries(series).BubbleRange);
  end
  else
    ReadChartCellRange(ANode, 'chart:values-cell-range-address', series.YRange);

  s := GetAttrValue(ANode, 'chart:attached-axis');
  if s = 'primary-y' then
    series.YAxis := calPrimary
  else if s = 'secondary-y' then
    series.YAxis := calSecondary;

  if series.XRange.IsEmpty then
    series.XRange.CopyFrom(series.Chart.XAxis.CategoryRange);

  xyCounter := 0;
  subnode := ANode.FirstChild;
  ptIndex := 0;
  while subnode <> nil do
  begin
    nodeName := subNode.NodeName;
    case nodeName of
      'chart:domain':
        begin
          if xyCounter = 0 then
          begin
            ReadChartCellRange(subnode, 'table:cell-range-address', series.XRange);
            inc(xyCounter);
          end else
          if xyCounter = 1 then
          begin
            series.YRange.CopyFrom(series.XRange);
            ReadChartCellRange(subnode, 'table:cell-range-address', series.XRange)
          end;
        end;
      'loext:property-mapping':
        begin
          s := GetAttrValue(subnode, 'loext:property');
          case s of
            'FillColor':
              ReadChartCellRange(subNode, 'loext:cell-range-address', series.FillColorRange);
            'BorderColor':
              ReadChartCellRange(subNode, 'loext:cell-range-address', series.LineColorRange);
          end;
        end;
      'chart:regression-curve':
        ReadChartRegressionProps(subNode, AStyleNode, AChart, series);
      'chart:data-point':
        begin
          fill := nil;
          line := nil;
          n := 1;
          pieOffset := 0;
          s := GetAttrValue(subnode, 'chart:style-name');
          if s <> '' then
          begin
            styleNode := FindStyleNode(AStyleNode, s);
            ReadChartSeriesDataPointStyle(styleNode, AChart, series, fill, line, pieOffset); // creates fill and line!
    //      end;   // <<<<<<<<<<<<<<<<<<<<               // !!! wp: putting these two lines back in creates a crash at the end with the stock series
    //      begin  // <<<<<<<<<<<<<<<<<<<<
            s := GetAttrValue(subnode, 'chart:repeated');
            if (s <> '') then
              n := StrToIntDef(s, 1);
            for i := 0 to n-1 do
            begin
              series.DataPointStyles.AddFillAndLine(ptIndex, fill, line, pieOffset);
              inc(ptIndex);
            end;
            fill.Free;  // the styles have been copied to the series datapoint style list and are not needed any more.
            line.Free;
          end;
        end;
      'chart:error-indicator':
        ReadChartSeriesErrorbarProps(subNode, AStyleNode, AChart, series);
    end;
    subnode := subNode.NextSibling;
  end;

  if series.XRange.IsEmpty and (AChart.Series.Count > 0) then
    series.XRange.CopyFrom(AChart.Series[0].XRange);

  if series.LabelRange.IsEmpty then
    series.LabelRange.CopyFrom(AChart.XAxis.CategoryRange);

  s := GetAttrValue(ANode, 'chart:style-name');
  if s <> '' then
  begin
    styleNode := FindStyleNode(AStyleNode, s);
    ReadChartSeriesStyle(styleNode, AChart, series);
  end;

  if (series is TsPieSeries) and (FPieSeriesStartAngle <> 999) then
    TsPieSeries(series).StartAngle := FPieSeriesStartAngle;
end;

procedure TsSpreadOpenDocChartReader.ReadChartSeriesStyle(AStyleNode: TDOMNode;
  AChart: TsChart; ASeries: TsChartSeries);
var
  nodeName: String;
  s: String;
  css: TsChartSeriesSymbol;
  value: Double;
  rel: Boolean;
  dataLabels: TsChartDataLabels = [];
  childNode1, childNode2, childNode3: TDOMNode;
begin
  // Defaults
  ASeries.LabelBorder.Style := clsNoLine;
  ASeries.LabelBackground.Style := cfsNoFill;

  nodeName := AStyleNode.NodeName;

  // Number format of labels as number...
  s := GetAttrValue(AStyleNode, 'style:data-style-name');
  if s <> '' then
    ASeries.LabelFormat := TsChartNumberFormatList(FNumberFormatList).FindFormatByName(s);
  // ... and as percentage
  s := GetAttrValue(AStyleNode, 'style:percentage-data-style-name');
  if s <> '' then
    ASeries.LabelFormatPercent := TsChartNumberFormatList(FNumberFormatList).FindFormatByName(s);

  AStyleNode := AStyleNode.FirstChild;
  while AStyleNode <> nil do begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:graphic-properties':
        begin
          if ASeries.ChartType in [ctBar] then
            ASeries.Line.Style := clsSolid;
          GetChartLineProps(AStyleNode, AChart, ASeries.Line);
          if ((ASeries is TsRadarSeries) and (ASeries.ChartType = ctRadar)) then //or (ASeries is TsCustomLineSeries) then
          begin
            // In ods, symbols and lines have the same color
            TsRadarSeries(ASeries).SymbolFill.Style := cfsSolid;
            TsRadarSeries(ASeries).SymbolFill.Color := ASeries.Line.Color;
            TsRadarSeries(ASeries).SymbolBorder.Style := clsNoLine;
          end else
          if (ASeries is TsScatterSeries) then
            GetChartFillProps(AStyleNode, AChart, TsScatterSeries(ASeries).SymbolFill)
          else
          if (ASeries is TsLineSeries) then
            GetChartFillProps(AStyleNode, AChart, TsLineSeries(ASeries).SymbolFill)
          else
            GetChartFillProps(AStyleNode, AChart, ASeries.Fill);
        end;
      'style:text-properties':
        GetChartTextProps(AStyleNode, ASeries.LabelFont);
      'style:chart-properties':
        begin
          s := GetAttrValue(AStyleNode, 'chart:label-position');
          case s of
            '': ASeries.LabelPosition := lpDefault;
            'outside': ASeries.LabelPosition := lpOutside;
            'inside': ASeries.LabelPosition := lpInside;
            'center': ASeries.LabelPosition := lpCenter;
            'top': ASeries.LabelPosition := lpAbove;
            'bottom': ASeries.LabelPosition := lpBelow;
            'near-origin': ASeries.LabelPosition := lpNearOrigin;
          end;

          // Label border color
          s := GetAttrValue(AStyleNode, 'loext:label-stroke-color');
          if s <> '' then
            ASeries.LabelBorder.Color := ChartColor(HTMLColorStrToColor(s));
          // Label border transparency
          s := GetAttrValue(AStyleNode, 'loext:label-stroke-opacity');
          if TryPercentStrToFloat(s, value) then
            ASeries.LabelBorder.Color.Transparency := 1.0 - value;
          // Label border line style
          s := GetAttrValue(AStyleNode, 'loext:label-stroke');
          if s <> '' then
            case s of
              'none': ASeries.LabelBorder.Style := clsNoLine;
              else    ASeries.LabelBorder.Style := clsSolid;
            end;

          // Items in data labels
          s := GetAttrValue(AStyleNode, 'chart:data-label-number');
          case s of
            'none': ;
            'value': Include(dataLabels, cdlValue);
            'percentage': Include(datalabels, cdlPercentage);
            'value-and-percentage': dataLabels := datalabels + [cdlValue, cdlPercentage];
          end;
          s := GetAttrValue(AStyleNode, 'chart:data-label-text');
          if s = 'true' then
            Include(dataLabels, cdlCategory);
          s := GetAttrValue(AStyleNode, 'chart:data-label-series');
          if s = 'true' then
            Include(dataLabels, cdlSeriesName);
          s := GetAttrValue(AStyleNode, 'chart:data-label-symbol');
          if s = 'true' then
            Include(dataLabels, cdlSymbol);
          ASeries.DataLabels := dataLabels;

          childNode1 := AStyleNode.FirstChild;
          while childNode1 <> nil do
          begin
            nodeName := childNode1.NodeName;
            if nodeName = 'chart:label-separator' then
            begin
              childNode2 := childNode1.FirstChild;
              while childNode2 <> nil do
              begin
                nodeName := childNode2.NodeName;
                if nodeName = 'text:p' then
                begin
                  ASeries.LabelSeparator := GetNodeValue(childNode2);
                  if ASeries.LabelSeparator = '' then
                  begin
                    childNode3 := childNode2.FirstChild;
                    while childNode3 <> nil do
                    begin
                      nodeName := childNode3.NodeName;
                      if nodeName = 'text:line-break' then
                      begin
                        ASeries.LabelSeparator := LineEnding;
                        break;
                      end;
                      childNode3 := childNode3.NextSibling;
                    end;
                  end;
                end;
                childNode2 := childNode2.NextSibling;
              end;
            end;
            childNode1 := childNode1.NextSibling;
          end;

          if (ASeries is TsCustomLineSeries) then
          begin
            s := GetAttrValue(AStyleNode, 'chart:symbol-name');
            if s <> '' then
            begin
              TsOpenedCustomLineSeries(ASeries).ShowSymbols := true;
              for css in TsChartSeriesSymbol do
                if SYMBOL_NAMES[css] = s then
                begin
                  TsOpenedCustomLineSeries(ASeries).Symbol := css;
                  break;
                end;
              s := GetAttrValue(AStyleNode, 'chart:symbol-width');
              if (s <> '') and EvalLengthStr(s, value, rel) then
                TsOpenedCustomLineSeries(ASeries).SymbolWidth := value;
              s := GetAttrValue(AStyleNode, 'chart:symbol-height');
              if (s <> '') and EvalLengthStr(s, value, rel) then
                TsOpenedCustomLineSeries(ASeries).SymbolHeight := value;
            end else
              TsOpenedCustomLineSeries(ASeries).ShowSymbols := false;
          end;
        end;

    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartStockSeriesStyle(AStyleNode: TDOMNode;
  AChart: TsChart; ASeries: TsStockSeries; ANodeName: String);
var
  nodeName: String;
begin
  nodeName := AStyleNode.NodeName;
  if nodeName = 'style:style' then
  begin
    AStyleNode := AStyleNode.Firstchild;
    while AStyleNode <> nil do
    begin
      nodeName := AStyleNode.NodeName;
      if nodeName = 'style:graphic-properties' then
      begin
        if ANodeName = 'chart:stock-gain-marker' then
        begin
          GetChartFillProps(AStyleNode, AChart, ASeries.CandleStickUpFill);
          GetChartLineProps(AStyleNode, AChart, ASeries.CandleStickUpBorder);
        end else
        if ANodeName = 'chart:stock-loss-marker' then
        begin
          GetChartFillProps(AStyleNode, AChart, ASeries.CandleStickDownFill);
          GetChartLineProps(AStyleNode, AChart, ASeries.CandleStickDownBorder);
        end else
        if ANodeName = 'chart:stock-range-line' then
          GetChartLineProps(AStyleNode, AChart, ASeries.RangeLine);
      end;
      AStyleNode := AStyleNode.NextSibling;
    end;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartTitleProps(ANode, AStyleNode: TDOMNode;
  AChart: TsChart; ATitle: TsChartText);
var
  textNode, childNode: TDOMNode;
  styleNode: TDOMNode;
  nodeName: String;
  s: String;
  value: Double;
  rel: Boolean;
begin
  s := '';
  textNode := ANode.FirstChild;
  while textNode <> nil do
  begin
    nodeName := textNode.NodeName;
    if nodeName = 'text:p' then
    begin
      // Each 'text:p' node is a paragraph --> we insert a line break except for the first paragraph
      if s <> '' then
        s := s + LineEnding;
      childNode := textNode.FirstChild;
      while childNode <> nil do
      begin
        nodeName := childNode.NodeName;
        case nodeName of
          '#text':
            s := s + childNode.TextContent;
          'text:s':
            s := s + ' ';
          'text:line-break':
            s := s + LineEnding;
          // to do: Is rtf formatting supported here? (text:span)
        end;
        childNode := childNode.NextSibling;
      end;
    end;
    textNode := textNode.NextSibling;
  end;
  ATitle.Caption := s;

  s := GetAttrValue(ANode, 'svg:x');
  if (s <> '') and EvalLengthStr(s, value, rel) then
    if not rel then
      ATitle.PosX := value;

  s := GetAttrValue(ANode, 'svg:y');
  if (s <> '') and EvalLengthStr(s, value, rel) then
    if not rel then
      AChart.Legend.PosY := value;

  s := GetAttrValue(ANode, 'chart:style-name');
  styleNode := FindStyleNode(AStyleNode, s);
  ReadChartTitleStyle(styleNode, AChart, ATitle);
end;

procedure TsSpreadOpenDocChartReader.ReadChartTitleStyle(AStyleNode: TDOMNode;
  AChart: TsChart; ATitle: TsChartText);
var
  nodeName: String;
  s: String;
  value: Double;
begin
  nodeName := AStyleNode.NodeName;
  AStyleNode := AStyleNode.FirstChild;
  while AStyleNode <> nil do begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:chart-properties':
        begin
          s := GetAttrValue(AStyleNode, 'style:rotation-angle');
          if (s <> '') and TryStrToFloat(s, value, FPointSeparatorSettings) then
            ATitle.RotationAngle := round(value);
        end;
      'style:graphic-properties':
        begin
          GetChartLineProps(AStyleNode, AChart, ATitle.Border);
          GetChartFillProps(AStyleNode, AChart, ATitle.Background);
        end;
      'style:text-properties':
        GetChartTextProps(AStyleNode, ATitle.Font);
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadCharts(AStream: TStream);
var
  i: Integer;
begin
  for i := 0 to FChartFiles.Count-1 do
    ReadChartFiles(AStream, FChartFiles[i]);
end;

{ Reads the styles stored in the Object files. }
procedure TsSpreadOpenDocChartReader.ReadObjectStyles(ANode: TDOMNode;
  AChart: TsChart; ARoot: String);
var
  nodeName: String;
begin
  nodeName := ANode.NodeName;
  ANode := ANode.FirstChild;
  while ANode <> nil do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'draw:stroke-dash':  // read line pattern
        ReadObjectLineStyles(ANode, AChart);
      'draw:hatch':        // read hatch pattern
        ReadObjectHatchStyles(ANode, AChart);
      'draw:gradient':     // gradient definition
        ReadObjectGradientStyles(ANode, AChart);
      'draw:fill-image':
        ReadObjectFillImages(ANode, AChart, ARoot);
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadObjectFillImages(ANode: TDOMNode;
  AChart: TsChart; ARoot: String);
var
  styleName: String;
  imgFileName: string;
  imgStream: TStream;
  img: TFPCustomImage;
begin
  styleName := GetAttrValue(ANode, 'draw:display-name');
  if styleName = '' then
    styleName := GetAttrValue(ANode, 'draw:name');

  imgFileName := GetAttrValue(ANode, 'xlink:href');
  if imgFileName = '' then
    exit;

  imgStream := TStreamList(FStreamList).FindByName(ARoot + imgFileName);
  if imgStream <> nil then
  begin
    img := TFPMemoryImage.Create(0, 0);     // do not destroy this image here!
    img.LoadFromStream(imgStream);
    AChart.Images.AddImage(styleName, img);
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadObjectGradientStyles(ANode: TDOMNode;
  AChart: TsChart);
var
  i: Integer;
  s: String;
  styleName: String;
  gs: TsChartGradientStyle;
  gradientStyle: TsChartGradientStyle = cgsLinear;
  startColor: TsChartColor;
  endColor: TsChartColor;
  startIntensity, endIntensity: Double;
  border, centerX, centerY: Double;
  angle: Double = 0.0;
begin
  styleName := GetAttrValue(ANode, 'draw:display-name');
  if styleName ='' then
    styleName := GetAttrValue(ANode, 'draw:name');

  s := GetAttrValue(ANode, 'draw:style');
  if s <> '' then
    for gs in TsChartGradientStyle do
      if GRADIENT_STYLES[gs] = s then
      begin
        gradientStyle := gs;
        break;
      end;

  s := GetAttrValue(ANode, 'draw:start-color');
  if s <> '' then
    startColor := ChartColor(HTMLColorStrToColor(s))
  else
    startColor := ChartColor(scSilver);

  s := GetAttrValue(ANode, 'draw:end-color');
  if s <> '' then
    endColor := ChartColor(HTMLColorStrToColor(s))
  else
    endColor := ChartColor(scWhite);

  s := GetAttrValue(ANode, 'draw:start-intensity');
  if not TryPercentStrToFloat(s, startIntensity) then
    startIntensity := 1.0;
  startIntensity := EnsureRange(startIntensity, 0.0, 1.0);

  s := GetAttrValue(ANode, 'draw:end-intensity');
  if not TryPercentStrToFloat(s, endIntensity) then
    endIntensity := 1.0;
  endIntensity := EnsureRange(endIntensity, 0.0, 1.0);

  s := GetAttrValue(ANode, 'draw:border');
  if not TryPercentStrToFloat(s, border) then
    border := 0.0;

  s := GetAttrValue(ANode, 'draw:angle');
  if s <> '' then begin
    for i := Length(s) downto 1 do
      if not (s[i] in ['0'..'9', '.', '+', '-']) then Delete(s, i, 1);
    angle := StrToFloatDef(s, 0.0, FPointSeparatorSettings);
    { ods has angle=0 in vertical direction, and orientation is CW
      --> We must transform to fps angular orientations (0° horizontal, CCW)
      But axial gradient uses "normal" angle }
    if gradientstyle <> cgsAxial then
      angle := FMod(90.0 + angle, 360.0)
    else
      angle := FMod(angle, 360.0);
  end;

  s := GetAttrValue(ANode, 'draw:cx');
  if not TryPercentStrToFloat(s, centerX) then
    centerX := 0.0;

  s := GetAttrValue(ANode, 'draw:cy');
  if not TryPercentStrToFloat(s, centerY) then
    centerY := 0.0;

  if gradientStyle <> cgsAxial then
    AChart.Gradients.AddGradient(styleName, gradientStyle,
      ModifyColor(startColor, startIntensity),
      ModifyColor(endColor, endIntensity),
      angle, centerX, centerY, border, 1.0)
  else
    AChart.Gradients.AddGradient(styleName, gradientStyle,
      ModifyColor(endColor, startIntensity),
      ModifyColor(startColor, endIntensity),
      angle, centerX, centerY, border, 1.0)
end;

{ Read the hatch pattern stored in the "draw:hatch" nodes of the chart's
  Object styles.xml file. }
procedure TsSpreadOpenDocChartReader.ReadObjectHatchStyles(ANode: TDOMNode; AChart: TsChart);
var
  s: String;
  styleName: String;
  hs, hatchStyle: TsChartHatchStyle;
  hatchColor: TsChartColor;
  hatchDist: Double;
  hatchAngle: Double;
  rel: Boolean;
begin
  styleName := GetAttrValue(ANode, 'draw:display-name');
  if styleName = '' then
    styleName := GetAttrValue(ANode, 'draw:name');

  s := GetAttrValue(ANode, 'draw:style');
  hatchStyle := chsSingle;
  for hs in TsChartHatchStyle do
    if HATCH_STYLES[hs] = s then
    begin
      hatchStyle := hs;
      break;
    end;

  s := GetAttrValue(ANode, 'draw:color');
  hatchColor := ChartColor(IfThen(s <> '', HTMLColorStrToColor(s), scBlack));

  s := GetAttrValue(ANode, 'draw:distance');
  if not EvalLengthStr(s, hatchDist, rel) then
    hatchDist := 2.0;

  s := GetAttrValue(ANode, 'draw:rotation');
  if TryStrToFloat(s, hatchAngle, FPointSeparatorSettings) then
    hatchAngle := hatchAngle / 10
  else
    hatchAngle := 0;

  AChart.Hatches.AddLineHatch(styleName, hatchStyle, hatchColor, hatchdist, 0.1, hatchAngle);
end;

{ Reads the line styles stored as "draw:stroke-dash" nodes in the chart's
  Object styles.xml file. }
procedure TsSpreadOpenDocChartReader.ReadObjectLineStyles(ANode: TDOMNode; AChart: TsChart);
var
  styleName: String;
  s: String;
  dots1: Integer;
  dots2: Integer = 0;
  dots1Length: double = 3.0;
  dots2Length: double = 0.0;
  distance: double = 3.0;
  rel1: Boolean = false;
  rel2: Boolean = false;
  relDist: Boolean = false;
begin
  styleName := GetAttrValue(ANode, 'draw:display-name');
  if styleName = '' then
    styleName := GetAttrValue(ANode, 'draw:name');

  s := GetAttrValue(ANode, 'draw:dots1');
  dots1 := StrToIntDef(s, 1);

  s := GetAttrValue(ANode, 'draw:dots2');
  dots2 := StrToIntDef(s, 0);

  s := GetAttrValue(ANode, 'draw:dots1-length');
  if not EvalLengthStr(s, dots1Length, rel1) then
    dots1Length := 3.0;

  s := GetAttrValue(ANode, 'draw:dots2-length');
  if not EvalLengthStr(s, dots2Length, rel2) then
    dots2Length := 0.0;

  s := GetAttrValue(ANode, 'draw:distance');
  if not EvalLengthstr(s, distance, relDist) then
    distance := 3.0;

  AChart.LineStyles.Add(styleName, dots1Length, dots1, dots2Length, dots2, distance, rel1 or rel2 or relDist);
end;

procedure TsSpreadOpenDocChartReader.ReadPictureFile(AStream: TStream;
  AFileName: String);
var
  memStream: TMemoryStream;
  img: TFPCustomImage;
  item: TStreamItem;
begin
  memStream := TMemoryStream.Create;
  try
    if UnzipToStream(AStream, AFileName, memStream) then
    begin
      memstream.Position := 0;
      item := TStreamItem.Create;
      item.Name := AFileName;
      item.Stream := TMemoryStream.Create;
      item.Stream.CopyFrom(memStream, memStream.Size);
      item.Stream.Position := 0;
      FStreamList.Add(item);
    end;
  finally
    memstream.Free;
  end;
end;


{------------------------------------------------------------------------------}
{                        TsSpreadOpenDocChartWriter                            }
{------------------------------------------------------------------------------}

constructor TsSpreadOpenDocChartWriter.Create(AWriter: TsBasicSpreadWriter);
begin
  inherited Create(AWriter);

  FPointSeparatorSettings := SysUtils.DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator:='.';

  FNumberFormatList := TsChartNumberFormatList.Create;
end;

destructor TsSpreadOpenDocChartWriter.Destroy;
begin
  FNumberFormatList.Free;
  inherited;
end;

procedure TsSpreadOpenDocChartWriter.AddChartsToZip(AZip: TZipper);
var
  i: Integer;
begin
  for i := 0 to TsWorkbook(Writer.Workbook).GetChartCount-1 do
  begin
    AZip.Entries.AddFileEntry(
      FSCharts[i], Format(OPENDOC_PATH_CHART_CONTENT, [i+1]));
    AZip.Entries.AddFileEntry(
      FSObjectStyles[i], Format(OPENDOC_PATH_CHART_STYLES, [i+1]));
  end;
end;

{ Writes the chart entries needed in the META-INF/manifest.xml file }
procedure TsSpreadOpenDocChartWriter.AddToMetaInfManifest(AStream: TStream);
var
  i: Integer;
begin
  for i:=0 to TsWorkbook(Writer.Workbook).GetChartCount-1 do
  begin
    AppendToStream(AStream, Format(
      '  <manifest:file-entry manifest:media-type="application/vnd.oasis.opendocument.chart" manifest:full-path="Object %d/" />' + LE,
      [i+1]
    ));
    AppendToStream(AStream, Format(
      '  <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="Object %d/content.xml" />' + LE,
      [i+1]
    ));
    AppendToStream(AStream, Format(
      ' <manifest:file-entry manifest:media-type="text/xml" manifest:full-path="Object %d/styles.xml" />' + LE,
      [i+1]
    ));

    // Object X/meta.xml and ObjectReplacement/Object X are not necessarily needed.
  end;
end;

procedure TsSpreadOpenDocChartWriter.CreateStreams;
var
  i, n: Integer;
begin
  n := TsWorkbook(Writer.Workbook).GetChartCount;
  SetLength(FSCharts, n);
  SetLength(FSObjectStyles, n);
  for i := 0 to n - 1 do
  begin
    FSCharts[i] := CreateTempStream(Writer.Workbook, 'fpsCh');
    FSObjectStyles[i] := CreateTempStream(Writer.Workbook, 'fpsOS');
  end;
end;

procedure TsSpreadOpenDocChartWriter.DestroyStreams;
var
  i: Integer;
begin
  for i := 0 to High(FSCharts) do
  begin
    DestroyTempStream(FSCharts[i]);
    DestroyTempStream(FSObjectStyles[i]);
  end;
  Setlength(FSCharts, 0);
  SetLength(FSObjectStyles, 0);
end;

function TsSpreadOpenDocChartWriter.GetChartAxisStyleAsXML(
  Axis: TsChartAxis; AIndent, AStyleID: Integer): String;
var
  chart: TsChart;
  indent: String;
  angle: single;
  textProps: String = '';
  graphProps: String = '';
  chartProps: String = '';
  numStyle: String = 'N0';
begin
  Result := '';
  if not Axis.Visible then
    exit;

  chart := Axis.Chart;

  // Get number format, use percent format for stacked percentage axis
  if (Axis = chart.YAxis) and (chart.StackMode = csmStackedPercentage) then
    numStyle := GetNumberFormatID(Axis.LabelFormatPercent)
  else
    numStyle := GetNumberFormatID(Axis.LabelFormat);
  if numStyle <> 'N0' then
    chartProps := chartProps + 'chart:link-data-style-to-source="false" ';

  // Show axis labels
  if Axis.ShowLabels then
    chartProps := chartProps + 'chart:display-label="true" ';

  // Logarithmic axis
  if Axis.Logarithmic then
    chartProps := chartProps + 'chart:logarithmic="true" ';

  // Axis scaling: minimum, maximum, tick intervals
  if not Axis.AutomaticMin then
    chartProps := chartProps + Format('chart:minimum="%g" ', [Axis.Min], FPointSeparatorSettings);
  if not Axis.AutomaticMax then
    chartProps := chartProps + Format('chart:maximum="%g" ', [Axis.Max], FPointSeparatorSettings);
  if not Axis.AutomaticMajorInterval then
    chartProps := chartProps + Format('chart:interval-major="%g" ', [Axis.MajorInterval], FPointSeparatorSettings);
  if not Axis.AutomaticMinorSteps then
    chartProps := chartProps + Format('chart:interval-minor-divisor="%d" ', [Axis.MinorCount]);

  // Position of the axis
  case Axis.Position of
    capStart: chartProps := chartProps + 'chart:axis-position="start" ';
    capEnd: chartProps := chartProps + 'chart:axis-position="end" ';
    capValue: chartProps := chartProps + Format('chart:axis-position="%g" ', [Axis.PositionValue], FPointSeparatorSettings);
  end;

  // Tick marks
  if (chart.GetChartType in [ctRadar, ctFilledRadar]) and (Axis = chart.YAxis) then
  begin
    // Radar series needs a "false" to hide the tick-marks
    chartProps := chartProps + Format('chart:tick-marks-major-inner="%s" ', [FALSE_TRUE[catInside in Axis.MajorTicks]]);
    chartProps := chartProps + Format('chart:tick-marks-major-outer="%s" ', [FALSE_TRUE[catOutside in Axis.MajorTicks]]);
    chartProps := chartProps + Format('chart:tick-marks-minor-inner="%s" ', [FALSE_TRUE[catInside in Axis.MinorTicks]]);
    chartProps := chartProps + Format('chart:tick-marks-minor-outer="%s" ', [FALSE_TRUE[catOutside in Axis.MinorTicks]]);
  end else
  begin
    // The other series hide the tick-marks by default.
    if (catInside in Axis.MajorTicks) then
      chartProps := chartProps + 'chart:tick-marks-major-inner="true" ';
    if (catOutside in Axis.MajorTicks) then
      chartProps := chartProps + 'chart:tick-marks-major-outer="true" ';
    if (catInside in Axis.MinorTicks) then
      chartProps := chartProps + 'chart:tick-marks-minor-inner="true" ';
    if (catOutside in Axis.MinorTicks) then
      chartProps := chartProps + 'chart:tick-marks-minor-outer="true" ';
  end;

  // Inverted axis direction
  if Axis.Inverted then
    chartProps := chartProps + 'chart:reverse-direction="true" ';

  // Rotated axis labels
  angle := Axis.LabelRotation;
  chartProps := chartProps + Format('style:rotation-angle="%.1f" ', [angle], FPointSeparatorSettings);

  // Bar series gap distance and over lap -- why did they put it here?
  if (chart.GetChartType = ctBar) and (Axis = chart.YAxis) then
    chartProps := chartProps + Format(
      'chart:gap-width="%d" chart:overlap="%d" ', [chart.BarGapWidthPercent, chart.BarOverlapPercent]);

  // Label color
  graphProps := 'svg:stroke-color="' + ColorToHTMLColorStr(Axis.AxisLine.Color.Color) + '" ';

  // Label font
  textProps := TsSpreadOpenDocWriter(Writer).WriteFontStyleXMLAsString(Axis.LabelFont);

  // Putting it all together...
  indent := DupeString(' ', AIndent);
  Result := Format(
    indent + '<style:style style:name="ch%d" style:family="chart" style:data-style-name="%s">' + LE +
    indent + '  <style:chart-properties %s/>' +  LE +
    indent + '  <style:graphic-properties %s/>' + LE +
    indent + '  <style:text-properties %s/>' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, numStyle, chartProps, graphProps, textProps ]
  );
end;

function TsSpreadOpenDocChartWriter.GetChartBackgroundStyleAsXML(
  AChart: TsChart; AFill: TsChartFill; ABorder: TsChartLine;
  AIndent, AStyleID: Integer): String;
var
  indent: String;
  fillStr: String = '';
  borderStr: String = '';
begin
  fillStr := GetChartFillStyleGraphicPropsAsXML(AChart, AFill);
  borderStr := GetChartLineStyleGraphicPropsAsXML(AChart, ABorder);
  indent := DupeString(' ', AIndent);
  Result := Format(
    indent + '<style:style style:name="ch%d" style:family="chart">' + LE +
    indent + '  <style:graphic-properties %s%s />' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, fillStr, borderStr ]
  );
end;

{ <style:style style:name="ch7" style:family="chart">
    <style:chart-properties chart:auto-position="true" style:rotation-angle="0"/>
    <style:text-properties fo:font-size="9pt" style:font-size-asian="9pt" style:font-size-complex="9pt"/>
  </style:style>

  ACaptionKind = 1 ---> Title
  ACaptionKind = 2 ---> SubTitle
  ACaptionKind = 3 ---> x xis
  ACaptionKind = 4 ---> y axis
  ACaptionKind = 5 ---> x2 axis
  ACaptionKind = 6 ---> y2 axis }
function TsSpreadOpenDocChartWriter.GetChartCaptionStyleAsXML(AChart: TsChart;
  ACaptionKind, AIndent, AStyleID: Integer): String;
var
  title: TsChartText;
  axis: TsChartAxis;
  font: TsFont;
  indent: String;
  rotAngle: Single;
  rotAngleStr: String = '';
  chartProps: String = '';
  textProps: String = '';
begin
  Result := '';

  case ACaptionKind of
    1, 2:
      begin
        if ACaptionKind = 1 then title := AChart.Title else title := AChart.Subtitle;
        font := title.Font;
        rotAngle := title.RotationAngle;
      end;
    3, 4, 5, 6:
      begin
        case ACaptionKind of
          3: axis := AChart.XAxis;
          4: axis := AChart.YAxis;
          5: axis := AChart.X2Axis;
          6: axis := AChart.Y2Axis;
        end;
        font := axis.Title.Font;
        rotAngle := axis.Title.RotationAngle;
        if not axis.DefaultTitleRotation then
        begin
          if AChart.RotatedAxes then
          begin
            if rotAngle = 0 then rotAngle := 90 else if rotAngle = 90 then rotAngle := 0;
          end;
          rotAngleStr := Format('%.1f', [rotangle], FPointSeparatorSettings);
        end;
      end;
    else
      raise Exception.Create('[GetChartCaptionStyleAsXML] Unknown caption.');
  end;

  chartProps := 'chart:auto-position="true" ';
  if rotAngleStr <> '' then
    chartProps := chartProps + Format('style:rotation-angle="%s" ', [rotAngleStr]);

  textProps := TsSpreadOpenDocWriter(Writer).WriteFontStyleXMLAsString(font);

  indent := DupeString(' ', AIndent);
  Result := Format(
    indent + '<style:style style:name="ch%d" style:family="chart">' + LE +
    indent + '  <style:chart-properties %s/>' + LE +
    indent + '  <style:text-properties %s/>' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, chartProps, textProps ]
  );
end;

function TsSpreadOpenDocChartWriter.GetChartErrorBarStyleAsXML(AChart: TsChart;
  AErrorBar: TsChartErrorBars; AIndent, AStyleID: Integer): String;
var
  graphProps: String;
  chartProps: String = '';
  indent: String;

  function GetCellRangeStr(ARange: TsChartRange): String;
  var
    sheet1, sheet2: String;
    r1, c1, r2, c2: Cardinal;
  begin
    sheet1 := ARange.GetSheet1Name;
    sheet2 := ARange.GetSheet2Name;
    r1 := ARange.Row1;
    c1 := ARange.Col1;
    r2 := ARange.Row2;
    c2 := ARange.Col2;
    Result := GetSheetCellRangeString_ODS(sheet1, sheet2, r1, c1, r2, c2, rfAllRel, false);
  end;

begin
  case AErrorBar.Kind of
    cebkConstant:
      begin
        chartProps := chartProps + 'chart:error-category="constant" ';
        if AErrorBar.ShowPos then
          chartProps := chartProps + Format('chart:error-upper-limit="%.9g" ', [ AErrorBar.ValuePos ], FPointSeparatorSettings);
        if AErrorBar.ShowNeg then
          chartProps := chartProps + Format('chart:error-lower-limit="%.9g" ', [ AErrorBar.ValueNeg ], FPointSeparatorSettings);
      end;
    cebkPercentage:
      begin
        chartProps := chartProps + 'chart:error-category="percentage" ';
        chartProps := chartProps + Format('chart:error-percentage="%.9g" ', [ AErrorBar.ValuePos ], FPointSeparatorSettings);
        chartProps := chartProps + 'loext:std-weight="1" ';
      end;
    cebkCellRange:
      begin
        chartProps := chartProps + 'chart:error-category="cell-range" ';
        if AErrorBar.ShowPos then
          chartProps := chartProps + 'chart:error-upper-range="' + GetCellRangeStr(AErrorBar.RangePos) + '" ';
        if AErrorBar.ShowNeg then
          chartProps := chartProps + 'chart:error-lower-range="' + GetCellRangeStr(AErrorBar.RangeNeg) + '" ';
        chartProps := chartProps + 'loext:std-weight="1" ';
      end;
  end;
  if AErrorBar.ShowPos then
    chartProps := chartProps + 'chart:error-upper-indicator="true" ';
  if AErrorBar.ShowNeg then
    chartProps := chartProps + 'chart:error-lower-indicator="true" ';

  graphProps := GetChartLineStyleGraphicPropsAsXML(AChart, AErrorBar.Line);

  indent := DupeString(' ', AIndent);
  Result := Format(
    indent + '<style:style style:name="ch%d" style:family="chart">' + LE +
    indent + '  <style:chart-properties %s/>' + LE +
    indent + '  <style:graphic-properties %s/>' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, chartProps, graphProps ]
  );
end;

function TsSpreadOpenDocChartWriter.GetChartFillStyleGraphicPropsAsXML(AChart: TsChart;
  AFill: TsChartFill): String;
var
  gradient: TsChartGradient;
  hatch: TsChartHatch;
  fillStr: String = '';
  opacityStr: String = '';
begin
  case AFill.Style of
    cfsNoFill:
      Result := 'draw:fill="none" ';
    cfsSolid:
      begin
        if (AFill.Color.Transparency > 0) then
          opacityStr := Format('draw:opacity="%d%%" ', [round(100*(1.0 - AFill.Color.Transparency))]);
        Result := Format(
          'draw:fill="solid" draw:fill-color="%s" %s',
          [ ColorToHTMLColorStr(AFill.Color.Color), opacityStr ]
        );
      end;
    cfsGradient:
      begin
        gradient := AChart.Gradients[AFill.Gradient];
        if (gradient.StartColor.Transparency > 0) then
          opacityStr := Format('draw:opacity="%d%%" ', [round(100*(1.0 - gradient.StartColor.Transparency))]);
        // to do: evaluate opacity of all gradient steps
        Result := Format(
          'draw:fill="gradient" ' +
          'draw:fill-gradient-name="%s" ' +
          'draw:gradient-step-count="0" %s',
          [ ASCIIName(gradient.Name), opacityStr ]
        );
      end;
    cfsHatched, cfsSolidHatched:
      begin
        hatch := AChart.Hatches[AFill.Hatch];
        if (hatch.PatternColor.Transparency > 0) then
          opacityStr := Format('draw:opacity="%d%%" ', [round(100*(1.0 - hatch.PatternColor.Transparency))]);
        if AFill.Style = cfsSolidHatched then
          fillStr := 'draw:fill-hatch-solid="true" ';
        Result := Format(
          'draw:fill="hatch" draw:fill-color="%s" %s' +
          'draw:fill-hatch-name="%s" %s',
          [ ColorToHTMLColorStr(AFill.Color.Color), opacityStr,
            ASCIIName(hatch.Name), fillStr
          ]
        );
      end;
  end;
end;

{
<style:style style:name="ch4" style:family="chart">
  <style:chart-properties chart:auto-position="true"/>
  <style:graphic-properties svg:stroke-color="#b3b3b3" draw:fill="none"
     draw:fill-color="#e6e6e6"/>
  <style:text-properties fo:font-family="Consolas"
     style:font-style-name="Standard" style:font-family-generic="modern"
     style:font-pitch="fixed" fo:font-size="12pt"
     style:font-size-asian="10pt" style:font-size-complex="10pt"/>
</style:style>
}
function TsSpreadOpenDocChartWriter.GetChartLegendStyleAsXML(AChart: TsChart;
  AIndent, AStyleID: Integer): String;
var
  indent: String;
  textProps: String = '';
  graphProps: String = '';
begin
  Result := '';

  if not AChart.Legend.Visible then
    exit;

  graphProps := GetChartLineStyleGraphicPropsAsXML(AChart, AChart.Legend.Border) +
                GetChartFillStyleGraphicPropsAsXML(AChart, AChart.Legend.Background);

  textProps := TsSpreadOpenDocWriter(Writer).WriteFontStyleXMLAsString(AChart.Legend.Font);

  indent := DupeString(' ', AIndent);
  Result := Format(
    indent + '<style:style style:name="ch%d" style:family="chart">' + LE +
    indent + '  <style:chart-properties />' + LE +
    indent + '  <style:graphic-properties %s/>' + LE +
    indent + '  <style:text-properties %s/>' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, graphProps, textProps ]
  );
end;

{ <style:style style:name="ch12" style:family="chart">
    <style:graphic-properties draw:stroke="dash" draw:stroke-dash="Dot"
       svg:stroke-color="#ff0000"/>
  </style:style> }
function TsSpreadOpenDocChartWriter.GetChartLineStyleAsXML(AChart: TsChart;
  ALine: TsChartLine; AIndent, AStyleID: Integer): String;
var
  ind: String;
  graphProps: String = '';
begin
  ind := DupeString(' ', AIndent);
  graphProps := GetChartLineStyleGraphicPropsAsXML(AChart, ALine);
  Result := Format(
    ind + '<style:style style:name="ch%d" style:family="chart">' + LE +
    ind + '  <style:graphic-properties %s/>' + LE +
    ind + '</style:style>' + LE,
    [ AStyleID, graphProps ]
  );
end;

{ Constructs the xml for a line style to be used in the <style:graphic-properties> }
function TsSpreadOpenDocChartWriter.GetChartLineStyleGraphicPropsAsXML(
  AChart: TsChart; ALine: TsChartLine; ForceNoLine: Boolean = false): String;
var
  strokeStr: String = '';
  widthStr: String = '';
  colorStr: String = '';
  opacityStr: String = '';
  linestyle: TsChartLineStyle;
begin
  if (ALine.Style = clsNoLine) or ForceNoLine then
  begin
    Result := 'draw:stroke="none" ';
    exit;
  end;

  strokeStr := 'draw:stroke="solid" ';
  if (ALine.Style <> clsSolid) then
  begin
    linestyle := AChart.GetLineStyle(ALine.Style);
    if linestyle <> nil then
      strokeStr := 'draw:stroke="dash" draw:stroke-dash="' + ASCIIName(linestyle.Name) + '" ';
  end;

  if ALine.Width > 0 then
    widthStr := Format('svg:stroke-width="%.1fmm" ', [ALine.Width], FPointSeparatorSettings);
  colorStr := Format('svg:stroke-color="%s" ', [ColorToHTMLColorStr(ALine.Color.Color)]);

  if ALine.Color.Transparency > 0 then
    opacityStr := Format('svg:stroke-opacity="%d%%" ', [round((1.0 - ALine.Color.Transparency)*100)], FPointSeparatorSettings);

  Result := strokeStr + widthStr + colorStr + opacityStr;
end;

function TsSpreadOpenDocChartWriter.GetChartPlotAreaStyleAsXML(AChart: TsChart;
  AIndent, AStyleID: Integer): String;
var
  indent: String;
  interpolation: TsChartInterpolation;
  interpolationStr: String = '';
  verticalStr: String = '';
  stackModeStr: String = '';
  rightAngledAxes: String = '';
  startAngleStr: String = '';
  candleStickStr: String = '';
  i: Integer;
begin
  indent := DupeString(' ', AIndent);

  if AChart.RotatedAxes then
    verticalStr := 'chart:vertical="true" ';

  case AChart.StackMode of
    csmDefault: ;
    csmStacked: stackModeStr := 'chart:stacked="true" ';
    csmStackedPercentage: stackModeStr := 'chart:percentage="true" ';
  end;

  if (AChart.Series.Count > 0) and (AChart.Series[0] is TsPieSeries) then
    startAngleStr := Format('chart:angle-offset="%d" ', [TsPieSeries(AChart.Series[0]).StartAngle]);

  // In FPSpreadsheet individual series can be "smooth", in Calc only all.
  // As a compromise, when we find at least one smooth series, all series are
  // treated as such by writing the "chart:interpolation" attribute
  for i := 0 to AChart.Series.Count-1 do
    if AChart.Series[i] is TsCustomLineSeries then
    begin
      interpolation := TsOpenedCustomLineSeries(AChart.Series[i]).Interpolation;
      case interpolation of
        ciLinear: Continue;
        ciCubicSpline: interpolationStr := 'chart:interpolation="cubic-spline" ';
        ciBSpline: interpolationStr := 'chart:interpolation="b-spline" ';
          // NOTE: LibreOffice v24.2.5.2 does not display the interpolated b-spline line any more...
        ciStepStart: interpolationStr := 'chart:interpolation="step-start" ';
        ciStepEnd: interpolationStr := 'chart:interpolation="step-end" ';
        ciStepCenterX: interpolationStr := 'chart:interpolation="step-center-x" ';
        ciStepCenterY: interpolationStr := 'chart:interpolation="step-center-y" ';
      end;
      break;
    end;

  if not (AChart.GetChartType in [ctRadar, ctFilledRadar, ctPie]) then
    rightAngledAxes := 'chart:right-angled-axes="true" ';

  for i := 0 to AChart.Series.Count-1 do
    if (AChart.Series[i] is TsStockSeries) and TsStockSeries(AChart.Series[i]).CandleStick then
    begin
      candleStickStr := 'chart:japanese-candle-stick="true" ';
      break;
    end;

  Result := Format(
    indent + '  <style:style style:name="ch%d" style:family="chart">', [ AStyleID ]) + LE +
    indent + '    <style:chart-properties ' +
                   interpolationStr +
                   verticalStr +
                   stackModeStr +
                   startAngleStr +
                   candleStickStr +
                   'chart:symbol-type="automatic" ' +
                   'chart:include-hidden-cells="false" ' +
                   'chart:auto-position="true" ' +
                   'chart:auto-size="true" ' +
                   'chart:treat-empty-cells="leave-gap" ' +
                   rightAngledAxes +
                  '/>' + LE +
    indent + '  </style:style>' + LE;
end;

function TsSpreadOpenDocChartWriter.GetChartRegressionEquationStyleAsXML(
    AChart: TsChart; AEquation: TsTrendlineEquation; AIndent, AStyleID: Integer): String;
var
  indent: String;
  numStyle: String = 'N0';
  chartprops: String = '';
  lineprops: String = '';
  fillprops: String = '';
  textprops: String = '';
begin
  Result := '';

  indent := DupeString(' ', AIndent);

  numStyle := GetNumberFormatID(AEquation.NumberFormat);

  if not AEquation.DefaultXName then
    chartprops := chartprops + Format('loext:regression-x-name="%s" ', [AEquation.XName]);
  if not AEquation.DefaultYName then
    chartprops := chartprops + Format('loext:regression-y-name="%s" ', [AEquation.YName])          ;

  if not AEquation.DefaultBorder then
    lineProps := GetChartLineStyleGraphicPropsAsXML(AChart, AEquation.Border);

  if not AEquation.DefaultFill then
    fillProps := GetChartFillStyleGraphicPropsAsXML(AChart, AEquation.Fill);

  if not AEquation.DefaultFont then
    textprops := TsSpreadOpenDocWriter(Writer).WriteFontStyleXMLAsString(AEquation.Font);

  Result := Format(
    indent + '<style:style style:name="ch%d" style:family="chart" style:data-style-name="%s">' + LE +
    indent + '  <style:chart-properties %s/>' + LE +
    indent + '  <style:graphic-properties %s/>' + LE +
    indent + '  <style:text-properties %s/>' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, numStyle, chartprops, fillprops + lineprops, textprops ]
  );
end;

function TsSpreadOpenDocChartWriter.GetChartRegressionStyleAsXML(AChart: TsChart;
  ASeriesIndex, AIndent, AStyleID: Integer): String;
var
  series: TsChartSeries;
  trendline: TsChartTrendline;
  indent: String;
  chartProps: String = '';
  graphProps: String = '';
begin
  Result := '';
  indent := DupeString(' ', AIndent);

  series := AChart.Series[ASeriesIndex];
  if not series.SupportsTrendline then
    exit;

  trendline := TsOpenedTrendlineSeries(series).Trendline;

  if trendline.TrendlineType = tltNone then
    exit;
  series := AChart.Series[ASeriesIndex] as TsScatterSeries;

  chartprops := Format(
    'chart:regression-name="%s" ' +
    'chart:regression-type="%s" ' +
    'chart:regression-extrapolate-forward="%g" ' +
    'chart:regression-extrapolate-backward="%g" ' +
    'chart:regression-force-intercept="%s" ' +
    'chart:regression-intercept-value="%g" ' +
    'chart:regression-max-degree="%d" ',
    [ trendline.Title,
      TRENDLINE_TYPE[trendline.TrendlineType] ,
      trendline.ExtrapolateForwardBy,
      trendline.ExtrapolateBackwardBy,
      FALSE_TRUE[trendline.ForceYIntercept],
      trendline.YInterceptValue,
      trendline.PolynomialDegree
    ], FPointSeparatorSettings
  );

  graphprops := GetChartLineStyleGraphicPropsAsXML(AChart, trendline.Line);

  Result := Format(
    indent + '<style:style style:name="ch%d" style:family="chart"> ' + LE +
    indent + '  <style:chart-properties %s/>' + LE +
    indent + '  <style:graphic-properties %s/>' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, chartprops, graphprops ]
  );
end;

{@@ ----------------------------------------------------------------------------
  Creates an xml string which contains the individual datapoint style with index
  ADataPointStyleIndex for the series with index ASeriesIndex.
-------------------------------------------------------------------------------}
function TsSpreadOpenDocChartWriter.GetChartSeriesDataPointStyleAsXML(AChart: TsChart;
  ASeriesIndex, ADataPointStyleIndex, AIndent, AStyleID: Integer): String;
var
  series: TsChartSeries;
  indent: String;
  chartProps: String;
  graphProps: String = '';
  dataPointStyle: TsChartDataPointStyle;
begin
  Result := '';
  indent := DupeString(' ', AIndent);

  series := AChart.Series[ASeriesIndex];

  if ADataPointStyleIndex > -1 then
    dataPointStyle := series.DataPointStyles[ADataPointStyleIndex]
  else
    dataPointStyle := nil;

  if dataPointStyle = nil then
  begin
    // No style information found. We write a node, nevertheless...  (maybe can be dropped?)
    Result := Format(
      indent + '<style:style style:name="ch%d" style:family="chart">' + LE +
      indent + '  <style:chart-properties/>' + LE +
      indent + '  <style:graphic-properties/>' + LE +
      indent + '</style:style>' + LE,
      [ AStyleID ]
    );
    exit;
  end;

  chartProps := 'chart:solid-type="cuboid" ';
  if datapointstyle.PieOffset > 0 then
    chartProps := chartProps + Format('chart:pie-offset="%d" ', [datapointStyle.PieOffset]);

  if dataPointStyle.Background <> nil then
    graphProps := graphProps + GetChartFillStyleGraphicPropsAsXML(AChart, dataPointStyle.Background);
  if dataPointStyle.Border <> nil then
    graphProps := graphProps + GetChartLineStyleGraphicPropsAsXML(AChart, dataPointStyle.Border);

  Result := Format(
    indent + '<style:style style:name="ch%d" style:family="chart">' + LE +
    indent + '  <style:chart-properties %s/>' + LE +
    indent + '  <style:graphic-properties %s/>' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, chartProps, graphProps ]
  );
end;

function TsSpreadOpenDocChartWriter.GetChartSeriesStyleAsXML(AChart: TsChart;
  ASeriesIndex, AIndent, AStyleID: Integer): String;
var
  series: TsChartSeries;
  lineser: TsOpenedCustomLineSeries = nil;
  indent: String;
  numStyle: String;
  forceNoLine: Boolean = false;
  chartProps: String = '';
  graphProps: String = '';
  textProps: String = '';
  lineProps: String = '';
  fillProps: String = '';
  labelSeparator: String = '';
begin
  Result := '';

  indent := DupeString(' ', AIndent);
  series := AChart.Series[ASeriesIndex];

  // Number format
  numStyle := GetNumberFormatID(series.LabelFormat);

  // Chart properties
  chartProps := 'chart:symbol-type="none" ';

  if ((series is TsLineSeries) and (series.ChartType <> ctFilledRadar)) or
     (series is TsScatterSeries) then
  begin
    lineser := TsOpenedCustomLineSeries(series);
    if lineser.ShowSymbols then
      chartProps := Format(
        'chart:symbol-type="named-symbol" chart:symbol-name="%s" chart:symbol-width="%.1fmm" chart:symbol-height="%.1fmm" ',
        [SYMBOL_NAMES[lineSer.Symbol], lineSer.SymbolWidth, lineSer.SymbolHeight ],
        FPointSeparatorSettings
      );
    forceNoLine := not lineSer.ShowLines;
  end;

  chartProps := chartProps + Format('chart:link-data-style-to-source="%s" ', [FALSE_TRUE[numStyle = 'N0']]);

  if ([cdlValue, cdlPercentage] * series.DataLabels = [cdlValue]) then
    chartProps := chartProps + 'chart:data-label-number="value" '
  else
  if ([cdlValue, cdlPercentage] * series.DataLabels = [cdlPercentage]) then
    chartProps := chartProps + 'chart:data-label-number="percentage" '
  else
  if ([cdlValue, cdlPercentage] * series.DataLabels = [cdlValue, cdlPercentage]) then
    chartProps := chartProps + 'chart:data-label-number="value-and-percentage" ';
  if (cdlCategory in series.DataLabels) then
    chartProps := chartProps + 'chart:data-label-text="true" ';
  if (cdlSeriesName in series.DataLabels) then
    chartProps := chartProps + 'chart:data-label-series="true" ';
  if (cdlSymbol in series.DataLabels) then
    chartProps := chartProps + 'chart:data-label-symbol="true" ';
  if series.LabelPosition <> lpDefault then
    chartProps := chartProps + 'chart:label-position="' + LABEL_POSITION[series.LabelPosition] + '" ';

  if series.LabelSeparator = ' ' then
    labelSeparator := ''
  else
  begin
    labelSeparator := series.LabelSeparator;
    if (pos('\n', labelSeparator) > 0) then
      labelSeparator := StringReplace(labelSeparator, '\n', '<text:line-break/>', [rfReplaceAll, rfIgnoreCase])
    else if (pos(#13#10, labelSeparator) > 0) then
      labelSeparator := StringReplace(labelSeparator, #13#10, '<text:line-break/>', [rfReplaceAll, rfIgnoreCase])
    else if (pos(#10, labelSeparator) > 0) then
      labelSeparator := StringReplace(labelSeparator, #10, '<text:line-break/>', [rfReplaceAll, rfIgnoreCase])
    else if (pos(#13, labelSeparator) > 0) then
      labelSeparator := StringReplace(labelSeparator, #13, '<text:line-break/>', [rfReplaceAll, rfIgnoreCase]);
    labelSeparator :=
      indent + '    <chart:label-separator>' + LE +
      indent + '      <text:p>' + labelSeparator + '</text:p>' + LE +
      indent + '    </chart:label-separator>' + LE;
  end;

  if series.LabelBorder.Style <> clsNoLine then
  begin
    chartProps := chartProps + 'loext:label-stroke="solid" ';
    chartProps := chartProps + 'loext:label-stroke-color="' + ColorToHTMLColorStr(series.LabelBorder.Color.Color) + '"';
    if series.LabelBorder.Color.Transparency > 0 then
      chartProps := chartProps + 'loext:label-stroke-opacity="' + IntToStr(round(100*(1.0 - series.LabelBorder.Color.Transparency))) + '"';
  end;

  if labelSeparator <> '' then
    chartProps := indent + '  <style:chart-properties ' + chartProps + '>' + LE + labelSeparator + indent + '  </style:chart-properties>'
  else
    chartProps := indent + '  <style:chart-properties ' + chartProps + '/>';

  // Graphic properties
  lineProps := GetChartLineStyleGraphicPropsAsXML(AChart, series.Line, forceNoLine);
  if (series is TsLineSeries) and (series.ChartType <> ctFilledRadar) then
  begin
    // NOTE: In LibreOffice lines and symbols have the same color. When different
    // colors are written here, the line color dominates.
    lineSer := TsOpenedCustomLineSeries(series);
    fillProps := GetChartFillStyleGraphicPropsAsXML(AChart, lineser.SymbolFill);
    if lineSer.ShowSymbols then
      graphProps := graphProps + fillProps;
    if lineSer.ShowLines and (lineser.Line.Style <> clsNoLine) then
      graphProps := graphProps + lineProps
    else
      graphProps := graphProps + 'draw:stroke="none" ';
  end else
  begin
    fillProps := GetChartFillStyleGraphicPropsAsXML(AChart, series.Fill);
    graphProps := fillProps + lineProps;
  end;

  // Text properties
  textProps := TsSpreadOpenDocWriter(Writer).WriteFontStyleXMLAsString(series.LabelFont);

  Result := Format(
    indent + '<style:style style:name="ch%d" style:family="chart" style:data-style-name="%s">' + LE +
    chartProps + LE +
    indent + '  <style:graphic-properties %s/>' + LE +
    indent + '  <style:text-properties %s/>' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, numstyle, graphProps, textProps ]
  );
end;

function TsSpreadOpenDocChartWriter.GetChartStockSeriesStyleAsXML(AChart: TsChart;
  ASeries: TsStockSeries; AKind: Integer; AIndent, AStyleID: Integer): String;
var
  indent: String;
  fillStr: String = '';
  lineStr: String = '';
begin
  case AKind of
    0: // gain marker
      begin
        fillStr := GetChartFillStyleGraphicPropsAsXML(AChart, ASeries.CandleStickUpFill);
        lineStr := GetChartLineStyleGraphicPropsAsXML(AChart, ASeries.CandleStickUpBorder);
      end;
    1: // loss marker
      begin
        fillStr := GetChartFillStyleGraphicPropsAsXML(AChart, ASeries.CandleStickDownFill);
        lineStr := GetChartLineStyleGraphicPropsAsXML(AChart, ASeries.CandleStickDownBorder);
      end;
    2:  // range line
      lineStr := GetChartLineStyleGraphicPropsAsXML(AChart, ASeries.RangeLine);
  end;

  if (fillStr <> '') or (lineStr <> '') then
  begin
    indent := DupeString(' ', AIndent);
    Result := Format(
      indent + '<style:style style:name="ch%d" style:family="chart">' + LE +
      indent + '  <style:graphic-properties ' + fillstr + lineStr + '/>' + LE +
      indent + '</style:style>' + LE,
      [ AStyleID ]
    );
  end else
    Result := '';
end;

function TsSpreadOpenDocChartWriter.GetNumberFormatID(ANumFormat: String): String;
var
  idx: Integer;
begin
  idx := TsChartNumberFormatList(FNumberFormatList).IndexOfFormat(ANumFormat);
  if idx > -1 then
    Result := Format('N%d', [idx])
  else
    Result := 'N0';
end;

procedure TsSpreadOpenDocChartWriter.ListAllNumberFormats(AChart: TsChart);
var
  i: Integer;
  series: TsChartSeries;
  trendline: TsChartTrendline;
begin
  FNumberFormatList.Clear;
  FNumberFormatList.Add('');

  // Formats of axis labels
  FNumberFormatList.Add(AChart.XAxis.LabelFormat);
  FNumberFormatList.Add(AChart.YAxis.LabelFormat);
  FNumberFormatList.Add(AChart.X2Axis.LabelFormat);
  FNumberFormatList.Add(AChart.Y2Axis.LabelFormat);
  if AChart.StackMode = csmStackedPercentage then
  begin
    FNumberFormatList.Add(AChart.YAxis.LabelFormatPercent);
    FNumberFormatList.Add(AChart.Y2Axis.LabelFormatPercent);
  end;

  // Formats of series labels
  for i := 0 to AChart.Series.Count-1 do
  begin
    series := AChart.Series[i];
    FNumberFormatList.Add(series.LabelFormat);
    // Format of fit equation
    if series.SupportsTrendline then
    begin
      trendline := TsOpenedTrendlineSeries(series).Trendline;
      if (trendline.TrendlineType <> tltNone) and
         (trendline.DisplayEquation or trendline.DisplayRSquare) then
      begin
        FNumberFormatList.Add(trendline.Equation.NumberFormat);
      end;
    end;
  end;
end;

{ Switches secondary axes to visible when there are series needing them. }
procedure TsSpreadOpenDocChartWriter.CheckAxis(AChart: TsChart; Axis: TsChartAxis);
var
  i: Integer;
begin
  if Axis = AChart.Y2Axis then
    for i := 0 to AChart.Series.Count - 1 do
      if AChart.Series[i].YAxis = calSecondary then
      begin
        Axis.Visible := true;
        break;
      end;
end;


(* DO NOT DELETE THIS! MAYBE NEEDED LATER...

{ Extracts the cells needed by the given chart from the chart's worksheet and
  copies their values into a temporary worksheet, AWorksheet, so that these
  data can be written to the xml immediately.
  Independently of the layout in the original worksheet, data are arranged in
  columns of AWorksheet, starting at cell A1.
  - First column: Categories (or index in case of scatter chart)
  - Second column:
      in case of category charts: y values of the first series,
      in case of scatter series: x values of the first series
  - Third column:
      in case of category charts: y values of the second series
      in case of scatter series. y values of the first series
  - etc.
  The first row contains
  - nothing in case of the first column
  - cell range reference in ODS syntax for the cells in the original worksheet.
  The aux worksheet should be contained in a separate workbook to avoid
  interfering with the writing process.
  }
procedure TsSpreadOpenDocChartWriter.PrepareChartTable(AChart: TsChart;
  AWorksheet: TsBasicWorksheet);
var
  isScatterChart: Boolean;
  series: TsChartSeries;
  seriesSheet: TsWorksheet;
  auxSheet: TsWorksheet;
  refStr, txt: String;
  i, j: Integer;
  srcCell, destCell: PCell;
  destCol: Cardinal;
  r1, c1, r2, c2: Cardinal;
  nRows: Integer;
begin
  if AChart.Series.Count = 0 then
    exit;

  auxSheet := TsWorksheet(AWorksheet);
  seriesSheet := TsWorkbook(Writer.Workbook).GetWorksheetByIndex(AChart.SheetIndex);
  isScatterChart := AChart.IsScatterChart;

  // Determine the number of rows in auxiliary output worksheet.
  nRows := 0;
  for i := 0 to AChart.Series.Count-1 do
  begin
    series := AChart.Series[i];
    j := series.GetXCount;
    if j > nRows then nRows := j;
    j := series.GetYCount;
    if j > nRows then nRows := j;
  end;

  // Write label column. If missing, write consecutive numbers 1, 2, 3, ...
  destCol := 0;
  for i := 0 to AChart.Series.Count-1 do
  begin
    series := AChart.Series[i];

    // Write the label column. Use consecutive numbers 1, 2, 3, ... if there
    // are no labels.
    if series.HasLabels then
    begin
      r1 := series.LabelRange.Row1;
      c1 := series.LabelRange.Col1;
      r2 := series.LabelRange.Row2;
      c2 := series.LabelRange.Col2;
      refStr := GetSheetCellRangeString_ODS(seriesSheet.Name, seriesSheet.Name, r1, c1, r2, c2, rfAllRel, false);
    end else
      refStr := '';

    auxSheet.WriteText(0, destCol, '');

    for j := 1 to nRows do
    begin
      if series.HasLabels then
      begin
        if series.LabelsInCol then
          srcCell := seriesSheet.FindCell(r1 + j - 1, c1)
        else
          srcCell := seriesSheet.FindCell(r1, c1 + j - 1);
      end else
        srcCell := nil;
      if srcCell <> nil then
      begin
        destCell := auxsheet.GetCell(j, destCol);
        seriesSheet.CopyValue(srcCell, destCell);
      end else
        destCell := auxSheet.WriteNumber(j, destCol, j);
    end;
    if (refStr <> '') then
      auxsheet.WriteComment(1, destCol, refStr);

    // In case of scatter plot write the x column. Use consecutive numbers 1, 2, 3, ...
    // if there are no x values.
    if isScatterChart then
    begin
      inc(destCol);

      if series.HasXValues then
      begin
        r1 := series.XRange.Row1;
        c1 := series.XRange.Col1;
        r2 := series.XRange.Row2;
        c2 := series.XRange.Col2;
        refStr := GetSheetCellRangeString_ODS(seriesSheet.Name, seriesSheet.Name, r1, c1, r2, c2, rfAllRel, false);
        if series.XValuesInCol then
          txt := 'Col ' + GetColString(c1)
        else
          txt := 'Row ' + GetRowString(r1);
      end else
      begin
        refStr := '';
        txt := '';
      end;

      auxSheet.WriteText(0, destCol, txt);

      for j := 1 to nRows do
      begin
        if series.HasXValues then
        begin
          if series.XValuesInCol then
            srcCell := seriesSheet.FindCell(r1 + j - 1, c1)
          else
            srcCell := seriesSheet.FindCell(r1, c1 + j - 1);
        end else
          srcCell := nil;
        if srcCell <> nil then
        begin
          destCell := auxsheet.GetCell(j, destCol);
          seriesSheet.CopyValue(srcCell, destCell);
        end else
          destCell := auxSheet.WriteNumber(j, destCol, j);
      end;
      if (refStr <> '') then
        auxsheet.WriteComment(1, destCol, refStr);
    end;

    // Write the y column
    if not series.HasYValues then
      Continue;

    inc(destCol);

    r1 := series.TitleAddr.Row;                   // Series title
    c1 := series.TitleAddr.Col;
    txt := seriesSheet.ReadAsText(r1, c1);
    auxsheet.WriteText(0, destCol, txt);
    refStr := GetSheetCellRangeString_ODS(seriesSheet.Name, seriesSheet.Name, r1, c1, r1, c1, rfAllRel, false);
    if (refStr <> '') then
      auxSheet.WriteComment(0, destCol, refStr);   // Store title reference as comment for svg node

    r1 := series.YRange.Row1;
    c1 := series.YRange.Col1;
    r2 := series.YRange.Row2;
    c2 := series.YRange.Col2;
    refStr := GetSheetCellRangeString_ODS(seriesSheet.Name, seriesSheet.Name, r1, c1, r2, c2, rfAllRel, false);
    for j := 1 to series.GetYCount do
    begin
      if series.YValuesInCol then
        srcCell := seriesSheet.FindCell(r1 + j - 1, c1)
      else
        srcCell := seriesSheet.FindCell(r1, c1 + j - 1);
      if srcCell <> nil then
      begin
        destCell := auxSheet.GetCell(j, destCol);
        seriesSheet.CopyValue(srcCell, destCell);
      end else
        destCell := auxSheet.WriteNumber(j, destCol, j);
    end;

    if (refStr <> '') then
      auxSheet.WriteComment(1, destCol, refStr);   // Store y range reference as comment for svg node
  end;
end;
*)

procedure TsSpreadOpenDocChartWriter.ResetStreams;
var
  i: Integer;
begin
  for i := 0 to High(FSCharts) do
  begin
    FSCharts[i].Position := 0;
    FSObjectStyles[i].Position := 0;
  end;
end;

{ Writes the chart to the specified stream.
  All chart elements are formatted by means of styles. To simplify assignment
  of styles to elements we first write the elements and create the style of
  the currently written chart element on the fly. Since styles must be written
  to the steam first, we write the chart elements to a separate stream which
  is appended to the main stream afterwards. }
procedure TsSpreadOpenDocChartWriter.WriteChart(AStream: TStream; AChart: TsChart);
var
  chartStream: TMemoryStream;
  styleStream: TMemoryStream;
  styleID: Integer;
begin
  chartStream := TMemoryStream.Create;
  styleStream := TMemoryStream.Create;
  try
    ListAllNumberFormats(AChart);
    WriteChartNumberStyles(styleStream, 4, AChart);

    styleID := 1;
    WriteChartBackground(chartStream, styleStream, 6, 4, AChart, styleID);
    WriteChartTitle(chartStream, styleStream, 6, 4, AChart, false, styleID);  // Title
    WriteChartTitle(chartStream, styleStream, 6, 4, AChart, true, styleID);   // Subtitle
    WriteChartLegend(chartStream, styleStream, 6, 4, AChart, styleID);        // Legend
    WriteChartPlotArea(chartStream, styleStream, 6, 4, AChart, styleID);      // Wall, axes, series

    // Here begins the main stream
    AppendToStream(AStream,
      XML_HEADER + LE);

    AppendToStream(AStream,
      '<office:document-content ' + LE +
      '    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' + LE +
      '    xmlns:ooo="http://openoffice.org/2004/office"' + LE +
      '    xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' + LE +
      '    xmlns:xlink="http://www.w3.org/1999/xlink"' + LE +
      '    xmlns:dc="http://purl.org/dc/elements/1.1/"' + LE +
      '    xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"' + LE +
      '    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' + LE +
      '    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' + LE +
      '    xmlns:rpt="http://openoffice.org/2005/report"' + LE +
      '    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' + LE +
      '    xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"' + LE +
      '    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' + LE +
      '    xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"' + LE +
      '    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' + LE +
      '    xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"' + LE +
      '    xmlns:ooow="http://openoffice.org/2004/writer"' + LE +
      '    xmlns:oooc="http://openoffice.org/2004/calc"' + LE +
      '    xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"' + LE +
      '    xmlns:xforms="http://www.w3.org/2002/xforms"' + LE +
      '    xmlns:tableooo="http://openoffice.org/2009/table"' + LE +
      '    xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"' + LE +
      '    xmlns:drawooo="http://openoffice.org/2010/draw"' + LE +
      '    xmlns:xhtml="http://www.w3.org/1999/xhtml"' + LE +
      '    xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"' + LE +
      '    xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"' + LE +
      '    xmlns:math="http://www.w3.org/1998/Math/MathML"' + LE +
      '    xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0"' + LE +
      '    xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"' + LE +
      '    xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0"' + LE +
      '    xmlns:dom="http://www.w3.org/2001/xml-events"' + LE +
      '    xmlns:xsd="http://www.w3.org/2001/XMLSchema"' + LE +
      '    xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' + LE +
      '    xmlns:grddl="http://www.w3.org/2003/g/data-view#"' + LE +
      '    xmlns:css3t="http://www.w3.org/TR/css3-text/"' + LE +
      '    xmlns:chartooo="http://openoffice.org/2010/chart" office:version="1.3">' + LE
    );

    // The file begins with the chart styles
    AppendToStream(AStream,
      '  <office:automatic-styles>' + LE
    );

    // Copy the styles from the temporary style stream.
    styleStream.Position := 0;
    AStream.CopyFrom(styleStream, stylestream.Size);

    // Now the chart part follows, after closing the styles part
    AppendToStream(AStream,
      '  </office:automatic-styles>' + LE +
      '  <office:body>' + LE +
      '    <office:chart>' + LE
    );

    // Copy the chart elements from the temporary chart stream
    chartStream.Position := 0;
    AStream.CopyFrom(chartStream, chartStream.Size);

    // After the chart elements we have the data to be plotted
//    WriteChartTable(AStream, AChart, 8);
    // wp: writing this makes no change except the series fills not being applied

    // Finally the footer.
    AppendToStream(AStream,
      '      </chart:chart>' + LE +
      '    </office:chart>' + LE +
      '  </office:body>' + LE +
      '</office:document-content>');
  finally
    chartStream.Free;
    styleStream.Free;
  end;
end;

procedure TsSpreadOpenDocChartWriter.WriteChartAxis(
  AChartStream, AStyleStream: TStream; AChartIndent, AStyleIndent: Integer;
  Axis: TsChartAxis; var AStyleID: Integer);
var
  indent: String;
  captionKind: Integer;
  chart: TsChart;
  series: TsChartSeries;
  sheet: TsWorksheet;
  refStr: String;
  r1, c1, r2, c2: Cardinal;
begin
  if not Axis.Visible then
    exit;

  chart := Axis.Chart;

  if Axis = chart.XAxis then
    captionKind := 3
  else if Axis = chart.YAxis then
    captionKind := 4
  else if Axis = chart.X2Axis then
    captionKind := 5
  else if Axis = chart.Y2Axis then
    captionKind := 6
  else
    raise Exception.Create('[WriteChartAxis] Unknown axis');

  // Write axis
  indent := DupeString(' ', AChartIndent);
  AppendToStream(AChartStream, Format(
    indent + '<chart:axis chart:style-name="ch%d" chart:dimension="%s" chart:name="%s-%s" chartooo:axis-type="auto">' + LE,
    [ AStyleID, AXIS_ID[captionKind], AXIS_LEVEL[captionKind], AXIS_ID[captionKind] ]
  ));

  if Axis.DateTime then
    AppendToStream(AChartStream,
      indent + '  <chartooo:date-scale/>' + LE
    );

  if (Axis = chart.XAxis) and (not chart.IsScatterChart) and (chart.Series.Count > 0) then
  begin
    series := chart.Series[0];
    sheet := TsWorksheet(chart.Worksheet);
    r1 := series.LabelRange.Row1;
    c1 := series.LabelRange.Col1;
    r2 := series.LabelRange.Row2;
    c2 := series.LabelRange.Col2;
    refStr := GetSheetCellRangeString_ODS(sheet.Name, sheet.Name, r1, c1, r2, c2, rfAllRel, false);
    AppendToStream(AChartStream, Format(
      indent + '  <chart:categories table:cell-range-address="%s"/>' + LE,
      [ refStr ]
    ));
  end;

  // Write axis style
  AppendToStream(AStyleStream,
    GetChartAxisStyleAsXML(Axis, AStyleIndent, AStyleID)
  );

  // Next style
  inc(AStyleID);

  // Axis title
  if Axis.Title.Visible and (Axis.Title.Caption <> '') then
  begin
    AppendToStream(AChartStream, Format(
      indent + '  <chart:title chart:style-name="ch%d">' + LE +
      indent + '    <text:p>%s</text:p>' + LE +
      indent + '  </chart:title>' + LE,
      [ AStyleID, Axis.Title.Caption ]
    ));

    // Axis title style
    AppendToStream(AStyleStream,
      GetChartCaptionStyleAsXML(chart, captionKind, AStyleIndent, AStyleID)
    );

    // Next style
    inc(AStyleID);
  end;

  // Major grid lines
  if Axis.MajorGridLines.Style <> clsNoLine then
  begin
    AppendToStream(AChartStream, Format(
      indent + '  <chart:grid chart:style-name="ch%d" chart:class="major"/>' + LE,
      [ AStyleID ]
    ));

    // Major grid lines style
    AppendToStream(AStyleStream,
      GetChartLineStyleAsXML(chart, Axis.MajorGridLines, AStyleIndent, AStyleID)
    );

    // Next style
    inc(AStyleID);
  end;

  // Minor grid lines
  if Axis.MinorGridLines.Style <> clsNoLine then
  begin
    AppendToStream(AChartStream, Format(
      indent + '  <chart:grid chart:style-name="ch%d" chart:class="minor"/>' + LE,
      [ AStyleID ]
    ));

    // Minor grid lines style
    AppendToStream(AStyleStream,
      GetChartLineStyleAsXML(chart, Axis.MinorGridLines, AStyleIndent, AStyleID)
    );

    // Next style
    inc(AStyleID);
  end;

  // Close the xml node
  AppendToStream(AChartStream,
    indent + '</chart:axis>' + LE
  );
end;

{ Writes the chart's background to the xml stream }
procedure TsSpreadOpenDocChartWriter.WriteChartBackground(
  AChartStream, AStyleStream: TStream; AChartIndent, AStyleIndent: Integer;
  AChart: TsChart; var AStyleID: Integer);
var
  indent: String;
  chartClass: String;
begin
  chartClass := CHART_TYPE_NAMES[AChart.GetChartType];
  if chartClass <>  '' then
    chartClass := 'chart:class="chart:' + chartClass + '"';

  indent := DupeString(' ', AChartIndent);
  AppendToStream(AChartStream, Format(
    indent + '<chart:chart chart:style-name="ch%d" %s' + LE +
    indent + '    svg:width="%.3fmm" svg:height="%.3fmm" ' + LE +
    indent + '    xlink:type="simple" xlink:href="..">' + LE, [
    AStyleID,
    chartClass,
    AChart.Width, AChart.Height      // Width, Height are in mm
    ],  FPointSeparatorSettings
  ));

  AppendToStream(AStyleStream,
    GetChartBackgroundStyleAsXML(AChart, AChart.Background, AChart.Border, AStyleIndent, AStyleID)
  );

  inc(AStyleID);
end;

{ Writes, for each gradient used by the chart, a node to the Object/styles xml file }
procedure TsSpreadOpenDocChartWriter.WriteObjectGradientStyles(AStream: TStream;
  AChart: TsChart; AIndent: Integer);
var
  i: Integer;
  gradient: TsChartGradient;
  style: String;
  indent: String;
  clr1, clr2: TsChartColor;
begin
  indent := DupeString(' ', AIndent);
  for i := 0 to AChart.Gradients.Count-1 do
  begin
    gradient := AChart.Gradients[i];
    clr1 := gradient.Startcolor;
    clr2 := gradient.EndColor;
    if gradient.Style = cgsAxial then
      SwapColors(clr1, clr2);
    style := indent + Format(
      '<draw:gradient draw:name="%s" draw:display-name="%s" ' +
        'draw:style="%s" ' +
        'draw:start-color="%s" draw:end-color="%s" ' +
        'draw:start-intensity="%.0f%%" draw:end-intensity="%.0f%%" ' +
        'draw:border="%.0f%%" ',
      [ ASCIIName(gradient.Name), gradient.Name,
        GRADIENT_STYLES[gradient.Style],
        ColorToHTMLColorStr(clr1.Color), ColorToHTMLColorStr(clr2.Color),
        100.0, 100.0,
        gradient.StartBorder * 100
      ]
    );
    case gradient.Style of
      cgsLinear:
        style := style + Format(
          'draw:angle="%.0fdeg" ',
          [ FMod(90 + gradient.Angle, 360.0) ],   // transform to fps angle orientations
          FPointSeparatorSettings
        );
      cgsAxial:
        style := style + Format(
          'draw:angle="%.0fdeg" ',
          [ FMod(gradient.Angle, 360.0) ],
          FPointSeparatorSettings
        );
      cgsElliptic, cgsSquare, cgsRectangular:
        style := style + Format(
          'draw:cx="%.0f%%" draw:cy="%.0f%%" draw:angle="%.0fdeg" ',
          [ gradient.CenterX * 100, gradient.CenterY * 100, gradient.Angle ],
          FPointSeparatorSettings
        );
      cgsRadial, cgsShape:
        style := style + Format(
          'draw:cx="%.0f%%" draw:cy="%.0f%%" ',
          [ gradient.CenterX * 100, gradient.CenterY * 100 ],
          FPointSeparatorSettings
        );
      else
        raise Exception.Create('Unsupported gradient style');
    end;
    style := style + '/>' + LE;

    AppendToStream(AStream, style);
  end;
end;

procedure TsSpreadOpenDocChartWriter.WriteObjectHatchStyles(AStream: TStream;
  AChart: TsChart; AIndent: Integer);
var
  indent: String;
  style: String;
  i: Integer;
  hatch: TsChartHatch;
begin
  indent := DupeString(' ', AIndent);
  for i := 0 to AChart.Hatches.Count-1 do
  begin
    hatch := AChart.Hatches[i];
    style := Format(indent +
      '<draw:hatch draw:name="%s" draw:display-name="%s" ' +
        'draw:style="%s" ' +
        'draw:color="%s" ' +
        'draw:distance="%.2fmm" ' +
        'draw:rotation="%.0f" />',
      [ ASCIIName(hatch.Name), hatch.Name,
        HATCH_STYLES[hatch.Style],
        ColorToHTMLColorStr(hatch.PatternColor.Color),
        hatch.PatternWidth,
        hatch.PatternAngle*10
      ],
      FPointSeparatorSettings
    );
    AppendToStream(AStream, style);
  end;
end;

procedure TsSpreadOpenDocChartWriter.WriteObjectLineStyles(AStream: TStream;
  AChart: TsChart; AIndent: Integer);
const
  LENGTH_UNIT: array[boolean] of string = ('mm', '%'); // relative to line width
  DECS: array[boolean] of Integer = (1, 0);            // relative to line width
var
  i: Integer;
  lineStyle: TsChartLineStyle;
  seg1, seg2: String;
  indent: String;
begin
  indent := DupeString(' ', AIndent);
  for i := 0 to AChart.NumLineStyles-1 do
  begin
    lineStyle := AChart.GetLineStyle(i);
    if linestyle.Segment1.Count > 0 then
      seg1 := Format('draw:dots1="%d" draw:dots1-length="%.*f%s" ', [
        lineStyle.Segment1.Count,
        DECS[linestyle.RelativeToLineWidth], linestyle.Segment1.Length, LENGTH_UNIT[linestyle.RelativeToLineWidth]
        ], FPointSeparatorSettings
      )
    else
      seg1 := '';

    if linestyle.Segment2.Count > 0 then
      seg2 := Format('draw:dots2="%d" draw:dots2-length="%.*f%s" ', [
        lineStyle.Segment2.Count,
        DECS[linestyle.RelativeToLineWidth], linestyle.Segment2.Length, LENGTH_UNIT[linestyle.RelativeToLineWidth]
        ], FPointSeparatorSettings
      )
    else
      seg2 := '';

    if (seg1 <> '') or (seg2 <> '') then
      AppendToStream(AStream, indent + Format(
        '<draw:stroke-dash draw:name="%s" draw:display-name="%s" draw:style="round" draw:distance="%.*f%s" %s%s/>' + LE, [
        ASCIIName(linestyle.Name), linestyle.Name,
        DECS[linestyle.RelativeToLineWidth], linestyle.Distance, LENGTH_UNIT[linestyle.RelativeToLineWidth],
        seg1, seg2
        ], FPointSeparatorSettings
      ));
  end;
end;

{ Writes the chart's legend to the xml stream }
procedure TsSpreadOpenDocChartWriter.WriteChartLegend(
  AChartStream, AStyleStream: TStream; AChartIndent, AStyleIndent: Integer;
  AChart: TsChart; var AStyleID: Integer);
var
  indent: String;
  canOverlap: String = '';
begin
  if (not AChart.Legend.Visible) then
    exit;

  if AChart.Legend.CanOverlapPlotArea then
    canOverlap := 'loext:overlay="true" ';

  // Write legend properties
  indent := DupeString(' ', AChartIndent);
  AppendToStream(AChartStream, Format(
    indent + '<chart:legend chart:style-name="ch%d" chart:legend-position="%s" style:legend-expansion="wide" %s/>' + LE,
    [ AStyleID, LEGEND_POSITION[AChart.Legend.Position], canOverlap ]
  ));

  // Write legend style
  AppendToStream(AStyleStream,
    GetChartLegendStyleAsXML(AChart, AStyleIndent, AStyleID)
  );

  // Next style
  inc(AStyleID);
end;

procedure TsSpreadOpenDocChartWriter.WriteChartNumberStyles(AStream: TStream;
  AIndent: Integer; AChart: TsChart);
var
  indent: String;
  numFmtName: String;
  numFmtStr: String;
  numFmtXML: String;
  i: Integer;
  parser: TsSpreadOpenDocNumFormatParser;
begin
  indent := DupeString(' ', AIndent);

  for i := 0 to FNumberFormatList.Count-1 do begin
    numFmtName := Format('N%d', [i]);
    numFmtStr := FNumberFormatList.ValueFromIndex[i];
    parser := TsSpreadOpenDocNumFormatParser.Create(numFmtStr, FWriter.Workbook.FormatSettings);
    try
      numFmtXML := parser.BuildXMLAsString(numFmtName);
      if numFmtXML <> '' then
        AppendToStream(AStream, indent + numFmtXML);
    finally
      parser.Free;
    end;
  end;
end;

{ Writes the file "Object N/styles.xml" (N = 1, 2, ...) which is needed by the
  charts since it defines the line dash patterns, or gradients. }
procedure TsSpreadOpenDocChartWriter.WriteObjectStyles(AStream: TStream;
  AChart: TsChart);
begin
  AppendToStream(AStream,
    XML_HEADER + LE);

  AppendToStream(AStream,
    '<office:document-styles ' + LE +
    '  xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"' + LE +
    '  xmlns:ooo="http://openoffice.org/2004/office"' + LE +
    '  xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"' + LE +
    '  xmlns:xlink="http://www.w3.org/1999/xlink"' + LE +
    '  xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"' + LE +
    '  xmlns:dc="http://purl.org/dc/elements/1.1/"' + LE +
    '  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"' + LE +
    '  xmlns:rpt="http://openoffice.org/2005/report"' + LE +
    '  xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"' + LE +
    '  xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0"' + LE +
    '  xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"' + LE +
    '  xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"' + LE +
    '  xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"' + LE +
    '  xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"' + LE +
    '  xmlns:ooow="http://openoffice.org/2004/writer"' + LE +
    '  xmlns:oooc="http://openoffice.org/2004/calc"' + LE +
    '  xmlns:css3t="http://www.w3.org/TR/css3-text/"' + LE +
    '  xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"' + LE +
    '  xmlns:tableooo="http://openoffice.org/2009/table"' + LE +
    '  xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"' + LE +
    '  xmlns:drawooo="http://openoffice.org/2010/draw"' + LE +
    '  xmlns:xhtml="http://www.w3.org/1999/xhtml"' + LE +
    '  xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"' + LE +
    '  xmlns:grddl="http://www.w3.org/2003/g/data-view#"' + LE +
    '  xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0"' + LE +
    '  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"' + LE +
    '  xmlns:dom="http://www.w3.org/2001/xml-events"' + LE +
    '  xmlns:chartooo="http://openoffice.org/2010/chart" office:version="1.3">' + LE
  );

  AppendToStream(AStream,
    '  <office:styles>' + LE
  );

  WriteObjectLineStyles(AStream, AChart, 4);
  WriteObjectGradientStyles(AStream, AChart, 4);
  WriteObjectHatchStyles(AStream, AChart, 4);

  AppendToStream(AStream,
    '  </office:styles>' + LE +
    '</office:document-styles>'
  );
end;

procedure TsSpreadOpenDocChartWriter.WriteChartPlotArea(
  AChartStream, AStyleStream: TStream; AChartIndent, AStyleIndent: Integer;
  AChart: TsChart; var AStyleID: Integer);
var
  indent: String;
  i: Integer;
begin
  indent := DupeString(' ', AChartIndent);

  // Plot area properties
  // (ods has a table:cell-range-address here but it is reconstructed by Calc)
  AppendToStream(AChartStream, Format(
    indent + '<chart:plot-area chart:style-name="ch%d" chart:data-source-has-labels="both">' + LE,
    [ AStyleID ]
  ));
  // Plot area style
  AppendToStream(AStyleStream,
    GetChartPlotAreaStyleAsXML(AChart, AStyleIndent, AStyleID)
  );

  // Next style
  inc(AStyleID);

  // Wall properties
  AppendToStream(AChartStream, Format(
    indent + '  <chart:wall chart:style-name="ch%d"/>' + LE,
    [ AStyleID ]
  ));
  // Wall style
  AppendToStream(AStyleStream,
    GetChartBackgroundStyleAsXML(AChart, AChart.PlotArea.Background, AChart.PlotArea.Border, AStyleIndent, AStyleID)
  );

  // Next style
  inc(AStyleID);

  // Floor properties
  AppendToStream(AChartStream, Format(
    indent + '  <chart:floor chart:style-name="ch%d"/>' + LE,
    [ AStyleID ]
  ));
  // Floor style
  AppendToStream(AStyleStream,
    GetChartBackgroundStyleAsXML(AChart, AChart.Floor.Background, AChart.Floor.Border, AStyleIndent, AStyleID)
  );

  // Next style
  inc(AStyleID);

  // primary x axis
  WriteChartAxis(AChartStream, AStyleStream, AChartIndent+2, AStyleIndent, AChart.XAxis, AStyleID);

  // primary y axis
  WriteChartAxis(AChartStream, AStyleStream, AChartIndent+2, AStyleIndent, AChart.YAxis, AStyleID);

  // secondary x axis
  CheckAxis(AChart, AChart.X2Axis);
  WriteChartAxis(AChartStream, AStyleStream, AChartIndent+2, AStyleIndent, AChart.X2Axis, AStyleID);

  // secondary y axis
  CheckAxis(AChart, AChart.Y2Axis);
  WriteChartAxis(AChartStream, AStyleStream, AChartIndent+2, AStyleIndent, AChart.Y2Axis, AStyleID);

  // series
  for i := 0 to AChart.Series.Count-1 do
    if AChart.Series[i].ChartType = ctStock then
      WriteChartStockSeries(AChartStream, AStyleStream, AChartIndent+2, AStyleIndent, AChart, i, AStyleID)
    else
      WriteChartSeries(AChartStream, AStyleStream, AChartIndent+2, AStyleIndent, AChart, i, AStyleID);

  // close xml node
  AppendToStream(AChartStream,
    indent + '</chart:plot-area>' + LE
  );
end;

procedure TsSpreadOpenDocChartWriter.WriteChartSeries(
  AChartStream, AStyleStream: TStream; AChartIndent, AStyleIndent: Integer;
  AChart: TsChart; ASeriesIndex: Integer; var AStyleID: Integer);
var
  indent: String;
  series: TsChartSeries;
  valuesRange: String = '';
  domainRangeX: String = '';
  domainRangeY: String = '';
  fillColorRange: String = '';
  lineColorRange: String = '';
  chartType: String = '';
  seriesYAxis: String = '';
  trendlineEquation: String = '';
  trendline: TsChartTrendline = nil;
  titleAddr: String;
  i, idx, count: Integer;
  nextStyleID, seriesStyleID, trendlineStyleID, trendlineEquStyleID: Integer;
  xErrStyleID, yErrStyleID, dataStyleID: Integer;
begin
  indent := DupeString(' ', AChartIndent);

  nextstyleID := AStyleID;
  seriesStyleID := AStyleID;
  trendlineStyleID := -1;
  trendlineEquStyleID := -1;
  xErrStyleID := -1;
  yErrStyleID := -1;
  dataStyleID := -1;

  series := AChart.Series[ASeriesIndex];

  // These are the x values of a scatter plot.
  if (series is TsCustomScatterSeries) then
  begin
    domainRangeX := GetSheetCellRangeString_ODS(
      series.XRange.GetSheet1Name, series.XRange.GetSheet2Name,
      series.XRange.Row1, series.XRange.Col1,
      series.XRange.Row2, series.XRange.Col2,
      rfAllRel, false
    );
  end;

  if series is TsBubbleSeries then
  begin
    // These are the y values of the in-plane coordinates of each bubble position.
    domainRangeY := GetSheetCellRangeString_ODS(
      series.YRange.GetSheet1Name, series.YRange.GetSheet2Name,
      series.YRange.Row1, series.YRange.Col1,
      series.YRange.Row2, series.YRange.Col2,
      rfAllRel, false
    );
    // These are the bubble radii
    with TsBubbleSeries(series) do
    begin
      valuesRange := GetSheetCellRangeString_ODS(
        BubbleRange.GetSheet1Name, BubbleRange.GetSheet2Name,
        BubbleRange.Row1, BubbleRange.Col1,
        BubbleRange.Row2, BubbleRange.Col2,
        rfAllRel, false
      );
    end
  end else
    // These are the y values of the non-bubble series
    valuesRange := GetSheetCellRangeString_ODS(
      series.YRange.GetSheet1Name, series.YRange.GetSheet2Name,
      series.YRange.Row1, series.YRange.Col1,
      series.YRange.Row2, series.YRange.Col2,
      rfAllRel, false
    );

  // Fill colors for bars, line series symbols, bubbles
  if (series.FillColorRange.Row1 <> series.FillColorRange.Row2) or
     (series.FillColorRange.Col1 <> series.FillColorRange.Col2)
  then
    fillColorRange := GetSheetCellRangeString_ODS(
      series.FillColorRange.GetSheet1Name, series.FillColorRange.GetSheet2Name,
      series.FillColorRange.Row1, series.FillColorRange.Col1,
      series.FillColorRange.Row2, series.FillColorRange.Col2,
      rfAllRel, false
    );

  // Line colors for bars, line series symbols, bubbles etc.
  if not series.LineColorRange.IsEmpty then
    lineColorRange := GetSheetCellRangeString_ODS(
      series.LineColorRange.GetSheet1Name, series.LineColorRange.GetSheet2Name,
      series.LineColorRange.Row1, series.LineColorRange.Col1,
      series.LineColorRange.Row2, series.LineColorRange.Col2,
      rfAllRel, false
    );

  // Axis of the series
  if AChart.Y2Axis.Visible then
    case series.YAxis of
      calPrimary  : seriesYAxis := 'chart:attached-axis="primary-y" ';
      calSecondary: seriesYAxis := 'chart:attached-axis="secondary-y" ';
    end;

  // And this is the title of the series for the legend
  if series.TitleAddr.IsUsed then
    titleAddr := GetSheetCellRangeString_ODS(
      series.TitleAddr.GetSheetName, series.TitleAddr.GetSheetName,
      series.TitleAddr.Row, series.TitleAddr.Col,
      series.TitleAddr.Row, series.TitleAddr.Col,
      rfAllRel, false
    )
  else
    titleAddr := '';

  // Number of data points
  if series.YValuesInCol then
    count := series.YRange.Row2 - series.YRange.Row1 + 1
  else
    count := series.YRange.Col2 - series.YRange.Col1 + 1;

  if series is TsPieSeries then
    chartType := 'circle'
  else
    chartType := CHART_TYPE_NAMES[series.ChartType];

  AppendToStream(AChartStream, Format(
    indent + '<chart:series chart:style-name="ch%d" ' +
               'chart:class="chart:%s" ' +                    // series type
               seriesYAxis +                                  // attached y axis
               'chart:values-cell-range-address="%s" ' +      // y values
               'chart:label-cell-address="%s">' + LE,         // series title
    [ seriesStyleID, chartType, valuesRange, titleAddr, chartType ]
  ));
  inc(nextStyleID);

  if domainRangeY <> '' then
    AppendToStream(AChartStream, Format(
      indent + '<chart:domain table:cell-range-address="%s"/>' + LE,
      [ domainRangeY ]
    ));
  if domainRangeX <> '' then
    AppendToStream(AChartStream, Format(
      indent + '<chart:domain table:cell-range-address="%s"/>' + LE,
      [ domainRangeX ]
    ));
  if fillColorRange <> '' then
    AppendToStream(AChartStream, Format(
      indent + '<loext:property-mapping loext:property="FillColor" loext:cell-range-address="%s"/>' + LE,
      [ fillColorRange ]
    ));
  if lineColorRange <> '' then
    AppendToStream(AChartStream, Format(
      indent + '<loext:property-mapping loext:property="BorderColor" loext:cell-range-address="%s"/>' + LE,
      [ lineColorRange ]
    ));

  // Error bars
  if series.XErrorBars.Visible then
  begin
    xErrStyleID := nextStyleID;
    AppendToStream(AChartStream, Format(
      indent + '<chart:error-indicator chart:style-name="ch%d" chart:dimension="x" />',
      [ xErrStyleID ]
    ));
    inc(nextStyleID);
  end;

  if series.YErrorBars.Visible then
  begin
    yErrStyleID := nextStyleID;
    AppendToStream(AChartStream, Format(
      indent + '<chart:error-indicator chart:style-name="ch%d" chart:dimension="y" />',
      [ yErrStyleID ]
    ));
    inc(nextStyleID);
  end;

  // Trend line
  if (series is TsScatterSeries) then
  begin
    trendline := TsScatterSeries(series).trendline;
    if trendline.TrendlineType <> tltNone then
    begin
      trendlineStyleID := nextStyleID;
      inc(nextStyleID);

      if trendline.DisplayEquation or trendline.DisplayRSquare then
      begin
        if (not trendline.Equation.DefaultXName) or (not trendline.Equation.DefaultYName) or
           (not trendline.Equation.DefaultBorder) or (not trendline.Equation.DefaultFill) or
           (not trendline.Equation.DefaultFont) or (not trendline.Equation.DefaultNumberFormat) or
           (not trendline.Equation.DefaultPosition) then
        begin
          trendlineEquStyleID := nextStyleID;
          trendlineEquation := trendlineEquation + Format('chart:style-name="ch%d" ', [ trendlineEquStyleID ]);
          inc(nextStyleID);
        end;
      end;
      if trendline.DisplayEquation then
        trendlineEquation := trendlineEquation + 'chart:display-equation="true" ';
      if trendline.DisplayRSquare then
        trendlineEquation := trendlineEquation + 'chart:display-r-square="true" ';

      if trendlineEquation <> '' then
      begin
        if not trendline.Equation.DefaultPosition then
          trendlineEquation := trendlineEquation + Format(
            'svg:x="%.2fmm" svg:y="%.2fmm" ',
            [ trendline.Equation.Left, trendline.Equation.Top ],
            FPointSeparatorSettings
          );

        AppendToStream(AChartStream, Format(
          indent + '  <chart:regression-curve chart:style-name="ch%d">' + LE +
          indent + '    <chart:equation %s />' + LE +
          indent + '  </chart:regression-curve>' + LE,
          [ trendlineStyleID, trendlineEquation ]
        ));
      end else
        AppendToStream(AChartStream, Format(
          indent + '  <chart:regression-curve chart:style-name="ch%d"/>',
          [ trendlineStyleID ]
        ));
    end;
  end;

  // Individual data point styles
  if series.DataPointStyles.Count = 0 then
    AppendToStream(AChartStream, Format(
      indent + '  <chart:data-point chart:repeated="%d" />' + LE,
      [ count ]
    ))
  else
  begin
    dataStyleID := nextStyleID;
    // Every data point gets a <chart:data-point> node with individual format
    for i := 0 to count - 1 do
    begin
      AppendToStream(AChartStream, Format(
        indent + '  <chart:data-point chart:style-name="ch%d"/>' + LE,
        [ dataStyleID + i ]
      ));
      inc(nextStyleID);
    end;
  end;

  AppendToStream(AChartStream,
    indent + '</chart:series>' + LE
  );

  // ---------------------------------------------------------------------------

  // Series style
  AppendToStream(AStyleStream,
    GetChartSeriesStyleAsXML(AChart, ASeriesIndex, AStyleIndent, seriesStyleID)
  );

  // Trend line style
  if trendlineStyleID <> -1 then
  begin
    AppendToStream(AStyleStream,
      GetChartRegressionStyleAsXML(AChart, ASeriesIndex, AStyleIndent, trendlineStyleID)
    );

    // Style of regression equation
    if trendlineEquStyleID <> -1 then
    begin
      AppendToStream(AStyleStream,
        GetChartRegressionEquationStyleAsXML(AChart, trendline.Equation, AStyleIndent, trendlineEquStyleID)
      );
    end;
  end;

  // Error bar styles
  if xErrStyleID <> -1 then
    AppendToStream(AStyleStream,
      GetChartErrorBarStyleAsXML(AChart, series.XErrorBars, AStyleIndent, xErrStyleID)
    );

  if yErrStyleID <> -1 then
    AppendToStream(AStyleStream,
      GetChartErrorBarStyleAsXML(AChart, series.YErrorBars, AStyleIndent, yErrStyleID)
    );

  // Data point styles
  if series.DataPointStyles.Count > 0 then
  begin
    for i := 0 to count - 1 do
    begin
      idx := series.DataPointStyles.IndexOfDatapoint(i);
      AppendToStream(AStyleStream,
        GetChartSeriesDataPointStyleAsXML(AChart, ASeriesIndex, idx, AStyleIndent, dataStyleID)
      );
      inc(dataStyleID);
    end;
  end;

  // Next style
  AStyleID := nextStyleID;
end;

procedure TsSpreadOpenDocChartWriter.WriteChartStockSeries(
  AChartStream, AStyleStream: TStream; AChartIndent, AStyleIndent: Integer;
  AChart: TsChart; ASeriesIndex: Integer; var AStyleID: Integer);

  procedure WriteRange(const AIndent, ARangeStr, ATitleStr, AYAxisStr: String; ACount: Integer);
  begin
    AppendToStream(AChartStream, Format(
      AIndent + '<chart:series ' + ARangeStr + AtitleStr + AYAxisStr + '>' + LE +
      AIndent + '  <chart:data-point chart:repeated="%d"/>' + LE +
      AIndent + '</chart:series>' + LE,
      [ ACount ] ));
  end;

var
  indent: String;
  openRange: String = '';
  highRange: String = '';
  lowRange: String = '';
  closeRange: String = '';
  titleAddr: String = '';
  seriesYAxis: String = '';
  series: TsStockSeries;
  count: Integer;
begin
  if not (AChart.Series[ASeriesIndex] is TsStockSeries) then
    exit;
  series := TsStockSeries(AChart.Series[ASeriesIndex]);
  indent := DupeString(' ', AChartIndent);

  // These are the open/high/low/close values of the OHLC series
  if series.CandleStick and (not series.OpenRange.IsEmpty) then
    openRange := Format('chart:values-cell-range-address="%s" ', [
      GetSheetCellRangeString_ODS(
        series.OpenRange.GetSheet1Name, series.OpenRange.GetSheet2Name,
        series.OpenRange.Row1, series.OpenRange.Col1,
        series.OpenRange.Row2, series.OpenRange.Col2,
        rfAllRel, false)
      ]);
  if not series.HighRange.IsEmpty then
    highRange := Format('chart:values-cell-range-address="%s" ', [
      GetSheetCellRangeString_ODS(
        series.HighRange.GetSheet1Name, series.HighRange.GetSheet2Name,
        series.HighRange.Row1, series.HighRange.Col1,
        series.HighRange.Row2, series.HighRange.Col2,
        rfAllRel, false)
      ]);
  if not series.LowRange.IsEmpty then
    lowRange := Format('chart:values-cell-range-address="%s" ', [
      GetSheetCellRangeString_ODS(
        series.LowRange.GetSheet1Name, series.LowRange.GetSheet2Name,
        series.LowRange.Row1, series.LowRange.Col1,
        series.LowRange.Row2, series.LowRange.Col2,
        rfAllRel, false)
      ]);
  if not series.CloseRange.IsEmpty then
    closeRange := Format('chart:values-cell-range-address="%s" ',[
      GetSheetCellRangeString_ODS(
        series.CloseRange.GetSheet1Name, series.CloseRange.GetSheet2Name,
        series.CloseRange.Row1, series.CloseRange.Col1,
        series.CloseRange.Row2, series.CloseRange.Col2,
        rfAllRel, false)
      ]);

  // Title of the series for the legend
  titleAddr := Format('chart:label-cell-address="%s" ', [
    GetSheetCellRangeString_ODS(
      series.TitleAddr.GetSheetName, series.TitleAddr.GetSheetName,
      series.TitleAddr.Row, series.TitleAddr.Col,
      series.TitleAddr.Row, series.TitleAddr.Col,
      rfAllRel, false)
    ]);

  // Axis of the series
  case series.YAxis of
    calPrimary  : seriesYAxis := 'chart:attached-axis="primary-y" ';
    calSecondary: seriesYAxis := 'chart:attached-axis="secondary-y" ';
  end;

  // Number of data points
  if series.YValuesInCol then
    count := series.YRange.Row2 - series.YRange.Row1 + 1
  else
    count := series.YRange.Col2 - series.YRange.Col1 + 1;

  // Store the series properties

  // "Open" values, only for CandleStick mode
  if series.CandleStick then
    WriteRange(indent, openRange, '', seriesYAxis, count);
  // "Low" values
  WriteRange(indent, lowRange, '', seriesYAxis, count);
  // "High" values
  WriteRange(indent, highRange, '', seriesYAxis, count);
  // "Close" values
  WriteRange(indent, closeRange, titleAddr, seriesYAxis, count);

  // Stock series styles
  AppendToStream(AChartStream, Format(
    indent + '<chart:stock-gain-marker chart:style-name="ch%d" />' + LE +
    indent + '<chart:stock-loss-marker chart:style-name="ch%d" />' + LE +
    indent + '<chart:stock-range-line chart:style-name="ch%d" />' + LE,
    [ AStyleID, AStyleID + 1, AStyleID + 2 ]));

  AppendToStream(AStyleStream,
    GetChartStockSeriesStyleAsXML(AChart, series, 0, AStyleIndent, AStyleID));
  AppendToStream(AStyleStream,
    GetChartStockSeriesStyleAsXML(AChart, series, 1, AStyleIndent, AStyleID + 1));
  AppendToStream(AStyleStream,
    GetChartStockSeriesStyleAsXML(AChart, series, 2, AStyleIndent, AStyleID + 2));

  inc(AStyleID, 3);
end;

procedure TsSpreadOpenDocChartWriter.WriteCharts;
var
  i: Integer;
  chart: TsChart;
begin
  for i := 0 to TsWorkbook(Writer.Workbook).GetChartCount - 1 do
    begin
      chart := TsWorkbook(Writer.Workbook).GetChartByIndex(i);
      WriteChart(FSCharts[i], chart);
      WriteObjectStyles(FSObjectStyles[i], chart);
    end;
end;

(* wp:
   DO NOT DELETE THIS - IT WAS A PAIN TO GET THIS, AND MAYBE IT WILL BE NEEDED
   LATER.
   AT THE MOMENT THIS IS NOT NEEDED, IN FACT, IT IS EVEN DETRIMENTAL:
   WITH THIS CODE INCLUDED, SERIES FILLS ARE IGNORED AND TITLES ARE NOT CORRECT.

{ Writes the chart's data table. NOTE: The chart gets its data from this table
  rather than from the worksheet! }
procedure TsSpreadOpenDocChartWriter.WriteChartTable(AStream: TStream;
  AChart: TsChart; AIndent: Integer);
var
  auxBook: TsWorkbook;
  auxSheet: TsWorksheet;

  procedure WriteAuxCell(AIndent: Integer; ACell: PCell);
  var
    ind: String;
    valueType: String;
    value: String;
    officeValue: String;
    draw: String;
  begin
    ind := DupeString(' ', AIndent);
    if (ACell = nil) or (ACell^.ContentType = cctEmpty) then
    begin
      AppendToStream(AStream,
        ind + '<table:table-cell>' + LE +
        ind + '  <text:p/>' + LE +
        ind + '</table:table-cell>' + LE
      );
      exit;
    end;

    case ACell^.ContentType of
      cctUTF8String:
        begin
          valueType := 'string';
          value := auxSheet.ReadAsText(ACell);
          officeValue := '';
        end;
      cctNumber, cctDateTime, cctBool:
        begin
          valueType := 'float';
          value := Format('%g', [auxSheet.ReadAsNumber(ACell)], FPointSeparatorSettings);
        end;
      cctError:
        begin
          valueType := 'float';
          value := 'NaN';
        end;
    end;

    if ACell^.ContentType = cctUTF8String then
      officeValue := ''
    else
      officeValue := ' office:value="' + value + '"';

    if auxSheet.HasComment(ACell) then
    begin
      draw := auxSheet.ReadComment(ACell);
      if draw <> '' then
        draw := Format(
          ind + '  <draw:g>' + LE +
          ind + '    <svg:desc>%s</svg:desc>' + LE +
          ind + '  </draw:g>' + LE, [draw]);
    end else
      draw := '';
    AppendToStream(AStream, Format(
      ind + '<table:table-cell office:value-type="%s"%s>' + LE +
        ind + '  <text:p>%s</text:p>' + LE +
        draw +
        ind + '</table:table-cell>' + LE,
        [ valueType, officevalue, value ]
      ));
  end;

var
  ind: String;
  colCountStr: String;
  n: Integer;
  r, c: Cardinal;
begin
  ind := DupeString(' ', AIndent);
  n := AChart.Series.Count;
  if n > 0 then
  begin
    if AChart.IsScatterChart then
      n := n * 2;
    colCountStr := Format('table:number-columns-repeated="%d"', [n]);
  end else
    colCountStr := '';

  AppendToStream(AStream, Format(
        ind + '<table:table table:name="local-table">' + LE +
        ind + '  <table:table-header-columns>' + LE +
        ind + '    <table:table-column/>' + LE +
        ind + '  </table:table-header-columns>' + LE +
        ind + '  <table:table-columns>' + LE +
        ind + '    <table:table-column %s/>' + LE +
        ind + '  </table:table-columns>' + LE, [ colCountStr ]
  ));

  auxBook := TsWorkbook.Create;
  try
    auxSheet := auxBook.AddWorksheet('chart');
    PrepareChartTable(AChart, auxSheet);

    // Header rows (containing the series names)
    AppendToStream(AStream,
        ind + '  <table:table-header-rows>' + LE +
        ind + '    <table:table-row>' + LE );
    for c := 0 to auxSheet.GetLastColIndex do
      WriteAuxCell(AIndent + 6, auxSheet.FindCell(0, c));
    AppendToStream(AStream,
        ind + '    </table:table-row>' + LE +
        ind + '  </table-header-rows>' + LE
    );

    // Write data rows
    AppendToStream(AStream,
        ind + '  <table:table-rows>' + LE
    );
    for r := 1 to auxSheet.GetLastRowIndex do
    begin
      AppendToStream(AStream,
        ind + '    <table:table-row>' + LE
      );
      for c := 0 to auxSheet.GetlastColIndex do
        WriteAuxCell(AIndent + 6, auxSheet.FindCell(r, c));
      AppendToStream(AStream,
        ind + '    </table:table-row>' + LE
      );
    end;
    AppendToStream(AStream,
        ind + '  </table:table-rows>' + LE +
        ind + '</table:table>' + LE
    );

    //auxBook.WriteToFile('table.ods', true);
  finally
    auxBook.Free;
  end;
end;
  *)
{ Writes the chart's title (or subtitle, depending on the value of IsSubTitle)
  to the xml stream (chart stream) and the corresponding style to the stylestream. }
procedure TsSpreadOpenDocChartWriter.WriteChartTitle(
  AChartStream, AStyleStream: TStream; AChartIndent, AStyleIndent: Integer;
  AChart: TsChart; IsSubtitle: Boolean; var AStyleID: Integer);
var
  title: TsChartText;
  captionKind: Integer;
  elementName: String;
  indent: String;
begin
  if IsSubTitle then
  begin
    title := AChart.SubTitle;
    elementName := 'subtitle';
    captionKind := 2;
  end else
  begin
    title := AChart.Title;
    elementName := 'title';
    captionKind := 1;
  end;

  if (not title.Visible) or (title.Caption = '') then
    exit;

  // Write title properties
  indent := DupeString(' ', AChartIndent);
  AppendToStream(AChartStream, Format(
    indent + '<chart:%s chart:style-name="ch%d">' + LE +
    indent + '  <text:p>%s</text:p>' + LE +
    indent + '</chart:%s>' + LE,
    [ elementName, AStyleID, title.Caption, elementName ], FPointSeparatorSettings
  ));

  // Write title style
  AppendToStream(AStyleStream,
    GetChartCaptionStyleAsXML(AChart, captionKind, AStyleIndent, AStyleID)
  );

  // Next style
  inc(AStyleID);
end;

{$ENDIF}

end.

