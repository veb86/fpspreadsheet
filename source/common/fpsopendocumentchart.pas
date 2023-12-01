unit fpsOpenDocumentChart;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, StrUtils, Contnrs, FPImage,
 {$IF FPC_FULLVERSION >= 20701}
  zipper,
 {$ELSE}
  fpszipper,
 {$ENDIF}
  laz2_xmlread, laz2_DOM,
  fpsTypes, fpSpreadsheet, fpsChart, fpsUtils, fpsReaderWriter, fpsXMLCommon;

type

  { TsSpreadOpenDocChartReader }

  TsSpreadOpenDocChartReader = class(TsBasicSpreadChartReader)
  private
    FChartFiles: TStrings;
    FPointSeparatorSettings: TFormatSettings;
    FNumberFormatList: TStrings;
    FPieSeriesStartAngle: Integer;
    FStreamList: TFPObjectList;
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
      ASeries: TsChartSeries; var AFill: TsChartFill; var ALine: TsChartLine);
    procedure ReadChartSeriesProps(ANode, AStyleNode: TDOMNode; AChart: TsChart);
    procedure ReadChartSeriesStyle(AStyleNode: TDOMNode; AChart: TsChart; ASeries: TsChartSeries);
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
    procedure ReadCharts(AStream: TStream); override;
  end;

  TsSpreadOpenDocChartWriter = class(TsBasicSpreadChartWriter)
  private
    FSCharts: array of TStream;
    FSObjectStyles: array of TStream;
    FNumberFormatList: TStrings;
    FPointSeparatorSettings: TFormatSettings;
    function GetChartAxisStyleAsXML(Axis: TsChartAxis; AIndent, AStyleID: Integer): String;
    function GetChartBackgroundStyleAsXML(AChart: TsChart; AFill: TsChartFill;
      ABorder: TsChartLine; AIndent: Integer; AStyleID: Integer): String;
    function GetChartCaptionStyleAsXML(AChart: TsChart; ACaptionKind, AIndent, AStyleID: Integer): String;
    function GetChartFillStyleGraphicPropsAsXML(AChart: TsChart;
      AFill: TsChartFill): String;
    function GetChartLegendStyleAsXML(AChart: TsChart;
      AIndent, AStyleID: Integer): String;
    function GetChartLineStyleAsXML(AChart: TsChart;
      ALine: TsChartLine; AIndent, AStyleID: Integer): String;
    function GetChartLineStyleGraphicPropsAsXML(AChart: TsChart;
      ALine: TsChartLine): String;
    function GetChartPlotAreaStyleAsXML(AChart: TsChart;
      AIndent, AStyleID: Integer): String;
    function GetChartRegressionEquationStyleAsXML(AChart: TsChart;
      AEquation: TsRegressionEquation; AIndent, AStyleID: Integer): String;
    function GetChartRegressionStyleAsXML(AChart: TsChart; ASeriesIndex, AIndent, AStyleID: Integer): String;
    function GetChartSeriesDataPointStyleAsXML(AChart: TsChart; ASeriesIndex, APointIndex, AIndent, AStyleID: Integer): String;
    function GetChartSeriesStyleAsXML(AChart: TsChart; ASeriesIndex, AIndent, AStyleID: integer): String;

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

implementation

uses
  fpsOpenDocument;

type
  TAxisKind = 3..6;

  TsCustomLineSeriesOpener = class(TsCustomLineSeries);

const
  OPENDOC_PATH_METAINF_MANIFEST = 'META-INF/manifest.xml';
  OPENDOC_PATH_CHART_CONTENT    = 'Object %d/content.xml';
  OPENDOC_PATH_CHART_STYLES     = 'Object %d/styles.xml';

  DEFAULT_FONT_NAME = 'Liberation Sans';

  CHART_TYPE_NAMES: array[TsChartType] of string = (
    '', 'bar', 'line', 'area', 'barLine', 'scatter', 'bubble',
    'radar', 'filled-radar', 'circle', 'ring'
  );

  SYMBOL_NAMES: array[TsChartSeriesSymbol] of String = (
    'square', 'diamond', 'arrow-up', 'arrow-down', 'arrow-left',
    'arrow-right', 'circle', 'star', 'x', 'plus', 'asterisk'
  );  // unsupported: bow-tie, hourglass, horizontal-bar, vertical-bar

  GRADIENT_STYLES: array[TsChartGradientStyle] of string = (
    'linear', 'axial', 'radial', 'ellipsoid', 'square', 'rectangular'
  );

  HATCH_STYLES: array[TsChartHatchStyle] of string = (
    'single', 'double', 'triple'
  );

  LABEL_POSITION: array[TsChartLabelPosition] of string = (
    '', 'outside', 'inside', 'center');

  LEGEND_POSITION: array[TsChartLegendPosition] of string = (
    'end', 'top', 'bottom', 'start'
  );

  AXIS_ID: array[TAxisKind] of string = ('x', 'y', 'x', 'y');
  AXIS_LEVEL: array[TAxisKind] of string = ('primary', 'primary', 'secondary', 'secondary');

  REGRESSION_TYPE: array [TsRegressionType] of string = (
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
  for i := Length(AText) downto 0 do
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


{------------------------------------------------------------------------------}
{                        internal number formats                               }
{------------------------------------------------------------------------------}

type
  TsChartNumberFormatList = class(TStringList)
  public
    constructor Create;
    function Add(const ANumFormat: String): Integer; override;
    function FindFormatByName(const AName: String): String;
  end;

constructor TsChartNumberFormatList.Create;
begin
  inherited;
  Add('');  // default number format
end;

// Adds a new format, but make sure to avoid duplicates.
function TsChartNumberFormatList.Add(const ANumFormat: String): Integer;
begin
  if (ANumFormat = '') and (Count > 0) then
    Result := 0
  else
  begin
    Result := IndexOf(ANumFormat);
    if Result = -1 then
      Result := inherited Add(ANumFormat);
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
    Result := Values[AName];
    if Result = 'General' then
      Result := '';
  end;
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
  FNumberFormatList.NameValueSeparator := ':';
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
          AFill.Color := HTMLColorStrToColor(sc);
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
          AFill.Color := HTMLColorStrToColor(sc);
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
    AFill.Transparency := 1.0 - opacity;

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
    ALine.Color := HTMLColorStrToColor(sc);

  sw := GetAttrValue(ANode, 'svg:stroke-width');
  if sw = '' then
    sw := GetAttrValue(ANode, 'draw:stroke-width');
  if (sw <> '') and EvalLengthStr(sw, value, rel) then
    ALine.Width := value;

  so := GetAttrValue(ANode, 'draw:stroke-opacity');
  if (so <> '') and TryPercentStrToFloat(so, value) then
    ALine.Transparency := 1.0 - value*0.01;

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
  grid.Color := $c0c0c0;

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
  styleName: String;
  styleNode: TDOMNode;
begin
  styleName := GetAttrValue(AChartNode, 'chart:style-name');
  styleNode := FindStyleNode(AStyleNode, styleName);
  ReadChartBackgroundStyle(styleNode, AChart, AChart);
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
          s := GetAttrValue(AStyleNode, 'chart:stacked');
          if s = 'true' then
            AChart.StackMode := csmStacked;
          s := GetAttrValue(AStyleNode, 'chart:percentage');
          if s = 'true' then
            AChart.StackMode := csmStackedPercentage;
          s := GetAttrValue(AStyleNode, 'chart:angle-offset');
          if s <> '' then
            FPieSeriesStartAngle := StrToInt(s);
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
  odsReader: TsSpreadOpenDocReader;
  s, nodeName: String;
begin
  if not (ASeries is TsScatterSeries) then
    exit;

  series := TsCustomScatterSeries(ASeries);
  odsReader := TsSpreadOpenDocReader(Reader);

  nodeName := AStyleNode.NodeName;
  s := GetAttrValue(AStyleNode, 'style:data-style-name');
  if s <> '' then
    s := TsChartNumberFormatList(FNumberFormatList).FindFormatByName(s);
  series.Regression.Equation.NumberFormat := s;

  AStyleNode := AStyleNode.FirstChild;
  while Assigned(AStyleNode) do
  begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:graphic-properties':
        begin
          GetChartLineProps(AStyleNode, AChart, series.Regression.Equation.Border);
          GetChartFillProps(AStyleNode, AChart, series.Regression.Equation.Fill);
        end;
      'style:text-properties':
        GetChartTextProps(AStyleNode, series.Regression.Equation.Font);
      'style:chart-properties':
        begin
          s := GetAttrValue(AStyleNode, 'loext:regression-x-name');
          if s <> '' then
            series.Regression.Equation.XName := s;

          s := GetAttrValue(AStyleNode, 'loext:regression-y-name');
          if s <> '' then
            series.Regression.Equation.YName := s;
        end;
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartRegressionProps(ANode, AStyleNode: TDOMNode;
  AChart: TsChart; ASeries: TsChartSeries);
var
  series: TsCustomScatterSeries;
  s, nodeName: String;
  styleNode: TDOMNode;
  subNode: TDOMNode;
begin
  if not (ASeries is TsCustomScatterSeries) then
    exit;

  series := TsCustomScatterSeries(ASeries);

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
      series.Regression.DisplayEquation := (s = 'true');

      s := GetAttrValue(subNode, 'chart:display-r-square');
      series.Regression.DisplayRSquare := (s = 'true');

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
  series: TsScatterSeries;
  s, nodeName: String;
  rt: TsRegressionType;
  value: Double;
  intValue: Integer;
begin
  if not (ASeries is TsScatterSeries) then
    exit;
  series := TsScatterSeries(ASeries);

  AStyleNode := AStyleNode.FirstChild;
  while Assigned(AStyleNode) do
  begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:graphic-properties':
        GetChartLineProps(AStyleNode, AChart, series.Regression.Line);
      'style:chart-properties':
        begin
          s := GetAttrValue(AStyleNode, 'chart:regression-name');
          series.Regression.Title := s;

          s := GetAttrValue(AStyleNode, 'chart:regression-type');
          for rt in TsRegressionType do
            if (s <> '') and (REGRESSION_TYPE[rt] = s) then
            begin
              series.Regression.RegressionType := rt;
              break;
            end;

          s := GetAttrValue(AStyleNode, 'chart:regression-max-degree');
          if (s <> '') and TryStrToInt(s, intValue) then
            series.Regression.PolynomialDegree := intValue;

          s := GetAttrValue(AStyleNode, 'chart:regression-extrapolate-forward');
          if (s <> '') and TryStrToFloat(s, value, FPointSeparatorSettings) then
            series.Regression.ExtrapolateForwardBy := value
          else
            series.Regression.ExtrapolateForwardBy := 0.0;

          s := GetAttrValue(AStyleNode, 'chart:regression-extrapolate-backward');
          if (s <> '') and TryStrToFloat(s, value, FPointSeparatorSettings) then
            series.Regression.ExtrapolateBackwardBy := value
          else
            series.Regression.ExtrapolateBackwardBy := 0.0;

          s := GetAttrValue(AStyleNode, 'chart:regression-force-intercept');
          series.Regression.ForceYIntercept := (s = 'true');

          s := GetAttrValue(AStyleNode, 'chart:regression-intercept-value');
          if (s <> '') and TryStrToFloat(s, value, FPointSeparatorSettings) then
            series.Regression.YInterceptValue := value;
        end;
    end;
    AStyleNode := AStyleNode.NextSibling;
  end;
end;

procedure TsSpreadOpenDocChartReader.ReadChartSeriesDataPointStyle(AStyleNode: TDOMNode;
  AChart: TsChart; ASeries: TsChartSeries; var AFill: TsChartFill; var ALine: TsChartLine);
var
  nodeName: string;
  grNode: TDOMNode;
begin
  AFill := nil;
  ALine := nil;

  nodeName := AStyleNode.NodeName;
  AStyleNode := AStyleNode.FirstChild;
  while AStyleNode <> nil do
  begin
    nodeName := AStyleNode.NodeName;
    if nodeName = 'style:graphic-properties' then
    begin
      AFill := TsChartFill.Create;
      if not GetChartFillProps(AStyleNode, AChart, AFill) then FreeAndNil(AFill);
      ALine := TsChartLine.Create;
      if not GetChartLineProps(AStyleNode, AChart, ALine) then FreeAndNil(ALine);
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
  n: Integer;
begin
  s := GetAttrValue(ANode, 'chart:class');
  case s of
    'chart:area': series := TsAreaSeries.Create(AChart);
    'chart:bar': series := TsBarSeries.Create(AChart);
    'chart:bubble': series := TsBubbleSeries.Create(AChart);
    'chart:circle': series := TsPieSeries.Create(AChart);
    'chart:filled-radar': series := TsRadarSeries.Create(AChart);
    'chart:line': series := TsLineSeries.Create(AChart);
    'chart:radar': series := TsRadarSeries.Create(AChart);
    'chart:ring': series := TsRingSeries.Create(AChart);
    'chart:scatter': series := TsScatterSeries.Create(AChart);
    else raise Exception.Create('Unknown/unsupported series type.');
  end;

  ReadChartCellAddr(ANode, 'chart:label-cell-address', series.TitleAddr);
  if (series is TsBubbleSeries) then
    ReadChartCellRange(ANode, 'chart:values-cell-range-address', TsBubbleSeries(series).BubbleRange)
  else
    ReadChartCellRange(ANode, 'chart:values-cell-range-address', series.YRange);

  xyCounter := 0;
  subnode := ANode.FirstChild;
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
            series.YRange.Assign(series.XRange);
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
          s := GetAttrValue(subnode, 'chart:style-name');
          if s <> '' then
          begin
            styleNode := FindStyleNode(AStyleNode, s);
            ReadChartSeriesDataPointStyle(styleNode, AChart, series, fill, line); // creates fill and line!
          end;
          s := GetAttrValue(subnode, 'chart:repeated');
          if (s <> '') then
            n := StrToIntDef(s, 1);
          series.AddDataPointStyle(fill, line, n);
          fill.Free;  // the styles have been copied to the series datapoint list and are not needed any more.
          line.Free;
        end;
    end;
    subnode := subNode.NextSibling;
  end;

  if series.LabelRange.IsEmpty then series.LabelRange.Assign(AChart.XAxis.CategoryRange);

  s := GetAttrValue(ANode, 'chart:style-name');
  styleNode := FindStyleNode(AStyleNode, s);
  ReadChartSeriesStyle(styleNode, AChart, series);

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
  s := GetAttrValue(AStyleNode, 'style:data-style-name');
  if s <> '' then
    s := TsChartNumberFormatList(FNumberFormatList).FindFormatByName(s);
  ASeries.LabelFormat := s;

  AStyleNode := AStyleNode.FirstChild;
  while AStyleNode <> nil do begin
    nodeName := AStyleNode.NodeName;
    case nodeName of
      'style:graphic-properties':
        begin
          if ASeries.ChartType in [ctBar] then
            ASeries.Line.Style := clsSolid;
          GetChartLineProps(AStyleNode, AChart, ASeries.Line);
          GetChartFillProps(AStyleNode, AChart, ASeries.Fill);
        end;
      'style:text-properties':
        GetChartTextProps(AStyleNode, ASeries.LabelFont);
      'style:chart-properties':
        begin
          s := GetAttrValue(AStyleNode, 'chart:label-position');
          case s of
            'outside': ASeries.LabelPosition := lpOutside;
            'inside': ASeries.LabelPosition := lpInside;
            'center': ASeries.LabelPosition := lpCenter;
          end;

          s := GetAttrValue(AStyleNode, 'loext:label-stroke-color');
          if s <> '' then
            ASeries.LabelBorder.Color := HTMLColorStrToColor(s);
          s := GetAttrValue(AStyleNode, 'loext:label-stroke');
          if s <> '' then
            case s of
              'none': ASeries.LabelBorder.Style := clsNoLine;
              else    ASeries.LabelBorder.Style := clsSolid;
            end;

          s := GetAttrValue(AStyleNode, 'chart:data-label-number');
          if (s <> '') and (s <> 'none') then
            Include(datalabels, cdlValue);
          s := GetAttrValue(AStyleNode, 'chart:data-label-number="percentage"');
          if s <> '' then
            Include(datalabels, cdlPercentage);
          s := GetAttrValue(AStyleNode, 'chart:data-label-number="value-and-percentage"');
          if s <> '' then
            dataLabels := dataLabels + [cdlValue, cdlPercentage];
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
              TsCustomLineSeriesOpener(ASeries).ShowSymbols := true;
              for css in TsChartSeriesSymbol do
                if SYMBOL_NAMES[css] = s then
                begin
                  TsCustomLineSeriesOpener(ASeries).Symbol := css;
                  break;
                end;
              s := GetAttrValue(AStyleNode, 'symbol-width');
              if (s <> '') and EvalLengthStr(s, value, rel) then
                TsCustomLineSeriesOpener(ASeries).SymbolWidth := value;
              s := GetAttrValue(AStyleNode, 'symbol-height');
              if (s <> '') and EvalLengthStr(s, value, rel) then
                TsCustomLineSeriesOpener(ASeries).SymbolHeight := value;
            end else
              TsCustomLineSeriesOpener(ASeries).ShowSymbols := false;
          end;
        end;

    end;
    AStyleNode := AStyleNode.NextSibling;
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
  startColor: TsColor = scSilver;
  endColor: TsColor = scWhite;
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
    startColor := HTMLColorStrToColor(s);

  s := GetAttrValue(ANode, 'draw:end-color');
  if s <> '' then
    endColor := HTMLColorStrToColor(s);

  s := GetAttrValue(ANode, 'draw:start-intensity');
  if not TryPercentStrToFloat(s, startIntensity) then
    startIntensity := 1.0;

  s := GetAttrValue(ANode, 'draw:end-intensity');
  if not TryPercentStrToFloat(s, endIntensity) then
    endIntensity := 1.0;

  s := GetAttrValue(ANode, 'draw:border');
  if not TryPercentStrToFloat(s, border) then
    border := 0.0;

  s := GetAttrValue(ANode, 'draw:angle');
  if s <> '' then begin
    for i := Length(s) downto 1 do
      if not (s[i] in ['0'..'9', '.', '+', '-']) then Delete(s, i, 1);
    angle := StrToFloatDef(s, 0.0, FPointSeparatorSettings);
  end;

  s := GetAttrValue(ANode, 'draw:cx');
  if not TryPercentStrToFloat(s, centerX) then
    centerX := 0.0;

  s := GetAttrValue(ANode, 'draw:cy');
  if not TryPercentStrToFloat(s, centerY) then
    centerY := 0.0;

  AChart.Gradients.AddGradient(styleName, gradientStyle, startColor, endColor,
    startIntensity, endIntensity, border, centerX, centerY, angle);
end;

{ Read the hatch pattern stored in the "draw:hatch" nodes of the chart's
  Object styles.xml file. }
procedure TsSpreadOpenDocChartReader.ReadObjectHatchStyles(ANode: TDOMNode; AChart: TsChart);
var
  s: String;
  styleName: String;
  hs, hatchStyle: TsChartHatchStyle;
  hatchColor: TsColor = scBlack;
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
  if s <> '' then
    hatchColor := HTMLColorStrToColor(s);

  s := GetAttrValue(ANode, 'draw:distance');
  if not EvalLengthStr(s, hatchDist, rel) then
    hatchDist := 2.0;

  s := GetAttrValue(ANode, 'draw:rotation');
  if TryStrToFloat(s, hatchAngle, FPointSeparatorSettings) then
    hatchAngle := hatchAngle / 10
  else
    hatchAngle := 0;

  AChart.Hatches.AddHatch(styleName, hatchStyle, hatchColor, hatchDist, hatchAngle);
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
  angle: Integer;
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
  chartProps := chartProps + Format('style:rotation-angle="%d" ', [angle]);

  // Label orientation
  graphProps := 'svg:stroke-color="' + ColorToHTMLColorStr(Axis.AxisLine.Color) + '" ';

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
  rotAngle: Integer;
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
        if AChart.RotatedAxes then
        begin
          if rotAngle = 0 then rotAngle := 90 else if rotAngle = 90 then rotAngle := 0;
        end;
      end;
    else
      raise Exception.Create('[GetChartCaptionStyleAsXML] Unknown caption.');
  end;

  chartProps := 'chart:auto-position="true" ';
  chartProps := chartProps + Format('style:rotation-angle="%d" ', [rotAngle]);

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

function TsSpreadOpenDocChartWriter.GetChartFillStyleGraphicPropsAsXML(AChart: TsChart;
  AFill: TsChartFill): String;
var
  gradient: TsChartGradient;
  hatch: TsChartHatch;
  fillStr: String = '';
begin
  case AFill.Style of
    cfsNoFill:
      Result := 'draw:fill="none" ';
    cfsSolid:
      Result := Format(
        'draw:fill="solid" draw:fill-color="%s" ',
        [ ColorToHTMLColorStr(AFill.Color) ]
      );
    cfsGradient:
      begin
        gradient := AChart.Gradients[AFill.Gradient];
        Result := Format(
          'draw:fill="gradient" ' +
          'draw:fill-gradient-name="%s" ' +
          'draw:gradient-step-count="0" ',
          [ ASCIIName(gradient.Name) ]
        );
      end;
    cfsHatched, cfsSolidHatched:
      begin
        hatch := AChart.Hatches[AFill.Hatch];
        if AFill.Style = cfsSolidHatched then
          fillStr := 'draw:fill-hatch-solid="true" ';
        Result := Format(
          'draw:fill="hatch" draw:fill-color="%s" ' +
          'draw:fill-hatch-name="%s" %s',
          [ ColorToHTMLColorStr(AFill.Color), ASCIIName(hatch.Name), fillStr ]
        );
      end;
  end;
  if (AFill.Style <> cfsNoFill) and (AFill.Transparency > 0) then
    Result := Result + Format('draw:opacity="%.0f%%" ',
      [ (1.0 - AFill.Transparency) * 100 ],
      FPointSeparatorSettings
    );
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
  AChart: TsChart; ALine: TsChartLine): String;
var
  strokeStr: String = '';
  widthStr: String = '';
  colorStr: String = '';
  linestyle: TsChartLineStyle;
begin
  if ALine.Style = clsNoLine then
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
  colorStr := Format('svg:stroke-color="%s" ', [ColorToHTMLColorStr(ALine.Color)]);

  Result := strokeStr + widthStr + colorStr;
end;

function TsSpreadOpenDocChartWriter.GetChartPlotAreaStyleAsXML(AChart: TsChart;
  AIndent, AStyleID: Integer): String;
var
  indent: String;
  interpolationStr: String = '';
  verticalStr: String = '';
  stackModeStr: String = '';
  rightAngledAxes: String = '';
  startAngleStr: String = '';
begin
  indent := DupeString(' ', AIndent);

  if AChart.RotatedAxes then
    verticalStr := 'chart:vertical="true" ';

  case AChart.StackMode of
    csmSideBySide: ;
    csmStacked: stackModeStr := 'chart:stacked="true" ';
    csmStackedPercentage: stackModeStr := 'chart:percentage="true" ';
  end;

  if (AChart.Series.Count > 0) and (AChart.Series[0] is TsPieSeries) then
    startAngleStr := Format('chart:angle-offset="%d" ', [TsPieSeries(AChart.Series[0]).StartAngle]);

  case AChart.Interpolation of
    ciLinear: ;
    ciCubicSpline: interpolationStr := 'chart:interpolation="cubic-spline" ';
    ciBSpline: interpolationStr := 'chart:interpolation="b-spline" ';
    ciStepStart: interpolationStr := 'chart:interpolation="step-start" ';
    ciStepEnd: interpolationStr := 'chart:interpolation="step-end" ';
    ciStepCenterX: interpolationStr := 'chart:interpolation="step-center-x" ';
    ciStepCenterY: interpolationStr := 'chart:interpolation="step-center-y" ';
  end;

  if not (AChart.GetChartType in [ctRadar, ctPie]) then
    rightAngledAxes := 'chart:right-angled-axes="true" ';

  Result := Format(
    indent + '  <style:style style:name="ch%d" style:family="chart">', [ AStyleID ]) + LE +
    indent + '    <style:chart-properties ' +
                   interpolationStr +
                   verticalStr +
                   stackModeStr +
                   startAngleStr +
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
    AChart: TsChart; AEquation: TsRegressionEquation; AIndent, AStyleID: Integer): String;
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
  series: TsScatterSeries;
  indent: String;
  chartProps: String = '';
  graphProps: String = '';
begin
  Result := '';
  series := AChart.Series[ASeriesIndex] as TsScatterSeries;
  if series.Regression.RegressionType = rtNone then
    exit;

  indent := DupeString(' ', AIndent);

  chartprops := Format(
    'chart:regression-name="%s" ' +
    'chart:regression-type="%s" ' +
    'chart:regression-extrapolate-forward="%g" ' +
    'chart:regression-extrapolate-backward="%g" ' +
    'chart:regression-force-intercept="%s" ' +
    'chart:regression-intercept-value="%g" ' +
    'chart:regression-max-degree="%d" ',
    [ series.Regression.Title,
      REGRESSION_TYPE[series.Regression.RegressionType] ,
      series.Regression.ExtrapolateForwardBy,
      series.Regression.ExtrapolateBackwardBy,
      FALSE_TRUE[series.Regression.ForceYIntercept],
      series.Regression.YInterceptValue,
      series.Regression.PolynomialDegree
    ], FPointSeparatorSettings
  );

  graphprops := GetChartLineStyleGraphicPropsAsXML(AChart, series.Regression.Line);

  Result := Format(
    indent + '<style:style style:name="ch%d" style:family="chart"> ' + LE +
    indent + '  <style:chart-properties %s/>' + LE +
    indent + '  <style:graphic-properties %s/>' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, chartprops, graphprops ]
  );
end;

function TsSpreadOpenDocChartWriter.GetChartSeriesDataPointStyleAsXML(AChart: TsChart;
  ASeriesIndex, APointIndex, AIndent, AStyleID: Integer): String;
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
  dataPointStyle := TsChartDataPointStyle(series.DataPointStyles[APointIndex]);

  chartProps := 'chart:solid-type="cuboid" ';

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


{ <style:style style:name="ch1400" style:family="chart" style:data-style-name="N0">
    <style:chart-properties
      chart:symbol-type="named-symbol"
      chart:symbol-name="arrow-down"
      chart:symbol-width="0.25cm"
      chart:symbol-height="0.25cm"
      chart:link-data-style-to-source="true"/>
    <style:graphic-properties
      svg:stroke-width="0.08cm"
      svg:stroke-color="#ffd320"
      draw:fill-color="#ffd320"
      dr3d:edge-rounding="5%"/>
    <style:text-properties fo:font-size="10pt" style:font-size-asian="10pt" style:font-size-complex="10pt"/>
  </style:style> }
function TsSpreadOpenDocChartWriter.GetChartSeriesStyleAsXML(AChart: TsChart;
  ASeriesIndex, AIndent, AStyleID: Integer): String;
var
  series: TsChartSeries;
  lineser: TsLineSeries = nil;
  indent: String;
  numStyle: String;
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
    lineser := TsLineSeries(series);
    if lineser.ShowSymbols then
      chartProps := Format(
        'chart:symbol-type="named-symbol" chart:symbol-name="%s" chart:symbol-width="%.1fmm" chart:symbol-height="%.1fmm" ',
        [SYMBOL_NAMES[lineSer.Symbol], lineSer.SymbolWidth, lineSer.SymbolHeight ],
        FPointSeparatorSettings
      );
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
    if pos('\n', labelSeparator) > 0 then
      labelSeparator := StringReplace(labelSeparator, '\n', '<text:line-break/>', [rfReplaceAll, rfIgnoreCase]);
    labelSeparator :=
      indent + '    <chart:label-separator>' + LE +
      indent + '      <text:p>' + labelSeparator + '</text:p>' + LE +
      indent + '    </chart:label-separator>' + LE;
  end;

  if series.LabelBorder.Style <> clsNoLine then
  begin
    chartProps := chartProps + 'loext:label-stroke="solid" ';
    chartProps := chartProps + 'loext:label-stroke-color="' + ColorToHTMLColorStr(series.LabelBorder.Color) + '"'
  end;

  if labelSeparator <> '' then
    chartProps := indent + '  <style:chart-properties ' + chartProps + '>' + LE + labelSeparator + indent + '  </style:chart-properties>'
  else
    chartProps := indent + '  <style:chart-properties ' + chartProps + '/>';

  // Graphic properties
  lineProps := GetChartLineStyleGraphicPropsAsXML(AChart, series.Line);
  fillProps := GetChartFillStyleGraphicPropsAsXML(AChart, series.Fill);
  if (series is TsLineSeries) and (series.ChartType <> ctFilledRadar) then
  begin
    if lineSer.ShowSymbols then
      graphProps := graphProps + fillProps;
    if lineSer.ShowLines and (lineser.Line.Style <> clsNoLine) then
      graphProps := graphProps + lineProps
    else
      graphProps := graphProps + 'draw:stroke="none" ';
  end else
    graphProps := fillProps + lineProps;

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

function TsSpreadOpenDocChartWriter.GetNumberFormatID(ANumFormat: String): String;
var
  idx: Integer;
begin
  idx := FNumberFormatList.IndexOf(ANumFormat);
  if idx > -1 then
    Result := Format('N%d', [idx])
  else
    Result := 'N0';
end;

procedure TsSpreadOpenDocChartWriter.ListAllNumberFormats(AChart: TsChart);
var
  i: Integer;
  series: TsChartSeries;
  regression: TsChartRegression;
begin
  FNumberFormatList.Clear;

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
    if (series is TsScatterSeries) then begin
      regression := TsScatterSeries(series).Regression;
      if (regression.RegressionType <> rtNone) and
         (regression.DisplayEquation or regression.DisplayRSquare) then
      begin
        FNumberFormatList.Add(regression.Equation.NumberFormat);
      end;
    end;
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
  AChartStream, AStyleStream: TStream;
  AChartIndent, AStyleIndent: Integer; Axis: TsChartAxis;
  var AStyleID: Integer);
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

  if (Axis = chart.XAxis) and (not chart.IsScatterChart) and (chart.Series.Count > 0) then
  begin
    series := chart.Series[0];
    sheet := TsWorkbook(Writer.Workbook).GetWorksheetByIndex(chart.SheetIndex);
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
  AChartStream, AStyleStream: TStream;
  AChartIndent, AStyleIndent: Integer; AChart: TsChart; var AStyleID: Integer);
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
    indent + '    xlink:href=".." xlink:type="simple">' + LE, [
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
begin
  indent := DupeString(' ', AIndent);
  for i := 0 to AChart.Gradients.Count-1 do
  begin
    gradient := AChart.Gradients[i];
    style := indent + Format(
      '<draw:gradient draw:name="%s" draw:display-name="%s" ' +
        'draw:style="%s" ' +
        'draw:start-color="%s" draw:end-color="%s" ' +
        'draw:start-intensity="%.0f%%" draw:end-intensity="%.0f%%" ' +
        'draw:border="%.0f%%" ',
      [ ASCIIName(gradient.Name), gradient.Name,
        GRADIENT_STYLES[gradient.Style],
        ColorToHTMLColorStr(gradient.StartColor), ColorToHTMLColorStr(gradient.EndColor),
        gradient.StartIntensity * 100, gradient.EndIntensity * 100,
        gradient.Border * 100
      ]
    );
    case gradient.Style of
      cgsLinear, cgsAxial:
        style := style + Format(
          'draw:angle="%.0fdeg" ',
          [ gradient.Angle ],
          FPointSeparatorSettings
        );
      cgsElliptic, cgsSquare, cgsRectangular:
        style := style + Format(
          'draw:cx="%.0f%%" draw:cy="%.0f%%" draw:angle="%.0fdeg" ',
          [ gradient.CenterX * 100, gradient.CenterY * 100, gradient.Angle ],
          FPointSeparatorSettings
        );
      cgsRadial:
        style := style + Format(
          'draw:cx="%.0f%%" draw:cy="%.0f%%" ',
          [ gradient.CenterX * 100, gradient.CenterY * 100 ],
          FPointSeparatorSettings
        );
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
        ColorToHTMLColorStr(hatch.LineColor),
        hatch.LineDistance,
        hatch.LineAngle*10
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
procedure TsSpreadOpenDocChartWriter.WriteChartLegend(AChartStream, AStyleStream: TStream;
  AChartIndent, AStyleIndent: Integer; AChart: TsChart; var AStyleID: Integer);
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
    numFmtStr := FNumberFormatList[i];
    parser := TsSpreadOpenDocNumFormatParser.Create(numFmtStr, FWriter.Workbook.FormatSettings);
    try
      numFmtXML := parser.BuildXMLAsString(numFmtName);
      if numFmtXML <> '' then
        AppendToStream(AStream, indent + numFmtXML);
    finally
      parser.Free;
    end;
  end;
  {

  AppendToStream(AStream,
    indent + '<number:number-style style:name="N0">' + LE +
    indent + '  <number:number number:min-integer-digits="1"/>' + LE +
    indent + '</number:number-style>' + LE
  );

  if AChart.StackMode = csmStackedPercentage then
    AppendToStream(AStream,
      indent + '<number:percentage-style style:name="N10010">' + LE +
      indent + '  <number:number number:decimal-places="0" number:min-decimal-places="0" number:min-integer-digits="1"/>' + LE +
      indent + '  <number:text>%</number:text>' + LE +
      indent + '</number:percentage-style>' + LE
    );
    }
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

procedure TsSpreadOpenDocChartWriter.WriteChartPlotArea(AChartStream, AStyleStream: TStream;
  AChartIndent, AStyleIndent: Integer; AChart: TsChart; var AStyleID: Integer);
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
  WriteChartAxis(AChartStream, AStyleStream, AChartIndent+2, AStyleIndent, AChart.X2Axis, AStyleID);

  // secondary y axis
  WriteChartAxis(AChartStream, AStyleStream, AChartIndent+2, AStyleIndent, AChart.Y2Axis, AStyleID);

  // series
  for i := 0 to AChart.Series.Count-1 do
    WriteChartSeries(AChartStream, AStyleStream, AChartIndent+2, AStyleIndent, AChart, i, AStyleID);

  // close xml node
  AppendToStream(AChartStream,
    indent + '</chart:plot-area>' + LE
  );
end;

procedure TsSpreadOpenDocChartWriter.WriteChartSeries(
  AChartStream, AStyleStream: TStream;
  AChartIndent, AStyleIndent: Integer; AChart: TsChart; ASeriesIndex: Integer;
  var AStyleID: Integer);
var
  indent: String;
  series: TsChartSeries;
  valuesRange: String = '';
  domainRangeX: String = '';
  domainRangeY: String = '';
  fillColorRange: String = '';
  lineColorRange: String = '';
  chartClass: String = '';
  regressionEquation: String = '';
  needRegressionStyle: Boolean = false;
  needRegressionEquationStyle: Boolean = false;
  regression: TsChartRegression = nil;
  titleAddr: String;
  i, count: Integer;
  styleID, dpStyleID: Integer;
begin
  indent := DupeString(' ', AChartIndent);
  styleID := AStyleID;

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

  // And this is the title of the series for the legend
  titleAddr := GetSheetCellRangeString_ODS(
    series.TitleAddr.GetSheetName, series.TitleAddr.GetSheetName,
    series.TitleAddr.Row, series.TitleAddr.Col,
    series.TitleAddr.Row, series.TitleAddr.Col,
    rfAllRel, false
  );

  // Number of data points
  if series.YValuesInCol then
    count := series.YRange.Row2 - series.YRange.Row1 + 1
  else
    count := series.YRange.Col2 - series.YRange.Col1 + 1;

  if series is TsRingSeries then
    chartClass := 'circle'
  else
    chartClass := CHART_TYPE_NAMES[series.ChartType];

  // Store the series properties
  AppendToStream(AChartStream, Format(
    indent + '<chart:series chart:style-name="ch%d" ' +
               'chart:class="chart:%s" ' +                    // series type
               'chart:values-cell-range-address="%s" ' +      // y values
               'chart:label-cell-address="%s">' + LE,         // series title
    [ AStyleID, chartClass, valuesRange, titleAddr, chartClass ]
  ));
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

  // Regression
  if (series is TsScatterSeries) then
  begin
    regression := TsScatterSeries(series).Regression;
    if regression.RegressionType <> rtNone then
    begin
      if regression.DisplayEquation or regression.DisplayRSquare then
      begin
        if (not regression.Equation.DefaultXName) or (not regression.Equation.DefaultYName) or
           (not regression.Equation.DefaultBorder) or (not regression.Equation.DefaultFill) or
           (not regression.Equation.DefaultFont) or (not regression.Equation.DefaultNumberFormat) or
           (not regression.Equation.DefaultPosition) then
        begin
          regressionEquation := regressionEquation + Format('chart:style-name="ch%d" ', [AStyleID + 2]);
          needRegressionEquationStyle := true;
          styleID := AStyleID + 2;
        end;
      end;
      if regression.DisplayEquation then
        regressionEquation := regressionEquation + 'chart:display-equation="true" ';
      if regression.DisplayRSquare then
        regressionEquation := regressionEquation + 'chart:display-r-square="true" ';

      if regressionEquation <> '' then
      begin
        if not regression.Equation.DefaultPosition then
          regressionEquation := regressionEquation + Format(
            'svg:x="%.2fmm" svg:y="%.2fmm" ',
            [ regression.Equation.Left, regression.Equation.Top ],
            FPointSeparatorSettings
          );

        AppendToStream(AChartStream, Format(
          indent + '  <chart:regression-curve chart:style-name="ch%d">' + LE +
          indent + '    <chart:equation %s />' + LE +
          indent + '  </chart:regression-curve>' + LE,
          [ AStyleID + 1, regressionEquation ]
        ));
      end else
        AppendToStream(AChartStream, Format(
          indent + '  <chart:regression-curve chart:style-name="ch%d"/>',
          [ AStyleID + 1 ]
        ));
      needRegressionStyle := true;
      if styleID = AStyleID then
        styleID := AStyleID + 1;
    end;
  end;

  // Individual data point styles
  if series.DataPointStyles.Count = 0 then
    AppendToStream(AChartStream, Format(
      indent + '  <chart:data-point chart:repeated="%d"/>' + LE,
      [ count ]
    ))
  else
  begin
    dpStyleID := styleID + 1;
    for i := 0 to count - 1 do
    begin
      if (i >= series.DataPointStyles.Count) or (series.DataPointStyles[i] = nil) then
        AppendToStream(AChartStream,
          indent + '  <chart:data-point chart:repeated="1">' + LE
        )
      else
      begin
        AppendToStream(AChartStream, Format(
          indent + '  <chart:data-point chart:style-name="ch%d" />' + LE,   // ToDo: could contain "chart:repeated"
          [ dpStyleID ]
        ));
        inc(dpStyleID);
      end;
    end;
  end;
  AppendToStream(AChartStream,
    indent + '</chart:series>' + LE
  );

  // Series style
  AppendToStream(AStyleStream,
    GetChartSeriesStyleAsXML(AChart, ASeriesIndex, AStyleIndent, AStyleID)
  );

  // Regression style
  if needRegressionStyle then
  begin
    inc(AStyleID);
    AppendToStream(AStyleStream,
      GetChartRegressionStyleAsXML(AChart, ASeriesIndex, AStyleIndent, AStyleID)
    );

    // Style of regression equation
    if needRegressionEquationStyle then
    begin
      inc(AStyleID);
      AppendToStream(AStyleStream,
        GetChartRegressionEquationStyleAsXML(AChart, regression.Equation, AStyleIndent, AStyleID)
      );
    end;
  end;

  // Data point styles
  for i := 0 to series.DataPointStyles.Count - 1 do
  begin
    inc(AStyleID);
    AppendToStream(AStyleStream,
      GetChartSeriesDataPointStyleAsXML(AChart, ASeriesIndex, i, AStyleIndent, AStyleID)
    );
  end;

  // Next style
  inc(AStyleID);
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
procedure TsSpreadOpenDocChartWriter.WriteChartTitle(AChartStream, AStyleStream: TStream;
  AChartIndent, AStyleIndent: Integer; AChart: TsChart; IsSubtitle: Boolean;
  var AStyleID: Integer);
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

end.

