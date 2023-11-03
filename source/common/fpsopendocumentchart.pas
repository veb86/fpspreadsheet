unit fpsOpenDocumentChart;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, StrUtils,
 {$IF FPC_FULLVERSION >= 20701}
  zipper,
 {$ELSE}
  fpszipper,
 {$ENDIF}
  fpsTypes, fpSpreadsheet, fpsChart, fpsUtils, fpsReaderWriter, fpsXMLCommon;

type
  TsSpreadOpenDocChartWriter = class(TsBasicSpreadChartWriter)
  private
    FSCharts: array of TStream;
    FSObjectStyles: array of TStream;
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
    function GetChartSeriesStyleAsXML(AChart: TsChart; ASeriesIndex, AIndent, AStyleID: integer): String;
//    function GetChartTitleStyleAsXML(AChart: TsChart; AStyleIndex, AIndent: Integer): String;
    procedure PrepareChartTable(AChart: TsChart; AWorksheet: TsBasicWorksheet);

  protected
    procedure WriteChart(AStream: TStream; AChart: TsChart);
    procedure WriteChartAxis(AChartStream, AStyleStream: TStream;
      AChartIndent, AStyleIndent: Integer; Axis: TsChartAxis; var AStyleID: Integer);
    procedure WriteChartBackground(AChartStream, AStyleStream: TStream;
      AChartIndent, AStyleIndent: Integer; AChart: TsChart; var AStyleID: Integer);
    procedure WriteChartLegend(AChartStream, AStyleStream: TStream;
      AChartIndent, AStyleIndent: Integer; AChart: TsChart; var AStyleID: Integer);
    procedure WriteChartNumberStyles(AStream: TStream;
      AIndent: Integer; AChart: TsChart);
    procedure WriteObjectStyles(AStream: TStream; AChart: TsChart);
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

const
  OPENDOC_PATH_CHART_CONTENT = 'Object %d/content.xml';
  OPENDOC_PATH_CHART_STYLES  = 'Object %d/styles.xml';

  CHART_TYPE_NAMES: array[TsChartType] of string = (
    '', 'bar', 'line', 'area', 'barLine', 'scatter', 'bubble',
    'radar', 'filled-radar', 'circle', 'ring'
  );

  SYMBOL_NAMES: array[TsChartSeriesSymbol] of String = (
    'square', 'diamond', 'arrow-up', 'arrow-down', 'arrow-left',
    'arrow-right', 'circle', 'star', 'x', 'plus', 'asterisk'
  );  // unsupported: bow-tie, hourglass, horizontal-bar, vertical-bar

  LABEL_POSITION: array[TsChartLabelPosition] of string = (
    '', 'outside', 'inside', 'center');

  LEGEND_POSITION: array[TsChartLegendPosition] of string = (
    'end', 'top', 'bottom', 'start'
  );

  AXIS_ID: array[TAxisKind] of string = ('x', 'y', 'x', 'y');
  AXIS_LEVEL: array[TAxisKind] of string = ('primary', 'primary', 'secondary', 'secondary');

  FALSE_TRUE: array[boolean] of string = ('false', 'true');

  LE = LineEnding;

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


{------------------------------------------------------------------------------}
{                        TsSpreadOpenDocChartWriter                            }
{------------------------------------------------------------------------------}

constructor TsSpreadOpenDocChartWriter.Create(AWriter: TsBasicSpreadWriter);
begin
  inherited Create(AWriter);
  FPointSeparatorSettings := SysUtils.DefaultFormatSettings;
  FPointSeparatorSettings.DecimalSeparator:='.';
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

  // Special number format for numerical axis labels
  if (Axis = chart.YAxis) and (chart.StackMode = csmStackedPercentage) then
    numStyle := 'N10010';

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
    chartProps := chartProps + Format('chart:interval-minor-divisor="%d" ', [Axis.MinorSteps]);

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
        font := axis.CaptionFont;
        rotAngle := axis.CaptionRotation;
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
  fillStr: String;
  fillColorStr: String;
  fillOpacity: String = '';
begin
  if AFill.Style = fsNoFill then
  begin
    Result := 'draw:fill="none" ';
    exit;
  end;

  // To do: extend with hatched and gradient fills
  fillStr := 'draw:fill="solid" ';
  fillColorStr := 'draw:fill-color="' + ColorToHTMLColorStr(AFill.FgColor) + '" ';
  if AFill.Transparency > 0 then
    fillOpacity := Format('draw:opacity="%.0f%%" ', [(1.0 - AFill.Transparency)*100], FPointSeparatorSettings);

  Result := fillStr + fillColorStr + fillOpacity;
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
  s: String;
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
begin
  indent := DupeString(' ', AIndent);

  if AChart.RotatedAxes then
    verticalStr := 'chart:vertical="true" ';

  case AChart.StackMode of
    csmSideBySide: ;
    csmStacked: stackModeStr := 'chart:stacked="true" ';
    csmStackedPercentage: stackModeStr := 'chart:percentage="true" ';
  end;

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

  // TO DO: Create chart number style list and find the current style there!
  if not AEquation.DefaultNumberFormat then
    numStyle := 'N0';

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
const
  REGRESSION_TYPE: array [TsRegressionType] of string = (
    '', 'linear', 'logarithmic', 'exponential', 'power', 'polynomial');
var
  series: TsScatterSeries;
  indent: String;
  chartProps: String = '';
  graphProps: String = '';
  textProps: String = '';
  lineProps: String = '';
  fillProps: String = '';
  labelSeparator: String = '';
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

  // Chart properties
  chartProps := 'chart:symbol-type="none" ';
  if (series is TsLineSeries) and (series.ChartType <> ctFilledRadar) then
  begin
    lineser := TsLineSeries(series);
    if lineser.ShowSymbols then
      chartProps := Format(
        'chart:symbol-type="named-symbol" chart:symbol-name="%s" chart:symbol-width="%.1fmm" chart:symbol-height="%.1fmm" ',
        [SYMBOL_NAMES[lineSer.Symbol], lineSer.SymbolWidth, lineSer.SymbolHeight ],
        FPointSeparatorSettings
      );
  end;

  chartProps := chartProps + 'chart:link-data-style-to-source="true" ';
  // to do: link-data-style-to-source must go to "false" in case of a specific
  // numeric label format. But that requires a number-style in the Object/style.xml

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
      '    <chart:label-separator>' + LE +
      '      <text:p>' + labelSeparator + '</text:p>' + LE +
      '    </chart:label-separator>' + LE;
  end;

  if labelSeparator <> '' then
    chartProps := '  <style:chart-properties ' + chartProps + '>' + LE + labelSeparator + '  </style:chart-properties>'
  else
    chartProps := '  <style:chart-properties ' + chartProps + '/>';

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
    indent + '<style:style style:name="ch%d" style:family="chart" style:data-style-name="N0">' + LE +
    indent + chartProps + LE +
    indent + '  <style:graphic-properties %s/>' + LE +
    indent + '  <style:text-properties %s/>' + LE +
    indent + '</style:style>' + LE,
    [ AStyleID, graphProps, textProps ]
  );
end;

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
//  FChartStyleList.Clear;

  chartStream := TMemoryStream.Create;
  styleStream := TMemoryStream.Create;
  try
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
  if Axis.ShowCaption and (Axis.Caption <> '') then
  begin
    AppendToStream(AChartStream, Format(
      indent + '  <chart:title chart:style-name="ch%d">' + LE +
      indent + '    <text:p>%s</text:p>' + LE +
      indent + '  </chart:title>' + LE,
      [ AStyleID, Axis.Caption ]
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
begin
  indent := DupeString(' ', AIndent);

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
end;

{ Writes the file "Object N/styles.xml" (N = 1, 2, ...) which is needed by the
  charts since it defines the line dash patterns. }
procedure TsSpreadOpenDocChartWriter.WriteObjectStyles(AStream: TStream;
  AChart: TsChart);
const
  LENGTH_UNIT: array[boolean] of string = ('mm', '%'); // relative to line width
  DECS: array[boolean] of Integer = (1, 0);            // relative to line width
var
  i: Integer;
  linestyle: TsChartLineStyle;
  seg1, seg2: String;
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
      AppendToStream(AStream, Format(
        '    <draw:stroke-dash draw:name="%s" draw:display-name="%s" draw:style="round" draw:distance="%.*f%s" %s%s/>' + LE, [
        ASCIIName(linestyle.Name), linestyle.Name,
        DECS[linestyle.RelativeToLineWidth], linestyle.Distance, LENGTH_UNIT[linestyle.RelativeToLineWidth],
        seg1, seg2
      ]));
  end;

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
  sheet: TsWorksheet;
  series: TsChartSeries;
  valuesRange: String = '';
  domainRangeX: String = '';
  domainRangeY: String = '';
  fillColorRange: String = '';
  chartClass: String = '';
  regressionEquation: String = '';
  needRegressionStyle: Boolean = false;
  needRegressionEquationStyle: Boolean = false;
  regression: TsChartRegression = nil;
  titleAddr: String;
  count: Integer;
begin
  indent := DupeString(' ', AChartIndent);

  series := AChart.Series[ASeriesIndex];
  sheet := TsWorkbook(Writer.Workbook).GetWorksheetByIndex(AChart.sheetIndex);

  // These are the x values of a scatter or bubble plot.
  if (series is TsScatterSeries) or (series is TsBubbleSeries) then
  begin
    domainRangeX := GetSheetCellRangeString_ODS(
      sheet.Name, sheet.Name,
      series.XRange.Row1, series.XRange.Col1,
      series.XRange.Row2, series.XRange.Col2,
      rfAllRel, false
    );
  end;

  if series is TsBubbleSeries then
  begin
    domainRangeY := GetSheetCellRangeString_ODS(
      sheet.Name, sheet.Name,
      series.YRange.Row1, series.YRange.Col1,
      series.YRange.Row2, series.YRange.Col2,
      rfAllRel, false
    );
    // These are the bubble radii
    valuesRange := GetSheetCellRangeString_ODS(
      sheet.Name, sheet.Name,
      TsBubbleSeries(series).BubbleRange.Row1, TsBubbleSeries(series).BubbleRange.Col1,
      TsBubbleSeries(series).BubbleRange.Row2, TsBubbleSeries(series).BubbleRange.Col2,
      rfAllRel, false
    );
  end else
    // These are the y values of the non-bubble series
    valuesRange := GetSheetCellRangeString_ODS(
      sheet.Name, sheet.Name,
      series.YRange.Row1, series.YRange.Col1,
      series.YRange.Row2, series.YRange.Col2,
      rfAllRel, false
    );

  // Fill colors for bars, line series symbols, bubbles
  if (series.FillColorRange.Row1 <> series.FillColorRange.Row2) or
     (series.FillColorRange.Col1 <> series.FillColorRange.Col2)
  then
    fillColorRange := GetSheetCellRangeString_ODS(
      sheet.Name, sheet.Name,
      series.FillColorRange.Row1, series.FillColorRange.Col1,
      series.FillColorRange.Row2, series.FillColorRange.Col2,
      rfAllRel, false
    );

  // And these are the data point labels.
  titleAddr := GetSheetCellRangeString_ODS(
    sheet.Name, sheet.Name,
    series.TitleAddr.Row, series.TitleAddr.Col,
    series.TitleAddr.Row, series.TitleAddr.Col,
    rfAllRel, false
  );
  count := series.YRange.Row2 - series.YRange.Row1 + 1;

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
      indent + '<loext:propertry-mapping loext:property="FillColor" loext:cell-range-address="%s"/>' + LE,
      [ fillColorRange ]
    ));

  { --- not working...
  if borderColorRange <> '' then
    AppendToStream(AChartStream, Format(
      indent + '<loext:property-mapping loext:property="BorderColor" loext:cell-range-address="%s"/>' + LE,
      [ borderColorRange ]
    ));
  }

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
        ))
      end else
        AppendToStream(AChartStream, Format(
          indent + '  <chart:regression-curve chart:style-name="ch%d"/>',
          [ AStyleID + 1 ]
        ));
      needRegressionStyle := true;
    end;
  end;

  AppendToStream(AChartStream, Format(
    indent + '  <chart:data-point chart:repeated="%d"/>' + LE,
    [ count ]
  ));
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
   DO NOT DELETE THIS - IT WAS A PAINT TO GET THIS, AND MAYBE IT WILL BE NEEDED
   LATER.
   AT THE MOMENT THIS IS NOT NEEDED, IN FACT, IT IS EVEN DETRIMENTAL:
   WITH THIS CODE INCLUDED, SERIES FILL ARE IGNORED AND TITLES ARE NOT CORRECT.

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

