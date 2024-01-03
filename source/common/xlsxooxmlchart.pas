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
    procedure ReadChartLineProps(ANode: TDOMNode; AChart: TsChart; AChartLine: TsChartLine);
  protected
    procedure ReadChart(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartAxis(ANode: TDOMNode; AChart: TsChart; AChartAxis: TsChartAxis);
    function ReadChartAxisTickMarks(ANode: TDOMNode): TsChartAxisTicks;
    procedure ReadChartBarSeries(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartLegend(ANode: TDOMNode; AChartLegend: TsChartLegend);
    procedure ReadChartPlotArea(ANode: TDOMNode; AChart: TsChart);
    procedure ReadChartSeriesProps(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartSeriesRange(ANode: TDOMNode; ARange: TsChartRange);
    procedure ReadChartSeriesTitle(ANode: TDOMNode; ASeries: TsChartSeries);
    procedure ReadChartText(ANode: TDOMNode; AText: TsChartText);
    procedure ReadChartTitle(ANode: TDOMNode; AChartTitle: TsChartText);

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

procedure TsSpreadOOXMLChartReader.ReadChartAxis(ANode: TDOMNode;
  AChart: TsChart; AChartAxis: TsChartAxis);
var
  nodeName, s: String;
  n: Integer;
  node: TDOMNode;
begin
  if ANode = nil then
    exit;

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
        ReadChartText(ANode.FindNode('c:tx'), AChartAxis.Title);
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
        ;
      'c:varyColors':
        ;
      'c:ser':
        begin
          ser := TsBarSeries.Create(AChart);
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
    end;
    ANode := ANode.NextSibling;
  end;
end;


procedure TsSpreadOOXMLChartReader.ReadChartLegend(ANode: TDOMNode;
  AChartLegend: TsChartLegend);
var
  child: TDOMNode;
  nodeName, s: String;
begin
  if ANode = nil then
    exit;

  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
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
      'c:legendEntry':
        ;
      'c:overlay':
        begin
          s := GetAttrValue(ANode, 'val');
          AChartLegend.canOverlapPlotArea := (s = '1');
        end;
      'c:spPr':
        ReadChartFillAndLineProps(ANode.FirstChild, AChartLegend.Chart, AChartLegend.Background, AChartLegend.Border);
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
  n: Integer;
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
                      s := getAttrValue(child2, 'sp');
                      if TryStrToInt64(s, sp) then
                        AChartLine.Style := AChart.LineStyles.Add('', d, 1, (d+sp), 0, 0, false);
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
      'c:barChart':
        ReadChartBarSeries(ANode.FirstChild, AChart);
      'c:catAx':
        ReadChartAxis(ANode.FirstChild, AChart, AChart.XAxis);
      'c:valAx':
        ReadChartAxis(ANode.FirstChild, AChart, AChart.YAxis);
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
  while ANode <> nil do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:idx': ;
      'c:order': ;
      'c:tx':
        ReadChartSeriesTitle(ANode.FirstChild, ASeries);
      'c:cat':
        ReadChartSeriesRange(ANode.FirstChild, ASeries.XRange);
      'c:val':
        ReadChartSeriesRange(ANode.FirstChild, ASeries.YRange);
      'c:spPr':
        ReadChartFillAndLineProps(ANode.FirstChild, ASeries.Chart, ASeries.Fill, ASeries.Line);
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

procedure TsSpreadOOXMLChartReader.ReadChartTitle(ANode: TDOMNode; AChartTitle: TsChartText);
var
  nodeName: String;
  child, child2, child3, child4: TDOMNode;
  s: String;
begin
  if ANode = nil then
    exit;
  while Assigned(ANode) do
  begin
    nodeName := ANode.NodeName;
    case nodeName of
      'c:tx':
        begin
          child := ANode.Firstchild;
          while Assigned(child) do
          begin
            nodeName := child.NodeName;
            case nodeName of
              'c:rich':
                begin
                  child2 := child.FirstChild;
                  while Assigned(child2) do
                  begin
                    nodeName := child2.NodeName;
                    case nodeName of
                      'a:p':
                        begin
                          child3 := child2.FirstChild;
                          while Assigned(child3) do
                          begin
                            nodeName := child3.NodeName;
                            case nodeName of
                              'a:r':
                                begin
                                  child4 := child3.FirstChild;
                                  while Assigned(child4) do
                                  begin
                                    nodeName := child4.NodeName;
                                    case nodeName of
                                      'a:t':
                                        AChartTitle.Caption := GetNodeValue(child4);
                                    end;
                                    child4 := child4.NextSibling;
                                  end;
                                end;
                            end;
                            child3 := child3.NextSibling;
                          end;
                        end;
                    end;
                    child2 := child2.NextSibling;
                  end;
                end;
            end;
            child := child.NextSibling;
          end;
        end;
    end;
    ANode := ANode.NextSibling;
  end;
end;

procedure TsSpreadOOXMLChartReader.ReadChartText(ANode: TDOMNode;
  AText: TsChartText);
begin
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

