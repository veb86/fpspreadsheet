unit fpsStockSeries;

{$mode objfpc}{$H+}

interface

uses
  LCLVersion, Classes, SysUtils, Graphics, Math,
  TAChartUtils, TAMath, TAGeometry, TADrawUtils, TALegend,
  TACustomSource, TACustomSeries, TAMultiSeries;

type
  {$IF LCL_FullVersion >= 3990000}
  TStockSeries = class(TOpenHighLowCloseSeries);
  {$ELSE}
  TOHLCBrushKind = (obkCandleUp, obkCandleDown);
  TOHLCPenKind = (opkCandleUp, opkCandleDown, opkCandleLine, opkLineUp, opkLineDown);

  TOHLCBrush = class(TBrush)
  private
    const
      DEFAULT_COLORS: array[TOHLCBrushKind] of TColor = (clLime, clRed);
  private
    FBrushKind: TOHLCBrushKind;
    function IsColorStored: Boolean;
    procedure SetBrushKind(AValue: TOHLCBrushKind);
  public
    property BrushKind: TOHLCBrushKind read FBrushKind write SetBrushKind;
  published
    property Color stored IsColorStored;
  end;

  TOHLCPen = class(TPen)
  private
    const
      DEFAULT_COLORS: array[TOHLCPenKind] of TColor = (clGreen, clMaroon, clDefault, clLime, clRed);
  private
    FPenKind: TOHLCPenKind;
    function IsColorStored: Boolean;
    procedure SetPenKind(AValue: TOHLCPenKind);
  public
    property PenKind: TOHLCPenKind read FPenKind write SetPenKind;
  published
    property Color stored IsColorStored;
  end;

  TOHLCMode = (mOHLC, mCandleStick);
  TTickWidthStyle = (twsPercent, twsPercentMin);

  TStockSeries = class(TBasicPointSeries)
  private
    FPen: array[TOHLCPenKind] of TOHLCPen;
    FBrush: array[TOHLCBrushKind] of TOHLCBrush;
    FTickWidth: Integer;
    FTickWidthStyle: TTickWidthStyle;
    FYIndexClose: Integer;
    FYIndexHigh: Integer;
    FYIndexLow: Integer;
    FYIndexOpen: Integer;
    FMode: TOHLCMode;
    function GetBrush(AIndex: TOHLCBrushKind): TOHLCBrush;
    function GetPen(AIndex: TOHLCPenKind): TOHLCPen;
    procedure SetBrush(AIndex: TOHLCBrushKind; AValue: TOHLCBrush);
    procedure SetPen(AIndex: TOHLCPenKind; AValue: TOHLCPen);
    procedure SetOHLCMode(AValue: TOHLCMode);
    procedure SetTickWidth(AValue: Integer);
    procedure SetTickWidthStyle(AValue: TTickWidthStyle);
    procedure SetYIndexClose(AValue: Integer);
    procedure SetYIndexHigh(AValue: Integer);
    procedure SetYIndexLow(AValue: Integer);
    procedure SetYIndexOpen(AValue: Integer);
  protected
    function CalcTickWidth(AX: Double; AIndex: Integer): Double;
    procedure GetLegendItems(AItems: TChartLegendItems); override;
    function GetSeriesColor: TColor; override;
    {$IF LCL_FullVersion >= 2020000}
    class procedure GetXYCountNeeded(out AXCount, AYCount: Cardinal); override;
    procedure UpdateLabelDirectionReferenceLevel(AIndex, AYIndex: Integer;
      var ALevel: Double); override;
    function SkipMissingValues(AIndex: Integer): Boolean; override;
    {$IFEND}
    function ToolTargetDistance(const AParams: TNearestPointParams;
      AGraphPt: TDoublePoint; APointIdx, AXIdx, AYIdx: Integer): Integer; override;
  public
    procedure Assign(ASource: TPersistent); override;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  public
    function AddXOHLC(
      AX, AOpen, AHigh, ALow, AClose: Double;
      ALabel: String = ''; AColor: TColor = clTAColor): Integer; inline;
    procedure Draw(ADrawer: IChartDrawer); override;
    function Extent: TDoubleRect; override;
    function GetNearestPoint(const AParams: TNearestPointParams;
      out AResults: TNearestPointResults): Boolean; override;
  published
    property CandlestickDownBrush: TOHLCBrush index obkCandleDown read GetBrush write SetBrush;
    property CandlestickDownPen: TOHLCPen index opkCandleDown read GetPen write SetPen;
    property CandlestickLinePen: TOHLCPen index opkCandleLine read GetPen write SetPen;
    property CandlestickUpBrush: TOHLCBrush index obkCandleUp read GetBrush write SetBrush;
    property CandlestickUpPen: TOHLCPen index opkCandleUp read GetPen write Setpen;
    property DownLinePen: TOHLCPen index opkLineDown read GetPen write SetPen;
    property LinePen: TOHLCPen index opkLineUp read GetPen write SetPen;
    property Mode: TOHLCMode read FMode write SetOHLCMode default mOHLC;
    property TickWidth: integer
      read FTickWidth write SetTickWidth default DEF_OHLC_TICK_WIDTH;
    property TickWidthStyle: TTickWidthStyle
      read FTickWidthStyle write SetTickWidthStyle default twsPercent;
    property ToolTargets default [nptPoint, nptYList, nptCustom];
    property YIndexClose: integer
      read FYIndexClose write SetYIndexClose default DEF_YINDEX_CLOSE;
    property YIndexHigh: Integer
      read FYIndexHigh write SetYIndexHigh default DEF_YINDEX_HIGH;
    property YIndexLow: Integer
      read FYIndexLow write SetYIndexLow default DEF_YINDEX_LOW;
    property YIndexOpen: Integer
      read FYIndexOpen write SetYIndexOpen default DEF_YINDEX_OPEN;
  published
    property AxisIndexX;
    property AxisIndexY;
    property MarkPositions;
    property Marks;
    property Source;
  end;
{$ENDIF}

implementation

{$IF LCL_FullVersion < 3990000}

uses
  FPCanvas;

type
  TLegendItemOHLCLine = class(TLegendItemLine)
  strict private
    FMode: TOHLCMode;
    FCandleStickUpColor: TColor;
    FCandleStickDownColor: TColor;
  public
    constructor Create(ASeries: TStockSeries; const AText: String);
    procedure Draw(ADrawer: IChartDrawer; const ARect: TRect); override;
  end;

constructor TLegendItemOHLCLine.Create(ASeries: TStockSeries; const AText: String);
var
  pen: TFPCustomPen;
begin
  case ASeries.Mode of
    mOHLC        : pen := ASeries.LinePen;
    mCandleStick : pen := ASeries.CandleStickLinePen;
  end;
  inherited Create(pen, AText);
  FMode := ASeries.Mode;
  FCandlestickUpColor := ASeries.CandlestickUpBrush.Color;
  FCandlestickDownColor := ASeries.CandlestickDownBrush.Color;
end;

procedure TLegendItemOHLCLine.Draw(ADrawer: IChartDrawer; const ARect: TRect);
const
  TICK_LENGTH = 3;
var
  dx, dy, x, y: Integer;
  pts: array[0..3] of TPoint;
begin
  inherited Draw(ADrawer, ARect);
  y := (ARect.Top + ARect.Bottom) div 2;
  dx := (ARect.Right - ARect.Left) div 3;
  x := ARect.Left + dx;
  case FMode of
    mOHLC:
      begin
        dy := ADrawer.Scale(TICK_LENGTH);
        ADrawer.Line(x, y, x, y + dy);
        x += dx;
        ADrawer.Line(x, y, x, y - dy);
      end;
    mCandlestick:
      begin
        dy := (ARect.Bottom - ARect.Top) div 4;
        pts[0] := Point(x, y-dy);
        pts[1] := Point(x, y+dy);
        pts[2] := Point(x+dx, y+dy);
        pts[3] := pts[0];
        ADrawer.SetBrushParams(bsSolid, FCandlestickUpColor);
        ADrawer.Polygon(pts, 0, 4);
        pts[0] := Point(x+dx, y+dy);
        pts[1] := Point(x+dx, y-dy);
        pts[2] := Point(x, y-dy);
        pts[3] := pts[0];
        ADrawer.SetBrushParams(bsSolid, FCandlestickDownColor);
        ADrawer.Polygon(pts, 0, 4);
      end;
  end;
end;

{ TOHLCBrush }

function TOHLCBrush.IsColorStored: Boolean;
begin
  Result := (Color = DEFAULT_COLORS[FBrushKind]);
end;

procedure TOHLCBrush.SetBrushKind(AValue: TOHLCBrushKind);
begin
  FBrushKind := AValue;
  Color := DEFAULT_COLORS[FBrushKind];
end;

{ TOHLCPen }

function TOHLCPen.IsColorStored: Boolean;
begin
  Result := (Color = DEFAULT_COLORS[FPenKind]);
end;

procedure TOHLCPen.SetPenKind(AValue: TOHLCPenKind);
begin
  FPenKind := AValue;
  Color := DEFAULT_COLORS[FPenKind];
end;

{ TStockSeries }

function TStockSeries.AddXOHLC(
  AX, AOpen, AHigh, ALow, AClose: Double;
  ALabel: String; AColor: TColor): Integer;
var
  y: Double;
begin
  if YIndexOpen = 0 then
    y := AOpen
  else if YIndexHigh = 0 then
    y := AHigh
  else if YIndexLow = 0 then
    y := ALow
  else if YIndexClose = 0 then
    y := AClose
  else
    raise Exception.Create('TOpenHighLowCloseSeries: Ordinary y value missing');

  Result := ListSource.Add(AX, y, ALabel, AColor);
  with ListSource.Item[Result]^ do begin
    SetY(YIndexOpen, AOpen);
    SetY(YIndexHigh, AHigh);
    SetY(YIndexLow, ALow);
    SetY(YIndexClose, AClose);
  end;
end;

procedure TStockSeries.Assign(ASource: TPersistent);
var
  bk: TOHLCBrushKind;
  pk: TOHLCPenKind;
begin
  if ASource is TStockSeries then
    with TOpenHighLowCloseSeries(ASource) do begin
      for bk in TOHLCBrushKind do
        Self.FBrush[bk] := FBrush[bk];
      for pk in TOHLCPenKind do
        Self.FPen[pk] := FPen[pk];
      Self.FMode := FMode;
      Self.FTickWidth := FTickWidth;
      Self.FYIndexClose := FYIndexClose;
      Self.FYIndexHigh := FYIndexHigh;
      Self.FYIndexLow := FYIndexLow;
      Self.FYIndexOpen := FYIndexOpen;
    end;
  inherited Assign(ASource);
end;

constructor TStockSeries.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);

  ToolTargets := [nptPoint, nptYList, nptCustom];
  FOptimizeX := false;
  FStacked := false;
  FTickWidth := DEF_OHLC_TICK_WIDTH;
  FYIndexClose := DEF_YINDEX_CLOSE;
  FYIndexHigh := DEF_YINDEX_HIGH;
  FYIndexLow := DEF_YINDEX_LOW;
  FYIndexOpen := DEF_YINDEX_OPEN;

  // Candlestick up brush
  FBrush[obkCandleUp] := TOHLCBrush.Create;
  FBrush[obkCandleUp].BrushKind := obkCandleUp;
  FBrush[obkCandleUp].OnChange := @StyleChanged;
  // Candlestick down brush
  FBrush[obkCandleDown] := TOHLCBrush.Create;
  FBrush[obkCandleDown].BrushKind := obkCandleDown;
  FBrush[obkCandleDown].OnChange := @StyleChanged;
  // Candlestick up border pen
  FPen[opkCandleUp] := TOHLCPen.Create;
  FPen[opkCandleUp].PenKind := opkCandleUp;
  FPen[opkCandleUp].OnChange := @StyleChanged;
  // Candlestick down border pen
  FPen[opkCandleDown] := TOHLCPen.Create;
  FPen[opkCandleDown].PenKind := opkCandleDown;
  FPen[opkCandleDown].OnChange := @StyleChanged;
  // Candlestick range pen
  FPen[opkCandleLine] := TOHLCPen.Create;
  FPen[opkCandleLine].PenKind := opkCandleLine;
  FPen[opkCandleLine].OnChange := @StyleChanged;
  // OHLC up pen
  FPen[opkLineUp] := TOHLCPen.Create;
  FPen[opkLineUp].PenKind := opkLineUp;
  FPen[opkLineUp].OnChange := @StyleChanged;
  // OHLC down pen
  FPen[opkLineDown] := TOHLCPen.Create;
  FPen[opkLineDown].PenKind := opkLineDown;
  FPen[opkLineDown].OnChange := @StyleChanged;
end;

destructor TStockSeries.Destroy;
var
  bk: TOHLCBrushKind;
  pk: TOHLCPenKind;
begin
  for bk in TOHLCBrushKind do
    FreeAndNil(FBrush[bk]);
  for pk in TOHLCPenKind do
    FreeAndNil(FPen[pk]);
  inherited;
end;

function TStockSeries.CalcTickWidth(AX: Double; AIndex: Integer): Double;
begin
  case FTickWidthStyle of
    twsPercent:
      Result := GetXRange(AX, AIndex) * PERCENT * TickWidth;
    twsPercentMin:
      begin
        if FMinXRange = 0 then
          UpdateMinXRange;
        Result := FMinXRange * PERCENT * TickWidth;
      end;
  end;
end;

procedure TStockSeries.Draw(ADrawer: IChartDrawer);

  function MaybeRotate(AX, AY: Double): TPoint;
  begin
    if IsRotated then
      Exchange(AX, AY);
    Result := ParentChart.GraphToImage(DoublePoint(AX, AY));
  end;

  procedure DoLine(AX1, AY1, AX2, AY2: Double);
  begin
    ADrawer.Line(MaybeRotate(AX1, AY1), MaybeRotate(AX2, AY2));
  end;

  procedure NoZeroRect(var R: TRect);
  begin
    if IsRotated then
    begin
      if R.Left = R.Right then inc(R.Right);
    end else
    begin
      if R.Top = R.Bottom then inc(R.Bottom);
    end;
  end;

  procedure DoRect(AX1, AY1, AX2, AY2: Double);
  var
    r: TRect;
  begin
    r.TopLeft := MaybeRotate(AX1, AY1);
    r.BottomRight := MaybeRotate(AX2, AY2);
    NoZeroRect(r);
    ADrawer.FillRect(r.Left, r.Top, r.Right, r.Bottom);
    ADrawer.Rectangle(r);
  end;

  procedure DrawOHLC(x, yopen, yhigh, ylow, yclose, tw: Double);
  begin
    DoLine(x, yhigh, x, ylow);
    DoLine(x, yclose, x + tw, yclose);
    if not IsNaN(yopen) then
      DoLine(x - tw, yopen, x, yopen);
  end;

  procedure DrawCandleStick(x, yopen, yhigh, ylow, yclose, tw: Double; APenIdx: Integer);
  begin
    if CandleStickLinePen.Color = clDefault then
      // use linepen and linedown pen for range line
      ADrawer.Pen := FPen[TOHLCPenKind(APenIdx + 3)]
    else
      ADrawer.Pen := CandleStickLinePen;
    DoLine(x, yhigh, x, ylow);
    ADrawer.Pen := FPen[TOHLCPenKind(APenIdx)];
    DoRect(x - tw, yopen, x + tw, yclose);
  end;

const
  UP_INDEX = 0;
  DOWN_INDEX = 1;
var
  my: Cardinal;
  ext2: TDoubleRect;
  i: Integer;
  x, tw, yopen, yhigh, ylow, yclose, prevclose: Double;
  idx: Integer;
  nx, ny: Cardinal;
begin
  if IsEmpty or (not Active) then exit;
  my := MaxIntValue([YIndexOpen, YIndexHigh, YIndexLow, YIndexClose]);
  if my >= Source.YCount then exit;

  ext2 := ParentChart.CurrentExtent;
  ExpandRange(ext2.a.X, ext2.b.X, 1.0);
  ExpandRange(ext2.a.Y, ext2.b.Y, 1.0);

  PrepareGraphPoints(ext2, true);

  prevclose := -Infinity;
  for i := FLoBound to FUpBound do begin
    x := GetGraphPointX(i);
    if IsNaN(x) then Continue;
    yopen := GetGraphPointY(i, YIndexOpen);
    if IsNaN(yopen) and (FMode = mCandleStick) then Continue;
    yhigh := GetGraphPointY(i, YIndexHigh);
    if IsNaN(yhigh) then Continue;
    ylow := GetGraphPointY(i, YIndexLow);
    if IsNaN(ylow) then Continue;
    yclose := GetGraphPointY(i, YIndexClose);
    if IsNaN(yclose) then Continue;
    tw := CalcTickWidth(x, i);

    if IsNaN(yopen) then
    begin
      // HLC chart: compare with close value of previous data point
      if prevclose < yclose then
        idx := UP_INDEX
      else
        idx := DOWN_INDEX;
    end else
    if (yopen <= yclose) then
      idx := UP_INDEX
    else
      idx := DOWN_INDEX;
    ADrawer.Brush := FBrush[TOHLCBrushKind(idx)];
    case FMode of
      mOHLC: ADrawer.Pen := FPen[TOHLCPenKind(idx + 3)];
      mCandlestick: ADrawer.Pen := FPen[TOHLCPenKind(idx)];
    end;
    if Source[i]^.Color <> clTAColor then
    begin
      ADrawer.SetPenParams(FPen[TOHLCPenKind(idx)].Style, Source[i]^.Color {$IF LCL_FUllVersion >= 2020000}, FPen[TOHLCPenKind(idx)].Width{$IFEND});
      ADrawer.SetBrushParams(FBrush[TOHLCBrushKind(idx)].Style, Source[i]^.Color);
    end;

    case FMode of
      mOHLC: DrawOHLC(x, yopen, yhigh, ylow, yclose, tw);
      mCandleStick: DrawCandleStick(x, yopen, yhigh, ylow, yclose, tw, idx);
    end;

    prevclose := yclose;
  end;

  {$IF LCL_FullVersion >= 2020000}
  GetXYCountNeeded(nx, ny);
  if Source.YCount > ny then
    for i := 0 to ny-1 do DrawLabels(ADrawer, i)
  else
  {$ENDIF}
    DrawLabels(ADrawer);
end;

function TStockSeries.Extent: TDoubleRect;
var
  x: Double;
  tw: Double;
  j: Integer;
begin
  Result := Source.ExtentList;                            // axis units

  // Enforce recalculation of tick/candlebox width
  FMinXRange := 0;

  // Show first and last open/close ticks and candle boxes fully.
  j := -1;
  x := NaN;
  while IsNaN(x) and (j < Source.Count-1) do begin
    inc(j);
    x := GetGraphPointX(j);                                 // graph units
  end;
  tw := CalcTickWidth(x, j);
  Result.a.X := Min(Result.a.X, GraphToAxisX(x - tw));    // axis units
//  Result.a.X := Min(Result.a.X, x - tw);
  j := Count;
  x := NaN;
  While IsNaN(x) and (j > 0) do begin
    dec(j);
    x := GetGraphPointX(j);
  end;
  tw := CalcTickWidth(x, j);
  Result.b.X := Max(Result.b.X, AxisToGraphX(x + tw));
//  Result.b.X := Max(Result.b.X, x + tw);
end;

function TStockSeries.GetBrush(AIndex: TOHLCBrushKind): TOHLCBrush;
begin
  Result := FBrush[AIndex];
end;

procedure TStockSeries.GetLegendItems(AItems: TChartLegendItems);
begin
  AItems.Add(TLegendItemOHLCLine.Create(Self, LegendTextSingle));
end;

function TStockSeries.GetNearestPoint(const AParams: TNearestPointParams;
  out AResults: TNearestPointResults): Boolean;
var
  i: Integer;
  graphClickPt, p: TDoublePoint;
  pImg: TPoint;
  x, yopen, yhigh, ylow, yclose, tw: Double;
  xImg, dist: Integer;
  R: TDoubleRect;
begin
  Result := inherited;

  if Result then begin
    if (nptPoint in AParams.FTargets) and (nptPoint in ToolTargets) then
      exit;
    if (nptYList in AParams.FTargets) and (nptYList in ToolTargets) then
      exit;
  end;
  if not ((nptCustom in AParams.FTargets) and (nptCustom in ToolTargets))
  then
    exit;

  graphClickPt := ParentChart.ImageToGraph(AParams.FPoint);
  pImg := AParams.FPoint;
  if IsRotated then begin
//    Exchange(pImg.X, pImg.Y);
    Exchange(graphclickpt.X, graphclickpt.Y);
    pImg := ParentChart.GraphToImage(graphClickPt);
  end;

  // Iterate through all points of the series
  for i := 0 to Count - 1 do begin
    x := GetGraphPointX(i);
    yopen := GetGraphPointY(i, YIndexOpen);
    yhigh := GetGraphPointY(i, YIndexHigh);
    ylow := GetGraphPointY(i, YIndexLow);
    yclose := GetGraphPointY(i, YIndexClose);
    tw := CalcTickWidth(x, i);

    dist := MaxInt;

    // click on vertical line
    if InRange(graphClickPt.Y, ylow, yhigh) then begin
      xImg := ParentChart.XGraphToImage(x);
      dist := sqr(pImg.X - xImg);
      AResults.FYIndex := -1;
    end;

    // click on candle box
    if FMode = mCandlestick then begin
      R.a := DoublePoint(x - tw, Min(yopen, yclose));
      R.b := DoublePoint(x + tw, Max(yopen, yclose));
      if InRange(graphClickPt.X, R.a.x, R.b.x) and InRange(graphClickPt.Y, R.a.Y, R.b.Y) then
      begin
        dist := 0;
        AResults.FYIndex := -1;
      end;
    end;

    // Sufficiently close?
    if dist < AResults.FDist then begin
      AResults.FDist := dist;
      AResults.FIndex := i;
      p := DoublePoint(x, yclose);   // "Close" value
      AResults.FValue := p;
      if IsRotated then Exchange(p.X, p.Y);
      AResults.FImg := ParentChart.GraphToImage(p);
      if dist = 0 then break;
    end;
  end;
  Result := AResults.FIndex > -1;
end;

function TStockSeries.GetPen(AIndex: TOHLCPenKind): TOHLCPen;
begin
  Result := FPen[AIndex];
end;

function TStockSeries.GetSeriesColor: TColor;
begin
  Result := LinePen.Color;
end;

{$IF LCL_FullVersion >= 2020000}
class procedure TStockSeries.GetXYCountNeeded(out AXCount, AYCount: Cardinal);
begin
  AXCount := 0;
  AYCount := 4;
end;
{$IFEND}

procedure TStockSeries.SetBrush(AIndex: TOHLCBrushKind; AValue: TOHLCBrush);
begin
  if GetBrush(AIndex) = AValue then exit;
  FBrush[AIndex].Assign(AValue);
  UpdateParentChart;
end;

procedure TStockSeries.SetPen(AIndex: TOHLCPenKind; AValue: TOHLCPen);
begin
  if GetPen(AIndex) = AValue then exit;
  FPen[AIndex].Assign(AValue);
  UpdateParentChart;
end;

procedure TStockSeries.SetOHLCMode(AValue: TOHLCMode);
begin
  if FMode = AValue then exit;
  FMode := AValue;
  UpdateParentChart;
end;

procedure TStockSeries.SetTickWidth(AValue: Integer);
begin
  if FTickWidth = AValue then exit;
  FTickWidth := AValue;
  UpdateParentChart;
end;

procedure TStockSeries.SetTickWidthStyle(AValue: TTickWidthStyle);
begin
  if FTickWidthStyle = AValue then exit;
  FTickWidthStyle := AValue;
  UpdateParentChart;
end;

procedure TStockSeries.SetYIndexClose(AValue: Integer);
begin
  if FYIndexClose = AValue then exit;
  FYIndexClose := AValue;
  UpdateParentChart;
end;

procedure TStockSeries.SetYIndexHigh(AValue: Integer);
begin
  if FYIndexHigh = AValue then exit;
  FYIndexHigh := AValue;
  UpdateParentChart;
end;

procedure TStockSeries.SetYIndexLow(AValue: Integer);
begin
  if FYIndexLow = AValue then exit;
  FYIndexLow := AValue;
  UpdateParentChart;
end;

procedure TStockSeries.SetYIndexOpen(AValue: Integer);
begin
  if FYIndexOpen = AValue then exit;
  FYIndexOpen := AValue;
  UpdateParentChart;
end;

{$IF LCL_FullVersion >= 2020000}
function TStockSeries.SkipMissingValues(AIndex: Integer): Boolean;
begin
  Result := IsNaN(Source[AIndex]^.Point);
  if not Result then
    Result := HasMissingYValue(AIndex, 4);
end;
{$IFEND}

function TStockSeries.ToolTargetDistance(
  const AParams: TNearestPointParams; AGraphPt: TDoublePoint;
  APointIdx, AXIdx, AYIdx: Integer): Integer;

  // All in image coordinates transformed to have a horizontal x axis
  function DistanceToLine(Pt: TPoint; x1, x2, y: Integer): Integer;
  begin
    if InRange(Pt.X, x1, x2) then     // FDistFunc does not calculate sqrt
      Result := sqr(Pt.Y - y)
    else
      Result := Min(
        AParams.FDistFunc(Pt, Point(x1, y)),
        AParams.FDistFunc(Pt, Point(x2, y))
      );
  end;

var
  x1, x2: Integer;
  w: Double;
  p, clickPt: TPoint;
  gp: TDoublePoint;
begin
  Unused(AXIdx);

  // Convert the "clicked" and "test" point to non-rotated axes
  if IsRotated then begin
    gp := ParentChart.ImageToGraph(AParams.FPoint);
    Exchange(gp.X, gp.Y);
    clickPt := ParentChart.GraphToImage(gp);
    Exchange(AGraphPt.X, AGraphPt.Y);
  end else
    clickPt := AParams.FPoint;

  w := CalcTickWidth(AGraphPt.X, APointIdx);
  x1 := ParentChart.XGraphToImage(AGraphPt.X - w);
  x2 := ParentChart.XGraphToImage(AGraphPt.X + w);
  p := ParentChart.GraphToImage(AGraphPt);

  case FMode of
    mOHLC:
      with ParentChart do
        if (AYIdx = YIndexOpen) then
          Result := DistanceToLine(clickPt, x1, p.x, p.y)
        else if (AYIdx = YIndexClose) then
          Result := DistanceToLine(clickPt, p.x, x2, p.y)
        else if (AYIdx = YIndexHigh) or (AYIdx = YIndexLow) then
          Result := AParams.FDistFunc(clickPt, p)
        else
          raise Exception.Create('TOpenHighLowCloseSeries.ToolTargetDistance: Illegal YIndex.');
    mCandleStick:
      with ParentChart do
        if (AYIdx = YIndexOpen) or (AYIdx = YIndexClose) then
          Result := DistanceToLine(clickPt, x1, x2, p.y)
        else if (AYIdx = YIndexHigh) or (AYIdx = YIndexLow) then
          Result := AParams.FDistFunc(clickPt, p)
        else
          raise Exception.Create('TOpenHighLowCloseSeries.ToolTargetDistance: Illegal YIndex.');
  end;
end;

{$IF LCL_FullVersion >= 2020000}
procedure TStockSeries.UpdateLabelDirectionReferenceLevel(
  AIndex, AYIndex: Integer; var ALevel: Double);
var
  item: PChartDataItem;
begin
  if AYIndex = FYIndexLow then
    ALevel := +Infinity
  else if AYIndex = FYIndexHigh then
    ALevel := -Infinity
  else begin
    item := Source.Item[AIndex];
    ALevel := (AxisToGraphY(item^.GetY(FYIndexLow)) + AxisToGraphY(item^.GetY(FYIndexHigh)))*0.5;
  end;
end;
{$ENDIF}

{$ENDIF}

end.

