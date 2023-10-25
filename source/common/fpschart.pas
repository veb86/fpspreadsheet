unit fpsChart;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

interface

uses
  Classes, SysUtils, Contnrs, fpsTypes, fpsUtils;

const
  clsNoLine = -2;
  clsSolid = -1;

{@@ Pre-defined chart line styles given as indexes into the chart's LineStyles
  list. Get their value in the constructor of TsChart. Default here to -1
  which is the code for a solid line, just in case that something goes wrong }
var
  clsFineDot: Integer = -1;
  clsDot: Integer = -1;
  clsDash: Integer = -1;
  clsDashDot: Integer = -1;
  clsLongDash: Integer = -1;
  clsLongDashDot: Integer = -1;
  clsLongDashDotDot: Integer = -1;

const
  DEFAULT_CHART_LINEWIDTH = 0.75;  // pts
  DEFAULT_CHART_FONT = 'Arial';

  DEFAULT_SERIES_COLORS: array[0..7] of TsColor = (
    scRed, scBlue, scGreen, scMagenta, scPurple, scTeal, scBlack, scGray
  );

type
  TsChart = class;

  TsChartLine = class
    Style: Integer;        // index into chart's LineStyle list or predefined clsSolid/clsNoLine
    Width: Double;         // mm
    Color: TsColor;        // in hex: $00bbggrr, r=red, g=green, b=blue
    Transparency: Double;  // in percent
  end;

  TsChartFill = class
    Style: TsFillStyle;
    FgColor: TsColor;
    BgColor: TsColor;
  end;

  TsChartLineSegment = record
    Length: Double;       // mm or % of linewidth
    Count: Integer;
  end;

  TsChartLineStyle = class
    Name: String;
    Segment1: TsChartLineSegment;
    Segment2: TsChartLineSegment;
    Distance: Double;     // mm or % of linewidth
    RelativeToLineWidth: Boolean;
    function GetID: String;
  end;

  TsChartLineStyleList = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartLineStyle;
    procedure SetItem(AIndex: Integer; AValue: TsChartLineStyle);
  public
    function Add(AName: String;
      ASeg1Length: Double; ASeg1Count: Integer;
      ASeg2Length: Double; ASeg2Count: Integer;
      ADistance: Double; ARelativeToLineWidth: Boolean): Integer;
    property Items[AIndex: Integer]: TsChartLineStyle read GetItem write SetItem; default;
  end;

  TsChartElement = class
  private
    FChart: TsChart;
    FVisible: Boolean;
  public
    constructor Create(AChart: TsChart);
    property Chart: TsChart read FChart;
    property Visible: Boolean read FVisible write FVisible;
  end;

  TsChartFillElement = class(TsChartElement)
  private
    FBackground: TsChartFill;
    FBorder: TsChartLine;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    property Background: TsChartFill read FBackground write FBackground;
    property Border: TsChartLine read FBorder write FBorder;
  end;

  TsChartText = class(TsChartFillElement)
  private
    FCaption: String;
    FShowCaption: Boolean;
    FFont: TsFont;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    property Caption: String read FCaption write FCaption;
    property Font: TsFont read FFont write FFont;
    property ShowCaption: Boolean read FShowCaption write FShowCaption;
  end;

  TsChartAxisPosition = (capStart, capEnd, capValue);
  TsChartType = (ctEmpty, ctBar, ctLine, ctArea, ctBarLine, ctScatter);

  TsChartAxis = class(TsChartFillElement)
  private
    FAutomaticMax: Boolean;
    FAutomaticMin: Boolean;
    FAutomaticMajorInterval: Boolean;
    FAutomaticMinorSteps: Boolean;
    FAxisLine: TsChartLine;
    FMajorGridLines: TsChartLine;
    FMinorGridLines: TsChartline;
    FInverted: Boolean;
    FCaption: String;
    FCaptionFont: TsFont;
    FCaptionRotation: Integer;
    FLabelFont: TsFont;
    FLabelFormat: String;
    FLabelRotation: Integer;
    FLogarithmic: Boolean;
    FMajorInterval: Double;
    FMajorTickLines: TsChartLine;
    FMax: Double;
    FMin: Double;
    FMinorSteps: Double;
    FMinorTickLines: TsChartLine;
    FPosition: TsChartAxisPosition;
    FPositionValue: Double;
    FShowCaption: Boolean;
    FShowLabels: Boolean;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    property AutomaticMax: Boolean read FAutomaticMax write FAutomaticMax;
    property AutomaticMin: Boolean read FAutomaticMin write FAutomaticMin;
    property AutomaticMajorInterval: Boolean read FAutomaticMajorInterval write FAutomaticMajorInterval;
    property AutomaticMinorSteps: Boolean read FAutomaticMinorSteps write FAutomaticMinorSteps;
    property AxisLine: TsChartLine read FAxisLine write FAxisLine;
    property Caption: String read FCaption write FCaption;
    property CaptionFont: TsFont read FCaptionFont write FCaptionFont;
    property CaptionRotation: Integer read FCaptionRotation write FCaptionRotation;
    property Inverted: Boolean read FInverted write FInverted;
    property LabelFont: TsFont read FLabelFont write FLabelFont;
    property LabelFormat: String read FLabelFormat write FLabelFormat;
    property LabelRotation: Integer read FLabelRotation write FLabelRotation;
    property Logarithmic: Boolean read FLogarithmic write FLogarithmic;
    property MajorGridLines: TsChartLine read FMajorGridLines write FMajorGridLines;
    property MajorInterval: Double read FMajorInterval write FMajorInterval;
    property MajorTickLines: TsChartLine read FMajorTickLines write FMajorTickLines;
    property Max: Double read FMax write FMax;
    property Min: Double read FMin write FMin;
    property MinorGridLines: TsChartLine read FMinorGridLines write FMinorGridLines;
    property MinorSteps: Double read FMinorSteps write FMinorSteps;
    property MinorTickLines: TsChartLine read FMinorTickLines write FMinorTickLines;
    property Position: TsChartAxisPosition read FPosition write FPosition;
    property PositionValue: Double read FPositionValue write FPositionValue;
    property ShowCaption: Boolean read FShowCaption write FShowCaption;
    property ShowLabels: Boolean read FShowLabels write FShowLabels;
  end;

  TsChartLegend = class(TsChartFillElement)
  private
    FFont: TsFont;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    property Font: TsFont read FFont write FFont;
  end;

  TsChartAxisLink = (alPrimary, alSecondary);

  TsChartSeries = class(TsChartElement)
  private
    FChartType: TsChartType;
    FXRange: TsCellRange;          // cell range containing the x data
    FYRange: TsCellRange;
    FLabelRange: TsCellRange;
    FYAxis: TsChartAxisLink;
    FTitleAddr: TsCellCoord;
    FLabelFormat: String;
    FLine: TsChartLine;
    FFill: TsChartFill;
    FBorder: TsChartLine;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    function GetCount: Integer;
    function GetXCount: Integer;
    function GetYCount: Integer;
    function HasLabels: Boolean;
    function HasXValues: Boolean;
    function HasYValues: Boolean;
    procedure SetTitleAddr(ARow, ACol: Cardinal);
    procedure SetLabelRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetXRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetYRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    function LabelsInCol: Boolean;
    function XValuesInCol: Boolean;
    function YValuesInCol: Boolean;
    property ChartType: TsChartType read FChartType;
    property Count: Integer read GetCount;
    property LabelFormat: String read FLabelFormat write FLabelFormat;  // Number format in Excel notation, e.g. '0.00'
    property LabelRange: TsCellRange read FLabelRange;
    property TitleAddr: TsCellCoord read FTitleAddr write FTitleAddr;
    property XRange: TsCellRange read FXRange;
    property YRange: TsCellRange read FYRange;
    property YAxis: TsChartAxisLink read FYAxis write FYAxis;

    property Border: TsChartLine read FBorder write FBorder;
    property Fill: TsChartFill read FFill write FFill;
    property Line: TsChartLine read FLine write FLine;
  end;

  TsChartSeriesSymbol = (
    cssRect, cssDiamond, cssTriangle, cssTriangleDown, cssTriangleLeft,
    cssTriangleRight, cssCircle, cssStar, cssX, cssPlus, cssAsterisk
  );

  TsLineSeries = class(TsChartSeries)
  private
    FSymbol: TsChartSeriesSymbol;
    FSymbolHeight: Double;  // in mm
    FSymbolWidth: Double;   // in mm
    FShowSymbols: Boolean;
    function GetSymbolBorder: TsChartLine;
    function GetSymbolFill: TsChartFill;
    procedure SetSymbolBorder(Value: TsChartLine);
    procedure SetSymbolFill(Value: TsChartFill);
  public
    constructor Create(AChart: TsChart);
    property Symbol: TsChartSeriesSymbol read FSymbol write FSymbol;
    property SymbolBorder: TsChartLine read GetSymbolBorder write SetSymbolBorder;
    property SymbolFill: TsChartFill read GetSymbolFill write SetSymbolFill;
    property SymbolHeight: double read FSymbolHeight write FSymbolHeight;
    property SymbolWidth: double read FSymbolWidth write FSymbolWidth;
    property ShowSymbols: Boolean read FShowSymbols write FShowSymbols;
  end;

  TsChartSeriesList = class(TFPList)
  private
    function GetItem(AIndex: Integer): TsChartSeries;
    procedure SetItem(AIndex: Integer; AValue: TsChartSeries);
  public
    property Items[AIndex: Integer]: TsChartSeries read GetItem write SetItem; default;
  end;

  TsChart = class(TsChartFillElement)
  private
    FIndex: Integer;             // Index in workbook's chart list
    FSheetIndex: Integer;
    FRow, FCol: Cardinal;
    FOffsetX, FOffsetY: Double;
    FWidth, FHeight: Double;     // Width, Height of the chart, in mm.

    FPlotArea: TsChartFillElement;
    FFloor: TsChartFillElement;
    FXAxis: TsChartAxis;
    FX2Axis: TsChartAxis;
    FYAxis: TsChartAxis;
    FY2Axis: TsChartAxis;

    FTitle: TsChartText;
    FSubTitle: TsChartText;
    FLegend: TsChartLegend;
    FSeriesList: TsChartSeriesList;

    FLineStyles: TsChartLineStyleList;
    function GetCategoryLabelRange: TsCellRange;

  public
    constructor Create;
    destructor Destroy; override;
    function AddSeries(ASeries: TsChartSeries): Integer;
    procedure DeleteSeries(AIndex: Integer);

    function GetChartType: TsChartType;
    function GetLineStyle(AIndex: Integer): TsChartLineStyle;
    function IsScatterChart: Boolean;
    function NumLineStyles: Integer;
    {
    function CategoriesInCol: Boolean;
    function CategoriesInRow: Boolean;
    function GetCategoryCount: Integer;
    function HasCategories: Boolean;
     }
    { Index of chart in workbook's chart list. }
    property Index: Integer read FIndex write FIndex;
    { Index of worksheet sheet which contains the chart. }
    property SheetIndex: Integer read FSheetIndex write FSheetIndex;
    { Row index of the cell in which the chart has its top/left corner (anchor) }
    property Row: Cardinal read FRow write FRow;
    { Column index of the cell in which the chart has its top/left corner (anchor) }
    property Col: Cardinal read FCol write FCol;
    { Offset of the left chart edge relative to the anchor cell, in mm }
    property OffsetX: double read FOffsetX write FOffsetX;
    { Offset of the top chart edge relative to the anchor cell, in mm }
    property OffsetY: double read FOffsetY write FOffsetY;
    { Width of the chart, in mm }
    property Width: double read FWidth write FWidth;
    { Height of the chart, in mm }
    property Height: double read FHeight write FHeight;

    { Attributes of the entire chart background }
    property Background: TsChartFill read FBackground write FBackground;
    property Border: TsChartLine read FBorder write FBorder;

    { Attributes of the plot area (rectangle enclosed by axes) }
    property PlotArea: TsChartFillElement read FPlotArea write FPlotArea;
    { Attributes of the floor of a 3D chart }
    property Floor: TsChartFillElement read FFloor write FFloor;

    { Attributes of the chart's title }
    property Title: TsChartText read FTitle write FTitle;
    { Attributes of the chart's subtitle }
    property Subtitle: TsChartText read FSubtitle write FSubTitle;
    { Attributs of the chart's legend }
    property Legend: TsChartLegend read FLegend write FLegend;

    { Attributes of the plot's primary x axis (bottom) }
    property XAxis: TsChartAxis read FXAxis write FXAxis;
    { Attributes of the plot's secondary x axis (top) }
    property X2Axis: TsChartAxis read FX2Axis write FX2Axis;
    { Attributes of the plot's primary y axis (left) }
    property YAxis: TsChartAxis read FYAxis write FYAxis;
    { Attributes of the plot's secondary y axis (right) }
    property Y2Axis: TsChartAxis read FY2Axis write FY2Axis;

    property CategoryLabelRange: TsCellRange read GetCategoryLabelRange;

    { Attributes of the series }
    property Series: TsChartSeriesList read FSeriesList write FSeriesList;
  end;

  TsChartList = class(TObjectList)
  private
    function GetItem(AIndex: Integer): TsChart;
    procedure SetItem(AIndex: Integer; AValue: TsChart);
  public
    property Items[AIndex: Integer]: TsChart read GetItem write SetItem; default;
  end;


implementation

{ TsChartLineStyle }

function TsChartLineStyle.GetID: String;
var
  i: Integer;
begin
  Result := Name;
  for i:=1 to Length(Result) do
    if Result[i] in [' ', '-'] then Result[i] := '_';
  Result := 'FPS' + Result;
end;


{ TsChartLineStyleList }

function TsChartLineStyleList.Add(AName: String;
  ASeg1Length: Double; ASeg1Count: Integer;
  ASeg2Length: Double; ASeg2Count: Integer;
  ADistance: Double; ARelativeToLineWidth: Boolean): Integer;
var
  ls: TsChartLineStyle;
begin
  ls := TsChartLineStyle.Create;
  ls.Name := AName;
  ls.Segment1.Count := ASeg1Count;
  ls.Segment1.Length := ASeg1Length;
  ls.Segment2.Count := ASeg2Count;
  ls.Segment2.Length := ASeg2Length;
  ls.Distance := ADistance;
  ls.RelativeToLineWidth := ARelativeToLineWidth;
  result := inherited Add(ls);
end;

function TsChartLineStyleList.GetItem(AIndex: Integer): TsChartLineStyle;
begin
  Result := TsChartLineStyle(inherited);
end;

procedure TsChartLineStyleList.SetItem(AIndex: Integer; AValue: TsChartLineStyle);
begin
  inherited Items[AIndex] := AValue;
end;


{ TsChartElement }

constructor TsChartElement.Create(AChart: TsChart);
begin
  inherited Create;
  FChart := AChart;
  FVisible := true;
end;


{ TsChartFillElement }

constructor TsChartFillElement.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FBackground := TsChartFill.Create;
  FBackground.Style := fsSolidFill;
  FBackground.BgColor := scWhite;
  FBackground.FgColor := scWhite;
  FBorder := TsChartLine.Create;
  FBorder.Style := clsSolid;
  FBorder.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  FBorder.Color := scBlack;
end;

destructor TsChartFillElement.Destroy;
begin
  FBorder.Free;
  FBackground.Free;
  inherited;
end;


{ TsChartText }

constructor TsChartText.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FShowCaption := true;
  FFont := TsFont.Create;
  FFont.Size := 10;
  FFont.Style := [];
  FFont.Color := scBlack;
end;

destructor TsChartText.Destroy;
begin
  FFont.Free;
  inherited;
end;


{ TsChartAxis }

constructor TsChartAxis.Create(AChart: TsChart);
begin
  inherited Create(AChart);

  FAutomaticMajorInterval := true;
  FAutomaticMinorSteps := true;

  FCaptionFont := TsFont.Create;
  FCaptionFont.Size := 10;
  FCaptionFont.Style := [];
  FCaptionFont.Color := scBlack;

  FLabelFont := TsFont.Create;
  FLabelFont.Size := 9;
  FLabelFont.Style := [];
  FLabelFont.Color := scBlack;

  FCaptionRotation := 0;
  FLabelRotation := 0;

  FShowCaption := true;
  FShowLabels := true;

  FAxisLine := TsChartLine.Create;
  FAxisLine.Color := scBlack;
  FAxisLine.Style := clsSolid;
  FAxisLine.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);

  FMajorTickLines := TsChartLine.Create;
  FMajorTickLines.Color := scBlack;
  FMajorTickLines.Style := clsSolid;
  FMajorTickLines.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);

  FMinorTickLines := TsChartLine.Create;
  FMinorTickLines.Color := scBlack;
  FMinorTickLines.Style := clsSolid;
  FMinorTickLines.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);

  FMajorGridLines := TsChartLine.Create;
  FMajorGridLines.Color := scSilver;
  FMajorGridLines.Style := clsSolid;
  FMajorGridLines.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);

  FMinorGridLines := TsChartLine.Create;
  FMinorGridLines.Color := scSilver;
  FMinorGridLines.Style := clsDash;
  FMinorGridLines.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
end;

destructor TsChartAxis.Destroy;
begin
  FMinorGridLines.Free;
  FMajorGridLines.Free;
  FMinorTickLines.Free;
  FMajorTickLines.Free;
  FAxisLine.Free;
  FLabelFont.Free;
  FCaptionFont.Free;
  inherited;
end;


{ TsChartLegend }

constructor TsChartLegend.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FFont := TsFont.Create;
  FFont.Size := 9;
  FVisible := true;
end;

destructor TsChartLegend.Destroy;
begin
  FFont.Free;
  inherited;
end;


{ TsChartSeries }

constructor TsChartSeries.Create(AChart: TsChart);
var
  idx: Integer;
begin
  inherited Create(AChart);

  idx := AChart.AddSeries(self);

  FBorder := TsChartLine.Create;
  FBorder.Style := clsSolid;
  FBorder.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  FBorder.Color := scBlack;

  FFill := TsChartFill.Create;
  FFill.Style := fsSolidFill;
  FFill.FgColor := DEFAULT_SERIES_COLORS[idx mod Length(DEFAULT_SERIES_COLORS)];
  FFill.BgColor := DEFAULT_SERIES_COLORS[idx mod Length(DEFAULT_SERIES_COLORS)];

  FLine := TsChartLine.Create;
  FLine.Style := clsSolid;
  FLine.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  FLine.Color := DEFAULT_SERIES_COLORS[idx mod Length(DEFAULT_SERIES_COLORS)];
end;

destructor TsChartSeries.Destroy;
begin
  FLine.Free;
  FFill.Free;
  FBorder.Free;
  inherited;
end;

function TsChartSeries.GetCount: Integer;
begin
  Result := GetYCount;
end;

function TsChartSeries.GetXCount: Integer;
begin
  if (FXRange.Row1 = FXRange.Row2) and (FXRange.Col1 = FXRange.Col2) then
    Result := 0
  else
  if (FXRange.Row1 = FXRange.Row2) then
    Result := FXRange.Col2 - FXRange.Col1 + 1
  else
    Result := FXRange.Row2 - FXRange.Row1 + 1;
end;

function TsChartSeries.GetYCount: Integer;
begin
  if YValuesInCol then
    Result := FYRange.Row2 - FYRange.Row1 + 1
  else
    Result := FYRange.Col2 - FYRange.Col1 + 1;
end;

function TsChartSeries.HasLabels: Boolean;
begin
  Result := not ((FLabelRange.Row1 = FLabelRange.Row2) and (FLabelRange.Col1 = FLabelRange.Col2));
end;

function TsChartSeries.HasXValues: Boolean;
begin
  Result := not ((FXRange.Row1 = FXRange.Row2) and (FXRange.Col1 = FXRange.Col2));
end;

function TsChartSeries.HasYValues: Boolean;
begin
  Result := not ((FYRange.Row1 = FYRange.Row2) and (FYRange.Col1 = FYRange.Col2));
end;

function TsChartSeries.LabelsInCol: Boolean;
begin
  Result := (FLabelRange.Col1 = FLabelRange.Col2) and (FLabelRange.Row1 <> FLabelRange.Row2);
end;

procedure TsChartSeries.SetTitleAddr(ARow, ACol: Cardinal);
begin
  FTitleAddr.Row := ARow;
  FTitleAddr.Col := ACol;
end;

procedure TsChartSeries.SetLabelRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Series labels can only be located in a single column or row.');
  FLabelRange.Row1 := ARow1;
  FLabelRange.Col1 := ACol1;
  FLabelRange.Row2 := ARow2;
  FLabelRange.Col2 := ACol2;
end;

procedure TsChartSeries.SetXRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Series x values can only be located in a single column or row.');
  FXRange.Row1 := ARow1;
  FXRange.Col1 := ACol1;
  FXRange.Row2 := ARow2;
  FXRange.Col2 := ACol2;
end;

procedure TsChartSeries.SetYRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Series y values can only be located in a single column or row.');
  FYRange.Row1 := ARow1;
  FYRange.Col1 := ACol1;
  FYRange.Row2 := ARow2;
  FYRange.Col2 := ACol2;
end;

function TsChartSeries.XValuesInCol: Boolean;
begin
  Result := (FXRange.Col1 = FXRange.Col2) and (FXRange.Row1 <> FXRange.Row2);
end;

function TsChartSeries.YValuesInCol: Boolean;
begin
  Result := (FYRange.Col1 = FYRange.Col2) and (FYRange.Row1 <> FYRange.Row2);
end;


{ TsChartSeriesList }

function TsChartSeriesList.GetItem(AIndex: Integer): TsChartSeries;
begin
  Result := TsChartSeries(inherited Items[AIndex]);
end;

procedure TsChartSeriesList.SetItem(AIndex: Integer; AValue: TsChartSeries);
begin
  inherited Items[AIndex] := AValue;
end;


{ TsLineSeries }

constructor TsLineSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctLine;
  FSymbolWidth := 2.5;
  FSymbolHeight := 2.5;
end;

function TsLineSeries.GetSymbolBorder: TsChartLine;
begin
  Result := FBorder;
end;

function TsLineSeries.GetSymbolFill: TsChartFill;
begin
  Result := FFill;
end;

procedure TsLineSeries.SetSymbolBorder(Value: TsChartLine);
begin
  FBorder := Value;
end;

procedure TsLineSeries.SetSymbolFill(Value: TsChartFill);
begin
  FFill := Value;
end;


{ TsChart }

constructor TsChart.Create;
begin
  inherited Create(nil);

  FLineStyles := TsChartLineStyleList.Create;
  clsFineDot := FLineStyles.Add('fine-dot', 100, 1, 0, 0, 100, false);
  clsDot := FLineStyles.Add('dot', 150, 1, 0, 0, 150, true);
  clsDash := FLineStyles.Add('dash', 300, 1, 0, 0, 150, true);
  clsDashDot := FLineStyles.Add('dash-dot', 300, 1, 100, 1, 150, true);
  clsLongDash := FLineStyles.Add('long dash', 400, 1, 0, 0, 200, true);
  clsLongDashDot := FLineStyles.Add('long dash-dot', 500, 1, 100, 1, 200, true);
  clsLongDashDotDot := FLineStyles.Add('long dash-dot-dot', 500, 1, 100, 2, 200, true);

  FSheetIndex := 0;
  FRow := 0;
  FCol := 0;
  FOffsetX := 0.0;
  FOffsetY := 0.0;
  FWidth := 12;
  FHeight := 9;

  // FBackground and FBorder already created by ancestor.

  FPlotArea := TsChartFillElement.Create(self);
  FFloor := TsChartFillElement.Create(self);
  FFloor.Background.Style := fsNoFill;

  FTitle := TsChartText.Create(self);
  FTitle.Font.Size := 14;

  FSubTitle := TsChartText.Create(self);
  FSubTitle.Font.Size := 12;

  FLegend := TsChartLegend.Create(self);

  FXAxis := TsChartAxis.Create(self);
  FXAxis.Caption := 'x axis';
  FXAxis.CaptionFont.Style := [fssBold];
  FXAxis.LabelFont.Size := 9;
  FXAxis.Position := capStart;

  FX2Axis := TsChartAxis.Create(self);
  FX2Axis.Caption := 'Secondary x axis';
  FX2Axis.CaptionFont.Style := [fssBold];
  FX2Axis.LabelFont.Size := 9;
  FX2Axis.Visible := false;
  FX2Axis.Position := capEnd;

  FYAxis := TsChartAxis.Create(self);
  FYAxis.Caption := 'y axis';
  FYAxis.CaptionFont.Style := [fssBold];
  FYAxis.CaptionRotation := 90;
  FYAxis.LabelFont.Size := 9;
  FYAxis.Position := capStart;

  FY2Axis := TsChartAxis.Create(self);
  FY2Axis.Caption := 'Secondary y axis';
  FY2Axis.CaptionFont.Style := [fssBold];
  FY2Axis.CaptionRotation := 90;
  FY2Axis.LabelFont.Size := 9;
  FY2Axis.Visible := false;
  FY2Axis.Position := capEnd;

  FSeriesList := TsChartSeriesList.Create;
end;

destructor TsChart.Destroy;
begin
  FSeriesList.Free;
  FXAxis.Free;
  FYAxis.Free;
  FY2Axis.Free;
  FLegend.Free;
  FTitle.Free;
  FSubtitle.Free;
  FLineStyles.Free;
  FFloor.Free;
  FPlotArea.Free;
  inherited;
end;

function TsChart.AddSeries(ASeries: TsChartSeries): Integer;
begin
  Result := FSeriesList.IndexOf(ASeries);
  if Result = -1 then
    Result := FSeriesList.Add(ASeries);
end;

procedure TsChart.DeleteSeries(AIndex: Integer);
begin
  if (AIndex >= 0) and (AIndex < FSeriesList.Count) then
    FSeriesList.Delete(AIndex);
end;

function TsChart.GetCategoryLabelRange: TsCellRange;
begin
  if FSeriesList.Count > 0 then
    Result := Series[0].LabelRange
  else
  begin
    Result.Row1 := 0;
    Result.Col1 := 0;
    Result.Row2 := 0;
    Result.Col2 := 0;
  end;
end;

function TsChart.GetChartType: TsChartType;
begin
  if FSeriesList.Count > 0 then
    Result := Series[0].ChartType
  else
    Result := ctEmpty;
end;

function TsChart.GetLineStyle(AIndex: Integer): TsChartLineStyle;
begin
  if AIndex >= 0 then
    Result := FLineStyles[AIndex]
  else
    Result := nil;
end;

function TsChart.IsScatterChart: Boolean;
begin
  Result := GetChartType = ctScatter;
end;

function TsChart.NumLineStyles: Integer;
begin
  Result := FLineStyles.Count;
end;


{ TsChartList }

function TsChartList.GetItem(AIndex: Integer): TsChart;
begin
  Result := TsChart(inherited Items[AIndex]);
end;

procedure TsChartlist.SetItem(AIndex: Integer; AValue: TsChart);
begin
  inherited Items[AIndex] := AValue;
end;

end.

