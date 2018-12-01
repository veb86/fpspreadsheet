unit fpschart;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Contnrs, fpsTypes, fpsUtils;

type
  TsChart = class;

  TsChartFill = record
    Style: TsFillStyle;
    FgColor: TsColor;
    BgColor: TsColor;
  end;

  TsChartPenStyle = (cpsSolid, cpsDashed, cpsDotted, cpsDashDot, cpsDashDotDot, cpsClear);

  TsChartLine = record
    Style: TsChartPenStyle;
    Width: Double;  // mm
    Color: TsColor;
  end;

  TsChartObj = class
  private
    FOwner: TsChartObj;
    FVisible: Boolean;
  public
    constructor Create(AOwner: TsChartObj);
    property Visible: Boolean read FVisible write FVisible;
  end;

  TsChartFillObj = class(TsChartObj)
  private
    FBackground: TsChartFill;
    FBorder: TsChartLine;
  public
    constructor Create(AOwner: TsChartObj);
    property Background: TsChartFill read FBackground write FBackground;
    property Border: TsChartLine read FBorder write FBorder;
  end;

  TsChartText = class(TsChartObj)
  private
    FCaption: String;
    FShowCaption: Boolean;
    FFont: TsFont;
  public
    constructor Create(AOwner: TsChartObj);
    property Caption: String read FCaption write FCaption;
    property Font: TsFont read FFont write FFont;
    property ShowCaption: Boolean read FShowCaption write FShowCaption;
  end;

  TsChartAxis = class(TsChartText)
  private
    FAutomaticMax: Boolean;
    FAutomaticMin: Boolean;
    FAutomaticMajorInterval: Boolean;
    FAutomaticMinorSteps: Boolean;
    FAxisLine: TsChartLine;
    FGridLines: TsChartLine;
    FInverted: Boolean;
    FLabelFont: TsFont;
    FLogarithmic: Boolean;
    FMajorInterval: Double;
    FMajorTickLines: TsChartLine;
    FMax: Double;
    FMin: Double;
    FMinorSteps: Double;
    FMinorTickLines: TsChartLine;
    FShowGrid: Boolean;
    FShowLabels: Boolean;
  public
    constructor Create(AOwner: TsChartObj);
    property AutomaticMax: Boolean read FAutomaticMax write FAutomaticMax;
    property AutomaticMin: Boolean read FAutomaticMin write FAutomaticMin;
    property AutomaticMajorInterval: Boolean read FAutomaticMajorInterval write FAutomaticMajorInterval;
    property AutomaticMinorSteps: Boolean read FAutomaticMinorSteps write FAutomaticMinorSteps;
    property AxisLine: TsChartLine read FAxisLine write FAxisLine;
    property GridLines: TsChartLine read FGridLines write FGridLines;
    property Inverted: Boolean read FInverted write FInverted;
    property LabelFont: TsFont read FLabelFont write FLabelFont;
    property Logarithmic: Boolean read FLogarithmic write FLogarithmic;
    property MajorInterval: Double read FMajorInterval write FMajorInterval;
    property MajorTickLines: TsChartLine read FMajorTickLines write FMajorTickLines;
    property Max: Double read FMax write FMax;
    property Min: Double read FMin write FMin;
    property MinorSteps: Double read FMinorSteps write FMinorSteps;
    property MinorTickLines: TsChartLine read FMinorTickLines write FMinorTickLines;
    property ShowGrid: Boolean read FShowGrid write FShowGrid;
    property ShowLabels: Boolean read FShowLabels write FShowLabels;
  end;

  TsChartLegend = class(TsChartFillObj)
  end;

  TsChartAxisLink = (alPrimary, alSecondary);

  TsChartSeries = class(TsChartObj)
  private
    FXRange: TsCellRange;          // cell range containing the x data
    FYRange: TsCellRange;
    FLabelsRange: TsCellRange;
    FXIndex: array of Integer;     // index of data point's x value within XRange
    FYIndex: array of Integer;
    FLabelsIndex: array of Integer;
    FYAxis: TsChartAxisLink;
    FTitle: String;
    function GetCount: Integer;
  public
    constructor Create(AChart: TsChart);
    property Count: Integer read GetCount;
    property LabelsRange: TsCellRange read FLabelsRange;
    property Title: String read FTitle;
    property XRange: TsCellRange read FXRange write FXRange;
    property YRange: TsCellRange read FYRange write FYRange;
    property YAxis: TsChartAxisLink read FYAxis write FYAxis;
  end;

  TsChartSeriesSymbol = (cssRect, cssDiamond, cssTriangle, cssTriangleDown,
    cssCircle, cssStar);

  TsLineSeries = class(TsChartSeries)
  private
    FLineStyle: TsChartLine;
    FShowLines: Boolean;
    FShowSymbols: Boolean;
    FSymbol: TsChartSeriesSymbol;
    FSymbolFill: TsChartFill;
    FSymbolBorder: TsChartLine;
    FSymbolHeight: Double;  // in mm
    FSymbolWidth: Double;   // in mm
  public
    constructor Create(AChart: TsChart);
    property LineStyle: TsChartLine read FLineStyle write FLineStyle;
    property ShowLines: Boolean read FShowLines write FShowLines;
    property ShowSymbols: Boolean read FShowSymbols write FShowSymbols;
    property Symbol: TsChartSeriesSymbol read FSymbol write FSymbol;
    property SymbolBorder: TsChartLine read FSymbolBorder write FSymbolBorder;
    property SymbolFill: TsChartFill read FSymbolFill write FSymbolFill;
    property SymbolHeight: double read FSymbolHeight write FSymbolHeight;
    property SymbolWidth: double read FSymbolWidth write FSymbolWidth;
  end;

  TsChartSeriesList = class(TFPList)
  private
    function GetItem(AIndex: Integer): TsChartSeries;
    procedure SetItem(AIndex: Integer; AValue: TsChartSeries);
  public
    property Items[AIndex: Integer]: TsChartSeries read GetItem write SetItem; default;
  end;

  TsChart = class(TsChartFillObj)
  private
    FSheetIndex: Integer;
    FRow, FCol: Cardinal;
    FOffsetX, FOffsetY: Double;
    FWidth, FHeight: Double;     // Width, Height of the chart, in mm.

    FPlotArea: TsChartFillObj;
    FXAxis: TsChartAxis;
    FYAxis: TsChartAxis;
    FY2Axis: TsChartAxis;

    FTitle: TsChartText;
    FSubTitle: TsChartText;
    FLegend: TsChartLegend;
    FSeriesList: TsChartSeriesList;
  public
    constructor Create;
    destructor Destroy; override;
    function AddSeries(ASeries: TsChartSeries): Integer;

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
    property Width: double read FWidth write FHeight;
    { Height of the chart, in mm }
    property Height: double read FHeight write FHeight;

    { Attributes of the plot area (rectangle enclosed by axes) }
    property PlotArea: TsChartFillObj read FPlotArea write FPlotArea;

    { Attributes of the chart's title }
    property Title: TsChartText read FTitle write FTitle;
    { Attributes of the chart's subtitle }
    property Subtitle: TsChartText read FSubtitle write FSubTitle;
    { Attributs of the chart's legend }
    property Legend: TsChartLegend read FLegend write FLegend;

    { Attributes of the plot's primary x axis (bottom) }
    property XAxis: TsChartAxis read FXAxis write FXAxis;
    { Attributes of the plot's primary y axis (left) }
    property YAxis: TsChartAxis read FYAxis write FYAxis;
    { Attributes of the plot's secondary y axis (right) }
    property Y2Axis: TsChartAxis read FY2Axis write FY2Axis;

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

const
  DEFAULT_LINE_WIDTH = 0.75;  // pts


{ TsChartObj }

constructor TsChartObj.Create(AOwner: TsChartObj);
begin
  inherited Create;
  FOwner := AOwner;
  FVisible := true;
end;


{ TsChartFillObj }

constructor TsChartFillObj.Create(AOwner: TsChartObj);
begin
  inherited Create(AOwner);
  FBackground.Style := fsSolidFill;
  FBackground.BgColor := scWhite;
  FBackground.FgColor := scWhite;
  FBorder.Style := cpsSolid;
  FBorder.Width := PtsToMM(DEFAULT_LINE_WIDTH);
  FBorder.Color := scBlack;
end;


{ TsChartText }

constructor TsChartText.Create(AOwner: TsChartObj);
begin
  inherited Create(AOwner);
  FShowCaption := true;
  FFont.FontName := '';  // replace by workbook's default font
  FFont.Size := 0;       // replace by workbook's default font size
  FFont.Style := [];
  FFont.Color := scBlack;
end;


{ TsChartAxis }

constructor TsChartAxis.Create(AOwner: TsChartObj);
begin
  inherited Create(AOwner);

  FAutomaticMajorInterval := true;
  FAutomaticMinorSteps := true;

  FLabelFont.FontName := '';  // replace by workbook's default font
  FLabelFont.Size := 0;       // Replace by workbook's default font size
  FLabelFont.Style := [];
  FLabelFont.Color := scBlack;

  FShowLabels := true;

  FAxisLine.Color := scBlack;
  FAxisLine.Style := cpsSolid;
  FAxisLine.Width := PtsToMM(DEFAULT_LINE_WIDTH);

  FMajorTickLines.Color := scBlack;
  FMajorTickLines.Style := cpsSolid;
  FMajorTickLines.Width := PtsToMM(DEFAULT_LINE_WIDTH);

  FMinorTickLines.Color := scBlack;
  FMinorTickLines.Style := cpsSolid;
  FMinorTickLines.Width := PtsToMM(DEFAULT_LINE_WIDTH);

  FGridLines.Color := scSilver;
  FGridLines.Style := cpsDotted;
  FGridLines.Width := PtsToMM(DEFAULT_LINE_WIDTH);
end;


{ TsChartSeries }

constructor TsChartSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  AChart.AddSeries(self);
  FTitle := 'Series ' + IntToStr(AChart.Series.Count);
end;

function TsChartSeries.GetCount: Integer;
begin
  Result := Length(FYIndex);
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

  FLineStyle.Color := scBlack;
  FLineStyle.Style := cpsSolid;
  FLineStyle.Width := PtsToMM(DEFAULT_LINE_WIDTH);

  FSymbolBorder.Color := scBlack;
  FSymbolBorder.Style := cpsSolid;
  FSymbolBorder.Width := PtsToMM(DEFAULT_LINE_WIDTH);

  FSymbolFill.FgColor := scWhite;
  FSymbolFill.BgColor := scWhite;
  FSymbolFill.Style := fsSolidFill;

  FSymbolWidth := 2.5;
  FSymbolHeight := 2.5;
end;


{ TsChart }

constructor TsChart.Create;
begin
  inherited Create(nil);

  FSheetIndex := 0;
  FRow := 0;
  FCol := 0;
  FOffsetX := 0.0;
  FOffsetY := 0.0;
  FWidth := 12;
  FHeight := 9;

  FTitle := TsChartText.Create(self);
  FTitle.Font.Size := 14;

  FSubTitle := TsChartText.Create(self);
  FSubTitle.Font.Size := 12;

  FLegend := TsChartLegend.Create(self);

  FXAxis := TsChartAxis.Create(self);
  FXAxis.Caption := 'x axis';
  FXAxis.LabelFont.Size := 9;
  FXAxis.Font.Size := 10;
  FXAxis.Font.Style := [fssBold];

  FYAxis := TsChartAxis.Create(self);
  FYAxis.Caption := 'y axis';
  FYAxis.LabelFont.Size := 9;
  FYAxis.Font.Size := 10;
  FYAxis.Font.Style := [fssBold];

  FY2Axis := TsChartAxis.Create(self);
  FY2Axis.Caption := 'Secondary y axis';
  FY2Axis.LabelFont.Size := 9;
  FY2Axis.Font.Size := 10;
  FY2Axis.Font.Style := [fssBold];
  FYAxis.Visible := false;

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
  inherited;
end;

function TsChart.AddSeries(ASeries: TsChartSeries): Integer;
begin
  Result := FSeriesList.Add(ASeries);
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

