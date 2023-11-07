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

  TsChartGradientStyle = (cgsLinear, cgsAxial, cgsRadial, cgsElliptic, cgsSquare, cgsRectangular);

  TsChartGradient = class
    Name: String;
    Style: TsChartGradientStyle;
    StartColor: TsColor;
    EndColor: TsColor;
    StartIntensity: Double;    // 0.0 ... 1.0
    EndIntensity: Double;      // 0.0 ... 1.0
    Border: Double;            // 0.0 ... 1.0
    CenterX, CenterY: Double;  // 0.0 ... 1.0
    Angle: Double;             // degrees
    constructor Create;
  end;

  TsChartGradientList = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartGradient;
    procedure SetItem(AIndex: Integer; AValue: TsChartGradient);
    function AddGradient(AName: String; AStyle: TsChartGradientStyle;
      AStartColor, AEndColor: TsColor; AStartIntensity, AEndIntensity: Double;
      ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
  public
    function AddAxialGradient(AName: String; AStartColor, AEndColor: TsColor;
      AStartIntensity, AEndIntensity, ABorder, AAngle: Double): Integer;
    function AddEllipticGradient(AName: String; AStartColor, AEndColor: TsColor;
      AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
    function AddLinearGradient(AName: String; AStartColor, AEndColor: TsColor;
      AStartIntensity, AEndIntensity, ABorder, AAngle: Double): Integer;
    function AddRadialGradient(AName: String;
      AStartColor, AEndColor: TsColor; AStartIntensity, AEndIntensity, ABorder: Double;
      ACenterX, ACenterY: Double): Integer;
    function AddRectangularGradient(AName: String; AStartColor, AEndColor: TsColor;
      AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
    function AddSquareGradient(AName: String; AStartColor, AEndColor: TsColor;
      AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
    function IndexOfName(AName: String): Integer;
    function FindByName(AName: String): TsChartGradient;
    property Items[AIndex: Integer]: TsChartGradient read GetItem write SetItem; default;
  end;

  TsChartHatchStyle = (chsSingle, chsDouble, chsTriple);

  TsChartHatch = class
    Name: String;
    Style: TsChartHatchStyle;
    LineColor: TsColor;
    LineDistance: Double;      // mm
    LineAngle: Double;         // degrees
    Filled: Boolean;           // filled with background color or not
  end;

  TsChartHatchList = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartHatch;
    procedure SetItem(AIndex: Integer; AValue: TsChartHatch);
  public
    function AddHatch(AName: String; AStyle: TsChartHatchStyle;
      ALineColor: TsColor; ALineDistance, ALineAngle: Double; AFilled: Boolean): Integer;
    function FindByName(AName: String): TsChartHatch;
    function IndexOfName(AName: String): Integer;
    property Items[AIndex: Integer]: TsChartHatch read GetItem write SetItem; default;
  end;

  TsChartFillStyle = (cfsNoFill, cfsSolid, cfsGradient, cfsHatched);

  TsChartFill = class
    Style: TsChartFillStyle;
    Color: TsColor;
    Gradient: Integer;
    Hatch: Integer;
    Transparency: Double;  // 0.0 ... 1.0
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
    FRotationAngle: Integer;
    FShowCaption: Boolean;
    FFont: TsFont;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    property Caption: String read FCaption write FCaption;
    property Font: TsFont read FFont write FFont;
    property ShowCaption: Boolean read FShowCaption write FShowCaption;
    property RotationAngle: Integer read FRotationAngle write FRotationAngle;
  end;

  TsChartAxisPosition = (capStart, capEnd, capValue);
  TsChartAxisTick = (catInside, catOutside);
  TsChartAxisTicks = set of TsChartAxisTick;
  TsChartType = (ctEmpty, ctBar, ctLine, ctArea, ctBarLine, ctScatter, ctBubble,
    ctRadar, ctFilledRadar, ctPie, ctRing);

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
    FLabelFormatPercent: String;
    FLabelRotation: Integer;
    FLogarithmic: Boolean;
    FMajorInterval: Double;
    FMajorTicks: TsChartAxisTicks;
    FMax: Double;
    FMin: Double;
    FMinorSteps: Double;
    FMinorTicks: TsChartAxisTicks;
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
    property LabelFormatPercent: String read FLabelFormatPercent write FLabelFormatPercent;
    property LabelRotation: Integer read FLabelRotation write FLabelRotation;
    property Logarithmic: Boolean read FLogarithmic write FLogarithmic;
    property MajorGridLines: TsChartLine read FMajorGridLines write FMajorGridLines;
    property MajorInterval: Double read FMajorInterval write FMajorInterval;
    property MajorTicks: TsChartAxisTicks read FMajorTicks write FMajorTicks;
    property Max: Double read FMax write FMax;
    property Min: Double read FMin write FMin;
    property MinorGridLines: TsChartLine read FMinorGridLines write FMinorGridLines;
    property MinorSteps: Double read FMinorSteps write FMinorSteps;
    property MinorTicks: TsChartAxisTicks read FMinorTicks write FMinorTicks;
    property Position: TsChartAxisPosition read FPosition write FPosition;
    property PositionValue: Double read FPositionValue write FPositionValue;
    property ShowCaption: Boolean read FShowCaption write FShowCaption;
    property ShowLabels: Boolean read FShowLabels write FShowLabels;
  end;

  TsChartLegendPosition = (lpRight, lpTop, lpBottom, lpLeft);

  TsChartLegend = class(TsChartFillElement)
  private
    FFont: TsFont;
    FCanOverlapPlotArea: Boolean;
    FPosition: TsChartLegendPosition;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    property CanOverlapPlotArea: Boolean read FCanOverlapPlotArea write FCanOverlapPlotArea;
    property Font: TsFont read FFont write FFont;
    property Position: TsChartLegendPosition read FPosition write FPosition;
    // There is also a "legend-expansion" but this does not seem to have a visual effect in Calc.
  end;

  TsChartAxisLink = (alPrimary, alSecondary);
  TsChartDataLabel = (cdlValue, cdlPercentage, cdlValueAndPercentage, cdlCategory, cdlSeriesName, cdlSymbol);
  TsChartDataLabels = set of TsChartDataLabel;
  TsChartLabelPosition = (lpDefault, lpOutside, lpInside, lpCenter);

  TsChartSeries = class(TsChartElement)
  private
    FChartType: TsChartType;
    FXRange: TsCellRange;          // cell range containing the x data
    FYRange: TsCellRange;
    FLabelRange: TsCellRange;
    FLabelFont: TsFont;
    FLabelPosition: TsChartLabelPosition;
    FLabelSeparator: string;
    FFillColorRange: TsCellRange;
    FYAxis: TsChartAxisLink;
    FTitleAddr: TsCellCoord;
    FLabelFormat: String;
    FLine: TsChartLine;
    FFill: TsChartFill;
    FDataLabels: TsChartDataLabels;
  protected
    function GetChartType: TsChartType; virtual;
  public
    constructor Create(AChart: TsChart); virtual;
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
    procedure SetFillColorRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    function LabelsInCol: Boolean;
    function XValuesInCol: Boolean;
    function YValuesInCol: Boolean;

    property ChartType: TsChartType read GetChartType;
    property Count: Integer read GetCount;
    property DataLabels: TsChartDataLabels read FDataLabels write FDataLabels;
    property FillColorRange: TsCellRange read FFillColorRange;
    property LabelFont: TsFont read FLabelFont write FLabelFont;
    property LabelFormat: String read FLabelFormat write FLabelFormat;  // Number format in Excel notation, e.g. '0.00'
    property LabelPosition: TsChartLabelPosition read FLabelPosition write FLabelPosition;
    property LabelRange: TsCellRange read FLabelRange;
    property LabelSeparator: string read FLabelSeparator write FLabelSeparator;
    property TitleAddr: TsCellCoord read FTitleAddr write FTitleAddr;  // use '\n' for line-break
    property XRange: TsCellRange read FXRange;
    property YRange: TsCellRange read FYRange;
    property YAxis: TsChartAxisLink read FYAxis write FYAxis;

    property Fill: TsChartFill read FFill write FFill;
    property Line: TsChartLine read FLine write FLine;
  end;
  TsChartSeriesClass = class of TsChartSeries;

  TsAreaSeries = class(TsChartSeries)
  public
    constructor Create(AChart: TsChart); override;
  end;

  TsBarSeries = class(TsChartSeries)
  public
    constructor Create(AChart: TsChart); override;
  end;

  TsBubbleSeries = class(TsChartSeries)
  private
    FBubbleRange: TsCellRange;
  public
    constructor Create(AChart: TsChart); override;
    procedure SetBubbleRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    property BubbleRange: TsCellRange read FBubbleRange;
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
    FShowLines: Boolean;
    FShowSymbols: Boolean;
    FBorder: TsChartLine;
    function GetSymbolFill: TsChartFill;
    procedure SetSymbolFill(Value: TsChartFill);
  public
    constructor Create(AChart: TsChart); override;
    destructor Destroy; override;
    property Symbol: TsChartSeriesSymbol read FSymbol write FSymbol;
    property SymbolBorder: TsChartLine read FBorder write FBorder;
    property SymbolFill: TsChartFill read GetSymbolFill write SetSymbolFill;
    property SymbolHeight: double read FSymbolHeight write FSymbolHeight;
    property SymbolWidth: double read FSymbolWidth write FSymbolWidth;
    property ShowLines: Boolean read FShowLines write FShowLines;
    property ShowSymbols: Boolean read FShowSymbols write FShowSymbols;
  end;

  TsPieSeries = class(TsChartSeries)
  public
    constructor Create(AChart: TsChart); override;
  end;

  TsRadarSeries = class(TsLineSeries)
  protected
    function GetChartType: TsChartType; override;
  end;

  TsRingSeries = class(TsChartSeries)
  public
    constructor Create(AChart: TsChart); override;
  end;

  TsRegressionType = (rtNone, rtLinear, rtLogarithmic, rtExponential, rtPower, rtPolynomial);

  TsRegressionEquation = class
    Fill: TsChartFill;
    Font: TsFont;
    Border: TsChartLine;
    NumberFormat: String;
    Left, Top: Double;  // mm, relative to outer chart boundaries!
    XName: String;
    YName: String;
    constructor Create;
    destructor Destroy; override;
    function DefaultBorder: Boolean;
    function DefaultFill: Boolean;
    function DefaultFont: Boolean;
    function DefaultNumberFormat: Boolean;
    function DefaultPosition: Boolean;
    function DefaultXName: Boolean;
    function DefaultYName: Boolean;
  end;

  TsChartRegression = class
    Title: String;
    RegressionType: TsRegressionType;
    ExtrapolateForwardBy: Double;
    ExtrapolateBackwardBy: Double;
    ForceYIntercept: Boolean;
    YInterceptValue: Double;
    PolynomialDegree: Integer;
    DisplayEquation: Boolean;
    DisplayRSquare: Boolean;
    Equation: TsRegressionEquation;
    Line: TsChartLine;
    constructor Create;
    destructor Destroy; override;
  end;

  TsScatterSeries = class(TsLineSeries)
  private
    FRegression: TsChartRegression;
  public
    constructor Create(AChart: TsChart); override;
    destructor Destroy; override;
    property Regression: TsChartRegression read FRegression write FRegression;
  end;

  TsChartSeriesList = class(TFPList)
  private
    function GetItem(AIndex: Integer): TsChartSeries;
    procedure SetItem(AIndex: Integer; AValue: TsChartSeries);
  public
    property Items[AIndex: Integer]: TsChartSeries read GetItem write SetItem; default;
  end;

  TsChartStackMode = (csmSideBySide, csmStacked, csmStackedPercentage);
  TsChartInterpolation = (
    ciLinear,
    ciCubicSpline, ciBSpline,
    ciStepStart, ciStepEnd, ciStepCenterX, ciStepCenterY
  );

  TsChart = class(TsChartFillElement)
  private
    FName: String;
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

    FRotatedAxes: Boolean;        // For bar series: vertical columns <--> horizontal bars
    FStackMode: TsChartStackMode; // For bar and area series
    FInterpolation: TsChartInterpolation; // For line/scatter series: data connection lines

    FTitle: TsChartText;
    FSubTitle: TsChartText;
    FLegend: TsChartLegend;
    FSeriesList: TsChartSeriesList;

    FLineStyles: TsChartLineStyleList;
    FGradients: TsChartGradientList;
    FHatches: TsChartHatchList;
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

    { Name for internal purposes to identify the chart during reading from file }
    property Name: String read FName write FName;
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

    { Connecting line between data points (for line and scatter series) }
    property Interpolation: TsChartInterpolation read FInterpolation write FInterpolation;
    { x and y axes exchanged (mainly for bar series, but works also for scatter and bubble series) }
    property RotatedAxes: Boolean read FRotatedAxes write FRotatedAxes;
    { Stacking of series (for bar and area series ) }
    property StackMode: TsChartStackMode read FStackMode write FStackMode;

    property CategoryLabelRange: TsCellRange read GetCategoryLabelRange;

    { Attributes of the series }
    property Series: TsChartSeriesList read FSeriesList write FSeriesList;

    { Style lists }
    property LineStyles: TsChartLineStyleList read FLineStyles;
    property Gradients: TsChartGradientList read FGradients;
    property Hatches: TsChartHatchList read FHatches;
  end;

  TsChartList = class(TObjectList)
  private
    function GetItem(AIndex: Integer): TsChart;
    procedure SetItem(AIndex: Integer; AValue: TsChart);
  public
    property Items[AIndex: Integer]: TsChart read GetItem write SetItem; default;
  end;


implementation

{ TsChartGradient }

constructor TsChartGradient.Create;
begin
  inherited Create;
  StartIntensity := 1.0;
  EndIntensity := 1.0;
end;


{ TsChartGradientList }

function TsChartGradientList.AddAxialGradient(AName: String;
  AStartColor, AEndColor: TsColor; AStartIntensity, AEndIntensity, ABorder: Double;
  AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsAxial, AStartColor, AEndColor,
    AStartIntensity, AEndIntensity, ABorder, 0.0, 0.0, AAngle);
end;

function TsChartGradientList.AddEllipticGradient(AName: String;
  AStartColor, AEndColor: TsColor; AStartIntensity, AEndIntensity, ABorder: Double;
  ACenterX, ACenterY, AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsElliptic, AStartColor, AEndColor,
    AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY, AAngle);
end;

function TsChartGradientList.AddGradient(AName: String; AStyle: TsChartGradientStyle;
  AStartColor, AEndColor: TsColor;
  AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
var
  item: TsChartGradient;
begin
  if AName = '' then
    AName := 'G' + IntToStr(Count+1);
  Result := IndexOfName(AName);
  if Result = -1 then
  begin
    item := TsChartGradient.Create;
    Result := inherited Add(item);
  end else
    item := Items[Result];
  item.Name := AName;
  item.Style := AStyle;
  item.StartColor := AStartColor;
  item.EndColor := AEndColor;
  item.StartIntensity := AStartIntensity;
  item.EndIntensity := AEndIntensity;
  item.Border := ABorder;
  item.Angle := AAngle;
  item.CenterX := ACenterX;
  item.CenterY := ACenterY;
end;

function TsChartGradientList.AddLinearGradient(AName: String;
  AStartColor, AEndColor: TsColor;
  AStartIntensity, AEndIntensity, ABorder,AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsLinear, AStartColor, AEndColor,
    AStartIntensity, AEndIntensity, ABorder, 0.0, 0.0, AAngle);
end;

function TsChartGradientList.AddRadialGradient(AName: String;
  AStartColor, AEndColor: TsColor;
  AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY: Double): Integer;
begin
  Result := AddGradient(AName, cgsRadial, AStartColor, AEndColor,
    AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY, 0);
end;

function TsChartGradientList.AddRectangularGradient(AName: String;
  AStartColor, AEndColor: TsColor;
  AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsRectangular, AStartColor, AEndColor,
    AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY, AAngle);
end;

function TsChartGradientList.AddSquareGradient(AName: String;
  AStartColor, AEndColor: TsColor;
  AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsSquare, AStartColor, AEndColor,
    AStartIntensity, AEndIntensity, ABorder, ACenterX, ACenterY, AAngle);
end;

function TsChartGradientList.FindByName(AName: String): TsChartGradient;
var
  idx: Integer;
begin
  idx := IndexOfName(AName);
  if idx > -1 then
    Result := Items[idx]
  else
    Result := nil;
end;

function TsChartGradientList.GetItem(AIndex: Integer): TsChartGradient;
begin
  Result := TsChartGradient(inherited Items[AIndex]);
end;

function TsChartGradientList.IndexOfName(AName: String): Integer;
begin
  for Result := 0 to Count-1 do
    if SameText(Items[Result].Name, AName) then
      exit;
  Result := -1;
end;

procedure TsChartGradientList.SetItem(AIndex: Integer; AValue: TsChartGradient);
begin
  inherited Items[AIndex] := AValue;
end;


{ TsChartHatchList }

function TsChartHatchList.AddHatch(AName: String; AStyle: TsChartHatchStyle;
  ALineColor: TsColor; ALineDistance, ALineAngle: Double; AFilled: Boolean): Integer;
var
  item: TsChartHatch;
begin
  if AName = '' then
    AName := 'Hatch' + IntToStr(Count+1);
  Result := IndexOfName(AName);
  if Result = -1 then
  begin
    item := TsChartHatch.Create;
    Result := inherited Add(item);
  end else
    item := Items[Result];
  item.Name := AName;
  item.Style := AStyle;
  item.LineColor := ALineColor;
  item.LineDistance := ALineDistance;
  item.LineAngle := ALineAngle;
  item.Filled := AFilled;
end;

function TsChartHatchList.FindByName(AName: String): TsChartHatch;
var
  idx: Integer;
begin
  idx := IndexOfName(AName);
  if idx > -1 then
    Result := Items[idx]
  else
    Result := nil;
end;

function TsChartHatchList.GetItem(AIndex: Integer): TsChartHatch;
begin
  Result := TsChartHatch(inherited Items[AIndex]);
end;

function TsChartHatchList.IndexOfName(AName: String): Integer;
begin
  for Result := 0 to Count-1 do
    if SameText(Items[Result].Name, AName) then
      exit;
  Result := -1;
end;

procedure TsChartHatchList.SetItem(AIndex: Integer; AValue: TsChartHatch);
begin
  inherited Items[AIndex] := AValue;
end;


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
  i: Integer;
begin
  Result := -1;
  for i := 0 to Count-1 do
    if TsChartLineStyle(Items[i]).Name = AName then
    begin
      Result := i;
      break;
    end;

  if Result = -1 then
  begin
    ls := TsChartLineStyle.Create;
    Result := inherited Add(ls);
  end else
    ls := TsChartlineStyle(Items[Result]);

  ls.Name := AName;
  ls.Segment1.Count := ASeg1Count;
  ls.Segment1.Length := ASeg1Length;
  ls.Segment2.Count := ASeg2Count;
  ls.Segment2.Length := ASeg2Length;
  ls.Distance := ADistance;
  ls.RelativeToLineWidth := ARelativeToLineWidth;
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
  FBackground.Style := cfsSolid;
  FBackground.Color := scWhite;
  FBackground.Gradient := -1;
  FBackground.Hatch := -1;
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

  FAutomaticMin := true;
  FAutomaticMax := true;
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

  FLabelFormatPercent := '0%';

  FCaptionRotation := 0;
  FLabelRotation := 0;

  FShowCaption := true;
  FShowLabels := true;

  FAxisLine := TsChartLine.Create;
  FAxisLine.Color := scBlack;
  FAxisLine.Style := clsSolid;
  FAxisLine.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);

  FMajorTicks := [catOutside];
  FMinorTicks := [];

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

  FFill := TsChartFill.Create;
  FFill.Style := cfsSolid;
  FFill.Color := DEFAULT_SERIES_COLORS[idx mod Length(DEFAULT_SERIES_COLORS)];
  FFill.Gradient := -1;
  FFill.Hatch := -1;

  FLine := TsChartLine.Create;
  FLine.Style := clsSolid;
  FLine.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  FLine.Color := DEFAULT_SERIES_COLORS[idx mod Length(DEFAULT_SERIES_COLORS)];

  FLabelFont := TsFont.Create;
  FLabelFont := TsFont.Create;
  FLabelFont.Size := 9;
end;

destructor TsChartSeries.Destroy;
begin
  FLabelFont.Free;
  FLine.Free;
  FFill.Free;
  inherited;
end;

function TsChartSeries.GetChartType: TsChartType;
begin
  Result := FChartType;
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

procedure TsChartSeries.SetFillColorRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Series fill color values can only be located in a single column or row.');
  FFillColorRange.Row1 := ARow1;
  FFillColorRange.Col1 := ACol1;
  FFillColorRange.Row2 := ARow2;
  FFillColorRange.Col2 := ACol2;
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


{ TsAreaSeries }

constructor TsAreaSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctArea;
end;


{ TsBarSeries }

constructor TsBarSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctBar;
end;


{ TsBubbleSeries }

constructor TsBubbleSeries.Create(AChart: TsChart);
begin
  inherited;
  FChartType := ctBubble;
end;

procedure TsBubbleSeries.SetBubbleRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Bubble series values can only be located in a single column or row.');
  FBubbleRange.Row1 := ARow1;
  FBubbleRange.Col1 := ACol1;
  FBubbleRange.Row2 := ARow2;
  FBubbleRange.Col2 := ACol2;
end;


{ TsLineSeries }

constructor TsLineSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctLine;
  FSymbolWidth := 2.5;
  FSymbolHeight := 2.5;
  FShowSymbols := false;
  FShowLines := true;

  FBorder := TsChartLine.Create;
  FBorder.Style := clsSolid;
  FBorder.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  FBorder.Color := scBlack;
end;

destructor TsLineSeries.Destroy;
begin
  FBorder.Free;
  inherited;
end;

function TsLineSeries.GetSymbolFill: TsChartFill;
begin
  Result := FFill;
end;

procedure TsLineSeries.SetSymbolFill(Value: TsChartFill);
begin
  FFill := Value;
end;


{ TsPieSeries }
constructor TsPieSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctPie;
  FLine.Color := scBlack;
end;


{ TsRadarSeries }
function TsRadarSeries.GetChartType: TsChartType;
begin
  if Fill.Style <> cfsNoFill then
    Result := ctFilledRadar
  else
    Result := ctRadar;
end;


{ TsRingSeries }
constructor TsRingSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctRing;
  FLine.Color := scBlack;
end;


{ TsRegressionEquation }
constructor TsRegressionEquation.Create;
begin
  inherited Create;
  Font := TsFont.Create;
  Font.Size := 9;
  Border := TsChartLine.Create;
  Border.Style := clsNoLine;
  Border.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  Border.Color := scBlack;
  Fill := TsChartFill.Create;
  Fill.Color := scWhite;
  XName := 'x';
  YName := 'f(x)';
end;

destructor TsRegressionEquation.Destroy;
begin
  Fill.Free;
  Border.Free;
  Font.Free;
  inherited;
end;

function TsRegressionEquation.DefaultBorder: Boolean;
begin
  Result := Border.Style = clsNoLine;
end;

function TsRegressionEquation.DefaultFill: Boolean;
begin
  Result := Fill.Style = cfsNoFill;
end;

function TsRegressionEquation.DefaultFont: Boolean;
begin
  Result := (Font.FontName = '') and (Font.Size = 9) and (Font.Style = []) and
            (Font.Color = scBlack);
end;

function TsRegressionEquation.DefaultNumberFormat: Boolean;
begin
  Result := NumberFormat = '';
end;

function TsRegressionEquation.DefaultPosition: Boolean;
begin
  Result := (Left = 0) and (Top = 0);
end;

function TsRegressionEquation.DefaultXName: Boolean;
begin
  Result := XName = 'x';
end;

function TsRegressionEquation.DefaultYName: Boolean;
begin
  Result := YName = 'f(x)';
end;


{ TsChartRegression }
constructor TsChartRegression.Create;
begin
  inherited Create;
  Line := TsChartLine.Create;
  Line.Style := clsSolid;
  Line.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  Line.Color := scBlack;

  Equation := TsRegressionEquation.Create;
end;

destructor TsChartRegression.Destroy;
begin
  Line.Free;
  inherited;
end;


{ TsScatterSeries }

constructor TsScatterSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctScatter;
  FRegression := TsChartRegression.Create;
end;

destructor TsScatterSeries.Destroy;
begin
  FRegression.Free;
  inherited;
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

  FGradients := TsChartGradientList.Create;
  FHatches := TsChartHatchList.Create;

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
  FFloor.Background.Style := cfsNoFill;

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
  FFloor.Free;
  FPlotArea.Free;
  FHatches.Free;
  FGradients.Free;
  FLineStyles.Free;
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

