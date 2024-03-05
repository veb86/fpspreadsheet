unit fpsChart;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

interface

uses
  Classes, SysUtils, Types, Contnrs, FPImage, fpsTypes, fpsUtils;

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
  TsChartColor = record
    Transparency: single;         // 0.0 (opaque) ... 1.0 (transparent)
    case Integer of
      0: (Red, Green, Blue, SystemColorIndex: Byte);
      1: (Color: TsColor);
    end;

const
  sccTransparent: TsChartColor = (Transparency: 255; Color: 0);

type
  TsChart = class;
  TsChartAxis = class;
  TsChartSeries = class;

  TsChartLine = class
    Style: Integer;        // index into chart's LineStyle list or predefined clsSolid/clsNoLine
    Width: Double;         // mm
    Color: TsChartColor;   // in hex: $00bbggrr, r=red, g=green, b=blue
    Transparency: Double;  // in percent
    constructor CreateSolid(AColor: TsChartColor; AWidth: Double);
    procedure CopyFrom(ALine: TsChartLine);
  end;

  TsChartGradientStyle = (cgsLinear, cgsAxial, cgsRadial, cgsElliptic, cgsSquare, cgsRectangular);

  TsChartGradientStep = record
    Value: Double;         // 0.0 ... 1.0
    Color: TsChartColor;
    Intensity: Double;     // 0.0 ... 1.0
  end;

  TsChartGradientSteps = array of TsChartGradientStep;

  TsChartGradient = class
  private
    FSteps: TsChartGradientSteps;
    function GetColor(AIndex: Integer): TsChartColor;
    function GetIntensity(AIndex: Integer): Double;
    function GetSteps(AIndex: Integer): TsChartGradientStep;
    procedure SetStep(AIndex: Integer; AValue: Double; AColor: TsChartColor; AIntensity: Double);
  public
    Name: String;
    Style: TsChartGradientStyle;
    Border: Double;            // 0.0 ... 1.0
    CenterX, CenterY: Double;  // 0.0 ... 1.0
    Angle: Double;             // degrees
    constructor Create;
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartGradient);
    procedure AddStep(AValue: Double; AColor: TsChartColor; AIntensity: Single = 1.0);
    function NumSteps: Integer;
    property Steps[AIndex: Integer]: TsChartGradientStep read GetSteps;
    property StartColor: TsChartColor index 0 read GetColor;
    property StartIntensity: Double index 0 read GetIntensity;
    property EndColor: TsChartColor index 1 read GetColor;
    property EndIntensity: Double index 1 read GetIntensity;
  end;

  TsChartGradientList = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartGradient;
    procedure SetItem(AIndex: Integer; AValue: TsChartGradient);
  public
    function AddGradient(AName: String; AGradient: TsChartGradient): Integer;
    function AddGradient(AName: String; AStyle: TsChartGradientStyle;
      AStartColor, AEndColor: TsChartColor; AStartIntensity, AEndIntensity: Double;
      ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
    function AddAxialGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AStartIntensity, AEndIntensity: Double;
      ABorder, AAngle: Double): Integer;
    function AddEllipticGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AStartIntensity, AEndIntensity: Double;
      ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
    function AddLinearGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AStartIntensity, AEndIntensity: Double;
      ABorder, AAngle: Double): Integer;
    function AddRadialGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AStartIntensity, AEndIntensity: Double;
      ABorder, ACenterX, ACenterY: Double): Integer;
    function AddRectangularGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AStartIntensity, AEndIntensity: Double;
      ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
    function AddSquareGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AStartIntensity, AEndIntensity: Double;
      ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
    function IndexOfName(AName: String): Integer;
    function FindByName(AName: String): TsChartGradient;
    property Items[AIndex: Integer]: TsChartGradient read GetItem write SetItem; default;
  end;

  TsChartHatchStyle = (chsDot, chsSingle, chsDouble, chsTriple);

  TSngPoint = record X, Y: Single; end;

  TsChartHatch = class
    Name: String;
    Style: TsChartHatchStyle;
    PatternColor: TsChartColor;
    PatternWidth: Double;         // Width of pattern (square), in mm if > 0, in px if < 0
    PatternHeight: Double;        // Height of pattern
    PatternAngle: Double;         // Rotation angle of pattern, in degrees
    NumDots: Integer;             // Number of dots within pattern
    DotPos: Array of TSngPoint;   // fraction of dot coordinates in pattern cell
    LineWidth: Single;            // Line width of line pattern, in mm
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartHatch);
  end;

  TsChartHatchList = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartHatch;
    procedure SetItem(AIndex: Integer; AValue: TsChartHatch);
  protected
    function NewPattern(AName: String): Integer;
  public
    function AddDotHatch(AName: String; ADotColor: TsChartColor;
      APatternWidth, APatternHeight: Double;
      ANumDots: Integer; const ADots: array of single): Integer;
    function AddDotHatch(AName: String; ADotColor: TsChartColor;
      APatternWidth, APatternHeight: Integer; ADots: String): Integer;
    function AddLineHatch(AName: String; AStyle: TsChartHatchStyle;
      ALineColor: TsChartColor; ALineDistance, ALineWidth, ALineAngle: Double): Integer;
    function FindByName(AName: String): TsChartHatch;
    function IndexOfName(AName: String): Integer;
    property Items[AIndex: Integer]: TsChartHatch read GetItem write SetItem; default;
  end;

  TsChartImage = class
    Name: String;
    Image: TFPCustomImage;
    Width, Height: Double;  // mm
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartImage);
  end;

  TsChartImagelist = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartImage;
    procedure SetItem(AIndex: Integer; AValue: TsChartImage);
  public
    function AddImage(AName: String; AImage: TFPCustomImage): Integer;
    function FindByName(AName: String): TsChartImage;
    function IndexOfName(AName: String): Integer;
    property Items[Aindex: Integer]: TsChartImage read GetItem write SetItem; default;
  end;

  TsChartFillStyle = (cfsNoFill, cfsSolid, cfsGradient, cfsHatched, cfsSolidHatched, cfsImage);

  TsChartFill = class
  public
    Style: TsChartFillStyle;
    Color: TsChartColor;
    Gradient: Integer;     // Index into chart's Gradients list
    Hatch: Integer;        // Index into chart's Hatches list
    Image: Integer;        // Index into chart's Images list
    constructor CreateSolidFill(AColor: TsChartColor);
    constructor CreateHatchFill(AHatchIndex: Integer; ABkColor: TsChartColor);
    procedure CopyFrom(AFill: TsChartFill);
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
    procedure CopyFrom(ASource: TsChartLineStyle);
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
    function IndexOfName(AName: String): Integer;
    property Items[AIndex: Integer]: TsChartLineStyle read GetItem write SetItem; default;
  end;

  TsChartCellAddr = class
  private
    FChart: TsChart;
  public
    Sheet: String;
    Row, Col: Cardinal;
    constructor Create(AChart: TsChart);
    procedure CopyFrom(ASource: TsChartCellAddr);
    function GetSheetName: String;
    function IsUsed: Boolean;
    property Chart: TsChart read FChart;
  end;

  TsChartRange = class
  private
    FChart: TsChart;
  public
    Sheet1, Sheet2: String;
    Row1, Col1, Row2, Col2: Cardinal;
    constructor Create(AChart: TsChart);
    procedure CopyFrom(ASource: TsChartRange);
    function NumCells: Integer;
    function GetSheet1Name: String;
    function GetSheet2Name: String;
    function IsEmpty: Boolean;
    property Chart: TsChart read FChart;
  end;

  TsChartElement = class
  private
    FChart: TsChart;
    FVisible: Boolean;
  protected
    function GetVisible: Boolean; virtual;
    procedure SetVisible(AValue: Boolean); virtual;
  public
    constructor Create(AChart: TsChart);
    procedure CopyFrom(ASource: TsChartElement); virtual;
    property Chart: TsChart read FChart;
    property Visible: Boolean read GetVisible write SetVisible;
  end;

  TsChartFillElement = class(TsChartElement)
  private
    FBackground: TsChartFill;
    FBorder: TsChartLine;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartElement); override;
    property Background: TsChartFill read FBackground write FBackground;
    property Border: TsChartLine read FBorder write FBorder;
  end;

  TsChartText = class(TsChartFillElement)
  private
    FCaption: String;
    FRotationAngle: single;
    FFont: TsFont;
    FPosX, FPosY: Double;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartElement); override;
    property Caption: String read FCaption write FCaption;
    property Font: TsFont read FFont write FFont;
    property RotationAngle: single read FRotationAngle write FRotationAngle;
    property PosX: Double read FPosX write FPosX;
    property PosY: Double read FPosY write FPosY;
    property Visible;
  end;

  TsChartAxisAlignment = (caaLeft, caaTop, caaRight, caaBottom);
  TsChartAxisPosition = (capStart, capEnd, capValue);
  TsChartAxisTick = (catInside, catOutside);
  TsChartAxisTicks = set of TsChartAxisTick;
  TsChartType = (ctEmpty, ctBar, ctLine, ctArea, ctBarLine, ctScatter, ctBubble,
    ctRadar, ctFilledRadar, ctPie, ctRing, ctStock);

  TsChartAxis = class(TsChartFillElement)
  private
    FAlignment: TsChartAxisAlignment;
    FAutomaticMax: Boolean;
    FAutomaticMin: Boolean;
    FAutomaticMajorInterval: Boolean;
    FAutomaticMinorInterval: Boolean;
    FAutomaticMinorSteps: Boolean;
    FAxisLine: TsChartLine;
    FCategoryRange: TsChartRange;
    FDefaultTitleRotation: Boolean;
    FMajorGridLines: TsChartLine;
    FMinorGridLines: TsChartline;
    FInverted: Boolean;
    FLabelFont: TsFont;
    FLabelFormat: String;
    FLabelFormatFromSource: Boolean;
    FLabelFormatDateTime: String;
    FLabelFormatPercent: String;
    FLabelRotation: Single;
    FLogarithmic: Boolean;
    FLogBase: Double;
    FMajorInterval: Double;
    FMajorTicks: TsChartAxisTicks;
    FMax: Double;
    FMin: Double;
    FMinorCount: Integer;
    FMinorInterval: Double;
    FMinorTicks: TsChartAxisTicks;
    FPosition: TsChartAxisPosition;
    FTitle: TsChartText;
    FPositionValue: Double;
    FShowLabels: Boolean;
    FDateTime: Boolean;
    function GetTitleRotationAngle: Single;
    procedure SetMax(AValue: Double);
    procedure SetMin(AValue: Double);
    procedure SetMinorCount(AValue: Integer);
    procedure SetMajorInterval(AValue: Double);
    procedure SetMinorInterval(AValue: Double);
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartElement); override;
    function GetOtherAxis: TsChartAxis;
    function GetRotatedAxis: TsChartAxis;
    procedure SetCategoryRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetCategoryRange(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    property Alignment: TsChartAxisAlignment read FAlignment write FAlignment;
    property AutomaticMax: Boolean read FAutomaticMax write FAutomaticMax;
    property AutomaticMin: Boolean read FAutomaticMin write FAutomaticMin;
    property AutomaticMajorInterval: Boolean read FAutomaticMajorInterval write FAutomaticMajorInterval;
    property AutomaticMinorInterval: Boolean read FAutomaticMinorInterval write FAutomaticMinorInterval;
    property AutomaticMinorSteps: Boolean read FAutomaticMinorSteps write FAutomaticMinorSteps;
    property AxisLine: TsChartLine read FAxisLine write FAxisLine;
    property CategoryRange: TsChartRange read FCategoryRange write FCategoryRange;
    property DateTime: Boolean read FDateTime write FDateTime;
    property DefaultTitleRotation: Boolean read FDefaultTitleRotation write FDefaultTitleRotation;
    property Inverted: Boolean read FInverted write FInverted;
    property LabelFont: TsFont read FLabelFont write FLabelFont;
    property LabelFormat: String read FLabelFormat write FLabelFormat;
    property LabelFormatDateTime: String read FLabelFormatDateTime write FLabelFormatDateTime;
    property LabelFormatFromSource: Boolean read FLabelFormatFromSource write FLabelFormatFromSource;
    property LabelFormatPercent: String read FLabelFormatPercent write FLabelFormatPercent;
    property LabelRotation: Single read FLabelRotation write FLabelRotation;
    property Logarithmic: Boolean read FLogarithmic write FLogarithmic;
    property LogBase: Double read FLogBase write FLogBase;
    property MajorGridLines: TsChartLine read FMajorGridLines write FMajorGridLines;
    property MajorInterval: Double read FMajorInterval write SetMajorInterval;
    property MajorTicks: TsChartAxisTicks read FMajorTicks write FMajorTicks;
    property Max: Double read FMax write SetMax;
    property Min: Double read FMin write SetMin;
    property MinorGridLines: TsChartLine read FMinorGridLines write FMinorGridLines;
    property MinorCount: Integer read FMinorCount write SetMinorCount;
    property MinorInterval: Double read FMinorInterval write SetMinorInterval;
    property MinorTicks: TsChartAxisTicks read FMinorTicks write FMinorTicks;
    // Position and PositionValue define where the axis is crossed by the other axis
    property Position: TsChartAxisPosition read FPosition write FPosition;
    property PositionValue: Double read FPositionValue write FPositionValue;
    property ShowLabels: Boolean read FShowLabels write FShowLabels;
    property Title: TsChartText read FTitle write FTitle;
    property TitleRotationAngle: Single read GetTitleRotationAngle;
    property Visible;
  end;

  TsChartLegendPosition = (lpRight, lpTop, lpBottom, lpLeft);

  TsChartLegend = class(TsChartFillElement)
  private
    FFont: TsFont;
    FCanOverlapPlotArea: Boolean;
    FPosition: TsChartLegendPosition;
    FPosX, FPosY: Double;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartElement); override;
    property CanOverlapPlotArea: Boolean read FCanOverlapPlotArea write FCanOverlapPlotArea;
    property Font: TsFont read FFont write FFont;
    property Position: TsChartLegendPosition read FPosition write FPosition;
    property PosX: Double read FPosX write FPosX;
    property PosY: Double read FPosY write FPosY;
    // There is also a "legend-expansion" but this does not seem to have a visual effect in Calc.
  end;

  TsChartAxisLink = (calPrimary, calSecondary);
  TsChartDataLabel = (cdlValue, cdlPercentage, cdlCategory, cdlSeriesName, cdlSymbol, cdlLeaderLines);
  TsChartDataLabels = set of TsChartDataLabel;
  TsChartLabelPosition = (lpDefault, lpOutside, lpInside, lpCenter, lpAbove, lpBelow, lpNearOrigin);
  TsChartLabelCalloutShape = (
    lcsRectangle, lcsRoundRect, lcsEllipse,
    lcsLeftArrow, lcsUpArrow, lcsRightArrow, lcsDownArrow,
    lcsRectangleWedge, lcsRoundRectWedge, lcsEllipseWedge
  );

  TsChartDataPointStyle = class(TsChartFillElement)
  private
    FDataPointIndex: Integer;
    FPieOffset: Integer;
  public
    procedure CopyFrom(ASource: TsChartElement);
    property DataPointIndex: Integer read FDataPointIndex write FDataPointIndex;
    property PieOffset: Integer read FPieOffset write FPieOffset;  // Percentage
  end;

  TsChartDataPointStyleList = class(TFPObjectList)
  private
    FChart: TsChart;
    function GetItem(AIndex: Integer): TsChartDataPointStyle;
    procedure SetItem(AIndex: Integer; AValue: TsChartDataPointStyle);
  public
    constructor Create(AChart: TsChart);
    function AddFillAndLine(ADataPointIndex: Integer; AFill: TsChartFill; ALine: TsChartline; APieOffset: Integer = 0): Integer;
    function AddSolidFill(ADataPointIndex: Integer; AColor: TsChartColor; ALine: TsChartLine = nil; APieOffset: Integer = 0): Integer;
    function IndexOfDataPoint(ADataPointIndex: Integer): Integer;
    property Items[AIndex: Integer]: TsChartDataPointStyle read GetItem write SetItem; default;
  end;

  TsTrendlineType = (tltNone, tltLinear, tltLogarithmic, tltExponential, tltPower, tltPolynomial);

  TsTrendlineEquation = class
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

  TsChartTrendline = class
    Title: String;
    TrendlineType: TsTrendlineType;
    ExtrapolateForwardBy: Double;
    ExtrapolateBackwardBy: Double;
    ForceYIntercept: Boolean;
    YInterceptValue: Double;
    PolynomialDegree: Integer;
    DisplayEquation: Boolean;
    DisplayRSquare: Boolean;
    Equation: TsTrendlineEquation;
    Line: TsChartLine;
    constructor Create;
    destructor Destroy; override;
  end;

  TsChartErrorBarKind = (cebkNone, cebkConstant, cebkPercentage, cebkCellRange);

  TsChartErrorBars = class(TsChartElement)
  private
    FSeries: TsChartSeries;
    FKind: TsChartErrorBarKind;
    FLine: TsChartLine;
    FRange: Array[0..1] of TsChartRange;
    FValue: Array[0..1] of Double;  // 0 = positive, 1 = negative error bar value
    FShow: Array[0..1] of Boolean;
    FShowEndCap: Boolean;
    function GetRange(AIndex: Integer): TsChartRange;
    function GetShow(AIndex: Integer): Boolean;
    function GetValue(AIndex: Integer): Double;
    procedure InternalSetErrorBarRange(AIndex: Integer;
      ASheet1: String; ARow1, ACol1: Cardinal;
      ASheet2: String; ARow2, ACol2: Cardinal);
    procedure SetLine(AValue: TsChartLine);
    procedure SetKind(AValue: TsChartErrorBarKind);
    procedure SetRange(AIndex: Integer; AValue: TsChartRange);
    procedure SetShow(AIndex: Integer; AValue: Boolean);
    procedure SetValue(AIndex: Integer; AValue: Double);
  protected
    function GetVisible: Boolean; override;
    procedure SetVisible(AValue: Boolean); override;
  public
    constructor Create(ASeries: TsChartSeries);
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartElement); override;
    procedure SetErrorBarRangePos(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetErrorBarRangePos(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    procedure SetErrorBarRangeNeg(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetErrorBarRangeNeg(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    property Kind: TsChartErrorBarKind read FKind write SetKind;
    property Line: TsChartLine read FLine write SetLine;
    property RangePos: TsChartRange index 0 read GetRange write SetRange;
    property RangeNeg: TsChartRange index 1 read GetRange write SetRange;
    property Series: TsChartSeries read FSeries;
    property ShowEndCap: Boolean read FShowEndCap write FShowEndCap;
    property ShowPos: Boolean index 0 read GetShow write SetShow;
    property ShowNeg: Boolean index 1 read GetShow write SetShow;
    property ValuePos: Double index 0 read GetValue write SetValue;
    property ValueNeg: Double index 1 read GetValue write SetValue;
  end;

  TsChartSeries = class(TsChartElement)
  private
    FChartType: TsChartType;
    FXRange: TsChartRange;          // cell range containing the x data
    FYRange: TsChartRange;          // ... and the y data
    FFillColorRange: TsChartRange;
    FLineColorRange: TsChartRange;
    FLabelBackground: TsChartFill;
    FLabelBorder: TsChartLine;
    FLabelRange: TsChartRange;
    FLabelFont: TsFont;
    FLabelPosition: TsChartLabelPosition;
    FLabelSeparator: string;
    FXAxis: TsChartAxisLink;
    FYAxis: TsChartAxisLink;
    FTitleAddr: TsChartCellAddr;
    FLabelFormat: String;
    FLabelFormatPercent: String;
    FDataLabels: TsChartDataLabels;
    FDataLabelCalloutShape: TsChartLabelCalloutShape;
    FDataPointStyles: TsChartDataPointStyleList;
    FOrder: Integer;
    FTrendline: TsChartTrendline;
    FSupportsTrendline: Boolean;
    FXErrorBars: TsChartErrorBars;
    FYErrorBars: TsChartErrorBars;
    FGroupIndex: Integer;  // series with the same GroupIndex can be stacked
    procedure SetXErrorBars(AValue: TsChartErrorBars);
    procedure SetYErrorBars(AValue: TsChartErrorBars);
  protected
    FLine: TsChartLine;
    FFill: TsChartFill;
    function GetChartType: TsChartType; virtual;
    property Trendline: TsChartTrendline read FTrendline write FTrendline;
  public
    constructor Create(AChart: TsChart); virtual;
    destructor Destroy; override;
    function GetCount: Integer;
    function GetXAxis: TsChartAxis;
    function GetYAxis: TsChartAxis;
    function GetXCount: Integer;
    function GetYCount: Integer;
    function HasLabels: Boolean;
    function HasXValues: Boolean;
    function HasYValues: Boolean;
    procedure SetTitleAddr(ARow, ACol: Cardinal);
    procedure SetTitleAddr(ASheet: String; ARow, ACol: Cardinal);
    procedure SetLabelRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetLabelRange(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    procedure SetXRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetXRange(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    procedure SetYRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetYRange(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    procedure SetFillColorRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetFillColorRange(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    procedure SetLineColorRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetLineColorRange(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    function LabelsInCol: Boolean;
    function XValuesInCol: Boolean;
    function YValuesInCol: Boolean;

    property ChartType: TsChartType read GetChartType;
    property Count: Integer read GetCount;
    property DataLabels: TsChartDataLabels read FDataLabels write FDataLabels;
    property DataLabelCalloutShape: TsChartLabelCalloutShape read FDataLabelCalloutShape write FDataLabelCalloutShape;
    property DataPointStyles: TsChartDatapointStyleList read FDataPointStyles;
    property FillColorRange: TsChartRange read FFillColorRange write FFillColorRange;
    property GroupIndex: Integer read FGroupIndex write FGroupIndex;
    property LabelBackground: TsChartFill read FLabelBackground write FLabelBackground;
    property LabelBorder: TsChartLine read FLabelBorder write FLabelBorder;
    property LabelFont: TsFont read FLabelFont write FLabelFont;
    property LabelFormat: String read FLabelFormat write FLabelFormat;  // Number format in Excel notation, e.g. '0.00'
    property LabelFormatPercent: String read FLabelFormatPercent write FLabelFormatPercent;
    property LabelPosition: TsChartLabelPosition read FLabelPosition write FLabelPosition;
    property LabelRange: TsChartRange read FLabelRange write FLabelRange;
    property LabelSeparator: string read FLabelSeparator write FLabelSeparator;
    property LineColorRange: TsChartRange read FLineColorRange write FLineColorRange;
    property Order: Integer read FOrder write FOrder;
    property TitleAddr: TsChartCellAddr read FTitleAddr write FTitleAddr;  // use '\n' for line-break
    property SupportsTrendline: Boolean read FSupportsTrendline;
    property XAxis: TsChartAxisLink read FXAxis write FXAxis;
    property XErrorBars: TsChartErrorBars read FXErrorBars write SetXErrorBars;
    property XRange: TsChartRange read FXRange write FXRange;
    property YAxis: TsChartAxisLink read FYAxis write FYAxis;
    property YErrorBars: TsChartErrorBars read FYErrorBars write SetYErrorBars;
    property YRange: TsChartRange read FYRange write FYRange;

    property Fill: TsChartFill read FFill write FFill;
    property Line: TsChartLine read FLine write FLine;
  end;
  TsChartSeriesClass = class of TsChartSeries;

  TsAreaSeries = class(TsChartSeries)
  public
    constructor Create(AChart: TsChart); override;
    property Trendline;
  end;

  TsBarSeries = class(TsChartSeries)
  public
    constructor Create(AChart: TsChart); override;
    property Trendline;
  end;

  TsChartSeriesSymbol = (
    cssRect, cssDiamond, cssTriangle, cssTriangleDown, cssTriangleLeft,
    cssTriangleRight, cssCircle, cssStar, cssX, cssPlus, cssAsterisk,
    cssDash, cssDot
  );

  TsChartInterpolation = (
    ciLinear,
    ciCubicSpline, ciBSpline,
    ciStepStart, ciStepEnd, ciStepCenterX, ciStepCenterY
  );

  TsCustomLineSeries = class(TsChartSeries)
  private
    FInterpolation: TsChartInterpolation;
    FSymbol: TsChartSeriesSymbol;
    FSymbolHeight: Double;  // in mm
    FSymbolWidth: Double;   // in mm
    FShowLines: Boolean;
    FShowSymbols: Boolean;
    FSymbolBorder: TsChartLine;
    FSymbolFill: TsChartFill;
    function GetSmooth: Boolean;
    procedure SetSmooth(AValue: Boolean);
  protected
    property Interpolation: TsChartInterpolation read FInterpolation write FInterpolation;
    property Smooth: Boolean read GetSmooth write SetSmooth;
    property Symbol: TsChartSeriesSymbol read FSymbol write FSymbol;
    property SymbolBorder: TsChartLine read FSymbolBorder write FSymbolBorder;
    property SymbolFill: TsChartFill read FSymbolFill write FSymbolFill;
    property SymbolHeight: double read FSymbolHeight write FSymbolHeight;
    property SymbolWidth: double read FSymbolWidth write FSymbolWidth;
    property ShowLines: Boolean read FShowLines write FShowLines;
    property ShowSymbols: Boolean read FShowSymbols write FShowSymbols;
  public
    constructor Create(AChart: TsChart); override;
    destructor Destroy; override;
  end;

  TsLineSeries = class(TsCustomLineSeries)
  public
    constructor Create(AChart: TsChart); override;
    property Interpolation;
    property Smooth;
    property Symbol;
    property SymbolBorder;
    property SymbolFill;
    property SymbolHeight;
    property SymbolWidth;
    property ShowLines;
    property ShowSymbols;
    property Trendline;
  end;

  TsSliceOrder = (soCCW, soCW);

  TsPieSeries = class(TsChartSeries)
  private
    FInnerRadiusPercent: Integer;
    FSliceOrder: TsSliceOrder;
    FStartAngle: Integer;         // degrees
    function GetSliceOffset(ASliceIndex: Integer): Integer;
  protected
    function GetChartType: TsChartType; override;
  public
    constructor Create(AChart: TsChart); override;
    property InnerRadiusPercent: Integer read FInnerRadiusPercent write FInnerRadiusPercent;
    property StartAngle: Integer read FStartAngle write FStartAngle;
    property SliceOffset[ASliceIndex: Integer]: Integer read GetSliceOffset;  // Percentage
    property SliceOrder: TsSliceOrder read FSliceOrder write FSliceOrder;
  end;

  TsRadarSeries = class(TsLineSeries)
  public
    constructor Create(AChart: TsChart); override;
  end;

  TsFilledRadarSeries = class(TsRadarSeries)
  public
    constructor Create(AChart: TsChart); override;
  end;

  TsCustomScatterSeries = class(TsCustomLineSeries)
  public
    constructor Create(AChart: TsChart); override;
    property Trendline;
  end;

  TsScatterSeries = class(TsCustomScatterSeries)
  public
    property Interpolation;
    property Smooth;
    property Symbol;
    property SymbolBorder;
    property SymbolFill;
    property SymbolHeight;
    property SymbolWidth;
    property ShowLines;
    property ShowSymbols;
  end;

  TsBubbleSizeMode = (bsmRadius, bsmArea);

  TsBubbleSeries = class(TsCustomScatterSeries)
  private
    FBubbleRange: TsChartRange;
    FBubbleScale: Double;
    FBubbleSizeMode: TsBubbleSizeMode;
  public
    constructor Create(AChart: TsChart); override;
    destructor Destroy; override;
    procedure SetBubbleRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetBubbleRange(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    property BubbleRange: TsChartRange read FBubbleRange;
    property BubbleScale: Double read FBubbleScale write FBubbleScale;
    property BubbleSizeMode: TsBubbleSizeMode read FBubbleSizeMode write FBubbleSizeMode;
  end;

  TsStockSeries = class(TsChartSeries)  //CustomScatterSeries)
  private
    FCandleStick: Boolean;
    FOpenRange: TsChartRange;
    FHighRange: TsChartRange;
    FLowRange: TsChartRange;  // close = normal y range
    FCandleStickDownFill: TsChartFill;
    FCandleStickDownBorder: TsChartLine;
    FCandleStickUpFill: TsChartFill;
    FCandleStickUpBorder: TsChartLine;
    FRangeLine: TsChartLine;
    FTickWidthPercent: Integer;
    // fill is CandleStickUpFill, line is RangeLine
  public
    constructor Create(AChart: TsChart); override;
    destructor Destroy; override;
    procedure SetOpenRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetOpenRange(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    procedure SetHighRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetHighRange(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    procedure SetLowRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetLowRange (ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    procedure SetCloseRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
    procedure SetCloseRange(ASheet1: String; ARow1, ACol1: Cardinal; ASheet2: String; ARow2, ACol2: Cardinal);
    property CandleStick: Boolean read FCandleStick write FCandleStick;
    property CandleStickDownFill: TsChartFill read FCandleStickDownFill write FCandleStickDownFill;
    property CandleStickUpFill: TsChartFill read FCandleStickUpFill write FCandleStickUpFill;
    property CandleStickDownBorder: TsChartLine read FCandleStickDownBorder write FCandleStickDownBorder;
    property CandleStickUpBorder: TsChartLine read FCandleStickUpBorder write FCandleStickUpBorder;
    property TickWidthPercent: Integer read FTickWidthPercent write FTickWidthPercent;
    property RangeLine: TsChartLine read FRangeLine write FRangeLine;
    property OpenRange: TsChartRange read FOpenRange;
    property HighRange: TsChartRange read FHighRange;
    property LowRange: TsChartRange read FLowRange;
    property CloseRange: TsChartRange read FYRange;
  end;

  TsChartSeriesList = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartSeries;
    procedure SetItem(AIndex: Integer; AValue: TsChartSeries);
  public
    property Items[AIndex: Integer]: TsChartSeries read GetItem write SetItem; default;
  end;

  TsChartStackMode = (csmDefault, csmStacked, csmStackedPercentage);

  TsChart = class(TsChartFillElement)
  private
    FName: String;
    FIndex: Integer;             // Index in workbook's chart list
    FWorkbook: TsBasicWorkbook;
    FWorksheet: TsBasicWorksheet;
//    FSheetIndex: Integer;
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
    FBarGapWidthPercent: Integer; // For bar series: distance between bars (relative to single bar width)
    FBarOverlapPercent: Integer;  // For bar series: overlap between bars

    FTitle: TsChartText;
    FSubTitle: TsChartText;
    FLegend: TsChartLegend;
    FSeriesList: TsChartSeriesList;

    FLineStyles: TsChartLineStyleList;
    FGradients: TsChartGradientList;
    FHatches: TsChartHatchList;
    FImages: TsChartImageList;

    function GetCategoryLabelRange: TsChartRange;

  protected
    function AddSeries(ASeries: TsChartSeries): Integer; virtual;

  public
    constructor Create;
    destructor Destroy; override;
//    function GetWorksheet: TsBasicWorksheet;

    procedure DeleteSeries(AIndex: Integer);

    function GetChartType: TsChartType;
    function GetLineStyle(AIndex: Integer): TsChartLineStyle;
    function IsScatterChart: Boolean;
    function NumLineStyles: Integer;

    { Name for internal purposes to identify the chart during reading from file }
    property Name: String read FName write FName;
    { Index of chart in workbook's chart list. }
    property Index: Integer read FIndex write FIndex;
    { Worksheet into which the chart is embedded }
    property Worksheet: TsBasicWorksheet read FWorksheet write FWorksheet;
    (*
    { Index of worksheet sheet which contains the chart. }
    property SheetIndex: Integer read FSheetIndex write FSheetIndex;
    *)
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
    { Workbook to which the chart belongs }
    property Workbook: TsBasicWorkbook read FWorkbook write FWorkbook;

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

    { Gap between bars/bar groups, as percentage of single bar width }
    property BarGapWidthPercent: Integer read FBarGapWidthPercent write FBarGapWidthPercent;
    { Overlapping of bars, as percentage of single bar width }
    property BarOverlapPercent: Integer read FBarOverlapPercent write FBarOverlapPercent;

    { Connecting line between data points (for line and scatter series) }
    property Interpolation: TsChartInterpolation read FInterpolation write FInterpolation;
    { x and y axes exchanged (mainly for bar series, but works also for scatter and bubble series) }
    property RotatedAxes: Boolean read FRotatedAxes write FRotatedAxes;
    { Stacking of series (for bar and area series ) }
    property StackMode: TsChartStackMode read FStackMode write FStackMode;

    property CategoryLabelRange: TsChartRange read GetCategoryLabelRange;

    { Attributes of the series }
    property Series: TsChartSeriesList read FSeriesList write FSeriesList;

    { Style lists }
    property LineStyles: TsChartLineStyleList read FLineStyles;
    property Gradients: TsChartGradientList read FGradients;
    property Hatches: TsChartHatchList read FHatches;
    property Images: TsChartImageList read FImages;
  end;

  TsChartList = class(TObjectList)
  private
    function GetItem(AIndex: Integer): TsChart;
    procedure SetItem(AIndex: Integer; AValue: TsChart);
  public
    property Items[AIndex: Integer]: TsChart read GetItem write SetItem; default;
  end;


function ChartColor(AColor: TsColor; ATransparency: Single = 0.0): TsChartColor;

implementation

uses
  Math, fpSpreadsheet;

{ TsChartColor }

function ChartColor(AColor: TsColor; ATransparency: Single = 0.0): TsChartColor;
begin
  Result.Color := AColor;
  Result.Transparency := ATransparency;
end;


{ TsChartLine }

constructor TsChartLine.CreateSolid(AColor: TsChartColor; AWidth: Double);
begin
  inherited Create;
  Style := clsSolid;
  Color := AColor;
  Width := AWidth;
end;

procedure TsChartLine.CopyFrom(ALine: TsChartLine);
begin
  if ALine <> nil then
  begin
    Style := ALine.Style;
    Width := ALine.Width;
    Color := ALine.Color;
    Transparency := ALine.Transparency;
  end;
end;


{ TsChartGradient }

constructor TsChartGradient.Create;
begin
  inherited Create;
  SetLength(FSteps, 2);
  SetStep(0, 0.0, ChartColor(scBlack), 1.0);
  SetStep(1, 1.0, ChartColor(scWhite), 1.0);
end;

destructor TsChartGradient.Destroy;
begin
  Name := '';
  inherited;
end;

{ Adds a new color step to the gradient. The new color is inserted at the
  correct index according to its value so that all values in the steps are
  ordered. If the exact value is already existing the gradient step is replaced.}
procedure TsChartGradient.AddStep(AValue: Double; AColor: TsChartColor;
  AIntensity: Single = 1.0);
var
  i, j, idx: Integer;
begin
  if AValue < 0 then AValue := 0.0;
  if AValue > 1 then AValue := 1.0;

  if Length(FSteps) > 0 then
  begin
    for i := 0 to High(FSteps) do
    begin
      if FSteps[i].Value = AValue then
      begin
        idx := i;
        break;
      end else
      if FSteps[i].Value > AValue then
      begin
        idx := i;
        SetLength(FSteps, Length(FSteps) + 1);
        for j := High(FSteps) downto i do
          FSteps[j] := FSteps[j-1];
        break;
      end;
    end;
  end else
  begin
    SetLength(FSteps, 1);
    idx := 0;
  end;
  SetStep(idx, AValue, AColor, AIntensity);
end;

procedure TsChartGradient.CopyFrom(ASource: TsChartGradient);
var
  i: Integer;
begin
  Name := ASource.Name;
  Style := ASource.Style;
  SetLength(FSteps, ASource.NumSteps);
  for i := 0 to Length(FSteps)-1 do
    FSteps[i] := ASource.Steps[i];
  Border := ASource.Border;
  CenterX := ASource.CenterX;
  CenterY := ASource.CenterY;
  Angle := ASource.Angle;
end;

function TsChartGradient.GetColor(AIndex: Integer): TsChartColor;
begin
  case AIndex of
    0: Result := FSteps[0].Color;
    1: Result := FSteps[High(FSteps)].Color;
  end;
end;

function TsChartGradient.GetIntensity(AIndex: Integer): Double;
begin
  case AIndex of
    0: Result := FSteps[0].Intensity;
    1: Result := FSteps[High(FSteps)].Intensity;
  end;
end;

function TsChartGradient.GetSteps(AIndex: Integer): TsChartGradientStep;
begin
  if AIndex < 0 then AIndex := 0;
  if AIndex >= Length(FSteps) then AIndex := Length(FSteps) - 1;
  Result := FSteps[AIndex];
end;

function TsChartGradient.NumSteps: Integer;
begin
  Result := Length(FSteps);
end;

procedure TsChartGradient.SetStep(AIndex: Integer; AValue: Double;
  AColor: TsChartColor; AIntensity: Double);
begin
  FSteps[AIndex].Value := AValue;
  FSteps[AIndex].Color := AColor;
  FSteps[AIndex].Intensity := AIntensity;
end;


{ TsChartGradientList }

function TsChartGradientList.AddAxialGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AStartIntensity, AEndIntensity: Double;
  ABorder, AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsAxial,
    AStartColor, AEndColor,
    AStartIntensity, AEndIntensity,
    ABorder, 0.0, 0.0, AAngle
  );
end;

function TsChartGradientList.AddEllipticGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AStartIntensity, AEndIntensity: Double;
  ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsElliptic,
    AStartColor, AEndColor,
    AStartIntensity, AEndIntensity,
    ABorder, ACenterX, ACenterY, AAngle
  );
end;

function TsChartGradientList.AddGradient(AName: String; AGradient: TsChartGradient): Integer;
begin
  if AName = '' then
    AName := 'G' + IntToStr(Count + 1);
  Result := IndexOfName(AName);
  if Result = -1 then
    Result := inherited Add(AGradient)
  else
    Items[Result].CopyFrom(AGradient);
end;

function TsChartGradientList.AddGradient(AName: String; AStyle: TsChartGradientStyle;
  AStartColor, AEndColor: TsChartColor; AStartIntensity, AEndIntensity: Double;
  ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
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
  item.AddStep(0.0, AStartColor,  AStartIntensity);
  item.AddStep(1.0, AEndColor, AEndIntensity);
  item.Border := ABorder;
  item.Angle := AAngle;
  item.CenterX := ACenterX;
  item.CenterY := ACenterY;
end;

function TsChartGradientList.AddLinearGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AStartIntensity, AEndIntensity: Double;
  ABorder,AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsLinear,
    AStartColor, AEndColor, AStartIntensity, AEndIntensity,
    ABorder, 0.0, 0.0, AAngle
  );
end;

function TsChartGradientList.AddRadialGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AStartIntensity, AEndIntensity: Double;
  ABorder, ACenterX, ACenterY: Double): Integer;
begin
  Result := AddGradient(AName, cgsRadial,
    AStartColor, AEndColor,
    AStartIntensity, AEndIntensity,
    ABorder, ACenterX, ACenterY, 0
  );
end;

function TsChartGradientList.AddRectangularGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AStartIntensity, AEndIntensity: Double;
  ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsRectangular,
    AStartColor, AEndColor,
    AStartIntensity, AEndIntensity,
    ABorder, ACenterX, ACenterY, AAngle
  );
end;

function TsChartGradientList.AddSquareGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AStartIntensity, AEndIntensity: Double;
  ABorder, ACenterX, ACenterY, AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsSquare,
    AStartColor, AEndColor,
    AStartIntensity, AEndIntensity,
    ABorder, ACenterX, ACenterY, AAngle
  );
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
  TsChartGradient(inherited Items[AIndex]).CopyFrom(AValue);
end;


{ TsChartHatch }

destructor TsChartHatch.Destroy;
begin
  Name := '';
  inherited;
end;

procedure TsChartHatch.CopyFrom(ASource: TsChartHatch);
var
  i: Integer;
begin
  Name := ASource.Name;
  Style := ASource.Style;
  PatternColor := ASource.PatternColor;
  PatternWidth := ASource.PatternWidth;
  PatternHeight := ASource.PatternHeight;
  PatternAngle := ASource.PatternAngle;
  NumDots := ASource.NumDots;
  SetLength(DotPos, Length(ASource.DotPos));
  for i := 0 to High(DotPos) do DotPos[i] := ASource.DotPos[i];
  LineWidth := ASource.LineWidth;
end;


{ TsChartHatchList }

function TsChartHatchList.AddDotHatch(AName: String; ADotColor: TsChartColor;
  APatternWidth, APatternHeight: Double;
  ANumDots: Integer; const ADots: array of single): Integer;
var
  item: TsChartHatch;
  i, j: Integer;
begin
  Result := NewPattern(AName);
  item := Items[Result];
  item.Name := AName;
  item.Style := chsDot;
  item.PatternColor := ADotColor;
  item.PatternWidth := APatternWidth;  // in millimeters (> 0), in px (< 0)
  item.PatternHeight := APatternHeight;
  item.PatternAngle:= 0.0;
  item.NumDots := ANumDots;
  j := 0;
  SetLength(item.DotPos, Length(ADots) div 2);
  for i := 0 to High(item.DotPos) do
  begin
    item.DotPos[i].X := ADots[j];
    item.DotPos[i].Y := ADots[j+1];
    inc(j, 2);
  end;
end;

function TsChartHatchList.AddDotHatch(AName: String; ADotColor: TsChartColor;
  APatternWidth, APatternHeight: Integer; ADots: String): Integer;
var
  i, x, y: Integer;
  w, h: Integer;
  dots: array of single = nil;
  nDots: Integer;
begin
  w := APatternWidth;
  if w < 0 then w := -w;

  h := APatternHeight;
  if h < 0 then h := -h;

  if Length(ADots) <> w*h then
    raise Exception.Create('Hatch pattern error.');

  x := 0;
  y := 0;
  nDots := 0;
  SetLength(dots, Length(ADots)*2);
  for i := 1 to Length(ADots) do
  begin
    if ADots[i] in ['x', 'X', '*'] then
    begin
      dots[nDots] := -1.0 * x;
      dots[nDots+1] := -1.0 * y;
      inc(nDots, 2);
    end;
    inc(x);
    if x = w then
    begin
      inc(y);
      x := 0;
    end;
  end;
  SetLength(dots, nDots);
  Result := AddDotHatch(AName, ADotColor, -w, -h, nDots div 2, dots);
end;

function TsChartHatchList.AddLineHatch(AName: String; AStyle: TsChartHatchStyle;
  ALineColor: TsChartColor; ALineDistance, ALineWidth, ALineAngle: Double): Integer;
var
  item: TsChartHatch;
begin
  if not (AStyle in [chsSingle, chsDouble, chsTriple]) then
    exit(-1);

  Result := NewPattern(AName);
  item := Items[Result];
  item.Name := AName;
  item.Style := AStyle;
  item.PatternColor := ALineColor;
  item.PatternWidth := ALineDistance;
  item.PatternAngle := ALineAngle;
  item.LineWidth := ALineWidth;
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

function TsChartHatchList.NewPattern(AName: String): integer;
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
  end;
end;

procedure TsChartHatchList.SetItem(AIndex: Integer; AValue: TsChartHatch);
begin
  TsChartHatch(inherited Items[AIndex]).CopyFrom(AValue);
end;


{ TsChartImage }

destructor TsChartImage.Destroy;
begin
  Name := '';
  //Image.Free;  --- created by caller --> must be destroyed by caller!
  inherited;
end;

procedure TsChartImage.CopyFrom(ASource: TsChartImage);
begin
  Name := ASource.Name;
  Image := ASource.Image;
  Width := ASource.Width;
  Height := ASource.Height;
end;


{ TsChartImageList }

function TsChartImageList.AddImage(AName: String; AImage: TFPCustomImage): Integer;
var
  item: TsChartImage;
begin
  if AName = '' then
    AName := 'Img' + IntToStr(Count + 1);
  Result := IndexOfName(AName);
  if Result = -1 then
  begin
    item := TsChartImage.Create;
    item.Name := AName;
    Result := inherited Add(item);
  end;
  Items[Result].Image := AImage;
end;

function TsChartImageList.FindByName(AName: String): TsChartImage;
var
  idx: Integer;
begin
  idx := IndexOfName(AName);
  if idx <> -1 then
    Result := Items[idx]
  else
    Result := nil;
end;

function TsChartImageList.GetItem(AIndex: Integer): TsChartImage;
begin
  Result := TsChartImage(inherited Items[AIndex]);
end;

function TsChartImageList.IndexOfName(AName: String): Integer;
begin
  for Result := 0 to Count-1 do
    if SameText(Items[Result].Name, AName) then
      exit;
  Result := -1;
end;

procedure TsChartImageList.SetItem(AIndex: Integer; AValue: TsChartImage);
begin
  TsChartImage(inherited Items[AIndex]).CopyFrom(AValue);
end;


{ TsChartFill }

constructor TsChartFill.CreateSolidFill(AColor: TsChartColor);
begin
  inherited Create;
  Style := cfsSolid;
  Color := AColor;
end;

constructor TsChartFill.CreateHatchFill(AHatchIndex: Integer; ABkColor: TsChartColor);
begin
  inherited Create;
  if aBkColor.Transparency = 1.0 then
    Style := cfsHatched
  else
    Style := cfsSolidHatched;
  Hatch := AHatchIndex;
  Color := ABkColor;
end;

procedure TsChartFill.CopyFrom(AFill: TsChartFill);
begin
  if AFill <> nil then
  begin
    Style := AFill.Style;
    Color := AFill.Color;
    Gradient := AFill.Gradient;
    Hatch := AFill.Hatch;
    Image := AFill.Image;
  end;
end;


{ TsChartLineStyle }

procedure TsChartLineStyle.CopyFrom(ASource: TsChartLineStyle);
begin
  Name := ASource.Name;
  Segment1 := ASource.Segment1;
  Segment2 := ASource.Segment2;
  Distance := ASource.Distance;
  RelativeToLineWidth := ASource.RelativeToLineWidth;
end;

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

function TsChartLineStyleList.IndexOfName(AName: String): Integer;
begin
  for Result := 0 to Count-1 do
    if Items[Result].Name = AName then
      exit;
  Result := -1;
end;

procedure TsChartLineStyleList.SetItem(AIndex: Integer; AValue: TsChartLineStyle);
begin
  TsChartLineStyle(inherited Items[AIndex]).CopyFrom(AValue);
end;


{ TsChartCellAddr }

constructor TsChartCellAddr.Create(AChart: TsChart);
begin
  FChart := AChart;
  Sheet := '';
  Row := UNASSIGNED_ROW_COL_INDEX;
  Col := UNASSIGNED_ROW_COL_INDEX;
end;

procedure TsChartCellAddr.CopyFrom(ASource: TsChartCellAddr);
begin
  Sheet := ASource.Sheet;
  Row := ASource.Row;
  Col := ASource.Col;
end;

function TsChartCellAddr.GetSheetName: String;
begin
  if Sheet <> '' then
    Result := Sheet
  else
    Result := FChart.Worksheet.Name;
end;

function TsChartCellAddr.IsUsed: Boolean;
begin
  Result := (Row <> UNASSIGNED_ROW_COL_INDEX) and (Col <> UNASSIGNED_ROW_COL_INDEX);
end;


{ TsChartRange }

constructor TsChartRange.Create(AChart: TsChart);
begin
  FChart := AChart;
  Sheet1 := '';
  Sheet2 := '';
  Row1 := UNASSIGNED_ROW_COL_INDEX;
  Col1 := UNASSIGNED_ROW_COL_INDEX;
  Row2 := UNASSIGNED_ROW_COL_INDEX;
  Col2 := UNASSIGNED_ROW_COL_INDEX;
end;

procedure TsChartRange.CopyFrom(ASource: TsChartRange);
begin
  Sheet1 := ASource.Sheet1;
  Sheet2 := ASource.Sheet2;
  Row1 := ASource.Row1;
  Col1 := ASource.Col1;
  Row2 := ASource.Row2;
  Col2 := ASource.Col2;
end;

function TsChartRange.GetSheet1Name: String;
begin
  if Sheet1 <> '' then
    Result := Sheet1
  else
  Result := FChart.Worksheet.Name;
  if SheetNameNeedsQuotes(Result) then
    Result := QuotedStr(Result);
end;

function TsChartRange.GetSheet2Name: String;
begin
  if Sheet2 <> '' then
    Result := Sheet2
  else
  Result := FChart.Worksheet.Name;
  if SheetNameNeedsQuotes(Result) then
    Result := QuotedStr(Result);
end;

function TsChartRange.IsEmpty: Boolean;
begin
  Result :=
    (Row1 = UNASSIGNED_ROW_COL_INDEX) and (Col1 = UNASSIGNED_ROW_COL_INDEX) and
    (Row2 = UNASSIGNED_ROW_COL_INDEX) and (Col2 = UNASSIGNED_ROW_COL_INDEX);
end;

function TsChartRange.NumCells: Integer;
begin
  if IsEmpty then
    Result := 0
  else
    Result := (Col2 - Col1 + 1) * (Row2 - Row1 + 1);
end;

{ TsChartElement }

constructor TsChartElement.Create(AChart: TsChart);
begin
  inherited Create;
  FChart := AChart;
  FVisible := true;
end;

procedure TsChartElement.CopyFrom(ASource: TsChartElement);
begin
  if ASource <> nil then
    Visible := ASource.Visible;
end;

function TsChartElement.GetVisible: Boolean;
begin
  Result := FVisible;
end;

procedure TsChartElement.SetVisible(AValue: Boolean);
begin
  FVisible := AValue;
end;


{ TsChartFillElement }

constructor TsChartFillElement.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FBackground := TsChartFill.Create;
  FBackground.Style := cfsSolid;
  FBackground.Color := ChartColor(scWhite);
  FBackground.Gradient := -1;
  FBackground.Hatch := -1;
  FBorder := TsChartLine.Create;
  FBorder.Style := clsSolid;
  FBorder.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  FBorder.Color := ChartColor(scBlack);
end;

destructor TsChartFillElement.Destroy;
begin
  FBorder.Free;
  FBackground.Free;
  inherited;
end;

procedure TsChartFillElement.CopyFrom(ASource: TsChartElement);
var
  srcElement: TsChartFillElement;
begin
  inherited CopyFrom(ASource);

  if ASource is TsChartFillElement then
  begin
    srcElement := TsChartFillElement(ASource);

    if srcElement.Background <> nil then
    begin
      if FBackground = nil then
        FBackground := TsChartFill.Create;
      FBackground.CopyFrom(srcElement.Background);
    end else
      FreeAndNil(FBackground);

    if srcElement.Border <> nil then
    begin
      if FBorder = nil then
        FBorder := TsChartLine.Create;
      FBorder.CopyFrom(srcElement.Border);
    end else
      FreeAndNil(FBorder);
  end;
end;


{ TsChartText }

constructor TsChartText.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FBorder.Style := clsNoLine;
  FBackground.Style := cfsNoFill;

  FFont := TsFont.Create;
  FFont.Size := 10;
  FFont.Style := [];
  FFont.Color := scBlack;

  FVisible := true;
end;

destructor TsChartText.Destroy;
begin
  FFont.Free;
  inherited;
end;

procedure TsChartText.CopyFrom(ASource: TsChartElement);
begin
  inherited CopyFrom(ASource);
  if ASource is TsChartText then
  begin
    FCaption := TsChartText(ASource).Caption;
    FRotationAngle := TsChartText(ASource).RotationAngle;
    FFont.CopyOf(TsChartText(ASource).Font);
    FPosX := TsChartText(ASource).PosX;
    FPosY := TsChartText(ASource).PosY;
  end;
end;


{ TsChartAxis }

constructor TsChartAxis.Create(AChart: TsChart);
begin
  inherited Create(AChart);

  FAutomaticMin := true;
  FAutomaticMax := true;
  FAutomaticMajorInterval := true;
  FAutomaticMinorInterval := true;
  FAutomaticMinorSteps := true;

  FCategoryRange := TsChartRange.Create(AChart);

  FTitle := TsChartText.Create(AChart);
  FDefaultTitleRotation := true;

  FLabelFont := TsFont.Create;
  FLabelFont.Size := 9;
  FLabelFont.Style := [];
  FLabelFont.Color := scBlack;

  FLabelFormatFromSource := true;
  FLabelFormatDateTime := '';
  FLabelFormatPercent := '0%';
  FLabelRotation := 0;
  FShowLabels := true;

  FAxisLine := TsChartLine.Create;
  FAxisLine.Color := ChartColor(scBlack);
  FAxisLine.Style := clsSolid;
  FAxisLine.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);

  FMajorTicks := [catOutside];
  FMinorTicks := [];

  FMajorGridLines := TsChartLine.Create;
  FMajorGridLines.Color := ChartColor(scSilver);
  FMajorGridLines.Style := clsSolid;
  FMajorGridLines.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);

  FMinorGridLines := TsChartLine.Create;
  FMinorGridLines.Color := ChartColor(scSilver);
  FMinorGridLines.Style := clsDash;
  FMinorGridLines.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);

  FLogarithmic := false;
  FLogBase := 10.0;
end;

destructor TsChartAxis.Destroy;
begin
  FMinorGridLines.Free;
  FMajorGridLines.Free;
  FAxisLine.Free;
  FLabelFont.Free;
  FTitle.Free;
  FCategoryRange.Free;
  inherited;
end;

procedure TsChartAxis.CopyFrom(ASource: TsChartElement);
begin
  inherited CopyFrom(ASource);
  if ASource is TsChartAxis then
  begin
    FAlignment := TsChartAxis(ASource).Alignment;
    FAutomaticMax := TsChartAxis(ASource).AutomaticMax;
    FAutomaticMin := TsChartAxis(ASource).AutomaticMin;
    FAutomaticMajorInterval := TsChartAxis(ASource).AutomaticMajorInterval;
    FAutomaticMinorInterval := TsChartAxis(ASource).AutomaticMinorInterval;
    FAutomaticMinorSteps := TsChartAxis(ASource).AutomaticMinorSteps;
    FAxisLine.CopyFrom(TsChartAxis(ASource).AxisLine);
    FCategoryRange.CopyFrom(TsChartAxis(ASource).CategoryRange);
    FMajorGridLines.CopyFrom(TsChartAxis(ASource).MajorGridLines);
    FMinorGridLines.CopyFrom(TsChartAxis(ASource).MinorGridLines);
    FInverted := TsChartAxis(ASource).Inverted;
    FLabelFont.CopyOf(TsChartAxis(ASource).LabelFont);
    FLabelFormat := TsChartAxis(ASource).LabelFormat;
    FLabelFormatFromSource := TsChartAxis(ASource).LabelFormatFromSource;
    FLabelFormatDateTime := TsChartAxis(ASource).LabelFormatDateTime;
    FLabelFormatPercent := TsChartAxis(ASource).LabelFormatPercent;
    FLabelRotation := TsChartAxis(ASource).LabelRotation;
    FLogarithmic := TsChartAxis(ASource).Logarithmic;
    FLogBase := TsChartAxis(ASource).LogBase;
    FMajorInterval := TsChartAxis(ASource).MajorInterval;
    FMajorTicks := TsChartAxis(ASource).MajorTicks;
    FMax := TsChartAxis(ASource).Max;
    FMin := TsChartAxis(ASource).Min;
    FMinorCount := TsChartAxis(ASource).MinorCount;
    FMinorInterval := TsChartAxis(ASource).MinorInterval;
    FMinorTicks := TsChartAxis(ASource).MinorTicks;
    FPosition := TsChartAxis(ASource).Position;
    FTitle.CopyFrom(TsChartAxis(ASource).Title);
    FPositionValue := TsChartAxis(ASource).PositionValue;
    FShowLabels := TsChartAxis(ASource).ShowLabels;
    FDateTime := TsChartAxis(ASource).DateTime;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Returns the other axis in the same direction, i.e. when the axis is the
  primary x axis the function returns the secondary x axis, etc.
-------------------------------------------------------------------------------}
function TsChartAxis.GetOtherAxis: TsChartAxis;
begin
  if Chart.XAxis = self then
    Result := Chart.X2Axis
  else if Chart.X2Axis = self then
    Result := Chart.XAxis
  else if Chart.YAxis = self then
    Result := Chart.Y2Axis
  else if Chart.Y2Axis = self then
    Result := Chart.YAxis;
end;

{@@ ----------------------------------------------------------------------------
  Returns the axis in the other direction when the chart is rotate.
-------------------------------------------------------------------------------}
function TsChartAxis.GetRotatedAxis: TsChartAxis;
begin
  if Chart.XAxis = self then
    Result := Chart.YAxis
  else if Chart.X2Axis = self then
    Result := Chart.Y2Axis
  else if Chart.YAxis = self then
    Result := Chart.XAxis
  else if Chart.Y2Axis = self then
    Result := Chart.X2Axis;
end;

{@@ ----------------------------------------------------------------------------
  Returns the text rotation angle of the axis title.
  When DefaultTitleRotation is true this is either 0 or 90, depending on the
  axis direction. Otherwise it is the title's RotationAngle.
-------------------------------------------------------------------------------}
function TsChartAxis.GetTitleRotationAngle: Single;
var
  rotated: Boolean;
begin
  if FDefaultTitleRotation then
  begin
    rotated := FChart.RotatedAxes;
    case FAlignment of
      caaLeft, caaRight: if rotated then Result := 0 else Result := 90;
      caaBottom, caaTop: if rotated then Result := 90 else Result := 0;
    end;
  end else
    Result := FTitle.RotationAngle;
end;

procedure TsChartAxis.SetCategoryRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetCategoryRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsChartAxis.SetCategoryRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Category values can only be located in a single column or row.');
  FCategoryRange.Sheet1 := ASheet1;
  FCategoryRange.Row1 := ARow1;
  FCategoryRange.Col1 := ACol1;
  FCategoryRange.Sheet2 := ASheet2;
  FCategoryRange.Row2 := ARow2;
  FCategoryRange.Col2 := ACol2;
end;

procedure TsChartAxis.SetMajorInterval(AValue: Double);
begin
  if IsNaN(AValue) or (AValue <= 0) then
    FAutomaticMajorInterval := true
  else
  begin
    FAutomaticMajorInterval := false;
    FMajorInterval := AValue;
  end;
end;

procedure TsChartAxis.SetMax(AValue: Double);
begin
  if IsNaN(AValue) then
    FAutomaticMax := true
  else
  begin
    FAutomaticMax := false;
    FMax := AValue;
  end;
end;

procedure TsChartAxis.SetMin(AValue: Double);
begin
  if IsNaN(AValue) then
    FAutomaticMin := true
  else
  begin
    FAutomaticMin := false;
    FMin := AValue;
  end;
end;

procedure TsChartAxis.SetMinorCount(AValue: Integer);
begin
  if IsNaN(AValue) or (AValue < 0) then
    FAutomaticMinorSteps := true
  else
  begin
    FAutomaticMinorSteps := false;
    FMinorCount := AValue;
  end;
end;


procedure TsChartAxis.SetMinorInterval(AValue: Double);
begin
  if IsNaN(AValue) or (AValue < 0) then
    FAutomaticMinorInterval := true
  else
  begin
    FAutomaticMinorInterval := false;
    FMinorInterval := AValue;
  end;
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

procedure TsChartLegend.CopyFrom(ASource: TsChartElement);
begin
  inherited CopyFrom(ASource);
  if ASource is TsChartLegend then
  begin
    FFont.CopyOf(TsChartLegend(ASource).Font);
    FCanOverlapPlotArea := TsChartLegend(ASource).CanOverlapPlotArea;
    FPosition := TsChartLegend(ASource).Position;
    FPosX := TsChartLegend(ASource).PosX;
    FPosY := TsChartLegend(ASource).PosY;
  end;
end;


{ TsChartDataPointStyle }

procedure TsChartDataPointStyle.CopyFrom(ASource: TsChartElement);
begin
  inherited CopyFrom(ASource);
  if ASource is TsChartDataPointStyle then
  begin
    FDataPointIndex := tsChartDataPointStyle(ASource).DataPointIndex;
    FPieOffset := TsChartDataPointStyle(ASource).PieOffset;
  end;
end;


{ TsChartDataPointStyleList }

constructor TsChartDataPointStyleList.Create(AChart: TsChart);
begin
  inherited Create;
  FChart := AChart;
end;

{ Note: You have the responsibility to destroy the AFill and ALine instances
  after calling AddFillAndLine ! }
function TsChartDataPointStyleList.AddFillAndLine(ADatapointIndex: Integer;
  AFill: TsChartFill; ALine: TsChartLine; APieOffset: Integer = 0): Integer;
var
  dataPointStyle: TsChartDataPointStyle;
begin
  dataPointStyle := TsChartDataPointStyle.Create(FChart);
  dataPointStyle.PieOffset := APieOffset;
  dataPointStyle.FDataPointIndex := ADataPointIndex;

  if AFill <> nil then
    dataPointStyle.Background.CopyFrom(AFill)
  else
  begin
    dataPointStyle.Background.Free;
    dataPointStyle.Background := nil;
  end;

  if ALine <> nil then
    dataPointStyle.Border.CopyFrom(ALine)
  else
  begin
    dataPointStyle.Border.Free;
    dataPointStyle.Border := nil;
  end;

  Result := inherited Add(dataPointStyle);
end;

function TsChartDataPointStyleList.AddSolidFill(ADataPointIndex: Integer;
  AColor: TsChartColor; ALine: TsChartLine = nil; APieOffset: Integer = 0): Integer;
var
  fill: TsChartFill;
begin
  fill := TsChartFill.Create;
  try
    fill.Style := cfsSolid;
    fill.Color := AColor;
    Result := AddFillAndLine(ADataPointIndex, fill, ALine, APieOffset);
  finally
    fill.Free;
  end;
end;

function TsChartDataPointStyleList.GetItem(AIndex: Integer): TsChartDataPointStyle;
begin
  if (AIndex >= 0) and (AIndex < Count) then
    Result := TsChartDataPointStyle(inherited Items[AIndex])
  else
    Result := nil;
end;

function TsChartDataPointStyleList.IndexOfDataPoint(ADataPointIndex: Integer): Integer;
begin
  for Result := 0 to Count - 1 do
    if Items[Result].DataPointIndex = ADataPointIndex then
      exit;
  Result := -1;
end;

procedure TsChartDataPointStyleList.SetItem(AIndex: Integer;
  AValue: TsChartDataPointStyle);
begin
  TsChartDataPointStyle(inherited Items[AIndex]).CopyFrom(AValue);
end;


{ TsChartErrorBars }

constructor TsChartErrorBars.Create(ASeries: TsChartSeries);
begin
  inherited Create(ASeries.Chart);
  FSeries := ASeries;
  FLine := TsChartLine.Create;
  FLine.Style := clsSolid;
  FLine.Color := ChartColor(scBlack);
  FLine.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  FRange[0] := TsChartRange.Create(ASeries.Chart);
  FRange[1] := TsChartRange.Create(ASeries.Chart);
  FShow[0] := false;
  FShow[1] := false;
  FShowEndCap := true;
end;

destructor TsChartErrorBars.Destroy;
begin
  FRange[1].Free;
  FRange[0].Free;
  FLine.Free;
  inherited;
end;

procedure TsChartErrorBars.CopyFrom(ASource: TsChartElement);
begin
  inherited CopyFrom(ASource);
  if ASource is TsChartErrorBars then
  begin
    FKind := TsChartErrorBars(ASource).Kind;
    FRange[0].CopyFrom(TsChartErrorBars(ASource).RangePos);
    FRange[1].CopyFrom(TsChartErrorBars(ASource).RangeNeg);
    FShow[0] := TsChartErrorBars(ASource).ShowPos;
    FShow[1] := TsChartErrorBars(ASource).ShowNeg;
    FShowEndCap := TsChartErrorBars(ASource).ShowEndCap;
    FValue[0] := TsChartErrorBars(ASource).ValuePos;
    FValue[1] := TsChartErrorBars(ASource).ValueNeg;
    FLine.CopyFrom(TsChartErrorBars(ASource).Line);
  end;
end;

function TsChartErrorBars.GetRange(AIndex: Integer): TsChartRange;
begin
  Result := FRange[AIndex];
end;

function TsChartErrorBars.GetVisible: Boolean;
begin
  Result := ShowPos or ShowNeg;
end;

function TsChartErrorBars.GetShow(AIndex: Integer): Boolean;
begin
  Result := FShow[AIndex];
end;

function TsChartErrorBars.GetValue(AIndex: Integer): Double;
begin
  result := FValue[AIndex];
end;

procedure TsChartErrorBars.InternalSetErrorBarRange(AIndex: Integer;
  ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Errorbar data can only be located in a single column or row.');
  FRange[AIndex].Sheet1 := ASheet1;
  FRange[AIndex].Row1 := ARow1;
  FRange[AIndex].Col1 := ACol1;
  FRange[AIndex].Sheet2 := ASheet2;
  FRange[AIndex].Row2 := ARow2;
  FRange[AIndex].Col2 := ACol2;
end;

procedure TsChartErrorBars.SetErrorBarRangePos(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  InternalSetErrorBarRange(0, '', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsChartErrorBars.SetErrorBarRangePos(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  InternalSetErrorBarRange(0, ASheet1, ARow1, ACol1, ASheet2, ARow2, ACol2);
end;

procedure TsChartErrorBars.SetErrorBarRangeNeg(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  InternalSetErrorBarRange(1, '', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsChartErrorBars.SetErrorBarRangeNeg(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  InternalSetErrorBarRange(1, ASheet1, ARow1, ACol1, ASheet2, ARow2, ACol2);
end;

procedure TsChartErrorBars.SetLine(AValue: TsChartLine);
begin
  FLine.CopyFrom(AValue);
end;

procedure TsChartErrorBars.SetKind(AValue: TsChartErrorBarKind);
begin
  FKind := AValue;
end;

procedure TsChartErrorBars.SetRange(AIndex: Integer; AValue: TsChartRange);
begin
  FRange[AIndex].CopyFrom(AValue);
end;

procedure TsChartErrorBars.SetShow(AIndex: Integer; AValue: Boolean);
begin
  FShow[AIndex] := AValue;
end;

procedure TsChartErrorBars.SetValue(AIndex: Integer; AValue: Double);
begin
  FValue[AIndex] := AValue;
end;

procedure TsChartErrorBars.SetVisible(AValue: Boolean);
begin
  ShowPos := AValue;
  ShowNeg := AValue;
end;


{ TsChartSeries }

constructor TsChartSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);

  FOrder := AChart.AddSeries(self);

  FXRange := TsChartRange.Create(AChart);
  FYRange := TsChartRange.Create(AChart);
  FFillColorRange := TsChartRange.Create(AChart);
  FLineColorRange := TsChartRange.Create(AChart);
  FLabelRange := TsChartRange.Create(AChart);
  FTitleAddr := TsChartCellAddr.Create(AChart);
  FGroupIndex := -1;

  FFill := TsChartFill.Create;
  FFill.Style := cfsSolid;
  FFill.Color := ChartColor(DEFAULT_SERIES_COLORS[FOrder mod Length(DEFAULT_SERIES_COLORS)]);
  FFill.Gradient := -1;
  FFill.Hatch := -1;

  FLine := TsChartLine.Create;
  FLine.Style := clsSolid;
  FLine.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  FLine.Color := ChartColor(DEFAULT_SERIES_COLORS[FOrder mod Length(DEFAULT_SERIES_COLORS)]);

  FDataPointStyles := TsChartDataPointStyleList.Create(AChart);

  FLabelFont := TsFont.Create;
  FLabelFont.Size := 9;

  FLabelBorder := TsChartLine.Create;
  FLabelBorder.Color := ChartColor(scBlack);
  FLabelBorder.Style := clsNoLine;

  FLabelBackground := TsChartFill.Create;
  FLabelBackground.Color := ChartColor(scWhite);
  FLabelBackground.Style := cfsNoFill;

  FLabelSeparator := ' ';
  FLabelFormatPercent := '0%';

  FTrendline := TsChartTrendline.Create;

  FXErrorBars := TsChartErrorBars.Create(Self);
  FYErrorBars := TsChartErrorBars.Create(Self);
end;

destructor TsChartSeries.Destroy;
begin
  FYErrorBars.Free;
  FXErrorBars.Free;
  FTrendline.Free;
  FLabelBackground.Free;
  FLabelBorder.Free;
  FLabelFont.Free;
  FDataPointStyles.Free;
  FLine.Free;
  FFill.Free;
  FTitleAddr.Free;
  FLabelRange.Free;
  FLineColorRange.Free;
  FFillColorRange.Free;
  FYRange.Free;
  FXRange.Free;
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

function TsChartSeries.GetXAxis: TsChartAxis;
begin
  if FXAxis = calPrimary then
    Result := Chart.XAxis
  else
    Result := Chart.X2Axis;
end;

function TsChartSeries.GetYAxis: TsChartAxis;
begin
  if FYAxis = calPrimary then
    Result := Chart.YAxis
  else
    Result := Chart.Y2Axis;
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
  SetTitleAddr('', ARow, ACol);
end;

procedure TsChartSeries.SetTitleAddr(ASheet: String; ARow, ACol: Cardinal);
begin
  FTitleAddr.Sheet := ASheet;
  FTitleAddr.Row := ARow;
  FTitleAddr.Col := ACol;
end;

procedure TsChartSeries.SetFillColorRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetFillColorRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsChartSeries.SetFillColorRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Series fill color values can only be located in a single column or row.');
  FFillColorRange.Sheet1 := ASHeet1;
  FFillColorRange.Row1 := ARow1;
  FFillColorRange.Col1 := ACol1;
  FFillColorRange.Sheet2 := ASheet2;
  FFillColorRange.Row2 := ARow2;
  FFillColorRange.Col2 := ACol2;
end;

procedure TsChartSeries.SetLabelRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetLabelRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsChartSeries.SetLabelRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Series labels can only be located in a single column or row.');
  FLabelRange.Sheet1 := ASheet1;
  FLabelRange.Row1 := ARow1;
  FLabelRange.Col1 := ACol1;
  FLabelRange.Sheet2 := ASheet2;
  FLabelRange.Row2 := ARow2;
  FLabelRange.Col2 := ACol2;
end;

procedure TsChartSeries.SetLineColorRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetLineColorRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsChartSeries.SetLineColorRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Series line color values can only be located in a single column or row.');
  FLineColorRange.Sheet1 := ASheet1;
  FLineColorRange.Row1 := ARow1;
  FLineColorRange.Col1 := ACol1;
  FLineColorRange.Sheet2 := ASheet2;
  FLineColorRange.Row2 := ARow2;
  FLineColorRange.Col2 := ACol2;
end;

procedure TsChartSeries.SetXErrorBars(AValue: TsChartErrorBars);
begin
  FXErrorBars.CopyFrom(AValue);
end;

procedure TsChartSeries.SetXRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetXRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsChartSeries.SetXRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Series x values can only be located in a single column or row.');
  FXRange.Sheet1 := ASheet1;
  FXRange.Row1 := ARow1;
  FXRange.Col1 := ACol1;
  FXRange.Sheet2 := ASheet2;
  FXRange.Row2 := ARow2;
  FXRange.Col2 := ACol2;
end;

procedure TsChartSeries.SetYErrorBars(AValue: TsChartErrorBars);
begin
  FYErrorBars.CopyFrom(AValue);
end;

procedure TsChartSeries.SetYRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetYRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsChartSeries.SetYRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Series y values can only be located in a single column or row.');
  FYRange.Sheet1 := ASheet1;
  FYRange.Row1 := ARow1;
  FYRange.Col1 := ACol1;
  FYRange.Sheet2 := ASheet2;
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
  FSupportsTrendline := true;
  FGroupIndex := 0;
end;


{ TsBarSeries }

constructor TsBarSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctBar;
  FSupportsTrendline := true;
  FGroupIndex := 0;
end;


{ TsBubbleSeries }

constructor TsBubbleSeries.Create(AChart: TsChart);
begin
  inherited;
  FBubbleRange := TsChartRange.Create(AChart);
  FBubbleScale := 1.0;
  FBubbleSizeMode := bsmArea;
  FChartType := ctBubble;
end;

destructor TsBubbleSeries.Destroy;
begin
  FBubbleRange.Free;
  inherited;
end;

{ Empty sheet name will be replaced by the name of the sheet containing the chart. }
procedure TsBubbleSeries.SetBubbleRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetBubbleRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsBubbleSeries.SetBubbleRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Bubble series values can only be located in a single column or row.');
  FBubbleRange.Sheet1 := ASheet1;
  FBubbleRange.Row1 := ARow1;
  FBubbleRange.Col1 := ACol1;
  FBubbleRange.Sheet2 := ASheet2;
  FBubbleRange.Row2 := ARow2;
  FBubbleRange.Col2 := ACol2;
end;


{ TsCustomLineSeries }

constructor TsCustomLineSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);

  FChartType := ctLine;
  FSupportsTrendline := true;

  FSymbolWidth := 2.5;
  FSymbolHeight := 2.5;
  FShowSymbols := false;
  FShowLines := true;

  FSymbolBorder := TsChartLine.Create;
  FSymbolBorder.Style := clsSolid;
  FSymbolBorder.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  FSymbolBorder.Color := ChartColor(scBlack);

  FSymbolFill := TsChartFill.Create;
  FSymbolFill.Style := cfsNoFill;
end;

destructor TsCustomLineSeries.Destroy;
begin
  FSymbolBorder.Free;
  FSymbolFill.Free;
  inherited;
end;

function TsCustomLineSeries.GetSmooth: Boolean;
begin
  Result := FInterpolation in [ciBSpline, ciCubicSpline];
end;

procedure TsCustomLineSeries.SetSmooth(AValue: Boolean);
begin
  if AValue then
  begin
    if not (FInterpolation in [ciBSpline, ciCubicSpline]) then
      FInterpolation := ciCubicSpline;
  end else
  begin
    if (FInterpolation in [ciBSpline, ciCubicSpline]) then
    FInterpolation := ciLinear;
  end;
end;

{ TsLineSeries }
constructor TsLineSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FGroupIndex := 0;
end;

{ TsPieSeries }
constructor TsPieSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctPie;
  FStartAngle := 90;
  FLine.Color := ChartColor(scBlack);
end;

function TsPieSeries.GetChartType: TsChartType;
begin
  if FInnerRadiusPercent > 0 then
    Result := ctRing
  else
    Result := ctPie;
end;

function TsPieSeries.GetSliceOffset(ASliceIndex: Integer): Integer;
var
  i: Integer;
  datapointstyle: TsChartDatapointStyle;
begin
  Result := 0;
  if (ASliceIndex >= 0) and (ASliceIndex < FDataPointStyles.Count) then
  begin
    datapointstyle := FDatapointStyles[ASliceIndex];
    if datapointstyle <> nil then
      Result := datapointstyle.PieOffset;
  end;
end;


{ TsRadarSeries }
constructor TsRadarSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctRadar;
  FFill.Style := cfsNoFill;  // to make the series default to ctRadar rather than ctFilledRadar
end;


{ TsFilledRadarSeries }
constructor TsFilledRadarSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctFilledRadar;
  Fill.Style := cfsSolid;
end;


{ TsTrendlineEquation }
constructor TsTrendlineEquation.Create;
begin
  inherited Create;
  Font := TsFont.Create;
  Font.Size := 9;
  Border := TsChartLine.Create;
  Border.Style := clsNoLine;
  Border.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  Border.Color := ChartColor(scBlack);
  Fill := TsChartFill.Create;
  Fill.Color := ChartColor(scWhite);
  XName := 'x';
  YName := 'f(x)';
end;

destructor TsTrendlineEquation.Destroy;
begin
  Fill.Free;
  Border.Free;
  Font.Free;
  inherited;
end;

function TsTrendlineEquation.DefaultBorder: Boolean;
begin
  Result := Border.Style = clsNoLine;
end;

function TsTrendlineEquation.DefaultFill: Boolean;
begin
  Result := Fill.Style = cfsNoFill;
end;

function TsTrendlineEquation.DefaultFont: Boolean;
begin
  Result := (Font.FontName = '') and (Font.Size = 9) and (Font.Style = []) and
            (Font.Color = scBlack);
end;

function TsTrendlineEquation.DefaultNumberFormat: Boolean;
begin
  Result := NumberFormat = '';
end;

function TsTrendlineEquation.DefaultPosition: Boolean;
begin
  Result := (Left = 0) and (Top = 0);
end;

function TsTrendlineEquation.DefaultXName: Boolean;
begin
  Result := XName = 'x';
end;

function TsTrendlineEquation.DefaultYName: Boolean;
begin
  Result := YName = 'f(x)';
end;


{ TsChartTrendline }
constructor TsChartTrendline.Create;
begin
  inherited Create;

  Line := TsChartLine.Create;
  Line.Style := clsSolid;
  Line.Width := PtsToMM(DEFAULT_CHART_LINEWIDTH);
  Line.Color := ChartColor(scBlack);

  Equation := TsTrendlineEquation.Create;
end;

destructor TsChartTrendline.Destroy;
begin
  Equation.Free;
  Line.Free;
  inherited;
end;


{ TsCustomScatterSeries }

constructor TsCustomScatterSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctScatter;
  FSupportsTrendline := true;
end;


{ TsStockSeries }

constructor TsStockSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctStock;
  FOpenRange := TsChartRange.Create(AChart);
  FHighRange := TsChartRange.Create(AChart);
  FLowRange := TsChartRange.Create(AChart);

  // FFill is CandleStickUp, FLine is RangeLine
  FCandleStickDownBorder := TsChartLine.CreateSolid(ChartColor(scBlack), PtsToMM(DEFAULT_CHART_LINEWIDTH));
  FCandleStickDownFill := TsChartFill.CreateSolidFill(ChartColor(scBlack)); // These are the Excel default colors
  FCandleStickUpBorder := TsChartLine.CreateSolid(ChartColor(scBlack), PtsToMM(DEFAULT_CHART_LINEWIDTH));
  FCandleStickUpFill := TsChartFill.CreateSolidFill(ChartColor(scWhite));
  FRangeLine := TsChartLine.CreateSolid(ChartColor(scBlack), PtsToMM(DEFAULT_CHART_LINEWIDTH));
  FTickWidthPercent := 50;
end;

destructor TsStockSeries.Destroy;
begin
  FRangeLine.Free;
  FCandleStickUpFill.Free;
  FCandleStickUpBorder.Free;
  FCandleStickDownBorder.Free;
  FCandleStickDownFill.Free;
  FOpenRange.Free;
  FHighRange.Free;
  FLowRange.Free;
  inherited;
end;

procedure TsStockSeries.SetOpenRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetOpenRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsStockSeries.SetOpenRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
 begin
   if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
     raise Exception.Create('Stock series values can only be located in a single column or row.');
   FOpenRange.Sheet1 := ASheet1;
   FOpenRange.Row1 := ARow1;
   FOpenRange.Col1 := ACol1;
   FOpenRange.Sheet2 := ASheet2;
   FOpenRange.Row2 := ARow2;
   FOpenRange.Col2 := ACol2;
 end;

procedure TsStockSeries.SetHighRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetHighRange('', ARow1, ACol1, '', ARow2, ACol2);
end;
procedure TsStockSeries.SetHighRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
    raise Exception.Create('Stock series values can only be located in a single column or row.');
  FHighRange.Sheet1 := ASheet1;
  FHighRange.Row1 := ARow1;
  FHighRange.Col1 := ACol1;
  FHighRange.Sheet2 := ASheet2;
  FHighRange.Row2 := ARow2;
  FHighRange.Col2 := ACol2;
end;

procedure TsStockSeries.SetLowRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetLowRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsStockSeries.SetLowRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
 begin
   if (ARow1 <> ARow2) and (ACol1 <> ACol2) then
     raise Exception.Create('Stock series values can only be located in a single column or row.');
   FLowRange.Sheet1 := ASheet1;
   FLowRange.Row1 := ARow1;
   FLowRange.Col1 := ACol1;
   FLowRange.Sheet2 := ASheet2;
   FLowRange.Row2 := ARow2;
   FLowRange.Col2 := ACol2;
 end;

procedure TsStockSeries.SetCloseRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetCloseRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

procedure TsStockSeries.SetCloseRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
 begin
   SetYRange(ASheet1, ARow1, ACol1, ASheet2, ARow2, ACol2);
 end;


{ TsChart }

constructor TsChart.Create;
begin
  inherited Create(nil);

  FLineStyles := TsChartLineStyleList.Create;
  clsFineDot := FLineStyles.Add('fine-dot', 100, 1, 0, 0, 100, false);
  clsDot := FLineStyles.Add('dot', 400, 1, 0, 0, 400, true);
  clsDash := FLineStyles.Add('dash', 1200, 1, 0, 0, 800, true);
  clsDashDot := FLineStyles.Add('dash-dot', 1000, 1, 300, 1, 2000, true);
  clsLongDash := FLineStyles.Add('long dash', 2400, 1, 0, 0, 800, true);
  clsLongDashDot := FLineStyles.Add('long dash-dot', 1600, 1, 800, 1, 800, true);
  clsLongDashDotDot := FLineStyles.Add('long dash-dot-dot', 1600, 1, 800, 2, 800, true);
   {
  clsFineDot := FLineStyles.Add('fine-dot', 100, 1, 0, 0, 100, false);
  clsDot := FLineStyles.Add('dot', 150, 1, 0, 0, 150, true);
  clsDash := FLineStyles.Add('dash', 300, 1, 0, 0, 150, true);
  clsDashDot := FLineStyles.Add('dash-dot', 300, 1, 100, 1, 150, true);
  clsLongDash := FLineStyles.Add('long dash', 400, 1, 0, 0, 200, true);
  clsLongDashDot := FLineStyles.Add('long dash-dot', 500, 1, 100, 1, 200, true);
  clsLongDashDotDot := FLineStyles.Add('long dash-dot-dot', 500, 1, 100, 2, 200, true);
  }
  FGradients := TsChartGradientList.Create;
  FHatches := TsChartHatchList.Create;
  FImages := TsChartImageList.Create;

  fgradients.AddLinearGradient('g1', ChartColor(scRed), ChartColor(scBlue), 1, 1, 0, 0);

  FWorksheet := nil;
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
  FXAxis.Alignment := caaBottom;
  FXAxis.Title.Font.Style := [fssBold];
  FXAxis.LabelFont.Size := 9;
  FXAxis.Position := capStart;

  FX2Axis := TsChartAxis.Create(self);
  FX2Axis.Alignment := caaTop;
  FX2Axis.Title.Font.Style := [fssBold];
  FX2Axis.LabelFont.Size := 9;
  FX2Axis.Visible := false;
  FX2Axis.Position := capEnd;

  FYAxis := TsChartAxis.Create(self);
  FYAxis.Alignment := caaLeft;
  FYAxis.Title.Font.Style := [fssBold];
  FYAxis.Title.RotationAngle := 90;
  FYAxis.LabelFont.Size := 9;
  FYAxis.Position := capStart;

  FY2Axis := TsChartAxis.Create(self);
  FY2Axis.Alignment := caaRight;
  FY2Axis.Title.Font.Style := [fssBold];
  FY2Axis.Title.RotationAngle := 90;
  FY2Axis.LabelFont.Size := 9;
  FY2Axis.Visible := false;
  FY2Axis.Position := capEnd;

  FSeriesList := TsChartSeriesList.Create;

  FBarGapWidthPercent := 50;
  FBarOverlapPercent := 0;
end;

destructor TsChart.Destroy;
begin
  FSeriesList.Free;
  FXAxis.Free;
  FX2Axis.Free;
  FYAxis.Free;
  FY2Axis.Free;
  FLegend.Free;
  FTitle.Free;
  FSubtitle.Free;
  FFloor.Free;
  FPlotArea.Free;
  FImages.Free;
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

function TsChart.GetCategoryLabelRange: TsChartRange;
begin
  Result := XAxis.CategoryRange;
end;

function TsChart.GetChartType: TsChartType;
var
  i: Integer;
begin
  if FSeriesList.Count > 0 then
  begin
    Result := Series[0].ChartType;
    for i := 0 to FSeriesList.Count-1 do
      if FSeriesList[i] is TsStockSeries then
      begin
        Result := ctStock;
        exit;
      end;
  end else
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

