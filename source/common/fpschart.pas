{@@ ----------------------------------------------------------------------------
  Unit **fpsChart** implements support for charts embedded in spreadsheets.

  AUTHOR:  Werner Pamler

  LICENSE: See the file COPYING.modifiedLGPL.txt, included in the Lazarus
           distribution, for details about the license.
-------------------------------------------------------------------------------}

unit fpsChart;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

interface

uses
  Classes, SysUtils, Contnrs, FPImage, fpsTypes, fpsUtils;

const
  {@@ Default width of lines in charts, in Points }
  DEFAULT_CHART_LINEWIDTH = 0.75;  // pts

  {@@ Default font in charts }
  DEFAULT_CHART_FONT = 'Arial';

  {@@ Default colors used by chart series (type @link(TsColor)) }
  DEFAULT_SERIES_COLORS: array[0..7] of TsColor = (
    scRed, scBlue, scGreen, scMagenta, scPurple, scTeal, scBlack, scGray
  );

type
  {@@ Data type alias for transparency of chart colors:
    Single precision value that should range between 0.0 (opaque) and 1.0 (fully transparent)
    (but this range is not enforced). }
  TsChartTransparency = single;

  {@@ Record describing a color used by charts (@link(TsColor)), includes a Transparency element (@link(TsChartTransparency))}
  TsChartColor = record
    Transparency: TsChartTransparency;
    case Integer of
      0: (Red, Green, Blue, SystemColorIndex: Byte);
      1: (Color: TsColor);
    end;

const
  {@@ Chart color which is fully transparent, i.e. inivisible. }
  sccTransparent: TsChartColor = (Transparency: 1.0; Color: 0);

type
  TsChart = class;
  TsChartAxis = class;
  TsChartSeries = class;

  {@@ Class containing the parameters of a line in charts }
  TsChartLine = class
  private
    function GetStyle: TsChartLinePatternStyle;
    procedure SetStyle(AValue: TsChartLinePatternStyle);
  public
    {@@ Index into the global raw line pattern list implemented in unit fpsPatterns. }
    PatternIndex: Integer;
    {@@ Line width, in millimeters }
    Width: Single;
    {@@ Color of the line. This is a @link(TsChartColor) including a Transparency value. }
    Color: TsChartColor;

    constructor Create;
    constructor CreateSolid(AColor: TsChartColor; AWidth: Double);
    procedure CopyFrom(ALine: TsChartLine);
    function IsHidden: Boolean;
    procedure SelectNoLine;
    procedure SelectPatternLine(ALineStyle: TsChartLinePatternStyle; AColor: TsChartColor; ALineWidth: Single = -1.0);
    procedure SelectPatternLine(APatternIndex: Integer; AColor: TsChartColor; ALineWidth: Single = -1.0);
    procedure SelectSolidLine(AColor: TsChartColor; ALineWidth: Single = -1.0);
    property Style: TsChartLinePatternStyle read GetStyle write SetStyle;
  end;

  {@@ Enumeration of the gradient style supported by charts
   @value  cgsLinear       Linear gradient
   @value  cgsAxial        Axial gradient
   @value  cgsRadial       Radial gradient
   @value  cgsElliptic     Elliptic gradient
   @value  cgsSquare       Special case of a rectangular gradient (squar)
   @value  cgsRectangular  Rectangular gradient
   @value  cgsShape        Gradient following the shape the filled chart element. }
  TsChartGradientStyle = (cgsLinear, cgsAxial, cgsRadial, cgsElliptic, cgsSquare, cgsRectangular, cgsShape);

  {@@ Record describing the properties of a gradient start, end, or intermediate point

   @member  Color   The color at this gradient step position. This is a @link(TsChartColor) including a Transparency value.
   @member(Value    A floating point value ranging between 0.0 and 1.0,
                    identifies the relative position in the filled area at which
                    the specified color is used.)
   @member  Intensity The intensity of the color a this gradient step, a floating point value between 0.0 and 1.0.
  }
  TsChartGradientStep = record
    Value: Single;         // 0.0 ... 1.0
    Color: TsChartColor;
    Intensity: Double;     // 0.0 ... 1.0
  end;

  {@@ Array of @link(TsChartGradientStep) elements }
  TsChartGradientSteps = array of TsChartGradientStep;

  {@@ Class containing all parameters of a gradient as used by spreadsheet charts. }
  TsChartGradient = class
  private
    FSteps: TsChartGradientSteps;
    function GetBorder(AIndex: Integer): Single;
    function GetColor(AIndex: Integer): TsChartColor;
    function GetSteps(AIndex: Integer): TsChartGradientStep;
    procedure SetBorder(AIndex: Integer; AValue: Single);
    procedure SetStep(AIndex: Integer; AValue: Double; AColor: TsChartColor);
  public
    {@@ Name of the gradient. Must be unique. }
    Name: String;
    {@@ A value of the @link(TsChartGradientStyle) enumeration to define the gradient type: linear, radial, etc.}
    Style: TsChartGradientStyle;
    {@@ For non-linear gradients the x coordinate of the gradient center point. Value 0.0 ... 1.0}
    CenterX: Single;
    {@@ For non-linear gradients the y coordinate of the gradient center point. Values 0.0 ... 1.0 }
    CenterY: Single;
    {@@  For linear gradients the direction of the gradient in degrees from start to end color. 0° is horizontal, growing in CCW direction }
    Angle: Double;
    constructor Create;
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartGradient);
    procedure AddStep(AValue: Single; AColor: TsChartColor);
    function NumSteps: Integer;
    {@@ A @link(TsChartGradientStep) array element containing all parameters of the start, end, or an intermediate step. The first array element is the starting, the last element the ending step of the gradien. }
    property Steps[AIndex: Integer]: TsChartGradientStep read GetSteps;
    {@@ The numerical value relative to the shape's size at which the gradient starts. }
    property StartBorder: Single index 0 read GetBorder write SetBorder;
    {@@ The numerical value relative to the shape's size at which the gradient ends. }
    property EndBorder: Single index 1 read GetBorder write SetBorder;
    {@@ The color at which the gradient starts (see @link(TsChartColor)).}
    property StartColor: TsChartColor index 0 read GetColor;
    {@@ The color at which the gradient ends (see @link(TsChartColor)). }
    property EndColor: TsChartColor index 1 read GetColor;
  end;

  {@@ A list collecting all gradient definitions used by a chart (see @link(TsChartGradient)). }
  TsChartGradientList = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartGradient;
    procedure SetItem(AIndex: Integer; AValue: TsChartGradient);
  public
    function AddGradient(AName: String; AGradient: TsChartGradient): Integer;
    function AddGradient(AName: String; AStyle: TsChartGradientStyle;
      AStartColor, AEndColor: TsChartColor; AAngle, ACenterX, ACenterY: Double;
      AStartBorder: Double = 0.0; AEndBorder: Double = 1.0): Integer;
    function AddAxialGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AAngle: Double): Integer;
    function AddEllipticGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AAngle, ACenterX, ACenterY: Double): Integer;
    function AddLinearGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AAngle: Double): Integer;
    function AddRadialGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      ACenterX, ACenterY: Double): Integer;
    function AddRectangularGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AAngle, ACenterX, ACenterY: Double): Integer;
    function AddSquareGradient(AName: String; AStartColor, AEndColor: TsChartColor;
      AAngle, ACenterX, ACenterY: Double): Integer;
    function IndexOfName(AName: String): Integer;
    function FindByName(AName: String): TsChartGradient;
    {@@ List element cast to the TsChartGradient class }
    property Items[AIndex: Integer]: TsChartGradient read GetItem write SetItem; default;
  end;

  {@@ Defines the style of hatch patterns as used by LibreOffice Calc
   @value chsDot     Hatch lines are dotted
   @value chsSingle  The pattern uses single hatch links.
   @value chsDouble  The pattern uses two hatch lines rotated by 90° to each other.
   @value chsTriple  The pattern uses three hatch lines rotated by 45 and 90°.
  }
  TsChartHatchStyle = (chsDot, chsSingle, chsDouble, chsTriple);

  {@@ Record for a point with coordinages X and Y given in single precision.
    @member X  X coordinate of the point
    @member Y  Y coordinate of the point
  }
  TSngPoint = record X, Y: Single; end;

  {@@ The class TsChartFillPattern represents a combination of pattern
    (taken from global raw fill pattern list) and foreground/background colors. }
  TsChartFillPattern = class
    {@@ Name of the fill pattern (must be unique).}
    Name: String;
    {@@ Pattern index into the global raw fill patterns list. }
    Index: Integer;
    {@@ Foreground color of the pattern (see @link(TsChartColor)) }
    FgColor: TsChartColor;
    {@@ Background color of the pattern (see @link(TsChartColor)). Use @link(sccTransparent) to achieve a non-filled background (not supported by all patterns and formats).}
    BgColor: TsChartColor;
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartFillPattern);
    function IsClearPattern: Boolean;
  end;

  {@@ List collecting all fill patterns used by the chart to which it belongs. }
  TsChartFillPatternList = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartFillPattern;
    procedure SetItem(AIndex: Integer; AValue: TsChartFillPattern);
  protected
    function DefaultPatternName: String;
    function NewPattern(AName: String): Integer;
  public
    function AddPattern(AName: String; APatternStyle: TsChartFillPatternStyle;
      APatternColor: TsChartColor): Integer;
    function AddPattern(AName: String; APatternStyle: TsChartFillPatternStyle;
      APatternColor, ABackColor: TsChartColor): Integer;
    function AddPattern(AName: String; ARawPatternIndex: Integer;
      APatternColor: TsChartColor): Integer;
    function AddPattern(AName: String; ARawPatternIndex: Integer;
      APatternColor, ABackColor: TsChartColor): Integer;
    function FindByName(AName: String): TsChartFillPattern;
    function IndexOfName(AName: String): Integer;
    function IndexOfNameAndBgColor(AName: String; AColor: TsChartColor): Integer;
    {@@ List element cast to the TsChartFillPattern class }
    property Items[AIndex: Integer]: TsChartFillPattern read GetItem write SetItem; default;
  end;

  {@@ Represents data for an image by which the associated shape is filled
    in a tiled way. }
  TsChartImage = class
    {@@ Name of the image instance. Must be unique. }
    Name: String;
    {@@ Index of the image stream into the workbook's EmbeddedObj list }
    EmbeddedObjIndex: Integer;
    {@@ Width as used in the chart, in millimeters }
    Width: Single;
    {@@ Height of the image as used in the chart, in millimeters }
    Height: Single;
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartImage);
  end;

  {@@ List collecting all images used in a chart for filling purposes. }
  TsChartImagelist = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartImage;
    procedure SetItem(AIndex: Integer; AValue: TsChartImage);
  public
    function AddImage(AName: String; AEmbeddedObjIndex: Integer;
      AImgWidth: Single = -1.0; AImgHeight: Single = -1.0): Integer;
    function FindByName(AName: String): TsChartImage;
    function IndexOfName(AName: String): Integer;
    {@@ List element cast to the TsChartImage class }
    property Items[Aindex: Integer]: TsChartImage read GetItem write SetItem; default;
  end;

  {@@ Specifies the way a shape is filled with color or structure
   @value  cfsNoFill     The shape is not filled at all.
   @value  cfsSolidFill  The shape is filled by a uniform color.
   @value  cfsGradient   The shape is filled by a color gradient.
   @value  cfsPattern    The shape is filled by a tiled colored pattern.
   @value  cfsImage      The shpae is filled by an arbitrary tiled image. }
  TsChartFillStyle = (cfsNoFill, cfsSolidFill, cfsGradient, cfsPattern, cfsImage);

  {@@ Class which defines how an area or shape is filled by color or by a pattern. }
  TsChartFill = class
  public
    {@@ Element of the @link(TsChartFillStyle) enumeration identifying the type of the fill }
    Style: TsChartFillStyle;
    {@@ Index into the chart's @link(TsChart.Gradients) list containing all fill gradient parameters }
    Gradient: Integer;
    {@@ Index into the chart's @link(TsChart.FillPatterns) list containing all (colored) fill patterns }
    Pattern: Integer;
    {@@ Index into the chart's @link(TsChart.Images) list containing all images used for filling }
    Image: Integer;
    {@@ Color to be used in the solid-fill case (@link(Style) = @link(cfsSolidFill))}
    Color: TsChartColor;
    constructor Create;
    constructor CreateSolidFill(AColor: TsChartColor);
    procedure CopyFrom(AFill: TsChartFill);
    procedure SelectGradientFill(AGradientIndex: Integer);
    procedure SelectImageFill(AImageIndex: Integer);
    procedure SelectNoFill;
    procedure SelectPatternFill(APatternIndex: Integer);
    procedure SelectSolidFill(AColor: TsChartColor);
  end;

  {@@ Class which defines a cell address which can be evaluated in the chart.
    Useful for series titles. }
  TsChartCellAddr = class
  private
    FChart: TsChart;
  public
    {@@ Name of the worksheet which contains the referenced cell address }
    Sheet: String;
    {@@ Row index of the referenced cell address }
    Row: Cardinal;
    {@@ Column index of the referenced cell address }
    Col: Cardinal;
    constructor Create(AChart: TsChart);
    procedure CopyFrom(ASource: TsChartCellAddr);
    function GetSheetName: String;
    function IsUsed: Boolean;
    property Chart: TsChart read FChart;
  end;

  {@@ Class which defines a 3d cell range.
    Needed by series x, y and label ranges. }
  TsChartRange = class
  private
    FChart: TsChart;
  public
    {@@ Name of the worksheet containing the first cell of the range (top/left corner) }
    Sheet1: String;
    {@@ Name of the worksheet containing the last cell of the range (bottom/right corner) }
    Sheet2: String;
    {@@ Row index of the top-most cell of the range }
    Row1: Cardinal;
    {@@ Column index of the left-most cell of the range }
    Col1: Cardinal;
    {@@ Row index of the bottom-most cell of the range }
    Row2: Cardinal;
    {@@ Column index of the right-most cell of the range }
    Col2: Cardinal;
    constructor Create(AChart: TsChart);
    procedure CopyFrom(ASource: TsChartRange);
    function GetSheet1Name: String;
    function GetSheet2Name: String;
    function IsEmpty: Boolean;
    function NumCellsPerSheet: Integer;
    {@@ Chart in which this cell range is used. }
    property Chart: TsChart read FChart;
  end;

  {@@ A chart consists of several elements: title, footer, plotarea, axes, series, legend. }
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
    {@@ Chart to which this element belongs. }
    property Chart: TsChart read FChart;
    {@@ Allows to show/hide the element. }
    property Visible: Boolean read GetVisible write SetVisible;
  end;

  {@@ General @link(TsChartElement) which has a border line and filled background. }
  TsChartFillElement = class(TsChartElement)
  private
    FBackground: TsChartFill;
    FBorder: TsChartLine;
  public
    constructor Create(AChart: TsChart);
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsChartElement); override;
    {@@ Fill properties for the background of the element }
    property Background: TsChartFill read FBackground write FBackground;
    {@@ Line properties for the border of the element }
    property Border: TsChartLine read FBorder write FBorder;
  end;

  {@@ General chart element which can display some text. It has a (rotatable)
    Font and, since it descends from TsChartFillElement, a background and a
    border. }
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
    {@@ Text displayed by this element }
    property Caption: String read FCaption write FCaption;
    {@@ Font used to draw the text. }
    property Font: TsFont read FFont write FFont;
    {@@ Rotation angle for text painting, in degrees }
    property RotationAngle: single read FRotationAngle write FRotationAngle;
    {@@ X coordinate of the position of the top/left corner of the element, in millimeters }
    property PosX: Double read FPosX write FPosX;
    {@@ Y coordinate of the position of the top/left corner of the element, in millimeters }
    property PosY: Double read FPosY write FPosY;
    property Visible;
  end;

  {@@ Enumeration which defines where an axis can be positioned in a chart
   @value caaLeft  Axis at the left (y axis)
   @value caaTop   Axis at the top (alternative x axis)
   @value caaRight Axis at the right (alternative y axis)
   @value caaBottom Axis at the bottom (x axis) }
  TsChartAxisAlignment = (caaLeft, caaTop, caaRight, caaBottom);

  {@@ Enumeration defining the position of an axis relative to the other orthogonal axis
   @value capStart   The other axis crosses at the start (Min).
   @value capEnd     The other axis crosses at the end (Max).
   @value capValue   The other axis crosses at the @link(TsChartAxis.PositionValue) of the axis. }
  TsChartAxisPosition = (capStart, capEnd, capValue);

  {@@ Enumeration of the position of axis ticks relative to the axis.
   @value catInside  The axis tick is drawn inside the chart area.
   @value catOutside The axis tick is drawn outside the chart area. }
  TsChartAxisTick = (catInside, catOutside);

  {@@ Set containing the possible axis tick positions at an axis, inside or outside the plot area. }
  TsChartAxisTicks = set of TsChartAxisTick;

  {@@ Enumeration of the basic chart types
   @value ctEmpty   Empty chart, i.e. no series definied so far
   @value ctBar     Bar chart
   @value ctLine    Line chart
   @value ctArea    Area chart
   @value ctBarLine The chart contains bar and line series
   @value ctScatter Scatter plot
   @value ctBubble  Bubble chart
   @value ctRadar   Radial chart
   @value ctFilledRadar Like ctRadar, but the area enclosed by the series is filled.
   @value ctPie     Pie chart
   @value ctRing    Like pie chart, but with an inner hole
   @value ctStock   Financial chart }
  TsChartType = (ctEmpty, ctBar, ctLine, ctArea, ctBarLine, ctScatter, ctBubble,
    ctRadar, ctFilledRadar, ctPie, ctRing, ctStock);

  {@@ The TsChartAxis class represents all elements required to characterize a
    chart axis. }
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
    {@@ Position of the axis: left, top, right, bottom side of the chart area }
    property Alignment: TsChartAxisAlignment read FAlignment write FAlignment;
    {@@ If true, the axis maximum is determined automatically from the plotted data. Otherwise the @link(Max) must be specified. }
    property AutomaticMax: Boolean read FAutomaticMax write FAutomaticMax;
    {@@ If true, the axis minimum is determined automatically from the plotted data. Otherwise the @link(Min) must be specified. }
    property AutomaticMin: Boolean read FAutomaticMin write FAutomaticMin;
    {@@ If true, the major axis tick interval is determined automatically from the plotted data. Otherwise the @link(MajorInterval) must be specified. }
    property AutomaticMajorInterval: Boolean read FAutomaticMajorInterval write FAutomaticMajorInterval;
    {@@ If true, the minor axis tick interval is determined automatically from the plotted data. Otherwise the @link(MinorInterval) must be specified. }
    property AutomaticMinorInterval: Boolean read FAutomaticMinorInterval write FAutomaticMinorInterval;
    {@@ If true, the count of minor axis ticks is determined automatically from the plotted data. Otherwise the @link(MinorCount) must be specified. }
    property AutomaticMinorSteps: Boolean read FAutomaticMinorSteps write FAutomaticMinorSteps;
    {@@ Describes the style and width of the axis line }
    property AxisLine: TsChartLine read FAxisLine write FAxisLine;
    {@@ Cell range which provides the axis categories (x axis text labels). }
    property CategoryRange: TsChartRange read FCategoryRange write FCategoryRange;
    {@@ When true the axis plots date/time values. }
    property DateTime: Boolean read FDateTime write FDateTime;
    {@@ When true the axis title is rotated by default. }
    property DefaultTitleRotation: Boolean read FDefaultTitleRotation write FDefaultTitleRotation;
    {@@ If true the axis runs in the opposite direction from maximum to minimum, rather than from minimum to maximum }
    property Inverted: Boolean read FInverted write FInverted;
    {@@ Defines the font used to write the axis labels. }
    property LabelFont: TsFont read FLabelFont write FLabelFont;
    {@@ Format mask used to convert the axis values to numeric strings. }
    property LabelFormat: String read FLabelFormat write FLabelFormat;
    {@@ Format mask used to convert the axis values to date/time strings. }
    property LabelFormatDateTime: String read FLabelFormatDateTime write FLabelFormatDateTime;
    {@@ If true, label formatting is taken from the corresponding cells. }
    property LabelFormatFromSource: Boolean read FLabelFormatFromSource write FLabelFormatFromSource;
    {@@ Format maks used to convert the axis values to percentage strings. }
    property LabelFormatPercent: String read FLabelFormatPercent write FLabelFormatPercent;
    {@@ Angle, in degrees, by which axis labels are rotated. }
    property LabelRotation: Single read FLabelRotation write FLabelRotation;
    {@@ When true the axis is labeled in a logarithmic scale, otherwise in a linear scale. }
    property Logarithmic: Boolean read FLogarithmic write FLogarithmic;
    {@@ Base of the logarithms when the axis is in @link(Logarithmic) mode. }
    property LogBase: Double read FLogBase write FLogBase;
    {@@ Style and width of the minor grid lines running perpendicular to the axis. }
    property MajorGridLines: TsChartLine read FMajorGridLines write FMajorGridLines;
    {@@ Interval between the major ticks. Used when @link(AutomaticMajorInterval) is true. }
    property MajorInterval: Double read FMajorInterval write SetMajorInterval;
    {@@ Defines whether major ticks are drawn inside, outside or at both sides of the plot area. }
    property MajorTicks: TsChartAxisTicks read FMajorTicks write FMajorTicks;
    {@@ Value where the axis is forced to end when @link(AutomaticMax) is true. }
    property Max: Double read FMax write SetMax;
    {@@ Value where the axis is forced go begin when @link(AutomaticMin) is true. }
    property Min: Double read FMin write SetMin;
    {@@ Style and width of the minor grid lines running perpendicular to the axis. }
    property MinorGridLines: TsChartLine read FMinorGridLines write FMinorGridLines;
    {@@ Number of minor ticks between consecutive major ticks. Used when @link(AutomaticMinorSteps) is true. }
    property MinorCount: Integer read FMinorCount write SetMinorCount;
    {@@ Interval between the minor ticks. Used when @link(AutomaticMinorInterval) is true. }
    property MinorInterval: Double read FMinorInterval write SetMinorInterval;
    {@@ Defines whether minor ticks are drawn inside, outside or at both sides of the plot area. }
    property MinorTicks: TsChartAxisTicks read FMinorTicks write FMinorTicks;
    {@@ Defines the position where the axis is crossed by the other (orthogonal) axis: start, end, or specific value }
    property Position: TsChartAxisPosition read FPosition write FPosition;
    {@@ Value at which the axis is crossed by the other (orthogonal) axis when Position is @link(capValue). }
    property PositionValue: Double read FPositionValue write FPositionValue;
    {@@ Allows to show/hide the axis labels. }
    property ShowLabels: Boolean read FShowLabels write FShowLabels;
    {@@ Text displayed as title along the axis }
    property Title: TsChartText read FTitle write FTitle;
    {@@ Angle by which the axis title is rotated. }
    property TitleRotationAngle: Single read GetTitleRotationAngle;
    {@@ Allows to show/hide the entire axis}
    property Visible;
  end;

  {@@ Enumeration of the possible legend positions
   @value legRight  Legend at the right side of the plot area
   @value legTop    Legend above the plot area
   @value legBottom Legend below the plot area
   @value legLeft   Legend at the left side of the plot area }
  TsChartLegendPosition = (legRight, legTop, legBottom, legLeft);

  {@@ TsChartLegend collects all parameters needed for displaying the chart legend.
    It descends from @link(TsChartFillElement). }
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
    {@@ When false the legend is drawn outside the plot area. }
    property CanOverlapPlotArea: Boolean read FCanOverlapPlotArea write FCanOverlapPlotArea;
    {@@ Font used to draw the legend item texts }
    property Font: TsFont read FFont write FFont;
    {@@ Position of the legend: at the right, top, bottom or left of the chart area. }
    property Position: TsChartLegendPosition read FPosition write FPosition;
    {@@ X shift of the legend from its default position }
    property PosX: Double read FPosX write FPosX;
    {@@ Y shift of the legend from its default position }
    property PosY: Double read FPosY write FPosY;
    // There is also a "legend-expansion" but this does not seem to have a visual effect in Calc.
  end;

  TsChartAxisLink = (calPrimary, calSecondary);

  {@@ Enumeration which determines the piece of information displayed as a series label
   @value cdlValue       displays the numerical y value
   @value cdlPercentage  in case of a stacked series, displays the percentages of each stack y value of the sum of all stacks
   @value cdlCategory    displays the category value of the series data point
   @value cdlSeriesName  displays the name of the series
   @value cdlSymbol      adds the series symbol to the series label display
   @value cdlLeaderLines displays a line from data point to label. }
  TsChartDataLabel = (cdlValue, cdlPercentage, cdlCategory, cdlSeriesName, cdlSymbol, cdlLeaderLines);

  {@@ Set of TsChartDatalabel elements which defines which information is displayed in series labels next to each data point. }
  TsChartDataLabels = set of TsChartDataLabel;

  {@@ Defines the position of series labels with respect to the data points. Not all items are available for all series types.
   @value lpDefault      Default position (depends on series)
   @value lpAbove        The labels are above the data points / filled area
   @value lpBelow        The labels are below the data points / filled area
   @value lpCenter       The labels are in the center of the data points / filled area
   @value lpOutside      In case of a bar series: the labels are outside the filled area
   @value lpInside       In case of a bar series: the labels are inside tha filled area
   @value lpNearOrigin   In case of bar series: the labels are placed near the origin.
   @value lpLeft         In case of point series: the labels are at the left of the data points.
   @value lpRight        In case of point series: the labels are at the right of the data points.
   @value lpAvoidOverlap In case of pie series: the labels are placed to avoid overlaps. }
  TsChartLabelPosition = (lpDefault, lpAbove, lpBelow, lpCenter, lpOutside, lpInside, lpNearOrigin, lpLeft, lpRight, lpAvoidOverlap);

  {@@ Defines the shape around the label text
   @value lcsRectangle   Rectangular box
   @value lcsRoundRect   Rectangle with round corners
   @value lcsEllipse     Ellipse
   @value lcsLeftArrow   Rectangle with arrow to the left
   @value lcsUpArrow     Rectangle with arrow upward
   @value lcsRightArrow  Rectangle with arrow to the right
   @value lcsDownArrow   Rectangle with arrow downward
   @value lcsRectangleWedge  Rectangle with callout arrow
   @value lcsRoundRectWedge  Rounded rectangle with callout arrow
   @value lcsEllipseWedge    Ellipse with callout arrow }
  TsChartLabelCalloutShape = (
    lcsRectangle, lcsRoundRect, lcsEllipse,
    lcsLeftArrow, lcsUpArrow, lcsRightArrow, lcsDownArrow,
    lcsRectangleWedge, lcsRoundRectWedge, lcsEllipseWedge
  );

  {@@ The TsChartDataPointStyle class combines all formatting parameters of data point labels. }
  TsChartDataPointStyle = class(TsChartFillElement)
  private
    FDataPointIndex: Integer;
    FPieOffset: Integer;
  public
    procedure CopyFrom(ASource: TsChartElement);
    {@@ Index of the datapoint to which the style is applied. }
    property DataPointIndex: Integer read FDataPointIndex write FDataPointIndex;
    {@@ In case of a pie series: Represents the percentage of the radius by which the pie with the DatapointIndex is moved away from tne center }
    property PieOffset: Integer read FPieOffset write FPieOffset;  // Percentage
  end;

  {@@ List containing the formatting parameters for each data point label. }
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
    {@@ List element cast to the TsChartDataPointStyle }
    property Items[AIndex: Integer]: TsChartDataPointStyle read GetItem write SetItem; default;
  end;

  {@@ Enumeration to define how a trend line is fitted to the data points of a series
    @value  tltNone     No trend line
    @value  tltLinear   A linear function is fitted (y = a + b x)
    @value  tltLogarithmic  A logarithmic function is fitted (y = a ln(x) )
    @value  tltExponential  An exponential function is fitted  (y = a exp(b x)
    @value  tltPower        A power function is fitted (y = a x^b)
    @value  tltPolynomial   A polynomial is fitted (y = a + b x + c x^2 + ...) }
  TsTrendlineType = (tltNone, tltLinear, tltLogarithmic, tltExponential, tltPower, tltPolynomial);

  {@@ Class which collects the parameters for drawing a mathemical equation in the chart.
    The equation represents the fit of a trend line to the data values. }
  TsTrendlineEquation = class
    {@@ Background of the equation output }
    Fill: TsChartFill;
    {@@ Font to be used for drawing the equation string }
    Font: TsFont;
    {@@ Border of the equation output }
    Border: TsChartLine;
    {@@ Excel-style format string for formatting numbers in the equation string. }
    NumberFormat: String;
    {@@ Position of the formula's left side on the chart, in millimeters, relative to the outer chart boundary }
    Left: Double;
    {@@ Position of the formula's upper side on the chart, in millimeters, relative to the outer chart boundary }
    Top: Double;
    {@@ Name of the independent variable in the expression string, by default "x" }
    XName: String;
    {@@ Name of the dependent variable in the expression string, by default "f(x)" }
    YName: String;

    constructor Create;
    destructor Destroy; override;
    function IsDefaultBorder: Boolean;
    function IsDefaultFill: Boolean;
    function IsDefaultFont: Boolean;
    function IsDefaultNumberFormat: Boolean;
    function IsDefaultPosition: Boolean;
    function IsDefaultXName: Boolean;
    function IsDefaultYName: Boolean;
  end;

  {@@ Class which represents a trend line fitted to a series of data points. }
  TsChartTrendline = class
    {@@ Legend text for the trend line }
    Title: String;
    {@@ Represents the mathematical equation fitted to the data points. }
    TrendlineType: TsTrendlineType;
    {@@ Extends the trend line by this distance in positive x direction }
    ExtrapolateForwardBy: Double;
    {@@ Extends the trend line by this distance in negative x direction }
    ExtrapolateBackwardBy: Double;
    {@@ Forces the trend line to intersect the y axis at the given YInterceptValue }
    ForceYIntercept: Boolean;
    {@@ y value at which the trend line intersects the y axis when ForceYIntercept is true }
    YInterceptValue: Double;
    {@@ Degree of the polynomial fitted to the data points when TrendLineType is tltPolynomial }
    PolynomialDegree: Integer;
    {@@ Displays the fitted equation in the chart }
    DisplayEquation: Boolean;
    {@@ Displays the "goodness of fit parameter" (R-squared) in the chart. }
    DisplayRSquare: Boolean;
    {@@ Describes the fitted equation }
    Equation: TsTrendlineEquation;
    {@@ Drawing parameters for the trend line (pattern, color, line width, ...) }
    Line: TsChartLine;
    constructor Create;
    destructor Destroy; override;
  end;

  {@@ Enumeration which determines the size of error bars
    @value cebkNone     Do not display error bars
    @value cebkConstant Error bars have a the same size for all data points
    @value cebkPercentage  The size of error bars is a percentage of the data point value
    @value cebkCellRange   The size of error bars is specified in a cell range separately for each data point. }
  TsChartErrorBarKind = (cebkNone, cebkConstant, cebkPercentage, cebkCellRange);

  {@@ Class describing error bars to be drawn at each data point }
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
    {@@ Characterizes the size of the error bars /none, constant, proportional, individual) }
    property Kind: TsChartErrorBarKind read FKind write SetKind;
    {@@ Style parameters for drawing the error bar line parts. }
    property Line: TsChartLine read FLine write SetLine;
    {@@ Cell range with the lengths of the positive error bars, used when Kind is cebkCellRange}
    property RangePos: TsChartRange index 0 read GetRange write SetRange;
    {@@ Cell range with the lengths of the negative error bars, used when Kind is cebkCellRange }
    property RangeNeg: TsChartRange index 1 read GetRange write SetRange;
    {@@ Series at which the error bars are to be displayed. }
    property Series: TsChartSeries read FSeries;
    {@@ If true, terminating cross bar is drawn at the ends of the error bars. }
    property ShowEndCap: Boolean read FShowEndCap write FShowEndCap;
    {@@ Allows to show/hide the positive error bars. }
    property ShowPos: Boolean index 0 read GetShow write SetShow;
    {@@ Allows to show/hide the negative error bars. }
    property ShowNeg: Boolean index 1 read GetShow write SetShow;
    {@@ Value of the positive error bar, used when Kind is cebkConstant or cebkPercentage }
    property ValuePos: Double index 0 read GetValue write SetValue;
    {@@ Value of the negative error bar, used when Kind is cebkConstant or cebkPercentage }
    property ValueNeg: Double index 1 read GetValue write SetValue;
  end;

  {@@ Ancestor of all spreadsheet series types }
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
    {@@ Trendline instance, fitted to the data points. Made public when this feature is useful. }
    property Trendline: TsChartTrendline read FTrendline write FTrendline;
  public
    constructor Create(AChart: TsChart); virtual;
    destructor Destroy; override;
    function GetCount: Integer;
    function GetDefaultSeriesColor: TsChartColor;
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

    {@@ Type of the series: bar, area, line, scatter, ... }
    property ChartType: TsChartType read GetChartType;
    {@@ Number of data points contained in the series }
    property Count: Integer read GetCount;
    {@@ Determines which kind of data point labels are displayed. }
    property DataLabels: TsChartDataLabels read FDataLabels write FDataLabels;
    {@@ Determines the shapes which are drawn around data point labels. }
    property DataLabelCalloutShape: TsChartLabelCalloutShape read FDataLabelCalloutShape write FDataLabelCalloutShape;
    {@@ Lists individual styles how individual data points are displayed. }
    property DataPointStyles: TsChartDatapointStyleList read FDataPointStyles;
    {@@ Cell range defining individual data point fill colors. }
    property FillColorRange: TsChartRange read FFillColorRange write FFillColorRange;
    {@@ Index of the group to which this series belongs. }
    property GroupIndex: Integer read FGroupIndex write FGroupIndex;
    {@@ Fill properties for the data point label backgrounds. }
    property LabelBackground: TsChartFill read FLabelBackground write FLabelBackground;
    {@@ Line properties for the data point label borders. }
    property LabelBorder: TsChartLine read FLabelBorder write FLabelBorder;
    {@@ Font used for drawing the data point labels. }
    property LabelFont: TsFont read FLabelFont write FLabelFont;
    {@@ Number format (in Excel notation) for drawing numerical data point labels. }
    property LabelFormat: String read FLabelFormat write FLabelFormat;  // Number format in Excel notation, e.g. '0.00'
    {@@ Number format (in Excel notation) for drawing percentage data point labels. }
    property LabelFormatPercent: String read FLabelFormatPercent write FLabelFormatPercent;
    {@@  Position of the data point labels relative to the series. }
    property LabelPosition: TsChartLabelPosition read FLabelPosition write FLabelPosition;
    {@@ Cell range containing data point labels. }
    property LabelRange: TsChartRange read FLabelRange write FLabelRange;
    {@@ Separator used when several pieces of information are combined in the data point label. }
    property LabelSeparator: string read FLabelSeparator write FLabelSeparator;
    {@@ Cell range containing data-point-related line colors. }
    property LineColorRange: TsChartRange read FLineColorRange write FLineColorRange;
    {@@ Index of the series in the chart's series list }
    property Order: Integer read FOrder write FOrder;
    {@@ Address of the cell which contains the series title (legend text). Use '\n' for line-breaks. }
    property TitleAddr: TsChartCellAddr read FTitleAddr write FTitleAddr;
    {@@ Information variable which is set to true when the series type supports trend lines. }
    property SupportsTrendline: Boolean read FSupportsTrendline;
    {@@ Determines whether the chart's XAxis or X2Axis is used for the independent axis of the series }
    property XAxis: TsChartAxisLink read FXAxis write FXAxis;
    {@@ Determines how error bars are drawn in the x direction. }
    property XErrorBars: TsChartErrorBars read FXErrorBars write SetXErrorBars;
    {@@ Cell range containing the x values of the data points. }
    property XRange: TsChartRange read FXRange write FXRange;
    {@@ Determines whether the chart's YAxis or Y2Axis is used for the dependent axis of the series }
    property YAxis: TsChartAxisLink read FYAxis write FYAxis;
    {@@ Determines how error bars are drawn in the y direction. }
    property YErrorBars: TsChartErrorBars read FYErrorBars write SetYErrorBars;
    {@@ Cell range containing the y values of the data points. }
    property YRange: TsChartRange read FYRange write FYRange;

    {@@ Parameters determining how the series is filled. What actually is filled depends on the series type. }
    property Fill: TsChartFill read FFill write FFill;
    {@@ Parameters determining how the series lines are drawn. Which lines are meant depends on the series type. }
    property Line: TsChartLine read FLine write FLine;
  end;

  {@@ Class type for all TsChartSeries classes }
  TsChartSeriesClass = class of TsChartSeries;

  {@@ Area series draws data values as a line and fills the area between the
    data point line and the zero value. }
  TsAreaSeries = class(TsChartSeries)
  public
    constructor Create(AChart: TsChart); override;
    property Trendline;
  end;

  {@@ TsBarSeries displays data values as bars. }
  TsBarSeries = class(TsChartSeries)
  public
    constructor Create(AChart: TsChart); override;
    property Trendline;
  end;

  {@@ Enumeration of the symbols used for data points in TsLineSeries or
    TsScatterSeries
    @value  ccsRect     Rectangular symbol
    @value  ccsDiamond  Diamond-shaped rectangle (rectangle rotated by 90 deg)
    @value  ccsTriangle Triangular symbol pointing upward
    @value  ccsTriangleDown Triangular symbol pointing downward
    @value  ccsTriangleLeft Triangular symbol pointing to the left
    @value  ccsTriangleRight  Triangular shape pointing to the right
    @value  ccsCircle  Circular symbol
    @value  ccsStar    Star-shaped symbol
    @value  ccsX       Diagonal cross, like character 'x'
    @value  ccsPlut    Horizontal/vertical cross, like character '+'
    @value  ccsAsterisk Asterisk-like character ('*'
    @value  ccsDash    Short horizontal line ('-');
    @value  ccsDot     Small dot }
  TsChartSeriesSymbol = (
    cssRect, cssDiamond, cssTriangle, cssTriangleDown, cssTriangleLeft,
    cssTriangleRight, cssCircle, cssStar, cssX, cssPlus, cssAsterisk,
    cssDash, cssDot
  );

  {@@ Enumeration for the ways how data points of a line or scatter series
    will be conntected
    @value  ciLinear  Two data points are connected by a straight line segment
    @value  ciCubicSpline  Data points are conencted by a cubic spline.
    @value  ciBSpline  Data points are connected by a B-spline curve.
    @value  ciStepStart Two data points are connected by a stepped curve: first vertically, then horizontally
    @value  ciStepEnd   Two data points are connected by a stepped curve: first horizontally, then vertically
    @value  ciStepCenterX Two data points are connected by a stepped curve where the step is in the middle of the x values
    @value  ciStepCenterY Two data points are connected by a stepped curve where the step is in the middle of the y values}
  TsChartInterpolation = (
    ciLinear,
    ciCubicSpline, ciBSpline,
    ciStepStart, ciStepEnd, ciStepCenterX, ciStepCenterY
  );

  {@@ Ancestor class of the "line-type" series }
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
  protected
    property Interpolation: TsChartInterpolation read FInterpolation write FInterpolation;
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
    {@@ Determines how data points are connected by lines }
    property Interpolation;
    {@@ Determines the kind of data point symbols: rectangular, circular, etc. }
    property Symbol;
    {@@ Defines the border of the data point symbols }
    property SymbolBorder;
    {@@ Defines how data point symbols are filles. }
    property SymbolFill;
    {@@ Defines the height of the data point symbols, in millimeters }
    property SymbolHeight;
    {@@ Defines the width of the data point symbols, in millimeters }
    property SymbolWidth;
    {@@ If true data points are connected by lines. }
    property ShowLines;
    {@@ If true symbols are displayed at each data point. }
    property ShowSymbols;
    {@@ Contains parameters determinng whether and how a trendline is displayed. }
    property Trendline;
  end;

  {@@ Enumeration for the order of slices in a TsPieSeries
   @param soCCW  Slices are arranged in counter-cclockwise order
   @param so CW  Slices are arranged in clockwise order.
  }
  TsSliceOrder = (soCCW, soCW);

  {@@ TsPieSeries displays data as slices of a pie. }
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

  {@@ TsRadarSeries is a radial series in which x is interpreted as the angle
   around the origin, and y as the distance from the center. }
  TsRadarSeries = class(TsLineSeries)
  public
    constructor Create(AChart: TsChart); override;
  end;

  {@@ Like TsRadarSeries, but the area between series and origin is filled }
  TsFilledRadarSeries = class(TsRadarSeries)
  public
    constructor Create(AChart: TsChart); override;
  end;

  {@@ Ancestor of the TsScatterSeries. The x values in this series are
   distributed randomly. }
  TsCustomScatterSeries = class(TsCustomLineSeries)
  public
    constructor Create(AChart: TsChart); override;
    property Trendline;
  end;

  {@@ Class containing the parameters of a "scatter series", i.e. a series with
    irregular distribution of the x values. }
  TsScatterSeries = class(TsCustomScatterSeries)
  public
   {@@ Determines how data points are connected by lines }
    property Interpolation;
   {@@ Determines the kind of data point symbols: rectangular, circular, etc. }
    property Symbol;
   {@@ Defines the border of the data point symbols }
    property SymbolBorder;
   {@@ Defines how data point symbols are filles. }
    property SymbolFill;
   {@@ Defines the height of the data point symbols, in millimeters }
    property SymbolHeight;
   {@@ Defines the width of the data point symbols, in millimeters }
    property SymbolWidth;
   {@@ If true data points are connected by lines. }
    property ShowLines;
   {@@ If true symbols are displayed at each data point. }
    property ShowSymbols;
  end;

  {@@ Enumeration to determine whether the bubble size values are aassumed to be
    proportional to the bubble radius or to the bubble area.
    @value bsmRadius  The radius of each bubble is assumed to be proprotional to the bubble value.
    @value bsmArea    The area of each bubble is assumed to be proportional to the bubble value. }
  TsBubbleSizeMode = (bsmRadius, bsmArea);

  {@@ Class containing the parameters of a "bubble" series, i.e. a series with
    irregularly distributed x values and in which the data points are drawn as
    circles with specific sizes. }
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
    {@@ Range of cells containing the bubble sizes for each data point }
    property BubbleRange: TsChartRange read FBubbleRange;
    {@@ Scaling factor for the bubble radii }
    property BubbleScale: Double read FBubbleScale write FBubbleScale;
    {@@ Determines whether the bubble sizes in the BubbleRange are interpreted as bubble radii or bubble areas }
    property BubbleSizeMode: TsBubbleSizeMode read FBubbleSizeMode write FBubbleSizeMode;
  end;

  { Series type for financial data of stock prices: opening, closing, highest
    and lowest price during a trading period. }
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
    {@@ If true the stock data are displayed in candlestick mode, otherwise only as vertical bars. }
    property CandleStick: Boolean read FCandleStick write FCandleStick;
    {@@ Fill properties applied to the candle box for falling stock price at this data point. }
    property CandleStickDownFill: TsChartFill read FCandleStickDownFill write FCandleStickDownFill;
    {@@ Fill properties applied to the candle box for rising the stock price at this data point. }
    property CandleStickUpFill: TsChartFill read FCandleStickUpFill write FCandleStickUpFill;
    {@@ Line properties applied to the border of the candle box for falling stock price at this data point. }
    property CandleStickDownBorder: TsChartLine read FCandleStickDownBorder write FCandleStickDownBorder;
    {@@ Line properties applied to the border of the candle box for rising stock price at this data point. }
    property CandleStickUpBorder: TsChartLine read FCandleStickUpBorder write FCandleStickUpBorder;
    {@@ Width of the Open/Close tick in non-candle-stick mode, given as percentage of the data point distance. }
    property TickWidthPercent: Integer read FTickWidthPercent write FTickWidthPercent;
    {@@ Line properties of the range line connecting highest and lowest stock prise in non-cancdle-stick mode }
    property RangeLine: TsChartLine read FRangeLine write FRangeLine;
    {@@ Cell range containing the "open" data values (starting stock price) }
    property OpenRange: TsChartRange read FOpenRange;
    {@@ Cell range containing the "high" data values (highest stock price) }
    property HighRange: TsChartRange read FHighRange;
    {@@ Cell range containing the "low" data values (lowest stock price) }
    property LowRange: TsChartRange read FLowRange;
    {@@ Cell range containing  the "close" data values (closing stock price) }
    property CloseRange: TsChartRange read FYRange;
  end;

  {@@ List class storing the series of a chart. }
  TsChartSeriesList = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsChartSeries;
    procedure SetItem(AIndex: Integer; AValue: TsChartSeries);
  public
    {@@ List elements cast to the TsChartSeries class }
    property Items[AIndex: Integer]: TsChartSeries read GetItem write SetItem; default;
  end;

  {@@ Enumeration for enabling stacking of bar and area series
   @value csmDefault  No stacking
   @value csmStacked  Data points are stacked at the same x value
   @value csmStackedPercentage Data points are stacked and recalculated as percentages of the total sum. }
  TsChartStackMode = (csmDefault, csmStacked, csmStackedPercentage);

  {@@ The main chart class. Here all elements and properties are put together. }
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

    FGradients: TsChartGradientList;
    FFillPatterns: TsChartFillPatternList;
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
    function IsScatterChart: Boolean;

    {@@ Name for internal purposes to identify the chart during reading from file }
    property Name: String read FName write FName;
    {@@ Index of chart in workbook's chart list. }
    property Index: Integer read FIndex write FIndex;
    {@@ Worksheet into which the chart is embedded }
    property Worksheet: TsBasicWorksheet read FWorksheet write FWorksheet;
    {@@ Row index of the cell in which the chart has its top/left corner (anchor) }
    property Row: Cardinal read FRow write FRow;
    {@@ Column index of the cell in which the chart has its top/left corner (anchor) }
    property Col: Cardinal read FCol write FCol;
    {@@ Offset of the left chart edge relative to the anchor cell, in mm }
    property OffsetX: double read FOffsetX write FOffsetX;
    {@@ Offset of the top chart edge relative to the anchor cell, in mm }
    property OffsetY: double read FOffsetY write FOffsetY;
    {@@ Width of the chart, in mm }
    property Width: double read FWidth write FWidth;
    {@@ Height of the chart, in mm }
    property Height: double read FHeight write FHeight;
    {@@ Workbook to which the chart belongs }
    property Workbook: TsBasicWorkbook read FWorkbook write FWorkbook;

    {@@ Attributes of the entire chart background }
    property Background: TsChartFill read FBackground write FBackground;
    property Border: TsChartLine read FBorder write FBorder;

    {@@ Attributes of the plot area (rectangle enclosed by axes) }
    property PlotArea: TsChartFillElement read FPlotArea write FPlotArea;
    {@@ Attributes of the floor of a 3D chart }
    property Floor: TsChartFillElement read FFloor write FFloor;

    {@@ Attributes of the chart's title }
    property Title: TsChartText read FTitle write FTitle;
    {@@ Attributes of the chart's subtitle }
    property Subtitle: TsChartText read FSubtitle write FSubTitle;
    {@@ Attributes of the chart's legend }
    property Legend: TsChartLegend read FLegend write FLegend;

    {@@ Attributes of the plot's primary x axis (bottom) }
    property XAxis: TsChartAxis read FXAxis write FXAxis;
    {@@ Attributes of the plot's secondary x axis (top) }
    property X2Axis: TsChartAxis read FX2Axis write FX2Axis;
    {@@ Attributes of the plot's primary y axis (left) }
    property YAxis: TsChartAxis read FYAxis write FYAxis;
    {@@ Attributes of the plot's secondary y axis (right) }
    property Y2Axis: TsChartAxis read FY2Axis write FY2Axis;

    {@@ Gap between bars/bar groups, as percentage of single bar width }
    property BarGapWidthPercent: Integer read FBarGapWidthPercent write FBarGapWidthPercent;
    {@@ Overlapping of bars, as percentage of single bar width }
    property BarOverlapPercent: Integer read FBarOverlapPercent write FBarOverlapPercent;

    {@@ Characterizes the connecting line between data points (for line and scatter series) }
    property Interpolation: TsChartInterpolation read FInterpolation write FInterpolation;
    {@@ When true, x and y axes are exchanged (mainly for bar series, but works also for scatter and bubble series) }
    property RotatedAxes: Boolean read FRotatedAxes write FRotatedAxes;
    {@@ Allows stacking of bar and area series }
    property StackMode: TsChartStackMode read FStackMode write FStackMode;

    {@@ Cell range which defines the labels at the category axis (x axis) }
    property CategoryLabelRange: TsChartRange read GetCategoryLabelRange;

    { Attributes of the series }
    {@@ List containing all the series which are displayed by the chart. }
    property Series: TsChartSeriesList read FSeriesList write FSeriesList;

    { Style lists }
    {@@ List containing all the fill gradients which have been defined for the chart. }
    property Gradients: TsChartGradientList read FGradients;
    {@@ List containing all the fill patterns which have been defined for the chart. }
    property FillPatterns: TsChartFillPatternList read FFillPatterns;
    {@@ List containing all the images which have been provided for filling purposes. }
    property Images: TsChartImageList read FImages;
  end;

  {@@ A list of charts. The stored charts are automatically destroyed by the list. }
  TsChartList = class(TObjectList)
  private
    function GetItem(AIndex: Integer): TsChart;
    procedure SetItem(AIndex: Integer; AValue: TsChart);
  public
    {@@ List elements cast to the TsChart class }
    property Items[AIndex: Integer]: TsChart read GetItem write SetItem; default;
  end;

  {@@ An dynamic array of charts }
  TsChartArray = array of TsChart;


function ChartColor(AColor: TsColor; ATransparency: TsChartTransparency = 0.0): TsChartColor;

implementation

uses
  Math, fpsPatterns;

{ TsChartColor }

{@@ ----------------------------------------------------------------------------
  Helper function to create a @link(TsChartColor) record as used for colors
  in the fpspreadsheet charts.
  The record contains a Transparency field in addition to the standard
  @link(TsColor) field.

  @param    AColor         RGB color
  @param    ATransparency  Transparency of the color, value between 0.0 and 1.0. Default: 0.0 (opaque)

  @returns  A TsChartColor record
-------------------------------------------------------------------------------}
function ChartColor(
  AColor: TsColor;
  ATransparency: TsChartTransparency = 0.0): TsChartColor;
begin
  Result.Color := AColor;
  Result.Transparency := ATransparency;
end;


{ TsChartLine }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChartLine class.

  Defaults to a solid black line of default line width.
-------------------------------------------------------------------------------}
constructor TsChartLine.Create;
begin
  inherited Create;
  SelectSolidLine(ChartColor(scBlack), PtsToMM(DEFAULT_CHART_LINEWIDTH));
end;

{@@ ----------------------------------------------------------------------------
  Creates a line with solid, un-patterned line.

  @param  AColor   Color of the line. This is a chart color containing a Transparency element.
  @param  AWidth   Linewidth, in millimeters. }
constructor TsChartLine.CreateSolid(AColor: TsChartColor; AWidth: Double);
begin
  inherited Create;
  SelectSolidLine(AColor, AWidth);
end;

{@@ ----------------------------------------------------------------------------
  Copies the parameters from the specified line

  @param  ALine   Line instance to be copied.
-------------------------------------------------------------------------------}
procedure TsChartLine.CopyFrom(ALine: TsChartLine);
begin
  if ALine <> nil then
  begin
    Style := ALine.Style;
    Width := ALine.Width;
    Color := ALine.Color;
  end;
end;

function TsChartLine.GetStyle: TsChartLinePatternStyle;
begin
  if (PatternIndex >= 0) and (PatternIndex < ord(clsCustom)) then
    Result := TsChartLinePatternStyle(PatternIndex)
  else
    Result := clsCustom;
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the line is hidden, i.e. its style is set to clsNoLine
-------------------------------------------------------------------------------}
function TsChartLine.IsHidden: Boolean;
begin
  Result := (PatternIndex = ord(clsNoLine));
end;

{@@ ----------------------------------------------------------------------------
  Switches the line to be hidden.
-------------------------------------------------------------------------------}
procedure TsChartLine.SelectNoLine;
begin
  Style := clsNoLine;
end;

{@@ ----------------------------------------------------------------------------
  Assigns a predefined patterned line to the TsChartLine instance.

  @param ALineStyle   A clsXXXX enumeration element of TsChartLinePatternStyle
  @param AColor       Color of the line
  @param ALineWidth   Line width, in mm. If omitted (or -1) the default linewidth is used.
-------------------------------------------------------------------------------}
procedure TsChartLine.SelectPatternLine(ALineStyle: TsChartLinePatternStyle;
  AColor: TsChartColor; ALineWidth: Single = -1.0);
begin
  Style := ALineStyle;
  Color := AColor;
  if ALineWidth < 0 then
    Width := PtsToMM(DEFAULT_CHART_LINEWIDTH)
  else
    Width := ALineWidth;
end;

{@@ ----------------------------------------------------------------------------
  Assigns a patterned line to the TsChartLine instance.

  @param  APatternIndex  Index into the global raw line patterns list, see also clsXXXX variables
  @param  AColor         Color of the line
  @param  ALineWidth     Line width, in mm. If omitted (or -1) the default linewidth is used.
-------------------------------------------------------------------------------}
procedure TsChartLine.SelectPatternLine(APatternIndex: Integer;
  AColor: TsChartColor; ALineWidth: Single = -1.0);
begin
  PatternIndex := APatternIndex;
  Color := AColor;
  if ALineWidth = -1.0 then
    Width := PtsToMM(DEFAULT_CHART_LINEWIDTH)
  else
    Width := ALineWidth;
end;

{@@ ----------------------------------------------------------------------------
  Makes the TsChartLine instance a solid line:

  @param  AColor      Color of the line
  @param  ALineWidth  Line width, in mm. If omitted (or -1) the default linewidth is used.
-------------------------------------------------------------------------------}
procedure TsChartLine.SelectSolidline(AColor: TsChartColor; ALineWidth: Single = -1.0);
begin
  SelectPatternLine(clsSolid, AColor, ALineWidth);
end;

procedure TsChartLine.SetStyle(AValue: TsChartLinePatternStyle);
begin
  PatternIndex := ord(AValue);
end;


{ TsChartGradient }

{@@ ----------------------------------------------------------------------------
  Constructor of the gradient. Adds a begin and an end step of the gradient.
-------------------------------------------------------------------------------}
constructor TsChartGradient.Create;
begin
  inherited Create;
  SetLength(FSteps, 2);
  SetStep(0, 0.0, ChartColor(scBlack));
  SetStep(1, 1.0, ChartColor(scWhite));
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsChartGradient class.
-------------------------------------------------------------------------------}
destructor TsChartGradient.Destroy;
begin
  Name := '';
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Adds a new color step to the gradient. The new color is inserted at the
  correct index according to its value so that all values in the steps are
  ordered. If the exact step value is already existing the gradient step is replaced.

  @param  AValue  A floating point value between 0 and 1 determining at which position the specified color is used.
  @param  AColor  Color used at position AValue of the gradient.
-------------------------------------------------------------------------------}
procedure TsChartGradient.AddStep(AValue: Single; AColor: TsChartColor);
const
  EPS = 1E-6;
var
  i, j, idx: Integer;
begin
  if AValue < 0 then AValue := 0.0;
  if AValue > 1 then AValue := 1.0;

  if Length(FSteps) > 0 then
  begin
    for i := 0 to High(FSteps) do
    begin
      if SameValue(FSteps[i].Value, AValue, EPS) then
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
  SetStep(idx, AValue, AColor);
end;

{@@ ----------------------------------------------------------------------------
  Copies the gradient parameters from ASource
-------------------------------------------------------------------------------}
procedure TsChartGradient.CopyFrom(ASource: TsChartGradient);
var
  i: Integer;
begin
  Name := ASource.Name;
  Style := ASource.Style;
  SetLength(FSteps, ASource.NumSteps);
  for i := 0 to Length(FSteps)-1 do
    FSteps[i] := ASource.Steps[i];
  CenterX := ASource.CenterX;
  CenterY := ASource.CenterY;
  Angle := ASource.Angle;
end;

function TsChartGradient.GetBorder(AIndex: Integer): Single;
begin
  case AIndex of
    0: Result := FSteps[0].Value;
    1: Result := FSteps[High(FSteps)].Value;
  end;
end;

function TsChartGradient.GetColor(AIndex: Integer): TsChartColor;
begin
  case AIndex of
    0: Result := FSteps[0].Color;
    1: Result := FSteps[High(FSteps)].Color;
  end;
end;

function TsChartGradient.GetSteps(AIndex: Integer): TsChartGradientStep;
begin
  if AIndex < 0 then AIndex := 0;
  if AIndex >= Length(FSteps) then AIndex := Length(FSteps) - 1;
  Result := FSteps[AIndex];
end;

{@@ ----------------------------------------------------------------------------
  Returns the number of steps (pre-defined colors) used by the gradient
-------------------------------------------------------------------------------}
function TsChartGradient.NumSteps: Integer;
begin
  Result := Length(FSteps);
end;

procedure TsChartGradient.SetBorder(AIndex: Integer; AValue: Single);
begin
  FSteps[AIndex].Value := AValue;
end;

procedure TsChartGradient.SetStep(AIndex: Integer; AValue: Double;
  AColor: TsChartColor);
begin
  FSteps[AIndex].Value := AValue;
  FSteps[AIndex].Color := AColor;
end;


{ TsChartGradientList }

{@@ ----------------------------------------------------------------------------
  Creates an axial gradient and adds it to the gradient list.
  This kind of gradient combines two linear gradient running from the outside
  and meeting at the center of shape

  Not supported by xlsx where it is replaced by a rectangular gradient.

  @param  AName   Name of the gradient. Must be unique. If a gradient with the same name already exists its parameters will be replaced by the new ones.
  @param  AStartColor  Color of the outermost gradient line
  @param  AEndColor    Color in the center of the gradient
  @param  AAngle       Rotation angle of the lines, in degrees. Measured relative to x axis and increases in CCW direction.
  @returns Index of the gradient in the gradient list.
-------------------------------------------------------------------------------}
function TsChartGradientList.AddAxialGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsAxial,
    AStartColor, AEndColor,
    AAngle, 0.0, 0.0
  );
end;

{@@ ----------------------------------------------------------------------------
  Creates an elliptical gradient and adds it to the gradient list.
  In this kind of gradient the lines of constant color are ellipses with size shrinking towards the center.

  @param  AName   Name of the gradient. Must be unique. If a gradient with the same name already exists its parameters will be replaced by the new ones.
  @param  AStartColor  Color of the outermost gradient ellipse (and beyond, if StartBorder > 0)
  @param  AEndColor    Color in the center of the gradient ellipse (and beyond, if EndBorder < 1)
  @param  ACenterX     Horizontal center point of the ellipses. Relative to the width of the filled shape (0.0 ... 1.0)
  @param  ACenterY     Vertical center point of the ellipses. Relative to the height of the filled shape (0.0 ... 1.0)
  @param  AAngle       Rotation angle of the ellipses, in degrees. Measured relative to x axis and increases in CCW direction.
  @returns Index of the gradient in the gradient list.
-------------------------------------------------------------------------------}
function TsChartGradientList.AddEllipticGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AAngle, ACenterX, ACenterY: Double): Integer;
begin
  Result := AddGradient(AName, cgsRectangular,
    AStartColor, AEndColor,
    AAngle,  // is ignored in xlsx
    ACenterX, ACenterY
  );
end;

{@@ ----------------------------------------------------------------------------
  Adds a pre-created gradient to the gradient list.

  @param  AName   Name of the gradient. Must be unique. If a gradient with the same name already exists its parameters will be replaced by the new ones.
  @param  AGradient    Instance of the gradient to be added.
  @returns Index of the gradient in the gradient list.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Creates a general gradient and adds it to the gradient list.

  Not supported by xlsx where it is replaced by the radial gradient.

  @param  AName   Name of the gradient. Must be unique. If a gradient with the same name already exists its parameters will be replaced by the new ones.
  @param  AStyle  Kind of gradient (linear, radial, rectangular, ...).
  @param  AStartColor  Color at the position where the gradient starts
  @param  AEndColor    Color at the position where the gradient ends
  @param  ACenterX     Horizontal center point of the gradient. Relative to the width of the filled shape (0.0 ... 1.0). Ignored by some gradient types.
  @param  ACenterY     Vertical center point of the gradient. Relative to the height of the filled shape (0.0 ... 1.0). Ignored by some gradient types.
  @param  AAngle       Direction of the color change, in degrees, increasing in CCW direction. Ignored by some gradient types.
  @param  AStartBorder Relative position along the gradient direction across the shape where the gradient starts.
  @param  AEndBorder   Relative position along the gradient direction across the shape where the gradient ends.
  @returns Index of the gradient in the gradient list.
-------------------------------------------------------------------------------}
function TsChartGradientList.AddGradient(AName: String; AStyle: TsChartGradientStyle;
  AStartColor, AEndColor: TsChartColor; AAngle, ACenterX, ACenterY: Double;
  AStartBorder: Double = 0.0; AEndBorder: Double = 1.0): Integer;
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
  item.AddStep(AStartBorder, AStartColor);
  item.AddStep(AEndBorder, AEndColor);
  item.Angle := AAngle;
  item.CenterX := ACenterX;
  item.CenterY := ACenterY;
end;

{@@ ----------------------------------------------------------------------------
  Creates a linear gradient and adds it to the gradient list.
  In this kind of gradient the contours of constant color are straight lines.

  @param  AName   Name of the gradient. Must be unique. If a gradient with the same name already exists its parameters will be replaced by the new ones.
  @param  AStartColor  Color at the position where the gradient starts
  @param  AEndColor    Color at the position where the gradient ends
  @param  AAngle       Direction of the color change, in degrees. When 0 the gradient begins at the left and ends at the right side of the filled shape. The angle increases in CCW direction.
  @returns Index of the gradient in the gradient list.
-------------------------------------------------------------------------------}
function TsChartGradientList.AddLinearGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AAngle: Double): Integer;
begin
  Result := AddGradient(AName, cgsLinear,
    AStartColor, AEndColor, AAngle, 0.0, 0.0
  );
end;

{@@ ----------------------------------------------------------------------------
  Creates a radial gradient and adds it to the gradient list.
  In this kind of gradient the lines of constant color are circles around a given center shrinking towards the center.

  @param  AName   Name of the gradient. Must be unique. If a gradient with the same name already exists its parameters will be replaced by the new ones.
  @param  AStartColor  Color of the outermost gradient circle
  @param  AEndColor    Color in the center of the gradient circles
  @param  ACenterX     Horizontal center point of the circles. Relative to the width of the filled shape (0.0 ... 1.0)
  @param  ACenterY     Vertical center point of the circles. Relative to the height of the filled shape (0.0 ... 1.0)
  @returns Index of the gradient in the gradient list.
-------------------------------------------------------------------------------}
function TsChartGradientList.AddRadialGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; ACenterX, ACenterY: Double): Integer;
begin
  Result := AddGradient(AName, cgsRadial,
    AStartColor, AEndColor,
    0,
    ACenterX, ACenterY
  );
end;

{@@ ----------------------------------------------------------------------------
  Creates a rectangular gradient and adds it to the gradient list.
  In this kind of gradient the lines of constant color are rectangles like the bounds of the shape shrinking towards the center.

  @param  AName   Name of the gradient. Must be unique. If a gradient with the same name already exists its parameters will be replaced by the new ones.
  @param  AStartColor  Color of the outermost gradient rectangle
  @param  AEndColor    Color in the center of the gradient rectangles
  @param  ACenterX     Horizontal center point of the rectangles. Relative to the width of the filled shape (0.0 ... 1.0)
  @param  ACenterY     Vertical center point of the rectangles. Relative to the height of the filled shape (0.0 ... 1.0)
  @param  AAngle       Rotation angle of the rectangles, in degrees. Measured relative to x axis and increases in CCW direction. Ignored by xlsx.
  @returns Index of the gradient in the gradient list.
-------------------------------------------------------------------------------}
function TsChartGradientList.AddRectangularGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AAngle, ACenterX, ACenterY: Double): Integer;
begin
  Result := AddGradient(AName, cgsRectangular,
    AStartColor, AEndColor,
    AAngle,  // is ignored in xlsx
    ACenterX, ACenterY
  );
end;

{@@ ----------------------------------------------------------------------------
  Creates a square gradient and adds it to the gradient list.
  In this kind of gradient the lines of constant color are squares with size shrinking towards the center.

  Not supported by xlsx where it is replaced by a rectangular gradient.

  @param  AName   Name of the gradient. Must be unique. If a gradient with the same name already exists its parameters will be replaced by the new ones.
  @param  AStartColor  Color of the outermost gradient square
  @param  AEndColor    Color in the center of the gradient square
  @param  ACenterX     Horizontal center point of the squares. Relative to the width of the filled shape (0.0 ... 1.0)
  @param  ACenterY     Vertical center point of the square. Relative to the height of the filled shape (0.0 ... 1.0)
  @param  AAngle       Rotation angle of the squares, in degrees. Measured relative to x axis and increases in CCW direction.
  @returns Index of the gradient in the gradient list.
-------------------------------------------------------------------------------}
function TsChartGradientList.AddSquareGradient(AName: String;
  AStartColor, AEndColor: TsChartColor; AAngle, ACenterX, ACenterY: Double): Integer;
begin
  Result := AddGradient(AName, cgsSquare,
    AStartColor, AEndColor,
    AAngle, ACenterX, ACenterY
  );
end;

{@@ ----------------------------------------------------------------------------
  Searches for a gradient in the list by its name.

  @param    AName   Name of the gradient
  @returns  The found gradient instance, or nil if not found.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Searches for a gradient in the list by its name.

  @param    AName   Name of the gradient
  @returns  Index of the found gradient instance, or -1 if not found.
-------------------------------------------------------------------------------}
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


{ TsChartFillPattern}

{@@ ----------------------------------------------------------------------------
  Destructor of the TsChartFillPattern class.
-------------------------------------------------------------------------------}
destructor TsChartFillPattern.Destroy;
begin
  Name := '';
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Copies the fill pattern from another @link(TsChartFillPattern) instance ASource.

  @param  ASource  Fill pattern instance from which the parameters are to be copied.
-------------------------------------------------------------------------------}
procedure TsChartFillPattern.CopyFrom(ASource: TsChartFillPattern);
var
  i: Integer;
begin
  Name := ASource.Name;
  Index := ASource.Index;
  FgColor := ASource.FgColor;
  BgColor := ASource.BgColor;
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the pattern has not background, i.e. when the shape is filled
  by an empty, clear pattern.

  @returns  true when the pattern's background color is fully transparent.
-------------------------------------------------------------------------------}
function TsChartFillPattern.IsClearPattern: Boolean;
begin
  Result := (BgColor.Transparency = 1.0);
end;


{ TsChartFillPatternList }

{@@ ----------------------------------------------------------------------------
  Adds a non-filled pre-defined pattern to the chart's FillPattern list.

  @param AName          Name of the pattern. Must be unique.
  @param(APatternStyle  A TsChartFillPatternStyle enumeration value
                        identifying the pre-defined pattern type.)
  @param APatternColor  Color of the pattern.
-------------------------------------------------------------------------------}
function TsChartFillPatternList.AddPattern(AName: String;
  APatternStyle: TsChartFillPatternStyle; APatternColor: TsChartColor): Integer;
begin
  Result := AddPattern(AName, APatternStyle, APatternColor, ChartColor(scWhite, 1.0));
end;

{@@ ----------------------------------------------------------------------------
  Adds a filled pre-defined pattern to the chart's FillPattern list.

  @param AName          Name of the pattern. Must be unique.
  @param(APatternStyle  A @link(TsChartFillPatternStyle) enumeration value identifying the pre-defined pattern type.)
  @param APatternColor  Color of the pattern.
  @param ABackColor     Color of the pattern background.
-------------------------------------------------------------------------------}
function TsChartFillPatternList.AddPattern(AName: String;
  APatternStyle: TsChartFillPatternStyle; APatternColor, ABackColor: TsChartColor): Integer;
var
  pattern: TsChartFillPattern;
  rawPatternIdx: Integer;
begin
  if AName = '' then
    AName := DefaultPatternName;
  rawPatternIdx := ord(APatternStyle);
  if rawPatternIdx >= GetRawFillPatternCount then
    raise Exception.Create('Raw fill pattern not defined.');
  Result := NewPattern(AName);
  pattern := Items[Result];
  pattern.Index := rawPatternIdx;
  pattern.FgColor := APatternColor;
  pattern.BgColor := ABackColor;
end;

{@@ ----------------------------------------------------------------------------
  Adds a non-filled user-defined pattern to the chart's FillPattern list.

  @param AName            Name of the pattern. Must be unique.
  @param(ARawPatternIndex Index of the user-defined fill-pattern (@link(TsChartFillPatternStyle))
                          into the raw fill pattern list. This index is
                          returned by the RegisterRawFillPattern routine.)
  @param APatternColor    Color of the pattern.
-------------------------------------------------------------------------------}
function TsChartFillPatternList.AddPattern(AName: String;
  ARawPatternIndex: Integer; APatternColor: TsChartColor): Integer;
begin
  Result := AddPattern(AName, ARawPatternIndex, APatternColor, ChartColor(scWhite, 1.0));
  // Do not use scBlack here - will hide the entire pattern in xlsx.
end;

{@@ ----------------------------------------------------------------------------
  Adds a filled user-defined pattern to the chart's FillPattern list.

  @param AName            Name of the pattern. Must be unique.
  @param(ARawPatternIndex Index of the user-defined fill-pattern (@link(TsChartFillPatternStyle))
                          into the raw fill pattern list.
                          This index is returned by the RegisterRawFillPattern routine.)
  @param APatternColor    Color of the pattern.
  @param ABackColor       Background color of the pattern.
-------------------------------------------------------------------------------}
function TsChartFillPatternList.AddPattern(AName: String;
  ArawPatternIndex: Integer; APatternColor, ABackColor: TsChartColor): Integer;
var
  pattern: TsChartFillPattern;
begin
  if AName = '' then
    AName := DefaultPatternName;
  Result := NewPattern(AName);
  pattern := Items[Result];
  pattern.Index := ARawPatternIndex;
  pattern.FgColor := APatternColor;
  pattern.BgColor := ABackColor;
end;

{@@ ----------------------------------------------------------------------------
  Returns the name of the pattern which is used when no specific name is defined.

  @returns  Name of the pattern
-------------------------------------------------------------------------------}
function TsChartFillPatternList.DefaultPatternName: String;
begin
  Result := 'Pattern' + IntToStr(Count+1);
end;

{@@ ----------------------------------------------------------------------------
 Finds the pattern having the specified name.

 @param    AName   Name of the pattern to be found.
 @returns          Instance of the pattern having the specified name.
-------------------------------------------------------------------------------}
function TsChartFillPatternList.FindByName(AName: String): TsChartFillPattern;
var
  idx: Integer;
begin
  idx := IndexOfName(AName);
  if idx > -1 then
    Result := Items[idx]
  else
    Result := nil;
end;

function TsChartFillPatternList.GetItem(AIndex: Integer): TsChartFillPattern;
begin
  Result := TsChartFillPattern(inherited Items[AIndex]);
end;

{@@ ----------------------------------------------------------------------------
 Finds the pattern having the specified name.

 @param AName  Name of the pattern to be found.
 @returns      Index of the pattern having the specified name.
-------------------------------------------------------------------------------}
function TsChartFillPatternList.IndexOfName(AName: String): Integer;
var
  s: String;
begin
  for Result := 0 to Count-1 do
  begin
    s := Items[Result].Name;
    if SameText(s, AName) then
      exit;
  end;
  Result := -1;
end;

{@@ ----------------------------------------------------------------------------
 Finds the pattern having a specific name and background color

 @param AName   Name of the pattern to be found.
 @param AColor  Background color of the pattern to be found.
 @returns       Index of the pattern having the specified name and background color.
-------------------------------------------------------------------------------}
function TsChartFillPatternList.IndexOfNameAndBgColor(AName: string;
  AColor: TsChartColor): Integer;
var
  s: String;
  bkClr: TsChartColor;
begin
  for Result := 0 to Count-1 do
  begin
    s := Items[Result].Name;
    bkClr := Items[Result].BgColor;
    if SameText(s, AName) and (bkClr.Color = AColor.Color) and (bkClr.Transparency = AColor.Transparency) then
      exit;
  end;
  Result := -1;
end;

{ ------------------------------------------------------------------------------
  Creates a new pattern with the specified name and adds it to the list.

  @param    AName   Name of the pattern to be added.
  @returns  Index of the nex pattern in the pattern list.
-------------------------------------------------------------------------------}
function TsChartFillPatternList.NewPattern(AName: String): integer;
var
  pattern: TsChartFillPattern;
begin
  Result := IndexOfName(AName);
  if Result = -1 then
  begin
    pattern := TsChartFillPattern.Create;
    pattern.Name := AName;
    Result := inherited Add(pattern);
  end;
end;

procedure TsChartFillPatternList.SetItem(AIndex: Integer; AValue: TsChartFillPattern);
begin
  TsChartFillPattern(inherited Items[AIndex]).CopyFrom(AValue);
end;


{ TsChartImage }

{@@ ----------------------------------------------------------------------------
  Destructor of the @link(TsChartImage) class
-------------------------------------------------------------------------------}
destructor TsChartImage.Destroy;
begin
  Name := '';
  //Image.Free;  --- created by caller --> must be destroyed by caller!
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Copies the parameters from another image, ASource.

  @param  ASource   @link(TsChartImage) instance to be copied
-------------------------------------------------------------------------------}
procedure TsChartImage.CopyFrom(ASource: TsChartImage);
begin
  Name := ASource.Name;
  EmbeddedObjIndex := ASource.EmbeddedObjIndex;
//  Image := ASource.Image;
  Width := ASource.Width;
  Height := ASource.Height;
end;


{ TsChartImageList }

{@@ ----------------------------------------------------------------------------
  Creates a @link(TsChartImage) instance and adds it to the chart's Images list.
  The image itself must have been added to the workbook's embedded objects list before.
  The returned index then is used here as AEmbeddedObjIndex.

  @param AName             Name of the image, must be unique.
  @param AEmbeddedObjIndex Index of the image in the workbook's embedded objects list
  @param(AImgWidth         Image width, in millimeters, in which the image
                           will appear in the chart. Use -1 to request original width.)
  @param(AImgHeight        Image height, in millimeters, that the image will
                           have in the chart. Use -1 to request original height.)
-------------------------------------------------------------------------------}
function TsChartImageList.AddImage(AName: String; AEmbeddedObjIndex: Integer;
  AImgWidth: Single = -1.0; AImgHeight: Single = -1.0): Integer;
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
  Items[Result].EmbeddedObjIndex := AEmbeddedObjIndex;
  Items[Result].Width := AImgWidth;
  Items[Result].Height := AImgHeight;
end;

{@@ ----------------------------------------------------------------------------
  Finds an image in the image list by its name.

  @param    AName   Name of the image to be found.
  @returns  Instance of the image having the requested name.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Finds an image in the image list by its name.

  @param    AName   Name of the image to be found.
  @returns  Index of the image instance having the requested name.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Constructor of the @link(TsChartFill) class

  Initializes as solid black fill.
-------------------------------------------------------------------------------}
constructor TsChartFill.Create;
begin
  inherited Create;
  Style := cfsSolidFill;
  Color := ChartColor(scBlack);
  Gradient := -1;
  Pattern := -1;
  Image := -1;
end;

{@@ ----------------------------------------------------------------------------
  Creates an instance of @link(TsChartFill) representing a uniform fill by the specified color.

  @param AColor @link(TsChartColor) record defining the uniform fill of the associated shape.
-------------------------------------------------------------------------------}
constructor TsChartFill.CreateSolidFill(AColor: TsChartColor);
begin
  inherited Create;
  Style := cfsSolidFill;
  Color := AColor;
end;

{@@ ----------------------------------------------------------------------------
  Copies the fill paramters from another @link(TsChartFill) instance.

  @param  AFill  @link(TsChartFill) instance to be copied.
-------------------------------------------------------------------------------}
procedure TsChartFill.CopyFrom(AFill: TsChartFill);
begin
  if AFill <> nil then
  begin
    Style := AFill.Style;
    Color := AFill.Color;
    Gradient := AFill.Gradient;
    Pattern := AFill.Pattern;
    Image := AFill.Image;
  end;
end;

{@@ ----------------------------------------------------------------------------
  Selects gradient fill parameters.
  The gradient is defined by its index into the chart's Gradients list.

  @param(AGradientIndex  Index of the gradient in the chart's Gradients list
                         to be used for filling the associated shape.) }
procedure TsChartFill.SelectGradientFill(AGradientIndex: Integer);
begin
  Style := cfsGradient;
  Gradient := AGradientIndex;
end;

{@@ ----------------------------------------------------------------------------
  Selects parameters for filling the shape by a tiled image.

  The image is specified by its index in the chart's Images list.

  @param  AImageIndex  Index of the image parameters in the chart's Images list
-------------------------------------------------------------------------------}
procedure TsChartFill.SelectImageFill(AImageIndex: Integer);
begin
  inherited Create;
  Style := cfsImage;
  Image := AImageIndex;
end;

{@@ ----------------------------------------------------------------------------
  Clears the associated area/shape from any filling.
-------------------------------------------------------------------------------}
procedure TsChartFill.SelectNoFill;
begin
  Style := cfsNoFill;
end;

{@@ ----------------------------------------------------------------------------
  Selects a pattern fill for the associated shape or area.

  The pattern is determined by its index in the chart's FillPatterns list.

  Pattern and background colors are already included in the pattern parameters.
  When the pattern's background color is fully transparent the pattern is
  shown without background.

  @param  APatternIndex  Index of the pattern in the chart's FillPattern list.
-------------------------------------------------------------------------------}
procedure TsChartFill.SelectPatternFill(APatternIndex: Integer);
begin
  Pattern := APatternIndex;
  Style := cfsPattern;
end;

{@@ ----------------------------------------------------------------------------
  Selects a uniform color for filling the shape / area.

  @param AColor @link(TsChartColor) record defining the color and transparency to be used for filling the shape.
-------------------------------------------------------------------------------}
procedure TsChartFill.SelectSolidFill(AColor: TsChartColor);
begin
  Color := AColor;
  Style := cfsSolidFill;
end;


{ TsChartCellAddr }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChartCellAddr class.
  Initializes the row and column indices of the referenced cell with the
  UNASSIGNED_ROW_COL_INDEX value.
-------------------------------------------------------------------------------}
constructor TsChartCellAddr.Create(AChart: TsChart);
begin
  FChart := AChart;
  Sheet := '';
  Row := UNASSIGNED_ROW_COL_INDEX;
  Col := UNASSIGNED_ROW_COL_INDEX;
end;

{@@ ----------------------------------------------------------------------------
  Copies the cell address properties from another instance of the class.
-------------------------------------------------------------------------------}
procedure TsChartCellAddr.CopyFrom(ASource: TsChartCellAddr);
begin
  Sheet := ASource.Sheet;
  Row := ASource.Row;
  Col := ASource.Col;
end;

{@@ ----------------------------------------------------------------------------
  Returns the name of the worksheet which contains the referenced cell.
  The returned string is quoted if the name contains illegal characters.
-------------------------------------------------------------------------------}
function TsChartCellAddr.GetSheetName: String;
begin
  if Sheet <> '' then
    Result := Sheet
  else
    Result := FChart.Worksheet.Name;
  if SheetNameNeedsQuotes(Result) then
    Result := QuotedStr(Result);
end;

{@@ ----------------------------------------------------------------------------
  Returns TRUE when the referenced cell address has been changed to be different
  from the default.
-------------------------------------------------------------------------------}
function TsChartCellAddr.IsUsed: Boolean;
begin
  Result := (Row <> UNASSIGNED_ROW_COL_INDEX) and (Col <> UNASSIGNED_ROW_COL_INDEX);
end;


{ TsChartRange }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChartRange class

  Initializes the worksheet names with empty strings and the row and column
  indices with the value UNASSIGNED_ROW_COL_INDEX.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Copies the range parameters from another range
-------------------------------------------------------------------------------}
procedure TsChartRange.CopyFrom(ASource: TsChartRange);
begin
  Sheet1 := ASource.Sheet1;
  Sheet2 := ASource.Sheet2;
  Row1 := ASource.Row1;
  Col1 := ASource.Col1;
  Row2 := ASource.Row2;
  Col2 := ASource.Col2;
end;

{@@ ----------------------------------------------------------------------------
  Returns the name of the worksheet which contains the first cell of the range.
  The returned string is quoted if the name contains illegal characters.
-------------------------------------------------------------------------------}
function TsChartRange.GetSheet1Name: String;
begin
  if Sheet1 <> '' then
    Result := Sheet1
  else
  Result := FChart.Worksheet.Name;
  if SheetNameNeedsQuotes(Result) then
    Result := QuotedStr(Result);
end;

{@@ ----------------------------------------------------------------------------
  Returns the name of the worksheet which contains the last cell of the range
  The returned string is quoted if the name contains illegal characters.
-------------------------------------------------------------------------------}
function TsChartRange.GetSheet2Name: String;
begin
  if Sheet2 <> '' then
    Result := Sheet2
  else
  Result := FChart.Worksheet.Name;
  if SheetNameNeedsQuotes(Result) then
    Result := QuotedStr(Result);
end;

{@@ ----------------------------------------------------------------------------
  Returns TRUE when the cell range has not been assigned so far.
-------------------------------------------------------------------------------}
function TsChartRange.IsEmpty: Boolean;
begin
  Result :=
    (Row1 = UNASSIGNED_ROW_COL_INDEX) and (Col1 = UNASSIGNED_ROW_COL_INDEX) and
    (Row2 = UNASSIGNED_ROW_COL_INDEX) and (Col2 = UNASSIGNED_ROW_COL_INDEX);
end;

{@@ ----------------------------------------------------------------------------
  Calculates the number of cells contained in each worksheet of the range.
-------------------------------------------------------------------------------}
function TsChartRange.NumCellsPerSheet: Integer;
begin
  if IsEmpty then
    Result := 0
  else
    Result := (Col2 - Col1 + 1) * (Row2 - Row1 + 1);
end;


{ TsChartElement }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChartElement

  @param  AChart  Chart to which the element belongs.
-------------------------------------------------------------------------------}
constructor TsChartElement.Create(AChart: TsChart);
begin
  inherited Create;
  FChart := AChart;
  FVisible := true;
end;

{@@ ----------------------------------------------------------------------------
  Copies all element properties from another TsChartElement
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChartFillElement class

  Initializes the background as white solid fill and the border a black solid line
-------------------------------------------------------------------------------}
constructor TsChartFillElement.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FBackground := TsChartFill.Create;
  FBackground.SelectSolidFill(ChartColor(scWhite));
  FBorder := TsChartLine.Create;
  FBorder.SelectSolidLine(ChartColor(scBlack));
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsChartFillElement class
-------------------------------------------------------------------------------}
destructor TsChartFillElement.Destroy;
begin
  FBorder.Free;
  FBackground.Free;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Copies the element properties from another TsChartElement instance
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChartText class

  Initializes the default font of the text and hides the background and border.
-------------------------------------------------------------------------------}
constructor TsChartText.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FBorder.SelectNoLine;
  FBackground.SelectNoFill;

  FFont := TsFont.Create;
  FFont.Size := 10;
  FFont.Style := [];
  FFont.Color := scBlack;

  FVisible := true;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsChartText class
-------------------------------------------------------------------------------}
destructor TsChartText.Destroy;
begin
  FFont.Free;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Copies the text style properties from another TsChartText instance
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChartAxis class
-------------------------------------------------------------------------------}
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
  FAxisLine.SelectSolidLine(ChartColor(scBlack));

  FMajorTicks := [catOutside];
  FMinorTicks := [];

  FMajorGridLines := TsChartLine.Create;
  FMajorGridLines.SelectSolidLine(ChartColor($b3b3b3));

  FMinorGridLines := TsChartLine.Create;
  FMinorGridLines.SelectSolidLine(ChartColor($dddddd));

  FLogarithmic := false;
  FLogBase := 10.0;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsChartAxis class
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Copies the axis parameters from another axis.
-------------------------------------------------------------------------------}
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
  Returns the axis in the other direction when the chart is rotated, i.e. when
  the primary x axis is rotated it returns the primary y axis.
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

{@@ ----------------------------------------------------------------------------
  Defines the cell range containing the axis categories

  @param  ARow1  Top row of the range
  @param  ACol1  Left column of the range
  @param  ARow2  Bottom row of the range
  @param  ACol2  Right column of the range
-------------------------------------------------------------------------------}
procedure TsChartAxis.SetCategoryRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetCategoryRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines a 3d cell range containing the axis categories

  @param  ASheet1 Name of the worksheet containing the cell at ARow1/ACol1
  @param  ARow1   Top row of the range
  @param  ACol1   Left column of the range
  @param  ASheet2 Name of the worksheet containing the cell at ARow2/ACol2
  @param  ARow2   Bottom row of the range
  @param  ACol2   Right column of the range
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Constructor of the chart legend.
-------------------------------------------------------------------------------}
constructor TsChartLegend.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FFont := TsFont.Create;
  FFont.Size := 9;
  FVisible := true;

  // FPosition := legBottom;
  // That's the default of xlsx, but TAChart has difficulties with automatically
  // arranging several items per row. And .ods uses lpRight anyway...
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsChartLegend class
-------------------------------------------------------------------------------}
destructor TsChartLegend.Destroy;
begin
  FFont.Free;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Copies all parameters from another legend instance
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Copies the parameters from another instance
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChartDatapointStyleList

  @param AChart  Identifies the chart in which the list is used.
-------------------------------------------------------------------------------}
constructor TsChartDataPointStyleList.Create(AChart: TsChart);
begin
  inherited Create;
  FChart := AChart;
end;

{@@ ----------------------------------------------------------------------------
  Adds specific fill and line instances as style for the data point having the given index

  @param ADataPointIndex  Index of the data point for which the style is to be provided
  @param AFill            TsChartFill instance determining how the background of the  data point label is filled
  @param ALine            TsChartLine instance determining how the border of the data point label is drawn
  @param APieOffset       In case of a pieseries, percentage of the pie radius by which the pie is moved away from the pie center.
  @returns  Index of the style entry created for the data point.
  @note(You have the responsibility to destroy the AFill and ALine
                          instances after calling AddFillAndLine !)
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Adds a solid fill as style for the data point having the given index

  @param ADataPointIndex  Index of the data point for which the style is to be provided
  @param AColor           Background color of the data point label
  @param ALine            Line style for the data point label border. Must be destroyed by the caller.
  @param APieOffset       In case of a pieseries, percentage of the pie radius by which the pie is moved away from the pie center.
  @returns  Index of the style entry created for the data point.
-------------------------------------------------------------------------------}
function TsChartDataPointStyleList.AddSolidFill(ADataPointIndex: Integer;
  AColor: TsChartColor; ALine: TsChartLine = nil; APieOffset: Integer = 0): Integer;
var
  fill: TsChartFill;
begin
  fill := TsChartFill.Create;
  try
    fill.Style := cfsSolidFill;
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

{@@-----------------------------------------------------------------------------
  Returns the list index of the datapoint having the specified ADataPointIndex

  @param    ADataPointIndex  Index of the datapoint for which the style is to be found.
  @returns  Index of the style assigned to the datapoint at the provided ADataPointIndex
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChartErrorBars class

  Initialized the line parameters as solid black line with end cap, but hidden by default.
-------------------------------------------------------------------------------}
constructor TsChartErrorBars.Create(ASeries: TsChartSeries);
begin
  inherited Create(ASeries.Chart);
  FSeries := ASeries;
  FLine := TsChartLine.Create;
  FLine.SelectSolidLine(ChartColor(scBlack));
  FRange[0] := TsChartRange.Create(ASeries.Chart);
  FRange[1] := TsChartRange.Create(ASeries.Chart);
  FShow[0] := false;
  FShow[1] := false;
  FShowEndCap := true;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsChartErrorBars class
-------------------------------------------------------------------------------}
destructor TsChartErrorBars.Destroy;
begin
  FRange[1].Free;
  FRange[0].Free;
  FLine.Free;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Copies the error bar parameters from another instance
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Sets the cell range which contains the individual sizes of the positive error bars for each data point.

  @param  ARow1  Top row of the cell range
  @param  ACol1  Left column of the cell range
  @param  ARow2  Bottom row of the cell range
  @param  ACol2  Right column of the cell range
-------------------------------------------------------------------------------}
procedure TsChartErrorBars.SetErrorBarRangePos(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  InternalSetErrorBarRange(0, '', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Sets the 3D cell range which contains the individual sizes of the positive error bars for each data point.

  @param  ASheet1 Name of the worksheet containing the first (top/left) cell of the range
  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ASheet2 Name of the worksheet containing the last (bottom/right) cell of the range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
procedure TsChartErrorBars.SetErrorBarRangePos(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
begin
  InternalSetErrorBarRange(0, ASheet1, ARow1, ACol1, ASheet2, ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Sets the cell range which contains the individual sizes of the negative error bars for each data point.

  @param  ARow1  Top row of the cell range
  @param  ACol1  Left column of the cell range
  @param  ARow2  Bottom row of the cell range
  @param  ACol2  Right column of the cell range
-------------------------------------------------------------------------------}
procedure TsChartErrorBars.SetErrorBarRangeNeg(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  InternalSetErrorBarRange(1, '', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Sets the 3D cell range which contains the individual sizes of the negative error bars for each data point.

  @param  ASheet1 Name of the worksheet containing the first (top/left) cell of the range
  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ASheet2 Name of the worksheet containing the last (bottom/right) cell of the range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChartSeries class
-------------------------------------------------------------------------------}
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
  FFill.SelectSolidFill(GetDefaultSeriesColor);
  FFill.Gradient := -1;
  FFill.Pattern := -1;

  FLine := TsChartLine.Create;
  FLine.SelectSolidLine(GetDefaultSeriesColor);

  FDataPointStyles := TsChartDataPointStyleList.Create(AChart);

  FLabelFont := TsFont.Create;
  FLabelFont.Size := 9;

  FLabelBorder := TsChartLine.Create;
  FLabelBorder.SelectNoLine;

  FLabelBackground := TsChartFill.Create;
  FLabelBackground.Color := ChartColor(scWhite);
  FLabelBackground.SelectNoFill;

  FLabelSeparator := ' ';
  FLabelFormatPercent := '0%';

  FTrendline := TsChartTrendline.Create;

  FXErrorBars := TsChartErrorBars.Create(Self);
  FYErrorBars := TsChartErrorBars.Create(Self);
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsChartSeries class
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Returns the number of data points contained in the series.
  This is the length of the YRange column or row
-------------------------------------------------------------------------------}
function TsChartSeries.GetCount: Integer;
begin
  Result := GetYCount;
end;

{@@ ----------------------------------------------------------------------------
  Returns the default color of the series.

  The default colors are stored in a global array. The color is picked based on
  the series index ("Order") wrapped around the end of the array.
-------------------------------------------------------------------------------}
function TsChartSeries.GetDefaultSeriesColor: TsChartColor;
begin
  Result := ChartColor(DEFAULT_SERIES_COLORS[FOrder mod Length(DEFAULT_SERIES_COLORS)]);
end;

{@@ ----------------------------------------------------------------------------
  Returns the instance of the TsChartAxis class which represents the x axis of
  the series. Depending of the TsChartAxisLink parameter "XAxis" this is either
  the chart's XAxis or X2Axis.
-------------------------------------------------------------------------------}
function TsChartSeries.GetXAxis: TsChartAxis;
begin
  if FXAxis = calPrimary then
    Result := Chart.XAxis
  else
    Result := Chart.X2Axis;
end;

{@@ ----------------------------------------------------------------------------
  Returns the instance of the TsChartAxis class which represents the y axis of
  the series. Depending of the TsChartAxisLink parameter "YAxis" this is either
  the chart's YAxis or Y2Axis.
-------------------------------------------------------------------------------}
function TsChartSeries.GetYAxis: TsChartAxis;
begin
  if FYAxis = calPrimary then
    Result := Chart.YAxis
  else
    Result := Chart.Y2Axis;
end;

{@@ ----------------------------------------------------------------------------
  Returns the number of x values contained in the series, based on the assigned
  XRange.

  Depending on whether the XRange is in a row or in a column the count of
  x values is determined from the corresponding dimension of the range.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Returns the number of y values contained in the series, based on the assigned
  YRange.

  Depending on whether the YRange is in a row or in a column the count of
  y values is determined from the corresponding dimension of the range.
-------------------------------------------------------------------------------}
function TsChartSeries.GetYCount: Integer;
begin
  if YValuesInCol then
    Result := FYRange.Row2 - FYRange.Row1 + 1
  else
    Result := FYRange.Col2 - FYRange.Col1 + 1;
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the series has data point labels, i.e. when the LabelRange
  is not empty.
-------------------------------------------------------------------------------}
function TsChartSeries.HasLabels: Boolean;
begin
  Result := not ((FLabelRange.Row1 = FLabelRange.Row2) and (FLabelRange.Col1 = FLabelRange.Col2));
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the series has x values, i.e. when the XRange is not empty.
-------------------------------------------------------------------------------}
function TsChartSeries.HasXValues: Boolean;
begin
  Result := not ((FXRange.Row1 = FXRange.Row2) and (FXRange.Col1 = FXRange.Col2));
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the series has y values, i.e. when the YRange is not empty.
-------------------------------------------------------------------------------}
function TsChartSeries.HasYValues: Boolean;
begin
  Result := not ((FYRange.Row1 = FYRange.Row2) and (FYRange.Col1 = FYRange.Col2));
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the data point label cells are arranged in a column.
-------------------------------------------------------------------------------}
function TsChartSeries.LabelsInCol: Boolean;
begin
  Result := (FLabelRange.Col1 = FLabelRange.Col2) and (FLabelRange.Row1 <> FLabelRange.Row2);
end;

{@@ ----------------------------------------------------------------------------
  Defines the cell which contains the title of the series (the text which appears
  in the legend.

  @param  ARow  Row index of the cell
  @param  ACol  Column index of the cell
-------------------------------------------------------------------------------}
procedure TsChartSeries.SetTitleAddr(ARow, ACol: Cardinal);
begin
  SetTitleAddr('', ARow, ACol);
end;

{@@ ----------------------------------------------------------------------------
  Defines the cell which contains the title of the series (the text which appears
  in the legend. In this overload of the procedure the cell can be in a different
  worksheet.

  @param  ASheet  Name of the worksheet containing the cell
  @param  ARow   Row index of the cell
  @param  ACol   Column index of the cell
-------------------------------------------------------------------------------}
procedure TsChartSeries.SetTitleAddr(ASheet: String; ARow, ACol: Cardinal);
begin
  FTitleAddr.Sheet := ASheet;
  FTitleAddr.Row := ARow;
  FTitleAddr.Col := ACol;
end;

{@@ ----------------------------------------------------------------------------
  Defines the cell range which contains the fill colors of the individual data points
  Only a single-column or single-row range is allowed.

  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
procedure TsChartSeries.SetFillColorRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetFillColorRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines the cell range which contains the fill colors of the individual data points.
  Only a single-column or single-row range is allowed.

  @param  ASheet1 Name of the worksheet containing the first (top/left) cell of the range
  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ASheet2 Name of the worksheet containing the last (bottom/right) cell of the range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Defines the cell range which contains the colors of the data point labels
  Only a single-column or single-row range is allowed.

  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
procedure TsChartSeries.SetLabelRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetLabelRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines the 3D cell range which contains the colors of the data point labels.
  Only a single-column or single-row range is allowed.

  @param  ASheet1 Name of the worksheet containing the first (top/left) cell of the range
  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ASheet2 Name of the worksheet containing the last (bottom/right) cell of the range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Defines the cell range which contains the colors of  data-point-related lines.
  Only a single-column or single-row range is allowed.

  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
procedure TsChartSeries.SetLineColorRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetLineColorRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines the 3D cell range which contains the colors of the data-point-related lines.
  Only a single-column or single-row range is allowed.

  @param  ASheet1 Name of the worksheet containing the first (top/left) cell of the range
  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ASheet2 Name of the worksheet containing the last (bottom/right) cell of the range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Defines the cell range which contains the x values of the data points.
  Only a single-column or single-row range is allowed.

  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
procedure TsChartSeries.SetXRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetXRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines the 3D cell range which contains the x values of the data points.
  Only a single-column or single-row range is allowed.

  @param  ASheet1 Name of the worksheet containing the first (top/left) cell of the range
  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ASheet2 Name of the worksheet containing the last (bottom/right) cell of the range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Defines the cell range which contains the y values of the data points.
  Only a single-column or single-row range is allowed.

  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
procedure TsChartSeries.SetYRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetYRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines the 3D cell range which contains the y values of the data points.
  Only a single-column or single-row range is allowed.

  @param  ASheet1 Name of the worksheet containing the first (top/left) cell of the range
  @param  ARow1   Top row of the cell range
  @param  ACol1   Left column of the cell range
  @param  ASheet2 Name of the worksheet containing the last (bottom/right) cell of the range
  @param  ARow2   Bottom row of the cell range
  @param  ACol2   Right column of the cell range
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Returns true when x values are arranged in a column.
-------------------------------------------------------------------------------}
function TsChartSeries.XValuesInCol: Boolean;
begin
  Result := (FXRange.Col1 = FXRange.Col2) and (FXRange.Row1 <> FXRange.Row2);
end;

{@@ ----------------------------------------------------------------------------
  Returns true when y values are arranged in a column.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Constructor of the TsAreaSeries class

  @param  AChart   Chart into which the series is inserted.
-------------------------------------------------------------------------------}
constructor TsAreaSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctArea;
  FSupportsTrendline := true;
  FGroupIndex := 0;
end;


{ TsBarSeries }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsBarSeries class

  @param  AChart   Chart into which the series is inserted.
-------------------------------------------------------------------------------}
constructor TsBarSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctBar;
  FSupportsTrendline := true;
  FGroupIndex := 0;
end;


{ TsBubbleSeries }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsBubbleSeries class

  @param  AChart   Chart into which the series is inserted.
-------------------------------------------------------------------------------}
constructor TsBubbleSeries.Create(AChart: TsChart);
begin
  inherited;
  FBubbleRange := TsChartRange.Create(AChart);
  FBubbleScale := 1.0;
  FBubbleSizeMode := bsmArea;
  FChartType := ctBubble;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsBubbleSeries class
-------------------------------------------------------------------------------}
destructor TsBubbleSeries.Destroy;
begin
  FBubbleRange.Free;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Defines the cell range which contains the bubble sizes

  @param  ARow1  Top row of the cell range
  @param  ACol1  Left column of the cell range
  @param  ARow2  Bottom row of the cell range
  @param  @Col2  Right column of the cell range
-------------------------------------------------------------------------------}
procedure TsBubbleSeries.SetBubbleRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  // Empty sheet name will be replaced by the name of the sheet containing the chart.
  SetBubbleRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines a 3D cell range which contains the bubble sizes

  @param  ASheet1  Name of the worksheet containing the top/left cell of the range (first cell)
  @param  ARow1    Top row of the cell range
  @param  ACol1    Left column of the cell range
  @param  ASheet2  Name of the worksheet containing the bottom/right cell of the range (last cell)
  @param  ARow2    Bottom row of the cell range
  @param  ACol2    Right column of the cell range
-------------------------------------------------------------------------------}
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

{@@-----------------------------------------------------------------------------
  Constructor of the TsCustomLineSeries class. It is the ancestor of
  TsLineSeries, TsScatterSeries and TsRadarSeries.

  Initializes with solid black lines and symbols turned off.
-------------------------------------------------------------------------------}
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
  FSymbolBorder.SelectSolidLine(ChartColor(scBlack));

  FSymbolFill := TsChartFill.Create;
  FSymbolFill.SelectNoFill;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsCustomLineSeries
-------------------------------------------------------------------------------}
destructor TsCustomLineSeries.Destroy;
begin
  FSymbolBorder.Free;
  FSymbolFill.Free;
  inherited;
end;


{ TsLineSeries }
{@@ ----------------------------------------------------------------------------
  Constructor of TsLineSeries
-------------------------------------------------------------------------------}
constructor TsLineSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FGroupIndex := 0;
end;

{ TsPieSeries }
{@@ ----------------------------------------------------------------------------
  Constructor of the TsPieSeries
-------------------------------------------------------------------------------}
constructor TsPieSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctPie;
  FStartAngle := 90;
  FLine.Color := ChartColor(scBlack);
end;

{@@ ----------------------------------------------------------------------------
  Returns the chart type: ctPie, or ctRing, depending on the value of hte
    InnerRadiusPercent property.
-------------------------------------------------------------------------------}
function TsPieSeries.GetChartType: TsChartType;
begin
  if FInnerRadiusPercent > 0 then
    Result := ctRing
  else
    Result := ctPie;
end;

{@@ ----------------------------------------------------------------------------
  Returns the fraction of the pie radius a slice at the given index is moved
  outwards from the center.

  @Param  ASliceIndex  Index of the data point considered here.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Constructor of the TsRadarSeries
-------------------------------------------------------------------------------}
constructor TsRadarSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctRadar;
  FFill.Style := cfsNoFill;  // to make the series default to ctRadar rather than ctFilledRadar
end;


{ TsFilledRadarSeries }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsFilledRadarSeries, a specialized, filled TsRadarSeries
-------------------------------------------------------------------------------}
constructor TsFilledRadarSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctFilledRadar;
  Fill.Style := cfsSolidFill;
end;


{ TsTrendlineEquation }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsTrendLineEquation class

  Initializes all parameters with reasonable values.
-------------------------------------------------------------------------------}
constructor TsTrendlineEquation.Create;
begin
  inherited Create;
  Font := TsFont.Create;
  Font.Size := 9;
  Border := TsChartLine.Create;
  Border.SelectNoLine;
  Fill := TsChartFill.Create;
  Fill.SelectSolidFill(ChartColor(scWhite));
  XName := 'x';
  YName := 'f(x)';
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsTrendLineEquation class
-------------------------------------------------------------------------------}
destructor TsTrendlineEquation.Destroy;
begin
  Fill.Free;
  Border.Free;
  Font.Free;
  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the trendline equation is displayed without border.
-------------------------------------------------------------------------------}
function TsTrendlineEquation.IsDefaultBorder: Boolean;
begin
  Result := (Border.Style = clsNoLine);
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the trendline equation is displayed without background fill.
--------------------------------------------------------------------------------}
function TsTrendlineEquation.IsDefaultFill: Boolean;
begin
  Result := (Fill.Style = cfsNoFill);
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the trendline equation is written with the default font
  (size 9, black color)
-------------------------------------------------------------------------------}
function TsTrendlineEquation.IsDefaultFont: Boolean;
begin
  Result := (Font.FontName = '') and (Font.Size = 9) and (Font.Style = []) and
            (Font.Color = scBlack);
end;

{@@ ----------------------------------------------------------------------------
  Returns true when numbers in the trendline expression are formatted with the
  "general" number format.
-------------------------------------------------------------------------------}
function TsTrendlineEquation.IsDefaultNumberFormat: Boolean;
begin
  Result := NumberFormat = '';
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the trendline equation is painted in the top/left corner of
  the entire chart area.
-------------------------------------------------------------------------------}
function TsTrendlineEquation.IsDefaultPosition: Boolean;
begin
  Result := (Left = 0) and (Top = 0);
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the independent variable in the trendline expression is
  named "x"
-------------------------------------------------------------------------------}
function TsTrendlineEquation.IsDefaultXName: Boolean;
begin
  Result := XName = 'x';
end;

{@@ ----------------------------------------------------------------------------
  Returns true when the dependent variable in the trendline expression is
  names "f(x)"
-------------------------------------------------------------------------------}
function TsTrendlineEquation.IsDefaultYName: Boolean;
begin
  Result := YName = 'f(x)';
end;


{ TsChartTrendline }

{@@ ----------------------------------------------------------------------------
  Constructor for the TsChartTrendLine class

  Initializes the trend line as a solid black line
-------------------------------------------------------------------------------}
constructor TsChartTrendline.Create;
begin
  inherited Create;

  Line := TsChartLine.Create;
  Line.SelectSolidLine(ChartColor(scBlack));

  Equation := TsTrendlineEquation.Create;
end;

{@@ ----------------------------------------------------------------------------
  Destructor of the TsChartTrendLine class
-------------------------------------------------------------------------------}
destructor TsChartTrendline.Destroy;
begin
  Equation.Free;
  Line.Free;
  inherited;
end;


{ TsCustomScatterSeries }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsCustomScatterSeries class
-------------------------------------------------------------------------------}
constructor TsCustomScatterSeries.Create(AChart: TsChart);
begin
  inherited Create(AChart);
  FChartType := ctScatter;
  FSupportsTrendline := true;
end;


{ TsStockSeries }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsStockSeries class
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Destructor of the TsStockSeries class
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Defines the cell range for the "open" values (starting stock prices for each
  date unit)

  @param  ARow1  Index of the top row of the range
  @param  ACol1  Index of the left column of the range
  @param  ARow2  Index of the bottom row of the range
  @param  ACol2  Index of the right column of the range
-------------------------------------------------------------------------------}
procedure TsStockSeries.SetOpenRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetOpenRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines a 3D cell range for the "open" values (starting stock prices for each
  date unit)

  @param  ASheet1 Name of the worksheet containing the first (= top/left) cell of the range
  @param  ARow1   Index of the top row of the range
  @param  ACol1   Index of the left column of the range
  @param  ASheet2 Name of the worksheet containing the last (=bottom/right) cell of the range
  @param  ARow2   Index of the bottom row of the range
  @param  ACol2   Index of the right column of the range
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Defines the cell range for the "high" values (highest stock price in each
  date unit)

  @param  ARow1  Index of the top row of the range
  @param  ACol1  Index of the left column of the range
  @param  ARow2  Index of the bottom row of the range
  @param  ACol2  Index of the right column of the range
-------------------------------------------------------------------------------}
procedure TsStockSeries.SetHighRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetHighRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines a 3D cell range for the "high" values (highest stock price in each
  date unit)

  @param  ASheet1 Name of the worksheet containing the first (= top/left) cell of the range
  @param  ARow1   Index of the top row of the range
  @param  ACol1   Index of the left column of the range
  @param  ASheet2 Name of the worksheet containing the last (=bottom/right) cell of the range
  @param  ARow2   Index of the bottom row of the range
  @param  ACol2   Index of the right column of the range
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Defines a cell range for the "low" values (lowest stock price in each
  date unit)

  @param  ARow1   Index of the top row of the range
  @param  ACol1   Index of the left column of the range
  @param  ARow2   Index of the bottom row of the range
  @param  ACol2   Index of the right column of the range
-------------------------------------------------------------------------------}
procedure TsStockSeries.SetLowRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetLowRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines a 3D cell range for the "low" values (lowest stock price in each
  date unit)

  @param  ASheet1 Name of the worksheet containing the first (= top/left) cell of the range
  @param  ARow1   Index of the top row of the range
  @param  ACol1   Index of the left column of the range
  @param  ASheet2 Name of the worksheet containing the last (=bottom/right) cell of the range
  @param  ARow2   Index of the bottom row of the range
  @param  ACol2   Index of the right column of the range
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Defines a cell range for the "close" values (closing stock price in each
  date unit)

  @param  ARow1   Index of the top row of the range
  @param  ACol1   Index of the left column of the range
  @param  ARow2   Index of the bottom row of the range
  @param  ACol2   Index of the right column of the range
-------------------------------------------------------------------------------}
procedure TsStockSeries.SetCloseRange(ARow1, ACol1, ARow2, ACol2: Cardinal);
begin
  SetCloseRange('', ARow1, ACol1, '', ARow2, ACol2);
end;

{@@ ----------------------------------------------------------------------------
  Defines a 3D cell range for the "close" values (closing stock price in each
  date unit)

  @param  ASheet1 Name of the worksheet containing the first (= top/left) cell of the range
  @param  ARow1   Index of the top row of the range
  @param  ACol1   Index of the left column of the range
  @param  ASheet2 Name of the worksheet containing the last (=bottom/right) cell of the range
  @param  ARow2   Index of the bottom row of the range
  @param  ACol2   Index of the right column of the range
-------------------------------------------------------------------------------}
procedure TsStockSeries.SetCloseRange(ASheet1: String; ARow1, ACol1: Cardinal;
  ASheet2: String; ARow2, ACol2: Cardinal);
 begin
   SetYRange(ASheet1, ARow1, ACol1, ASheet2, ARow2, ACol2);
 end;


{ TsChart }

{@@ ----------------------------------------------------------------------------
  Constructor of the TsChart class
-------------------------------------------------------------------------------}
constructor TsChart.Create;
begin
  inherited Create(nil);

  CreateRawFillPatterns;
  CreateRawLinePatterns;

  FGradients := TsChartGradientList.Create;
  FFillPatterns := TsChartFillPatternList.Create;
  FImages := TsChartImageList.Create;

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

{@@-----------------------------------------------------------------------------
  Destructor of the TsChart class
-------------------------------------------------------------------------------}
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
  FFillPatterns.Free;
  FGradients.Free;

  DestroyRawLinePatterns;
  DestroyRawFillPatterns;

  inherited;
end;

{@@ ----------------------------------------------------------------------------
  Adds the specified series to the internal series list and returns the index
  in this list which has been assigned to the series.

  @param    ASeries  Series to be added.
  @returns  Index into the chart's Series list.
-------------------------------------------------------------------------------}
function TsChart.AddSeries(ASeries: TsChartSeries): Integer;
begin
  Result := FSeriesList.IndexOf(ASeries);
  if Result = -1 then
    Result := FSeriesList.Add(ASeries);
end;

{@@ ----------------------------------------------------------------------------
  Removes the series with the specified index from the internal series list.
  The series itself is not destroyed.

  @param  AIndex   Indes into the chart's Series list of the series to be deleted.
-------------------------------------------------------------------------------}
procedure TsChart.DeleteSeries(AIndex: Integer);
begin
  if (AIndex >= 0) and (AIndex < FSeriesList.Count) then
    FSeriesList.Delete(AIndex);
end;

{@@ ----------------------------------------------------------------------------
  Returns the cell range containing the labels for the x axis categories.
-------------------------------------------------------------------------------}
function TsChart.GetCategoryLabelRange: TsChartRange;
begin
  Result := XAxis.CategoryRange;
end;

{@ -----------------------------------------------------------------------------
  Returns the type of the chart. It is derived from the type assigned to the
  first series.
-------------------------------------------------------------------------------}
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

{@@ ----------------------------------------------------------------------------
  Returns true if the chart is a scatter chart.
-------------------------------------------------------------------------------}
function TsChart.IsScatterChart: Boolean;
begin
  Result := GetChartType = ctScatter;
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

