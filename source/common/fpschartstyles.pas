unit fpsChartStyles;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

interface

uses
  Classes, SysUtils, fpsTypes, fpsChart;

type
  // Identifiers for what is stored in a style
  TsChartStyleElement = (
    ceBackground, ceWall, ceFloor, ceLegend, ceTitle, ceSubTitle, cePlotArea,
    ceXAxis, ceXAxisCaption, ceXAxisMajorGrid, ceXAxisMinorGrid,
    ceX2Axis, ceX2AxisCaption, ceX2AxisMajorGrid, ceX2AxisMinorGrid,
    ceYAxis, ceYAxisCaption, ceYAxisMajorGrid, ceYAxisMinorGrid,
    ceY2Axis, ceY2AxisCaption, ceY2AxisMajorGrid, ceY2AxisMinorGrid,
    ceSeries, ceSeriesBorder, ceSeriesFill, ceSeriesLine
  );

type
  TsChartLineRec = record
    Style: Integer;        // index into chart's LineStyle list or predefined clsSolid/clsNoLine
    Width: Double;         // mm
    Color: TsColor;        // in hex: $00bbggrr, r=red, g=green, b=blue
    Transparency: Double;  // in percent
    procedure FromChart(AChart: TsChart; AElement: TsChartStyleElement; AIndex: Integer);
    function GetChartLine(AChart: TsChart; AElement: TsChartStyleElement; AIndex: Integer): TsChartLine;
    procedure ToChart(AChart: TsChart; AElement: TsChartStyleElement; AIndex: Integer);
    class operator = (A, B: TsChartLineRec): Boolean;
  end;

  TsChartFillRec = record
    Style: TsFillStyle;
    FgColor: TsColor;
    BgColor: TsColor;
    procedure FromChart(AChart: TsChart; AElement: TsChartStyleElement; AIndex: Integer);
    function GetChartFill(AChart: TsChart; AElement: TsChartStyleElement; AIndex: Integer): TsChartFill;
    procedure ToChart(AChart: TsChart; AElement: TsChartStyleElement; AIndex: Integer);
    class operator = (A, B: TsChartFillRec): Boolean;
  end;

  TsChartFontRec = record
    FontName: String;
    Size: Double;
    Style: TsFontStyles;
    Color: TsColor;
    Position: TsFontPosition;
    procedure FromChart(AChart: TsChart; AElement: TsChartStyleElement);
    function GetChartFont(AChart: TsChart; AElement: TsChartStyleElement): TsFont;
    procedure ToChart(AChart: TsChart; AElement: TsChartStyleElement);
    procedure ToFont(AFont: TsFont);
    class operator = (A, B: TsChartFontRec): Boolean;
  end;

  TsChartAxisRec = record
    AutomaticMax: Boolean;
    AutomaticMin: Boolean;
    AutomaticMajorInterval: Boolean;
    AutomaticMinorInterval: Boolean;
    AxisLine: TsChartLineRec;
    MajorGridLines: TsChartLineRec;
    MinorGridLines: TsChartLineRec;
    MajorTickLines: TsChartLineRec;
    MinorTickLines: TsChartLineRec;
    Inverted: Boolean;
//    CaptionFont: TsChartFontRec;
    LabelFont: TsChartFontRec;
    LabelFormat: String;
    LabelRotation: Integer;
    Logarithmic: Boolean;
    MajorInterval: Double;
    MinorInterval: Double;
    Position: TsChartAxisPosition;
//    ShowCaption: Boolean;
    ShowLabels: Boolean;
    Visible: Boolean;
    procedure FromChart(AChart: TsChart; AElement: TsChartStyleElement);
    function GetChartAxis(AChart: TsChart; AElement: TsChartStyleElement): TsChartAxis;
    function GetChartLine(AChart: TsChart; AElement: TsChartStyleElement): TsChartLine;
    procedure ToChart(AChart: TsChart; AElement: TsChartStyleElement);
    class operator = (A, B: TsChartAxisRec): Boolean;
  end;

  TsChartTextRec = record
    Font: TsChartFontRec;
    Rotation: Integer;
    Visible: Boolean;
    procedure FromChart(AChart: TsChart; AElement: TsChartStyleElement);
    procedure ToChart(AChart: TsChart; AElement: TsChartStyleElement);
    class operator = (A, B: TsChartTextRec): Boolean;
  end;

  TsChartLegendRec = record
    Font: TsChartFontRec;
    Border: TsChartLineRec;
    Fill: TsChartFillRec;
    Visible: Boolean;
    procedure FromChart(AChart: TsChart; AElement: TsChartStyleElement);
    procedure ToChart(AChart: TsChart; AElement: TsChartStyleElement);
    class operator = (A, B: TsChartLegendRec): Boolean;
  end;

  TsChartPlotAreaRec = record
    FChart: TsChart;
    procedure FromChart(AChart: TsChart; AElement: TsChartStyleElement);
    procedure ToChart(AChart: TsChart; AElement: TsChartStyleElement);
    class operator = (A, B: TsChartPlotAreaRec): Boolean;
  end;

  TsChartSeriesRec = record
    Line: TsChartLineRec;
    Fill: TsChartFillRec;
    Border: TsChartFillRec;
    procedure FromChart(AChart: TsChart; AElement: TsChartStyleElement; AIndex: Integer);
    procedure ToChart(AChart: TsChart; AElement: TsChartStyleElement; AIndex: Integer);
    class operator = (A, B: TsChartSeriesRec): Boolean;
  end;

  {----------------------------------------------------------------------------}

  TsChartStyle = class
  private
    FName: String;
    FElement: TsChartStyleElement;
  public
    constructor Create(AElement: TsChartStyleElement); virtual;
    procedure ApplyToChart(AChart: TsChart; AIndex: Integer); virtual; abstract;
    procedure ExtractFromChart(AChart: TsChart; AIndex: Integer); virtual; abstract;
    property Element: TsChartStyleElement read FElement;
    property Name: String read FName;
  end;

  TsChartStyleClass = class of TsChartStyle;

  TsChartStyle_Background = class(TsChartStyle)
  private
    FBackground: TsChartFillRec;
    FBorder: TsChartLineRec;
  public
    procedure ApplyToChart(AChart: TsChart; AIndex: Integer); override;
    procedure ExtractFromChart(AChart: TsChart; AIndex: Integer); override;
    property Background: TsChartFillRec read FBackground;
    property Border: TsChartLineRec read FBorder;
  end;

  TsChartStyle_Line = class(TsChartStyle)
  private
    FLine: TsChartLineRec;
  public
    procedure ApplyToChart(AChart: TsChart; AIndex: Integer); override;
    procedure ExtractFromChart(AChart: TsChart; AIndex: Integer); override;
    property Line: TsChartLineRec read FLine;
  end;

  TsChartStyle_Axis = class(TsChartStyle)
  private
    FAxis: TsChartAxisRec;
  public
    procedure ApplyToChart(AChart: TsChart; AIndex: Integer); override;
    procedure ExtractFromChart(AChart: TsChart; AIndex: Integer); override;
    property Axis: TsChartAxisRec read FAxis write FAxis;
  end;

  TsChartStyle_Caption = class(TsChartStyle)
  private
    FCaption: TsChartTextRec;
  public
    procedure ApplyToChart(AChart: TsChart; AIndex: Integer); override;
    procedure ExtractFromChart(AChart: TsChart; AIndex: Integer); override;
    property Caption: TsChartTextRec read FCaption write FCaption;
  end;

  TsChartStyle_Legend = class(TsChartStyle)
  private
    FLegend: TsChartLegendRec;
  public
    procedure ApplyToChart(AChart: TsChart; AIndex: Integer); override;
    procedure ExtractFromChart(AChart: TsChart; AIndex: Integer); override;
    property Legend: TsChartLegendRec read FLegend write FLegend;
  end;

  TsChartStyle_PlotArea = class(TsChartStyle)
  private
    FPlotArea: TsChartPlotAreaRec;
  public
    procedure ApplyToChart(AChart: TsChart; AIndex: Integer); override;
    procedure ExtractFromChart(AChart: TsChart; AIndex: Integer); override;
  end;

  TsChartStyle_Series = class(TsChartStyle)
  private
    // for all series types
    FLine: TsChartLineRec;
    FFill: TsChartFillRec;
    FBorder: TsChartLineRec;
    // for TsLineSeries
    FSymbol: TsChartSeriesSymbol;
    FSymbolHeight: Double;  // in mm
    FSymbolWidth: Double;   // in mm
    FShowSymbols: Boolean;
  public
    procedure ApplyToChart(AChart: TsChart; AIndex: Integer); override;
    procedure ExtractFromChart(AChart: TsChart; AIndex: Integer); override;
    property Line: TsChartLineRec read FLine write FLine;        // lineseries lines
    property Fill: TsChartFillRec read FFill write FFill;        // symbol fill, bar fill, area border
    property Border: TsChartLineRec read FBorder write FBorder;  // symbol border, bar border, area border
    // for line series only
    property ShowSymbols: Boolean read FShowSymbols write FShowSymbols;
    property Symbol: TsChartSeriesSymbol read FSymbol write FSymbol;
    property SymbolHeight: Double read FSymbolHeight write FSymbolHeight;
    property SymbolWidth: Double read FSymbolWidth write FSymbolWidth;
  end;

  { ---------------------------------------------------------------------------}

  TsChartStyleList = class(TFPList)
  protected

  public
    destructor Destroy; override;
    function AddChartStyle(AName: String; AChart: TsChart;
      AStyleClass: TsChartStyleClass; AElement: TsChartStyleElement;
      AIndex: Integer = -1): Integer;
    procedure Clear;
    procedure Delete(AIndex: Integer);
    function FindStyleIndexByName(const AName: String): Integer;
    {
    function FindChartStyle(AChart: TsChart; AStyleClass: TsChartStyleClass;
      AElement: TsChartStyleElement; AIndex: Integer = -1): Integer;
    }
  end;

implementation

{==============================================================================}
{                             Style records                                    }
{     Copies of the chart properties to simplify handling in the style.        }
{==============================================================================}

{ TsFontRec }
procedure TsChartFontRec.FromChart(AChart: TsChart; AElement: TsChartStyleElement);
var
  fnt: TsFont;
begin
  fnt := GetChartFont(AChart, AElement);
  FontName := fnt.FontName;
  Size := fnt.Size;
  Style := fnt.Style;
  Color := fnt.Color;
  Position := fnt.Position;
end;

function TsChartFontRec.GetChartFont(AChart: TsChart; AElement: TsChartStyleElement): TsFont;
begin
  case AElement of
    ceXAxis: Result := AChart.XAxis.LabelFont;
    ceYAxis: Result := AChart.YAxis.LabelFont;
    ceX2Axis: Result := AChart.X2Axis.LabelFont;
    ceY2Axis: Result := AChart.Y2Axis.LabelFont;
    ceXAxisCaption: Result := AChart.XAxis.CaptionFont;
    ceYAxisCaption: Result := AChart.YAxis.CaptionFont;
    ceX2AxisCaption: Result := AChart.X2Axis.CaptionFont;
    ceY2AxisCaption: Result := AChart.Y2Axis.CaptionFont;
    ceTitle: Result := AChart.Title.Font;
    ceSubtitle: Result := AChart.Subtitle.Font;
    ceLegend: Result := AChart.Legend.Font;
  else
    raise Exception.Create('[TsChartFontRec] Font not supported.');
  end;
end;

procedure TsChartFontRec.ToChart(AChart: TsChart; AElement: TsChartStyleElement);
var
  fnt: TsFont;
begin
  fnt := GetChartFont(AChart, AElement);
  fnt.FontName := FontName;
  fnt.Size := Size;
  fnt.Style := Style;
  fnt.Color := Color;
  fnt.Position := Position;
end;

procedure TsChartFontRec.ToFont(AFont: TsFont);
begin
  AFont.FontName := FontName;
  AFont.Size := Size;
  AFont.Style := Style;
  AFont.Color := Color;
  AFont.Position := Position;
end;

class operator TsChartFontRec.= (A, B: TsChartFontRec): Boolean;
begin
  Result := (A.FontName = B.FontName) and (A.Size = B.Size) and
    (A.Style = B.Style) and (A.Color = B.Color) and
    (A.Position = B.Position);
end;

{ TsChartLineRec }
procedure TsChartLineRec.FromChart(AChart: TsChart; AElement: TsChartStyleElement;
  AIndex: Integer);
var
  L: TsChartLine;
begin
  L := GetChartLine(AChart, AElement, AIndex);
  Style := L.Style;
  Width := L.Width;
  Color := L.Color;
  Transparency := L.Transparency;
end;

function TsChartLineRec.GetChartLine(AChart: TsChart;
  AElement: TsChartStyleElement; AIndex: Integer): TsChartline;
begin
  case AElement of
    ceBackground: Result := AChart.Border;
    ceWall: Result := AChart.PlotArea.Border;
    ceFloor: Result := AChart.Floor.Border;
    ceTitle: Result := AChart.Title.Border;
    ceSubTitle: Result := AChart.SubTitle.Border;
    ceLegend: Result := AChart.Legend.Border;
    ceXAxis: Result := AChart.XAxis.AxisLine;
    ceYAxis: Result := AChart.YAxis.AxisLine;
    ceX2Axis: Result := AChart.X2Axis.AxisLine;
    ceY2Axis: Result := AChart.Y2Axis.AxisLine;
    ceXAxisMajorGrid: Result := AChart.XAxis.MajorGridLines;
    ceXAxisMinorGrid: Result := AChart.XAxis.MinorGridLines;
    ceYAxisMajorGrid: Result := AChart.YAxis.MajorGridLines;
    ceYAxisMinorGrid: Result := AChart.YAxis.MinorGridLines;
    ceX2AxisMajorGrid: Result := AChart.X2Axis.MajorGridLines;
    ceX2AxisMinorGrid: Result := AChart.X2Axis.MinorGridLines;
    ceY2AxisMajorGrid: Result := AChart.Y2Axis.MajorGridLines;
    ceY2AxisMinorGrid: Result := AChart.Y2Axis.MinorGridLines;
    ceSeriesBorder: Result := AChart.Series[AIndex].Border;
    ceSeriesLine: Result := AChart.Series[AIndex].Line;
    else
      raise Exception.Create('[TsChartLineRec.GetChartLine] Line not supported.');
  end;
end;

procedure TsChartLineRec.ToChart(AChart: TsChart; AElement: TsChartStyleElement;
  AIndex: Integer);
var
  L: TsChartLine;
begin
  L := GetChartLine(AChart, AElement, AIndex);
  L.Style := Style;
  L.Width := Width;
  L.Color := Color;
  L.Transparency := Transparency;
end;

class operator TsChartLineRec.= (A, B: TsChartLineRec): Boolean;
begin
  Result := (A.Style = B.Style) and (A.Width = B.Width) and
    (A.Color = B.Color) and (A.Transparency = B.Transparency);
end;

{ TsChartFillRec }
procedure TsChartFillRec.FromChart(AChart: TsChart;
  AElement: TsChartStyleElement; AIndex: Integer);
var
  f: TsChartFill;
begin
  f := GetChartFill(AChart, AElement, AIndex);
  Style := f.Style;
  FgColor := f.FgColor;
  BgColor := f.BgColor;
end;

function TsChartFillRec.GetChartFill(AChart: TsChart;
  AElement: TsChartStyleElement; AIndex: Integer): TsChartFill;
begin
  case AElement of
    ceBackground: Result := AChart.Background;
    ceWall: Result := AChart.PlotArea.Background;
    ceFloor: Result := AChart.Floor.Background;
    ceLegend: Result := AChart.Legend.Background;
    ceTitle: Result := AChart.Title.Background;
    ceSubTitle: Result := AChart.SubTitle.Background;
    ceSeriesFill: Result := AChart.Series[AIndex].Fill;
    else
      raise Exception.Create('[TsChartFillRec.GetChartFill] Fill not supported.');
  end;
end;

procedure TsChartFillRec.ToChart(AChart: TsChart; AElement: TsChartStyleElement;
  AIndex: Integer);
var
  f: TsChartFill;
begin
  f := GetChartFill(AChart, AElement, AIndex);
  f.Style := Style;
  f.FgColor := FgColor;
  f.BgColor := BgColor;
end;

class operator TsChartFillRec.= (A, B: TsChartFillRec): Boolean;
begin
  Result := (A.Style = B.Style) and (A.FgColor = B.FgColor) and (A.BgColor = B.BgColor);
end;

{ TsChartAxisRec }
procedure TsChartAxisRec.FromChart(AChart: TsChart; AElement: TsChartStyleElement);
var
  axis: TsChartAxis;
begin
  axis := GetChartAxis(AChart, AElement);
  AutomaticMax := axis.AutomaticMax;
  AutomaticMin := axis.AutomaticMin;
  AutomaticMajorInterval := axis.AutomaticMajorInterval;
  AutomaticMinorInterval := axis.AutomaticMinorSteps;
  //AxisLine.FromChart(axis.AxisLine);
  //MajorGridLines.FromChart(axis.MajorGridLines);
  //MinorGridLines.FromChart(axis.MinorGridLines);
  //MajorTickLines.FromChart(axis.MajorTickLines);
  //MinorTickLines.FromChart(axis.MinorTickLines);
  Inverted := axis.Inverted;
//  CaptionFont.FromFont(Axis.Font);
  //LabelFont.FromFont(axis.LabelFont);
  LabelFormat := axis.LabelFormat;
  LabelRotation := axis.LabelRotation;
  Logarithmic := axis.Logarithmic;
  MajorInterval := axis.MajorInterval;
  MinorInterval := axis.MinorSteps;
  Position := axis.Position;
//  ShowCaption := Axis.ShowCaption;
  ShowLabels := axis.ShowLabels;
  Visible := axis.Visible;
end;

function TsChartAxisRec.GetChartAxis(AChart: TsChart; AElement: TsChartStyleElement): TsChartAxis;
begin
  case AElement of
    ceXAxis, ceXAxisCaption, ceXAxisMajorGrid, ceXAxisMinorGrid:
      Result := AChart.XAxis;
    ceYAxis, ceYAxisCaption, ceYAxisMajorGrid, ceYAxisMinorGrid:
      Result := AChart.YAxis;
    ceX2Axis, ceX2AxisCaption, ceX2AxisMajorGrid, ceX2AxisMinorGrid:
      Result := AChart.X2Axis;
    ceY2Axis, ceY2AxisCaption, ceY2AxisMajorGrid, ceY2AxisMinorGrid:
      Result := AChart.Y2Axis;
    else
      raise Exception.Create('[TsChartAxisRec.GetChartAxis] Element not supported.');
  end;
end;

function TsChartAxisRec.GetChartLine(AChart: TsChart; AElement: TsChartStyleElement): TsChartLine;
var
  axis: TsChartAxis;
begin
  axis := GetChartAxis(AChart, AElement);
  case AElement of
    ceXAxis, ceX2Axis, ceYAxis, ceY2Axis:
      Result := axis.AxisLine;
    ceXAxisMajorGrid, ceX2AxisMajorGrid, ceYAxisMajorGrid, ceY2AxisMajorGrid:
      Result := axis.MajorGridLines;
    ceXAxisMinorGrid, ceX2AxisMinorGrid, ceYAxisMinorGrid, ceY2AxisMinorGrid:
      Result := axis.MinorGridLines;
    else
      raise Exception.Create('[TsChartAxisRec.GetChartLine] Element not supported.');
  end;
end;

procedure TsChartAxisRec.ToChart(AChart: TsChart; AElement: TsChartStyleElement);
var
  axis: TsChartAxis;
begin
  axis := GetChartAxis(AChart, AElement);
  axis.AutomaticMax := AutomaticMax;
  axis.AutomaticMin := AutomaticMin;
  axis.AutomaticMajorInterval := AutomaticMajorInterval;
  axis.AutomaticMinorSteps := AutomaticMinorInterval;
  //AxisLine.ToChart(Axis.AxisLine);
  //MajorGridLines.ToChart(Axis.MajorGridLines);
  //MinorGridLines.ToChart(Axis.MinorGridLines);
  //MajorTickLines.ToChart(Axis.MajorTickLines);
  //MinorTickLines.ToChart(Axis.MinorTickLines);
  Axis.Inverted := Inverted;
//  CaptionFont.ToFont(Axis.Font);
  //LabelFont.ToFont(Axis.LabelFont);
  Axis.LabelFormat := LabelFormat;
  Axis.LabelRotation := LabelRotation;
  Axis.Logarithmic := Logarithmic;
  Axis.MajorInterval := MajorInterval;
  Axis.MinorSteps := MinorInterval;
  Axis.Position := Position;
//  Axis.ShowCaption := ShowCaption;
  Axis.Visible := Visible;
  Axis.ShowLabels := ShowLabels;
end;

class operator TsChartAxisRec.= (A, B: TsChartAxisRec): Boolean;
begin
  Result := (A.AutomaticMax = B.AutomaticMax) and (A.AutomaticMin = B.AutomaticMin) and
    (A.AutomaticMajorInterval = B.AutomaticMajorInterval) and
    (A.AutomaticMinorInterval = B.AutomaticMinorInterval) and
    (A.AxisLine = B.AxisLine) and
    (A.MajorGridLines = B.MajorGridLines) and
    (A.MinorGridLines = B.MinorGridLines) and
    (A.MajorTickLines = B.MajorTickLines) and
    (A.MinorTickLines = B.MinorTickLines) and
    (A.Inverted = B.Inverted) and
//    (A.CaptionFont = B.CaptionFont) and
    (A.LabelFont = B.LabelFont) and
    (A.LabelFormat = B.LabelFormat) and
    (A.LabelRotation = B.LabelRotation) and
    (A.Logarithmic = B.Logarithmic) and
    (A.MajorInterval = B.MajorInterval) and
    (A.MinorInterval = B.MinorInterval) and
    (A.Position = B.Position) and
//    (A.ShowCaption = B.ShowCaption) and
    (A.ShowLabels = B.ShowLabels) and
    (A.Visible = B.Visible);
end;

{ TsChartTextRec }

procedure TsChartTextRec.FromChart(AChart: TsChart; AElement: TsChartStyleElement);
begin
  case AElement of
    ceTitle:
      begin
        Font.FromChart(AChart, ceTitle);
        Visible := AChart.Title.ShowCaption;
      end;
    ceSubtitle:
      begin
        Font.FromChart(AChart, ceSubTitle);
        Visible := AChart.Subtitle.ShowCaption;
      end;
    ceXAxisCaption:
      begin
        Font.FromChart(AChart, ceXAxisCaption);
        Visible := AChart.XAxis.ShowCaption;
      end;
    ceYAxisCaption:
      begin
        Font.FromChart(AChart, ceYAxisCaption);
        Visible := AChart.YAxis.ShowCaption;
      end;
    ceX2AxisCaption:
      begin
        Font.FromChart(AChart, ceX2AxisCaption);
        Visible := AChart.X2Axis.ShowCaption;
      end;
    ceY2AxisCaption:
      begin
        Font.FromChart(AChart, ceY2AxisCaption);
        Visible := AChart.Y2Axis.ShowCaption;
      end;
  end;
end;

procedure TsChartTextRec.ToChart(AChart: TsChart; AElement: TsChartStyleElement);
begin
  case AElement of
    ceTitle:
      begin
        Font.ToChart(AChart, ceTitle);
        AChart.Title.ShowCaption := Visible;
      end;
    ceSubtitle:
      begin
        Font.ToChart(AChart, ceSubtitle);
        AChart.Subtitle.ShowCaption := Visible;
      end;
    ceXAxisCaption:
      begin
        Font.ToChart(AChart, ceXAxisCaption);
        AChart.XAxis.ShowCaption := Visible;
      end;
    ceYAxisCaption:
      begin
        Font.ToChart(AChart, ceYAxisCaption);
        AChart.YAxis.ShowCaption := Visible;
      end;
    ceX2AxisCaption:
      begin
        Font.ToChart(AChart, ceX2AxisCaption);
        AChart.X2Axis.ShowCaption := Visible;
      end;
    ceY2AxisCaption:
      begin
        Font.ToChart(AChart, ceY2AxisCaption);
        AChart.Y2Axis.ShowCaption := Visible;
      end;
  end;
end;

class operator TsChartTextRec.= (A, B: TsChartTextRec): Boolean;
begin
  Result := (A.Font = B.Font) and (A.Visible = B.Visible);
end;

{ TsChartLegendRec }
procedure TsChartLegendRec.FromChart(AChart: TsChart; AElement: TsChartStyleElement);
begin
  Font.FromChart(AChart, ceLegend);
  Border.FromChart(AChart, ceLegend, 0);
  Fill.FromChart(AChart, ceLegend, 0);
  Visible := AChart.Legend.Visible;
end;

procedure TsChartLegendRec.ToChart(AChart: TsChart; AElement: TsChartStyleElement);
begin
  Font.ToChart(AChart, ceLegend);
  Border.ToChart(AChart, ceLegend, 0);
  Fill.ToChart(AChart, ceLegend, 0);
  AChart.Legend.Visible := Visible;
end;

class operator TsChartLegendRec.= (A, B: TsChartLegendRec): Boolean;
begin
  Result := (A.Font = B.Font) and (A.Border = B.Border) and (A.Fill = B.Fill) and
    (A.Visible = B.Visible);
end;

{ TsChartPlotAreaRec }
procedure TsChartPlotAreaRec.FromChart(AChart: TsChart; AElement: TsChartStyleElement);
begin
  FChart := AChart;
end;

procedure TSChartPlotAreaRec.ToChart(AChart: TsChart; AElement: TsChartStyleElement);
begin
end;

class operator TsChartPlotAreaRec.= (A, B: TsChartPlotAreaRec): Boolean;
begin
  Result := A.FChart = B.FChart;
end;

{ TsChartSeriesRec }
procedure TsChartSeriesRec.FromChart(AChart: TsChart; AElement: TsChartStyleElement;
  AIndex: Integer);
begin
  Line.FromChart(AChart, AElement, AIndex);
  Fill.FromChart(AChart, AElement, AIndex);
  Border.FromChart(AChart, AElement, AIndex);
end;

procedure TsChartSeriesRec.ToChart(AChart: TsChart; AElement: TsChartStyleElement;
  AIndex: Integer);
begin
  Line.ToChart(AChart, ceSeriesLine, AIndex);
  Fill.ToChart(AChart, ceSeriesFill, AIndex);
  Border.ToChart(AChart, ceSeriesBorder, AIndex);
end;

class operator TsChartSeriesRec.= (A, B: TsChartSeriesRec): Boolean;
begin
  Result := (A.Line = B.Line) and (A.Fill = B.Fill) and (A.Border = B.Border);
end;


{==============================================================================}
{                 Style classes to be listed in ChartStyleList                 }
{==============================================================================}

{ TsChartStyle }

constructor TsChartStyle.Create(AElement: TsChartStyleElement);
begin
  inherited Create;
  FElement := AElement;
end;

{ TsChartStyle_Background }

procedure TsChartStyle_Background.ApplyToChart(AChart: TsChart; AIndex: Integer);
begin
  if (FElement in [ceBackground, ceWall, ceFloor]) then
  begin
    FBackground.ToChart(AChart, FElement, AIndex);
    FBorder.ToChart(AChart, FElement, AIndex);
  end else
    raise Exception.Create('[TsChartStyle_Background.ApplyToChart] Unknown background style.');
end;

procedure TsChartStyle_Background.ExtractFromChart(AChart: TsChart; AIndex: Integer);
begin
  if (FElement in [ceBackground, ceWall, ceFloor]) then
  begin
    FBackground.FromChart(AChart, FElement, AIndex);
    FBorder.FromChart(AChart, FElement, AIndex);
  end else
    raise Exception.Create('[TsChartStyle_Background.ExtractFromChart] Unknown background style.');
end;


{ TsChartStyle_Line }

const
  ALLOWED_LS = [
    ceXAxisMajorGrid, ceXAxisMinorGrid, ceX2AxisMajorGrid, ceX2AxisMinorGrid,
    ceYAxisMajorGrid, ceYAxisMinorGrid, ceY2AxisMajorGrid, ceY2AxisMinorGrid
  ];

procedure TsChartStyle_Line.ApplyToChart(AChart: TsChart; AIndex: Integer);
begin
  if (FElement in ALLOWED_LS) then
    Line.ToChart(AChart, FElement, AIndex)
  else
    raise Exception.Create('[TsChartStyle_Line.ApplytoGrid] Unknown line');
end;

procedure TsChartStyle_Line.ExtractFromChart(AChart: TsChart; AIndex: Integer);
begin
  if (FElement in ALLOWED_LS) then
    Line.FromChart(AChart, FElement, AIndex)
  else
    raise Exception.Create('[TsChartStyle_Line.ExtractFromChart] Unknown line');
end;


{ TsChartStyle_Axis }

procedure TsChartStyle_Axis.ApplyToChart(AChart: TsChart; AIndex: Integer);
begin
  Axis.ToChart(AChart, FElement);
  Axis.AxisLine.ToChart(AChart, FElement, -1);
  Axis.LabelFont.ToChart(AChart, FElement);
end;

procedure TsChartStyle_Axis.ExtractFromChart(AChart: TsChart; AIndex: Integer);
begin
  Axis.FromChart(AChart, FElement);
  Axis.AxisLine.FromChart(AChart, FElement, -1);
  Axis.LabelFont.FromChart(AChart, FElement);
end;


{ TsChartStyle_Caption }

const
  ALLOWED_CAPTIONS = [ceXAxisCaption, ceX2AxisCaption, ceYAxisCaption, ceY2AxisCaption, ceTitle, ceSubTitle];

procedure TsChartStyle_Caption.ApplyToChart(AChart: TsChart; AIndex: Integer);
begin
  if (FElement in ALLOWED_CAPTIONS) then
    Caption.ToChart(AChart, FElement)
  else
    raise Exception.Create('[TsChartstyle_Caption.ApplyToChart] Unknown caption');
end;

procedure TsChartStyle_Caption.ExtractFromChart(AChart: TsChart; AIndex: Integer);
begin
  if (FElement in ALLOWED_CAPTIONS) then
    Caption.FromChart(AChart, FElement)
  else
    raise Exception.Create('[TsChartStyle_Caption.ExtractFromChart] Unknown caption');
end;


{ TsChartStyle_Legend }

procedure TsChartStyle_Legend.ApplyToChart(AChart: TsChart; AIndex: Integer);
begin
  FLegend.ToChart(AChart, ceLegend);
end;

procedure TsChartStyle_Legend.ExtractFromChart(AChart: TsChart; AIndex: Integer);
begin
  FLegend.FromChart(AChart, ceLegend);
end;


{ TsChartStyle_PlotArea }
{ For the moment, this is a dummy style because I don't know how the plotarea
  parameters are changed in ODS. }

procedure TsChartStyle_PlotArea.ApplyToChart(AChart: TsChart; AIndex: Integer);
begin
  FPlotArea.ToChart(AChart, cePlotArea);
end;

procedure TsChartStyle_PlotArea.ExtractFromChart(AChart: TsChart; AIndex: Integer);
begin
  FPlotArea.FromChart(AChart, cePlotArea);
end;


{ TsChartStyle_Series }

procedure TsChartStyle_Series.ApplyToChart(AChart: TsChart; AIndex: Integer);
begin
  Line.ToChart(AChart, ceSeriesLine, AIndex);
  Fill.ToChart(AChart, ceSeriesFill, AIndex);
  Border.ToChart(AChart, ceSeriesBorder, AIndex);

  if (AChart.Series[AIndex] is TsLineSeries) then
  begin
    TsLineSeries(AChart.Series[AIndex]).Symbol := FSymbol;
    TsLineSeries(AChart.Series[AIndex]).SymbolHeight := FSymbolHeight;
    TsLineSeries(AChart.Series[AIndex]).SymbolWidth := FSymbolWidth;
    TsLineSeries(AChart.Series[AIndex]).ShowSymbols := FShowSymbols;
  end;
end;

procedure TsChartStyle_Series.ExtractFromChart(AChart: TsChart; AIndex: Integer);
begin
  Line.FromChart(AChart, ceSeriesLine, AIndex);
  Fill.FromChart(AChart, ceSeriesFill, AIndex);
  Border.FromChart(AChart, ceSeriesBorder, AIndex);

  if (AChart.Series[AIndex] is TsLineSeries) then
  begin
    FSymbol := TsLineSeries(AChart.Series[AIndex]).Symbol;
    FSymbolHeight := TsLineSeries(AChart.Series[AIndex]).SymbolHeight;
    FSymbolWidth := TsLineSeries(AChart.Series[AIndex]).SymbolWidth;
    FShowSymbols := TsLineSeries(AChart.Series[AIndex]).ShowSymbols;
  end else
    FShowSymbols := false;
end;


{==============================================================================}
{                          TsChartStyleList                                    }
{==============================================================================}

destructor TsChartStyleList.Destroy;
begin
  Clear;
  inherited;
end;

{ Adds a new style to the style list. The style is created as the given style
  class. Which piece of chart formatting is included in the style, is determined
  by the AElement parameter. In case of series styles, the index of the series
  must be provided as parameter AIndex. }
function TsChartStyleList.AddChartStyle(AName: String; AChart: TsChart;
  AStyleClass: TsChartStyleClass; AElement: TsChartStyleElement;
  AIndex: Integer = -1): Integer;
var
  newStyle: TsChartStyle;
begin
  newStyle := AStyleClass.Create(AElement);
  newStyle.ExtractFromChart(AChart, AIndex);
  newStyle.FName := AName;
  Result := Add(newStyle);
end;

{ Clears the chart style list. Destroys the individual items. }
procedure TsChartStyleList.Clear;
var
  j: Integer;
begin
  for j := 0 to Count-1 do
    TsChartStyle(Items[j]).Free;
  inherited Clear;
end;

procedure TsChartStyleList.Delete(AIndex: Integer);
begin
  TsChartStyle(Items[AIndex]).Free;
  inherited;
end;

function TsChartStyleList.FindStyleIndexByName(const AName: String): Integer;
var
  i: Integer;
begin
  for i := 0 to Count-1 do
  begin
    if TsChartStyle(Items[i]).Name = AName then
    begin
      Result := i;
      exit;
    end;
  end;
  Result := -1;
end;
(*
{ Adds a new style to the style list. The style is created as the given style
  class. Which piece of chart formattting is included in the style, is determined
  by the AElement parameter. In case of series styles, the index of the series
  must be provided as parameter AIndex. }
function TsChartStyleList.AddChartStyle(AChart: TsChart;
  AStyleClass: TsChartStyleClass; AElement: TsChartStyleElement;
  AIndex: Integer = -1): Integer;
var
  newStyle, style: TsChartStyle;
  i: Integer;
begin
//  Result := -1;

  newStyle := AStyleClass.Create(AElement);
  newStyle.ExtractFromChart(AChart, AIndex);
  Result := Add(newStyle);
  newStyle.FStyleID := Result;
  {
  for i := 0 to Count-1 do
  begin
    if (TsChartStyle(Items[i]) is AStyleClass) then
    begin
      style := TsChartStyle(Items[i]);
      if style.EqualTo(newStyle) then
      begin
        Result := i;
        break;
      end;
    end;
  end;

  if Result = -1 then
    Result := Add(newStyle)
  else
    newStyle.Free;
    }
end;
*)
                (*
{ Finds the index of the style matching the formatting of the given
  chart element (AElement); in case of series styles the series index must be
  provided as AIndex.
  Returns -1 if there is no such style. }
function TsChartStyleList.FindChartStyle(AChart: TsChart;
  AStyleClass: TsChartStyleClass; AElement: TsChartStyleElement;
  AIndex: Integer = -1): Integer;
var
  newStyle, style: TsChartStyle;
  i: Integer;
begin
  Result := -1;

  newStyle := AStyleClass.Create(AElement);
  try
    newStyle.ExtractFromChart(AChart, AIndex);

    for i := 0 to Count-1 do
    begin
      if (TsChartStyle(Items[i]) is AStyleClass) then
      begin
        style := TsChartStyle(Items[i]);
        if style.EqualTo(newStyle) then
        begin
          Result := i;
          exit;
        end;
      end;
    end;
  finally
    newStyle.Free;
  end;
end;
      *)
(*
{ Searches whether the style of the specified axis is already in the
  list. If not, a new style is created and added.
  The type of the requested axis must be provided as parameter.
  Returns the index of the style. }
function TsChartStyleList.FindChartAxisStyle(AChart: TsChart;
  AType: TsChartAxisType): Integer;
var
  newStyle, style: TsChartAxisStyle;
  i: Integer;
begin
  Result := -1;
  newStyle := TsChartAxisStyle.Create(AType);
  newStyle.ExtractFromChart(AChart);
  for i := 0 to Count-1 do
  begin
    if (TsChartStyle(Items[i]) is TsChartAxisStyle) then
    begin
      style := TsChartAxisStyle(Items[i]);
      if (style.AxisType = AType) and (style.Axis = newStyle.Axis) then
      begin
        Result := i;
        break;
      end;
    end;
  end;
  if Result = -1 then
    Result := Add(newStyle)
  else
    newStyle.Free;
end;

{ Searches whether the style of the specified line is already in the
  list. If not, a new style is created and added.
  Returns the index of the style. }
function TsChartStyleList.FindChartLineStyle(AChart: TsChart): Integer;
//  AType: TsChartLineType): Integer;
var
  newStyle, style: TsChartLineStyle;
  i: Integer;
begin
  Result := -1;
  newStyle := TsChartLineStyle.Create; //(AType);
  newStyle.ExtractFromChart(AChart);
  if newStyle.Line.Style = clsNoLine then
    exit;
  for i := 0 to Count-1 do
  begin
    if (TsChartStyle(Items[i]) is TsChartLineStyle) then
    begin
      style := TsChartLineStyle(Items[i]);
      if (style.Line = newStyle.Line) then
      begin
        Result := i;
        break;
      end;
    end;
  end;
  if Result = -1 then
    Result := Add(newStyle)
  else
    newStyle.Free;
end;

{ Searches whether the background style of the specified chart is already in the
  list. If not, a new style is created and added.
  The type of the requested background must be provided as parameter.
  Returns the index of the style. }
function TsChartStyleList.FindChartBackgroundStyle(AChart: TsChart;
  AType: TsChartBackgroundType): Integer;
var
  newStyle, style: TsChartBackgroundStyle;
  i: Integer;
begin
  Result := -1;
  newStyle := TsChartBackgroundStyle.Create(AType);
  newStyle.ExtractFromChart(AChart);
  for i := 0 to Count-1 do
  begin
    if (TsChartStyle(Items[i]) is TsChartBackgroundStyle) then
    begin
      style := TsChartBackgroundStyle(Items[i]);
      if (style.BackgroundType = AType) and (style.Background = newStyle.Background) then
      begin
        Result := i;
        break;
      end;
    end;
  end;
  if Result = -1 then
    Result := Add(newStyle)
  else
    newStyle.Free;
end;

{ Searches whether the style of the specified caption is already in the
  list. If not, a new style is created and added.
  The type of the requested axis must be provided as parameter.
  Returns the index of the style. }
function TsChartStyleList.FindChartCaptionStyle(AChart: TsChart;
  AType: TsChartCaptionType): Integer;
var
  newStyle, style: TsChartCaptionStyle;
  i: Integer;
begin
  Result := -1;
  newStyle := TsChartCaptionStyle.Create(AType);
  newStyle.ExtractFromChart(AChart);
  for i := 0 to Count-1 do
  begin
    if (TsChartStyle(Items[i]) is TsChartCaptionStyle) then
    begin
      style := TsChartCaptionStyle(Items[i]);
      if {(style.AxisType = AType) and} (style.Caption = newStyle.Caption) then
      begin
        Result := i;
        break;
      end;
    end;
  end;
  if Result = -1 then
    Result := Add(newStyle)
  else
    newStyle.Free;
end;

{ Searches whether the legend style of the given chart is already in the
  list. If not, a new style is created and added.
  The type of the requested axis must be provided as parameter.
  Returns the index of the style. }
function TsChartStyleList.FindChartLegendStyle(AChart: TsChart): Integer;
var
  newStyle, style: TsChartLegendStyle;
  i: Integer;
begin
  Result := -1;
  newStyle := TsChartLegendStyle.Create;
  newStyle.ExtractFromChart(AChart);
  for i := 0 to Count-1 do
  begin
    if (TsChartStyle(Items[i]) is TsChartLegendStyle) then
    begin
      style := TsChartLegendStyle(Items[i]);
      if (style.Legend = newStyle.Legend) then
      begin
        Result := i;
        break;
      end;
    end;
  end;
  if Result = -1 then
    Result := Add(newStyle)
  else
    newStyle.Free;
end;
*)
end.

