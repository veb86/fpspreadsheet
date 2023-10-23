unit fpsChartStyles;

{$mode objfpc}{$H+}
{$modeswitch advancedrecords}

interface

uses
  Classes, SysUtils, fpsTypes, fpsChart;

type
  TsChartBackgroundType = (cbtBackground, cbtWall, cbtFloor);

  TsChartAxisType = (catPrimaryX, catPrimaryY, catSecondaryX, catSecondaryY);

  TsChartCaptionType = (cctTitle, cctSubtitle,
    cctPrimaryX, cctPrimaryY, cctSecondaryX, cctSecondaryY);

  TsChartLineRec = record
    Style: Integer;        // index into chart's LineStyle list or predefined clsSolid/clsNoLine
    Width: Double;         // mm
    Color: TsColor;        // in hex: $00bbggrr, r=red, g=green, b=blue
    Transparency: Double;  // in percent
    procedure FromLine(ALine: TsChartline);
    procedure ToLine(ALine: TsChartLine);
    class operator = (A, B: TsChartLineRec): Boolean;
  end;

  TsChartFillRec = record
    Style: TsFillStyle;
    FgColor: TsColor;
    BgColor: TsColor;
    procedure FromFill(AFill: TsChartFill);
    procedure ToFill(AFill: TsChartFill);
    class operator = (A, B: TsChartFillRec): Boolean;
  end;

  TsChartFontRec = record
    FontName: String;
    Size: Double;
    Style: TsFontStyles;
    Color: TsColor;
    Position: TsFontPosition;
    procedure FromFont(AFont: TsFont);
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
    procedure FromAxis(Axis: TsChartAxis);
    procedure ToAxis(Axis: TsChartAxis);
    class operator = (A, B: TsChartAxisRec): Boolean;
  end;

  TsChartCaptionRec = record
    Font: TsChartFontRec;
    Rotation: Integer;
    Visible: Boolean;
    procedure FromChart(AChart: TsChart; AType: TsChartCaptionType);
    procedure ToChart(AChart: TsChart; AType: TsChartCaptionType);
    class operator = (A, B: TsChartCaptionRec): Boolean;
  end;

  {----------------------------------------------------------------------------}

  TsChartStyle = class
  public
    procedure ApplyToChart(AChart: TsChart); virtual; abstract;
    procedure ExtractFromChart(AChart: TsChart); virtual; abstract;
  end;

  TsChartBackgroundStyle = class(TsChartStyle)
  private
    FBackgroundType: TsChartBackgroundType;
    FBackground: TsChartFillRec;
    FBorder: TsChartLineRec;
  public
    constructor Create(AType: TsChartBackgroundType);
    procedure ApplyToChart(AChart: TsChart); override;
    procedure ExtractFromChart(AChart: TsChart); override;
    property BackgroundType: TsChartBackgroundType read FBackgroundType write FBackgroundType;

    property Background: TsChartFillRec read FBackground;
    property Border: TsChartLineRec read FBorder;
  end;

  TsChartAxisStyle = class(TsChartStyle)
  private
    FAxis: TsChartAxisRec;
    FAxisType: TsChartAxisType;
  public
    constructor Create(AType: TsChartAxisType);
    procedure ApplyToChart(AChart: TsChart); override;
    procedure ExtractFromChart(AChart: TsChart); override;

    property Axis: TsChartAxisRec read FAxis write FAxis;
    property AxisType: TsChartAxisType read FAxisType write FAxisType;
  end;

  TsChartCaptionStyle = class(TsChartStyle)
  private
    FCaption: TsChartCaptionRec;
    FCaptionType: TsChartCaptionType;
  public
    constructor Create(AType: TsChartCaptionType);
    procedure ApplyToChart(AChart: TsChart); override;
    procedure ExtractFromChart(AChart: TsChart); override;

    property Caption: TsChartCaptionRec read FCaption write FCaption;
    property CaptionType: TsChartCaptionType read FCaptionType write FCaptionType;
  end;

  { ---------------------------------------------------------------------------}

  TsChartStyleList = class(TFPList)
  protected

  public
    destructor Destroy; override;
    procedure AddChartAxisStyle(AChart: TsChart; AType: TsChartAxisType);
    procedure AddChartBackgroundStyle(AChart: TsChart; AType: TsChartBackgroundType);
    procedure AddChartCaptionStyle(AChart: TsChart; AType: TsChartCaptionType);
    procedure Clear;
    function FindChartAxisStyle(AChart: TsChart; AType: TsChartAxisType): Integer;
    function FindChartBackgroundStyle(AChart: TsChart; AType: TsChartBackgroundType): Integer;
    function FindChartCaptionStyle(AChart: TsChart; AType: TsChartCaptionType): Integer;
  end;

implementation

{==============================================================================}
{                             Style records                                    }
{==============================================================================}

{ TsFontRec }
procedure TsChartFontRec.FromFont(AFont: TsFont);
begin
  FontName := AFont.FontName;
  Size := AFont.Size;
  Style := AFont.Style;
  Color := AFont.Color;
  Position := AFont.Position;
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
procedure TsChartLineRec.FromLine(ALine: TsChartLine);
begin
  Style := ALine.Style;
  Width := ALine.Width;
  Color := ALine.Color;
  Transparency := ALine.Transparency;
end;

procedure TsChartLineRec.ToLine(ALine: TsChartLine);
begin
  ALine.Style := Style;
  ALine.Width := Width;
  ALine.Color := Color;
  ALine.Transparency := Transparency;
end;

class operator TsChartLineRec.= (A, B: TsChartLineRec): Boolean;
begin
  Result := (A.Style = B.Style) and (A.Width = B.Width) and
    (A.Color = B.Color) and (A.Transparency = B.Transparency);
end;

{ TsChartFillRec }
procedure TsChartFillRec.FromFill(AFill: TsChartFill);
begin
  Style := AFill.Style;
  FgColor := AFill.FgColor;
  BgColor := AFill.BgColor;
end;

procedure TsChartFillRec.ToFill(AFill: TsChartFill);
begin
  AFill.Style := Style;
  AFill.FgColor := FgColor;
  AFill.BgColor := BgColor;
end;

class operator TsChartFillRec.= (A, B: TsChartFillRec): Boolean;
begin
  Result := (A.Style = B.Style) and (A.FgColor = B.FgColor) and (A.BgColor = B.BgColor);
end;

{ TsChartAxisRec }
procedure TsChartAxisRec.FromAxis(Axis: TsChartAxis);
begin
  AutomaticMax := Axis.AutomaticMax;
  AutomaticMin := Axis.AutomaticMin;
  AutomaticMajorInterval := Axis.AutomaticMajorInterval;
  AutomaticMinorInterval := Axis.AutomaticMinorSteps;
  AxisLine.FromLine(Axis.AxisLine);
  MajorGridLines.FromLine(Axis.MajorGridLines);
  MinorGridLines.FromLine(Axis.MinorGridLines);
  MajorTickLines.FromLine(Axis.MajorTickLines);
  MinorTickLines.FromLine(Axis.MinorTickLines);
  Inverted := Axis.Inverted;
//  CaptionFont.FromFont(Axis.Font);
  LabelFont.FromFont(Axis.LabelFont);
  LabelFormat := Axis.LabelFormat;
  LabelRotation := Axis.LabelRotation;
  Logarithmic := Axis.Logarithmic;
  MajorInterval := Axis.MajorInterval;
  MinorInterval := Axis.MinorSteps;
  Position := Axis.Position;
//  ShowCaption := Axis.ShowCaption;
  ShowLabels := Axis.ShowLabels;
  Visible := Axis.Visible;
end;

procedure TsChartAxisRec.ToAxis(Axis: TsChartAxis);
begin
  Axis.AutomaticMax := AutomaticMax;
  Axis.AutomaticMin := AutomaticMin;
  Axis.AutomaticMajorInterval := AutomaticMajorInterval;
  Axis.AutomaticMinorSteps := AutomaticMinorInterval;
  AxisLine.ToLine(Axis.AxisLine);
  MajorGridLines.ToLine(Axis.MajorGridLines);
  MinorGridLines.ToLine(Axis.MinorGridLines);
  MajorTickLines.ToLine(Axis.MajorTickLines);
  MinorTickLines.ToLine(Axis.MinorTickLines);
  Axis.Inverted := Inverted;
//  CaptionFont.ToFont(Axis.Font);
  LabelFont.ToFont(Axis.LabelFont);
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

{ TsChartCaptionRec }

procedure TsChartCaptionRec.FromChart(AChart: TsChart; AType: TsChartCaptionType);
begin
  case AType of
    cctTitle:
      begin
        Font.FromFont(AChart.Title.Font);
        Visible := AChart.Title.ShowCaption;
      end;
    cctSubtitle:
      begin
        Font.FromFont(AChart.Subtitle.Font);
        Visible := AChart.Subtitle.ShowCaption;
      end;
    cctPrimaryX:
      begin
        Font.FromFont(AChart.XAxis.CaptionFont);
        Visible := AChart.XAxis.ShowCaption;
      end;
    cctPrimaryY:
      begin
        Font.FromFont(AChart.YAxis.CaptionFont);
        Visible := AChart.YAxis.ShowCaption;
      end;
    cctSecondaryX:
      begin
        Font.FromFont(AChart.X2Axis.CaptionFont);
        Visible := AChart.X2Axis.ShowCaption;
      end;
    cctSecondaryY:
      begin
        Font.FromFont(AChart.Y2Axis.CaptionFont);
        Visible := AChart.Y2Axis.ShowCaption;
      end;
  end;
end;

procedure TsChartCaptionRec.ToChart(AChart: TsChart; AType: TsChartCaptionType);
begin
  case AType of
    cctTitle:
      begin
        Font.ToFont(AChart.Title.Font);
        AChart.Title.ShowCaption := Visible;
      end;
    cctSubtitle:
      begin
        Font.ToFont(AChart.Subtitle.Font);
        AChart.Subtitle.ShowCaption := Visible;
      end;
    cctPrimaryX:
      begin
        Font.ToFont(AChart.XAxis.CaptionFont);
        AChart.XAxis.ShowCaption := Visible;
      end;
    cctPrimaryY:
      begin
        Font.ToFont(AChart.YAxis.CaptionFont);
        AChart.YAxis.ShowCaption := Visible;
      end;
    cctSecondaryX:
      begin
        Font.ToFont(AChart.X2Axis.CaptionFont);
        AChart.X2Axis.ShowCaption := Visible;
      end;
    cctSecondaryY:
      begin
        Font.ToFont(AChart.Y2Axis.CaptionFont);
        AChart.Y2Axis.ShowCaption := Visible;
      end;
  end;
end;

class operator TsChartCaptionRec.= (A, B: TsChartCaptionRec): Boolean;
begin
  Result := (A.Font = B.Font) and (A.Visible = B.Visible);
end;


{ TsChartBackgroundstyle }

constructor TsChartBackgroundStyle.Create(AType: TsChartBackgroundType);
begin
  inherited Create;
  FBackgroundType := AType;
end;

procedure TsChartBackgroundStyle.ApplyToChart(AChart: TsChart);
begin
  case FBackgroundType of
    cbtBackground:
      begin
        FBackground.ToFill(AChart.Background);
        FBorder.ToLine(AChart.Border);
      end;
    cbtWall:
      begin
        FBackground.ToFill(AChart.PlotArea.Background);
        FBorder.ToLine(AChart.PlotArea.Border);
      end;
    cbtFloor:
      begin
        FBackground.ToFill(AChart.Floor.Background);
        FBorder.ToLine(AChart.Floor.Border);
      end;
    else
      raise Exception.Create('Unknown background style.');
  end;
end;

procedure TsChartBackgroundStyle.ExtractFromChart(AChart: TsChart);
begin
  case FBackgroundType of
    cbtBackground:
      begin
        FBackground.FromFill(AChart.Background);
        FBorder.FromLine(AChart.Border);
      end;
    cbtWall:
      begin
        FBackground.FromFill(AChart.PlotArea.Background);
        FBorder.FromLine(AChart.PlotArea.Border);
      end;
    cbtFloor:
      begin
        FBackground.FromFill(AChart.Floor.Background);
        FBorder.FromLine(AChart.Floor.Border);
      end;
  end;
end;

{ TsChartAxisStyle }

constructor TsChartAxisStyle.Create(AType: TsChartAxisType);
begin
  inherited Create;
  FAxisType := AType;
end;

procedure TsChartAxisStyle.ApplyToChart(AChart: TsChart);
begin
  case FAxisType of
    catPrimaryX: Axis.ToAxis(AChart.XAxis);
    catPrimaryY: Axis.ToAxis(AChart.YAxis);
    catSecondaryX: Axis.ToAxis(AChart.X2Axis);
    catSecondaryY: Axis.ToAxis(AChart.Y2Axis);
  end;
end;

procedure TsChartAxisStyle.ExtractFromChart(AChart: TsChart);
begin
  case FAxisType of
    catPrimaryX: Axis.FromAxis(AChart.XAxis);
    catPrimaryY: Axis.FromAxis(AChart.YAxis);
    catSecondaryX: Axis.FromAxis(AChart.X2Axis);
    catSecondaryY: Axis.FromAxis(AChart.Y2Axis);
  end;
end;


{ TsChartCaptionStyle }

constructor TsChartCaptionStyle.Create(AType: TsChartCaptionType);
begin
  inherited Create;
  FCaptionType := AType;
end;

procedure TsChartCaptionStyle.ApplyToChart(AChart: TsChart);
begin
  Caption.ToChart(AChart, FCaptionType);
end;

procedure TsChartCaptionStyle.ExtractFromChart(AChart: TsChart);
begin
  Caption.FromChart(AChart, FCaptionType);
end;


{ TsChartStyleList }

destructor TsChartStyleList.Destroy;
begin
  Clear;
  inherited;
end;

{ Adds the style of the specified axis type in the given chart as new style to
  the style list. But only if the same style does not yet exist. }
procedure TsChartStyleList.AddChartAxisStyle(AChart: TsChart;
  AType: TsChartAxisType);
begin
  FindChartAxisStyle(AChart, AType);
end;

{ Adds the style of the specified caption in the given chart as new style to
  the style list. But only if the same style does not yet exist. }
procedure TsChartStyleList.AddChartCaptionStyle(AChart: TsChart;
  AType: TsChartCaptionType);
begin
  FindChartCaptionStyle(AChart, AType);
end;

{ Adds the style of the specified type in the given chart as new style to the
  style list. But only if the same style does not yet exist. }
procedure TsChartStyleList.AddChartBackgroundStyle(AChart: TsChart;
  AType: TsChartBackgroundType);
begin
  FindChartBackgroundStyle(AChart, AType);
end;

procedure TsChartStyleList.Clear;
var
  j: Integer;
begin
  for j := 0 to Count-1 do
    TsChartStyle(Items[j]).Free;
  inherited Clear;
end;

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


end.

