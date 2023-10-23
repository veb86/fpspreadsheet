unit fpsChartStyles;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpsTypes, fpsChart;

type
  TsChartStyleType = (cstBackground, cstWall, cstFloor);

  TsChartStyle = class
  private
    FStyleType: TsChartStyleType;
  public
    constructor Create(AStyleType: TsChartStyleType); virtual;
    procedure ApplyToChart(AChart: TsChart); virtual; abstract;
    procedure ExtractFromChart(AChart: TsChart); virtual; abstract;
    property StyleType: TsChartStyleType read FStyleType;
  end;

  TsChartBackgroundStyle = class(TsChartStyle)
  private
    FBackground: TsChartFillRec;
    FBorder: TsChartLineRec;
  public
    procedure ApplyToChart(AChart: TsChart); override;
    procedure ExtractFromChart(AChart: TsChart); override;
    property Background: TsChartFillRec read FBackground;
    property Border: TsChartLineRec read FBorder;
  end;

  TsChartStyleList = class(TFPList)
  protected

  public
    destructor Destroy; override;
    procedure AddChartBackgroundStyle(AChart: TsChart; AStyleType: TsChartStyleType);
    procedure Clear;
    function FindChartBackgroundStyle(AChart: TsChart; AStyleType: TsChartStyleType): Integer;
  end;

implementation

{ TsChartStyle }

constructor TsChartStyle.Create(AStyleType: TsChartStyleType);
begin
  FStyleType := AStyleType;
end;

{ TsChartBackgroundstyle }

procedure TsChartBackgroundStyle.ApplyToChart(AChart: TsChart);
begin
  case FStyleType of
    cstBackground:
      begin
        AChart.Background.FromRecord(FBackground);
        AChart.Border.FromRecord(FBorder);
      end;
    cstWall:
      begin
        AChart.PlotArea.Background.FromRecord(FBackground);
        AChart.PlotArea.Border.FromRecord(FBorder);
      end;
    cstFloor:
      begin
        AChart.Floor.Background.FromRecord(FBackGround);
        AChart.Floor.Border.FromRecord(FBorder);
      end;
  end;
end;

procedure TsChartBackgroundStyle.ExtractFromChart(AChart: TsChart);
begin
  case FStyleType of
    cstBackground:
      begin
        FBackground := AChart.Background.ToRecord;
        FBorder := AChart.Border.ToRecord;
      end;
    cstWall:
      begin
        FBackground := AChart.PlotArea.Background.ToRecord;
        FBorder := AChart.PlotArea.Border.ToRecord;
      end;
    cstFloor:
      begin
        FBackground := AChart.Floor.Background.ToRecord;
        FBorder := AChart.Floor.Border.ToRecord;
      end;
  end;
end;

{ TsChartStyleList }

destructor TsChartStyleList.Destroy;
begin
  Clear;
  inherited;
end;

{ Adds the style of the specified type in the given chart as new style to the
  style list. But only if the same style does not yet exist. }
procedure TsChartStyleList.AddChartBackgroundStyle(AChart: TsChart;
  AStyleType: TsChartStyleType);
begin
  FindChartBackgroundStyle(AChart, AStyleType);
end;

procedure TsChartStyleList.Clear;
var
  j: Integer;
begin
  for j := 0 to Count-1 do
    TsChartStyle(Items[j]).Free;
  inherited Clear;
end;

{ Searches whether the background style of the specified chart is already in the
  list. If not, a new style is created and added.
  The type of the requested background must be provided as parameter.
  Returns the index of the style. }
function TsChartStyleList.FindChartBackgroundStyle(AChart: TsChart;
  AStyleType: TsChartStyleType): Integer;
var
  newStyle, style: TsChartBackgroundStyle;
  i: Integer;
begin
  Result := -1;
  newStyle := TsChartBackgroundStyle.Create(AStyleType);
  newStyle.ExtractFromChart(AChart);
  for i := 0 to Count-1 do
  begin
    if (TsChartStyle(Items[i]) is TsChartBackgroundStyle) and
       (TsChartStyle(Items[i]).StyleType = AStyleType) then
    begin
      style := TsChartBackgroundStyle(Items[i]);
      if (style.Background = newStyle.Background) then
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

