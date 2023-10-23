unit fpsChartStyles;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, fpsTypes, fpsChart;

type
  TsChartStyle = class
  public
    procedure ApplyToChart(AChart: TsChart); virtual; abstract;
    procedure ExtractFromChart(AChart: TsChart); virtual; abstract;
  end;

  TsChartBackgroundStyle = class(TsChartStyle)
  private
    FBackground: TsChartFill;
    FBorder: TsChartLine;
  public
    procedure ApplyToChart(AChart: TsChart); override;
    procedure ExtractFromChart(AChart: TsChart); override;
    property Background: TsChartFill read FBackground;
    property Border: TsChartLine read FBorder;
  end;

  TsChartStyleList = class(TFPList)
  public
    destructor Destroy; override;
    procedure Clear;
    function FindChartBackgroundStyle(AChart: TsChart): Integer;
  end;

implementation

{ TsChartBackgroundstyle }

procedure TsChartBackgroundStyle.ApplyToChart(AChart: TsChart);
begin
  AChart.Background := FBackground;
  AChart.Border := FBorder;
end;

procedure TsChartBackgroundStyle.ExtractFromChart(AChart: TsChart);
begin
  FBackground := AChart.Background;
  FBorder := AChart.Border;
end;

{ TsChartStyleList }

destructor TsChartStyleList.Destroy;
begin
  Clear;
  inherited;
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
  Returns the index of the style. }
function TsChartStyleList.FindChartBackgroundStyle(AChart: TsChart): Integer;
var
  newStyle, style: TsChartBackgroundStyle;
  i: Integer;
begin
  Result := -1;
  newStyle := TsChartBackgroundStyle.Create;
  newStyle.ExtractFromChart(AChart);
  for i := 0 to Count-1 do
  begin
    if (TsChartStyle(Items[i]) is TsChartBackgroundStyle) then
    begin
       style := TsChartBackgroundStyle(Items[i]);
       if style.FBackground = newStyle.FBackground then
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

