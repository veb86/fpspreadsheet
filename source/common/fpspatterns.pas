unit fpsPatterns;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, Classes, Contnrs, Math,
  fpsTypes, fpsChart;

{ Fill patterns }

type
  TsLineFillPatternMultiplier = (lfpmSingle, lfpmDouble, lfpmTriple);

  TsLineFillPattern = class
    Distance: Single;   // Distance of lines, in mm if > 0, in px if < 0
    Angle: Single;      // Rotation angle of pattern, in degrees
    LineWidth: Single;  // Line width of pattern, in mm if > 0, in px if < 0
    Multiplier: TsLineFillPatternMultiplier;
  end;

  TsDotFillPattern = array[0..7] of Byte;

const
  SOLID_DOT_FILL_PATTERN: TsDotFillPattern = ($FF, $FF, $FF, $FF, $FF, $FF, $FF, $FF);

type
  { TsRawFillPattern
    Combines the same visual pattern as a dot matrix and as vector strokes.
    Vector strokes are supported only by ODS.
    "Raw" patterns do not carry any color - this is added by TsChartFillPattern.
    Used by charts only. }
  TsRawFillPattern = class
  private
    FName: String;
    FExcelName: string;   // style name used in Excel xlsx
    FDotPattern: TsDotFillPattern;
    FLinePattern: TsLineFillPattern;
    procedure SetLinePattern(AValue: TsLineFillPattern);
  public
    constructor Create(AName: String; ADotPattern: TsDotFillPattern;
      ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier);
    constructor Create(AName: String; ADotPattern: TsDotFillPattern);
    constructor Create(AName: String;
      ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier);
    destructor Destroy; override;
    procedure CopyFrom(ASource: TsRawFillPattern); virtual;
    property ExcelName: String read FExcelName;
    property Name: String read FName;
    property DotPattern: TsDotFillPattern read FDotPattern write FDotPattern;
    property LinePattern: TsLineFillPattern read FLinePattern write SetLinePattern;
  end;

  TsRawFillPatternList = class(TFPObjectlist)
  private
    function GetItem(AIndex: Integer): TsRawFillPattern;
    procedure SetItem(AIndex: Integer; AValue: TsRawFillPattern);
  protected
    function AddOrReplace(APattern: TsRawFillPattern): Integer;
    function FindSimilarDotPattern(ALineDistance, ALineAngle, ALineWidth: Single;
      AMultiplier: TsLineFillPatternMultiplier): TsDotFillPattern;
  public
    procedure AddBuiltinPatterns;
    function AddFillPattern(AName, AExcelName: String; ADotPattern: TsDotFillPattern;
      ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier): Integer;
    function AddDotFillPattern(AName, AExcelName: String; APattern: TsDotFillPattern): Integer; overload;
    function AddLineFillPattern(AName, AExcelName: String; ALineDistance, ALineAngle, ALineWidth: Single;
      AMultiplier: TsLineFillPatternMultiplier): Integer;
    function ClonePattern(AIndex: Integer; AName: String): Integer;
    function FindByName(AName: String): TsRawFillPattern;
    function FindLinePatternIndex(ALineDistance, ALineAngle, ALineWidth: Single;
      AMultiplier: TsLineFillPatternMultiplier): Integer;
    function IndexOfName(AName: String): Integer;
    property Items[AIndex: Integer]: TsRawFillPattern read GetItem write SetItem; default;
  end;


  { Line patterns }

  TsChartLengthUnit = (cluMillimeters, cluPercentage);

  TsRawLinePatternElement = record
    Length: Single;      // mm or % of linewidth
    Count: Integer;
  end;

  TsRawLinePattern = class
    Name: String;
    ExcelName: String;
    Element1: TsRawLinePatternElement;
    Element2: TsRawLinePatternElement;
    DistanceLength: Single;       // Space between elements, mm or % of linewidth
    LengthUnit: TsChartLengthUnit;
    constructor Create(AName: String;
      AElement1Length: Single; AElement1Count: Integer;
      ADistanceLength: Single; AUnit: TsChartLengthUnit); overload;
    constructor Create(AName: String;
      AElement1Length: Single; AElement1Count: Integer;
      AElement2Length: Single; AElement2Count: Integer;
      ADistanceLength: Single; AUnit: TsChartLengthUnit); overload;
    procedure CopyFrom(ASource: TsRawLinePattern);
  end;

  TsRawLinePatternList = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsRawLinePattern;
    procedure SetItem(AIndex: Integer; AValue: TsRawLinePattern);
  protected
    function AddOrReplace(APattern: TsRawLinePattern): Integer;
  public
    procedure AddBuiltinPatterns;
    function AddPattern(AName: String;
      AElement1Length: Single; AElement1Count: Integer;
      AElement2Length: Single; AElement2Count: Integer;
      ADistanceLength: Single; ALengthUnit: TsChartLengthUnit): Integer;
    function FindPatternIndex(AElement1Length: Single; AElement1Count: Integer;
      AElement2Length: Single; AElement2Count: Integer;
      ADistanceLength: Single; ALengthUnit: TsChartLengthUnit): Integer;
    function IndexOfName(AName: String): Integer;
    property Items[AIndex: Integer]: TsRawLinePattern read GetItem write SetItem; default;
  end;


{ global fill pattern procedures }

procedure CreateRawFillPatterns;
procedure DestroyRawFillPatterns;

function GetRawFillPattern(APatternIndex: Integer): TsRawFillPattern;
function GetRawFillPatternCount: Integer;
function GetRawFillPatternIndex(ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier): Integer;
function GetRawFillPatternName(APatternStyle: TsChartFillPatternStyle): String;

function RegisterRawFillPattern(AName: String; ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;


{ global line pattern procedures }

procedure CreateRawLinePatterns;
procedure DestroyRawLinePatterns;

function GetRawLinePattern(APatternIndex: Integer): TsRawLinePattern;
function GetRawLinePatternCount: Integer;
function GetRawLinePatternIndex(APatternName: String): Integer;
function GetRawLinePatternIndex(AElement1Length: Single; AElement1Count: Integer;
  AElement2Length: Single; AElement2Count: Integer;
  ADistanceLength: Single; AUnit: TsChartLengthUnit): Integer;
function GetRawLinePatternName(APatternStyle: TsChartLinePatternStyle): String;

function RegisterRawLinePattern(AName: String;
  AElementLength: Single; AElementCount: Integer;
  ADistanceLength: Single; AUnit: TsChartLengthUnit): Integer;
function RegisterRawLinePattern(AName: String;
  AElement1Length: Single; AElement1Count: Integer;
  AElement2Length: Single; AElement2Count: Integer;
  ADistanceLength: Single; AUnit: TsChartLengthUnit): Integer;


implementation

const
  WIDE_DISTANCE = 3.0;
  NARROW_DISTANCE = 1.5;
  THIN_LINE = 0.1;
  THICK_LINE = 0.3;

{$include fpspatterns_dotpatterns.inc}

function StringToDotPattern(APattern: String): TsDotFillPattern;
const
  w = 8;
  h = 8;
var
  i, j, n, b: Integer;
begin
  n := Length(APattern);
  if n > w*h then
    n := w*h;

  FillChar(Result, Sizeof(Result), 0);
  b := 0;
  i := 0;
  for j := 1 to n do
  begin
    if APattern[j] in ['x', 'X'] then
      Result[i] := Result[i] or (1 shl b);
    inc(b);
    if b = w then
    begin
      inc(i);
      b := 0;
    end;
  end;
end;


{ TsRawFillPattern }

constructor TsRawFillPattern.Create(AName: String; ADotPattern: TsDotFillPattern);
begin
  inherited Create;
  FName := AName;
  FDotPattern := ADotPattern;
  FLinePattern := nil;  // no line pattern
end;

constructor TsRawFillPattern.Create(AName: String;
  ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier);
begin
  inherited Create;
  FName := AName;
  FDotPattern := SOLID_DOT_FILL_PATTERN;
  FLinePattern := TsLineFillPattern.Create;
  FLinePattern.Distance := ALineDistance;
  FLinePattern.Angle := ALineAngle;
  FLinePattern.LineWidth := ALineWidth;
  FLinePattern.Multiplier := AMultiplier;
end;

constructor TsRawFillPattern.Create(AName: String; ADotPattern: TsDotFillPattern;
  ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier);
begin
  inherited Create;
  FName := AName;
  FDotPattern := ADotPattern;
  FLinePattern := TsLineFillPattern.Create;
  FLinePattern.Distance := ALineDistance;
  FLinePattern.Angle := ALineAngle;
  FLinePattern.LineWidth := ALineWidth;
  FLinePattern.Multiplier := AMultiplier;
end;

destructor TsRawFillPattern.Destroy;
begin
  FName := '';
  FLinePattern.Free;
  inherited;
end;

procedure TsRawFillPattern.CopyFrom(ASource: TsRawFillPattern);
begin
  FName := ASource.Name;
  FExcelName := ASource.ExcelName;
  FDotPattern := ASource.DotPattern;
  SetLinePattern(ASource.LinePattern);
end;

procedure TsRawFillPattern.SetLinePattern(AValue: TsLineFillPattern);
begin
  if AValue = nil then
  begin
    FLinePattern.Free;
    FLinePattern := nil;
  end else
  begin
    if FLinePattern = nil then
      FLinePattern := TsLineFillPattern.Create;
    FLinePattern.Distance := AValue.Distance;
    FLinePattern.Angle := AValue.Angle;
    FLinePattern.LineWidth := AValue.LineWidth;
    FLinePattern.Multiplier := AValue.Multiplier;
  end;
end;
              (*
{ TsLineFillPatternStyle }

constructor TsLineFillPatternStyle.Create(AName: String;
  ALineDistance, ALineAngle, ALineWidth: Single);
begin
  inherited Create(AName);
  FLineDistance := ALineDistance;
  FLineAngle := ALineAngle;
  FLineWidth := ALineWidth;
end;

procedure TsLineFillPatternStyle.CopyFrom(ASource: TsFillPatternStyle);
begin
  if (ASource is TsLineFillPatternStyle) then
  begin
    FLineDistance := TsLineFillPatternStyle(ASource).LineDistance;
    FLineAngle := TsLineFillPatternStyle(ASource).LineAngle;
    FLineWidth := TsLineFillPatternStyle(ASource).LineWidth;
  end;
  inherited CopyFrom(ASource);
end;

// The pattern itself goes to infinity, limited by size of area to be filled.
function TsLineFillPatternStyle.GetHeight: Integer;
begin
  Result := -1;
end;

function TsLineFillPatternStyle.GetWidth: Integer;
begin
  Result := -1;
end;


{ TsDotFillPatternStyle }

constructor TsDotFillPatternStyle.Create(AName: String; APattern: String);
begin
  inherited Create(AName);
  StringToPattern(APattern);
end;

constructor TsDotFillPatternStyle.Create(AName: String; APattern: TBrushPattern);
begin
  inherited Create(AName);
  FPattern := APattern;
end;

procedure TsDotFillPatternStyle.CopyFrom(ASource: TsFillPatternStyle);
begin
  if ASource is TsDotFillPatternStyle then
    FPattern := TsDotFillPatternStyle(ASource).Pattern;
  inherited CopyFrom(ASource);
end;

function TsDotFillPatternStyle.GetWidth: Integer;
begin
  Result := SizeOf(TPenPattern);
end;

function TsDotFillPatternStyle.GetHeight: Integer;
begin
  Result := SizeOf(TBrushPattern) div SizeOf(TPenPattern);
end;

procedure TsDotFillPatternStyle.StringToPattern(APattern: String);
var
  L: TStrings;
  w, h, i, j, x, y: Integer;
  patt: String;
begin
  L := TStringList.Create;
  try
    L.Text := APattern;
    h := L.Count;
    w := Length(L[0]);
    for i := 1 to L.Count-1 do
      if Length(L[i]) <> w then
        raise Exception.Create('The lines in the pattern string must have the same lengths.');

    FillChar(FPattern, Sizeof(FPattern), 0);
    y := 0;
    i := 0;
    while y < SizeOf(TPenPattern) do
    begin
      if i >= L.Count then
        i := 0;
      patt := L[i];
      x := 0;
      j := 1;
      while x < SizeOf(TPenPattern) do
      begin
        if j > w then
          j := 1;
        if (patt[j] in ['x', 'X']) then
          FPattern[i] := FPattern[i] or (1 shl x);
        inc(x);
        inc(j);
      end;
      inc(y);
      inc(i);
    end;
  finally
    L.Free;
  end;
end;
              *)

{ TsRawFillPatternList }

{ Creates the built-in dot patterns as used by Excel. Not all of them are
  available in ODS. }
procedure TsRawFillPatternList.AddBuiltinPatterns;
var
  i: Integer;
  fps: TsChartFillpatternStyle;
begin
  // IMPORTANT: Add all predefined fill patterns in correct order as defined
  // by the type TsChartfillPatternStyle.
  AddDotFillPattern(GetRawFillPatternName(fpsGray05), 'pct5',                   // 0
    StringToDotPattern(GRAY05_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray10), 'pct10',
    StringToDotPattern(GRAY20_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray20), 'pct20',
    StringToDotPattern(GRAY20_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray25), 'pct25',
    StringToDotPattern(GRAY25_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray30), 'pct30',
    StringToDotPattern(GRAY30_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray40), 'pct40',                  // 5
    StringToDotPattern(GRAY40_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray50), 'pct50',
    StringToDotPattern(GRAY50_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray60), 'pct60',
    StringToDotPattern(GRAY60_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray70), 'pct70',
    StringToDotPattern(GRAY70_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray75), 'pct75',
    StringToDotPattern(GRAY75_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray80), 'pct80',                  // 10
    StringToDotPattern(GRAY80_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsGray90), 'pct90',
    StringToDotPattern(GRAY90_PATTERN));

  AddFillPattern(GetRawFillPatternName(fpsHorThick), 'dkHorz',                  // 12
    StringToDotPattern(HOR_PATTERN_THICK),
    WIDE_DISTANCE, 0.0, THICK_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsVertThick), 'dkVert',                 // 13
    StringToDotPattern(VERT_PATTERN_THICK),
    WIDE_DISTANCE, 90.0, THICK_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsDiagUpThick), 'dkUpDiag',             // 14
    StringToDotPattern(DIAG_UP_PATTERN_THICK),
    WIDE_DISTANCE, 45.0, THICK_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsDiagDownThick), 'dkDnDiag',           // 15
    StringToDotPattern(DIAG_DOWN_PATTERN_THICK),
    WIDE_DISTANCE, -45.0, THICK_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsHatchThick), 'openDmnd',              // 16  -- replacement
    StringToDotPattern(HATCH_PATTERN_THICK),
    WIDE_DISTANCE, 45.0, THICK_LINE, lfpmDouble);
  AddFillPattern(GetRawFillPatternName(fpsCrossThick), 'smGrid',                // 17  -- replacement
    StringToDotPattern(CROSS_PATTERN_THICK),
    WIDE_DISTANCE, 0.0, THICK_LINE, lfpmDouble);

  AddFillPattern(GetRawFillPatternName(fpsHorThin), 'ltHorz',                   // 18
    StringToDotPattern(HOR_PATTERN_THIN),
    WIDE_DISTANCE, 0.0, THIN_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsVertThin), 'ltVert',                  // 19
    StringToDotPattern(VERT_PATTERN_THIN),
    WIDE_DISTANCE, 90.0, THIN_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsDiagUpThin), 'wdUpDiag',              // 20
    StringToDotPattern(DIAG_UP_PATTERN_THIN),
    WIDE_DISTANCE, 45.0, THIN_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsDiagDownThin), 'wdDnDiag',            // 21
    StringToDotPattern(DIAG_DOWN_PATTERN_THIN),
    WIDE_DISTANCE, -45.0, THIN_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsHatchThin), 'openDmnd',               // 22
    StringToDotPattern(HATCH_PATTERN_THIN),
    WIDE_DISTANCE, 45.0, THIN_LINE, lfpmDouble);
  AddFillPattern(GetRawFillPatternName(fpsCrossThin), 'lgGrid',                 // 23
    StringToDotPattern(CROSS_PATTERN_THIN),
    WIDE_DISTANCE, 0.0, THIN_LINE, lfpmDouble);

  AddFillPattern(GetRawFillPatternName(fpsHorNarrow), 'narHorz',                // 24
    StringToDotPattern(HOR_PATTERN_NARROW),
    NARROW_DISTANCE, 0.0, THIN_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsVertNarrow), 'narVert',               // 25
    StringToDotPattern(VERT_PATTERN_NARROW),
    NARROW_DISTANCE, 90.0, THIN_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsDiagUpNarrow), 'ltUpDiag',            // 26
    StringToDotPattern(DIAG_UP_PATTERN_NARROW),
    NARROW_DISTANCE, 45.0, THIN_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsDiagDownNarrow), 'ltDnDiag',          // 27
    StringToDotPattern(DIAG_DOWN_PATTERN_NARROW),
    NARROW_DISTANCE, -45.0, THIN_LINE, lfpmSingle);
  AddFillPattern(GetRawFillPatternName(fpsHatchNarrow), 'openDmnd',             // 28   -- replacement
    StringToDotPattern(HATCH_PATTERN_NARROW),
    NARROW_DISTANCE, 45.0, THIN_LINE, lfpmDouble);
  AddFillPattern(GetRawFillPatternName(fpsCrossNarrow), 'smGrid',               // 29
    StringToDotPattern(CROSS_PATTERN_NARROW),
    NARROW_DISTANCE, 0.0, THIN_LINE, lfpmDouble);

  AddDotFillPattern(GetRawFillPatternName(fpsHorDash), 'dashHorz',              // 30
    StringToDotPattern(HOR_DASH_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsVertDash), 'dashVert',             // 31
    StringToDotPattern(VERT_DASH_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsDiagUpDash), 'dashUpDiag',         // 32
    StringToDotPattern(DIAG_UP_DASH_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsDiagDownDash), 'dashDnDiag',       // 33
    StringToDotPattern(DIAG_DOWN_DASH_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsCrossDot), 'dotGrid',              // 34
    StringToDotPattern(CROSS_PATTERN_DOT));
  AddDotFillPattern(GetRawFillPatternName(fpsHatchDot), 'dotDmnd',              // 35
    StringToDotPattern(HATCH_PATTERN_DOT));

  AddDotFillPattern(GetRawFillPatternName(fpsBrickHor), 'horzBrick',            // 36
    StringToDotPattern(BRICK_HOR_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsBrickDiag), 'diagBrick',           // 37
    StringToDotPattern(BRICK_DIAG_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsCheckerboardLarge), 'lgCheck',     // 38
    StringToDotPattern(CHECKERBOARD_LARGE_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsCheckerboardSmall), 'smCheck',     // 39
    StringToDotPattern(CHECKERBOARD_SMALL_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsConfettiLarge), 'lgConfetti',      // 40
    StringToDotPattern(CONFETTI_LARGE_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsConfettiSmall), 'smConfetti',      // 41
    StringToDotPattern(CONFETTI_SMALL_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsDiamond), 'solidDmnd',             // 42
    StringToDotPattern(DIAMOND_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsDivot), 'divot',                   // 43
    StringToDotPattern(DIVOT_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsPlaid), 'plaid',                   // 44
    StringToDotPattern(PLAID_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsShingle), 'shingle',               // 45
    StringToDotPattern(SHINGLE_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsSphere), 'sphere',                 // 46
    StringToDotPattern(SPHERE_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsTrellis), 'trellis',               // 47
    StringToDotPattern(TRELLIS_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsWave), 'wave',                     // 48
    StringToDotPattern(WAVE_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsWeave), 'weave',                   // 49
    StringToDotPattern(WEAVE_PATTERN));
  AddDotFillPattern(GetRawFillPatternName(fpsZigZag), 'zigZag',                 // 50
    StringToDotPattern(ZIGZAG_PATTERN));

  {
  // Check for completeness
  for i := 0 to Count - 1 do
    with Items[i] do
      WriteLn(
        'TsRawFillPatternList.AddBuiltinPatterns] i=', i,
        ' style=', TsChartFillPatternStyle(i),
        ' Name=', Name,
        ' ExcelName=', ExcelName
      );
  }
end;

function TsRawFillPatternList.AddDotFillPattern(AName, AExcelName: String;
  APattern: TsDotFillPattern): Integer;
var
  patt: TsRawFillPattern;
begin
  patt := TsRawFillPattern.Create(AName, APattern);
  patt.FExcelName := AExcelName;
  Result := AddOrReplace(patt);
end;

function TsRawFillPatternList.AddFillPattern(AName, AExcelName: String;
  ADotPattern: TsDotFillPattern; ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
var
  patt: TsRawFillPattern;
begin
  patt := TsRawFillPattern.Create(AName,
    ADotPattern,
    ALineDistance, ALineAngle, ALineWidth, AMultiplier
  );
  patt.FExcelName := AExcelName;
  Result := AddOrReplace(patt);
end;

function TsRawFillPatternList.AddLineFillPattern(AName, AExcelName: String;
  ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
var
  patt: TsRawFillPattern;
begin
  patt := TsRawFillPattern.Create(AName, ALineDistance, ALineAngle, ALineWidth, AMultiplier);
  patt.DotPattern := FindSimilarDotPattern(ALineDistance, ALineAngle, ALineWidth, AMultiplier);
  patt.FExcelName := AExcelName;
  Result := AddOrReplace(patt);
end;

function TsRawFillPatternList.AddOrReplace(APattern: TsRawFillPattern): Integer;
var
  idx: Integer;
begin
  idx := IndexOfName(APattern.Name);
  if idx = -1 then
    Result := Add(APattern)
  else
  begin
    Items[idx].CopyFrom(APattern);
//    Insert(idx, APattern);
    Result := idx;
  end;
end;

{ Adds a copy of the pattern at the given index and gives it a new name. }
function TsRawFillPatternList.ClonePattern(AIndex: Integer; AName: String): Integer;
var
  patt: TsRawFillPattern;
begin
  patt := Items[AIndex];
  Result := AddFillPattern(AName, '',
    patt.DotPattern,
    patt.LinePattern.Distance, patt.LinePattern.Angle, patt.LinePattern.LineWidth, patt.LinePattern.Multiplier
  );
end;

function TsRawFillPatternList.FindByName(AName: String): TsRawFillPattern;
var
  i: Integer;
begin
  for i := 0 to Count-1 do
    if Items[i].Name = AName then
    begin
      Result := Items[i];
      exit;
    end;
  Result := nil;
end;

function TsRawFillPatternList.FindLinePatternIndex(
  ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
var
  i: Integer;
  patt: TsRawFillPattern;
begin
  for i := 0 to Count-1 do
  begin
    patt := Items[i];
    if patt.LinePattern = nil then
      Continue;
    if SameValue(ALineDistance, patt.LinePattern.Distance, 0.1) and
       SameValue(ALineAngle, patt.LinePattern.Angle, 0.1) and
       SameValue(ALineWidth, patt.LinePattern.LineWidth, 0.1) and
       (AMultiplier = patt.LinePattern.Multiplier) then
    begin
      Result := i;
      exit;
    end;
  end;
  Result := -1;
end;

function TsRawFillPatternList.FindSimilarDotPattern(ALineDistance, ALineAngle,
  ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier): TsDotFillPattern;
var
  sinAngle, cosAngle: Single;
  similarIdx: Integer;
  singleLine: Boolean;
begin
  similarIdx := -1;
  SinCos(DegToRad(ALineAngle), sinAngle, cosAngle);
  singleLine := (AMultiplier = lfpmSingle);
  if SameValue(sinAngle, 0.0, 0.2) then  // about +/- 12°
  begin
    if (ALineDistance >= WIDE_DISTANCE) or (ALineDistance <= -8) then  // wider than 3mm or 8px
      similarIdx := IfThen(singleLine, ord(fpsHorThin), ord(fpsCrossThin))
    else
      similarIdx := IfThen(singleLine, ord(fpsHorNarrow), ord(fpsCrossNarrow));
  end else
  if SameValue(cosAngle, 0.0, 0.2) then  // about 90 +/- 12°
  begin
    if (ALineDistance >= WIDE_DISTANCE) or (ALineDistance <= -8) then
      similarIdx := IfThen(singleLine, ord(fpsVertThin), ord(fpsCrossThin))
    else
      similarIdx := IfThen(singleLine, ord(fpsVertNarrow), ord(fpsCrossNarrow));
  end else
  if ALineAngle > 0 then
  begin
    if (ALineDistance >= WIDE_DISTANCE) or (ALineDistance <= -8) then
      similarIdx := IfThen(singleLine, ord(fpsDiagUpThin), ord(fpsHatchThin))
    else
      similarIdx := IfThen(singleLine, ord(fpsDiagUpNarrow), ord(fpsHatchNarrow));
  end else
  begin
    if (ALineDistance >= WIDE_DISTANCE) or (ALineDistance <= -8) then
      similarIdx := IfThen(singleLine, ord(fpsDiagDownThin), ord(fpsHatchThin))
    else
      similarIdx := IfThen(singleLine, ord(fpsDiagDownNarrow), ord(fpsHatchNarrow));
  end;
  if similarIdx > -1 then
    Result := Items[similarIdx].DotPattern
  else
    Result := SOLID_DOT_FILL_PATTERN;
end;

function TsRawFillPatternList.IndexOfName(AName: String): Integer;
begin
  for Result := 0 to Count-1 do
    if Items[Result].Name = AName then
      exit;
  Result := -1;
end;

function TsRawFillPatternList.GetItem(AIndex: Integer): TsRawFillPattern;
begin
  Result := TsRawFillPattern(inherited Items[AIndex]);
end;

procedure TsRawFillPatternList.SetItem(AIndex: Integer; AValue: TsRawFillPattern);
begin
  TsRawFillPattern(inherited Items[AIndex]).CopyFrom(AValue);
end;

{ ------------------------------------------------------------------------------
                        global fill pattern procedures
-------------------------------------------------------------------------------}
var
  RawFillPatterns: TsRawFillPatternList = nil;
  RawFillPatterns_ReferenceCounter: Integer = 0;

procedure CreateRawFillPatterns;
begin
  if RawFillPatterns_ReferenceCounter = 0 then
  begin
    RawFillPatterns := TsRawFillPatternList.Create;
    RawFillPatterns.AddBuiltinPatterns;
  end;
  inc(RawFillPatterns_ReferenceCounter);
end;

procedure DestroyRawFillPatterns;
begin
  dec(RawFillPatterns_ReferenceCounter);
  if RawFillPatterns_ReferenceCounter <= 0 then
  begin
    FreeAndNil(RawfillPatterns);
    RawFillPatterns_ReferenceCounter := 0;
  end;
end;

function GetRawFillPattern(APatternIndex: Integer): TsRawFillPattern;
begin
  if Assigned(RawFillPatterns) then
    Result := RawFillPatterns.Items[APatternIndex]
  else
    Result := nil;
end;

{ Finds the index of the line fill pattern having the specified parameters.
  Returns -1 if not found. }
function GetRawFillPatternIndex(ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
begin
  if Assigned(RawFillPatterns) then
    Result := RawFillPatterns.FindLinePatternIndex(ALineDistance, ALineAngle, ALineWidth, AMultiplier)
  else
    Result := -1;
end;

function GetRawFillPatternName(APatternStyle: TsChartFillPatternStyle): String;
const
  PatternName: Array[TsChartFillPatternStyle] of string = (
    'GRAY05', 'GRAY10', 'GRAY20', 'GRAY25', 'GRAY30', 'GRAY40',
    'GRAY50', 'GRAY60', 'GRAY70', 'GRAY75', 'GRAY80', 'GRAY90',
    'HOR_THICK', 'VERT_THICK', 'DIAG_UP_THICK', 'DIAG_DOWN_THICK', 'HATCH_THICK', 'CROSS_THICK',
    'HOR_THIN', 'VERT_THIN', 'DIAG_UP_THIN', 'DIAG_DOWN_THIN', 'HATCH_THIN', 'CROSS_THIN',
    'HOR_NARROW', 'VERT_NARROW', 'DIAG_UP_NARROW', 'DIAG_DOWN_NARROW', 'HATCH_NARROW', 'CROSS_NARROW',
    'HOR_DASH', 'VERT_DASH', 'DIAG_UP_DASH', 'DIAG_DOWN_DASH', 'CROSS_DOT', 'HATCH_DOT',
    'BRICK_HOR', 'BRICK_DIAG', 'CHECKERBOARD_LARGE', 'CHECKERBOARD_SMALL', 'CONFETTI_LARGE', 'CONFETTI_SMALL',
    'DIAMOND', 'DIVOT', 'PLAID', 'SHINGLE', 'SPHERE', 'TRELLIS', 'WAVE', 'WEAVE', 'ZIGZAG'
  );
begin
  Result := PatternName[APatternStyle];
end;

function GetRawFillPatternCount: Integer;
begin
  if Assigned(RawFillPatterns) then
    Result := RawFillPatterns.Count
  else
    Result := 0;
end;

function RegisterRawFillPattern(AName: String; ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
begin
  if not Assigned(RawFillPatterns) then
    CreateRawFillPatterns;
  Result := RawFillPatterns.AddLineFillPattern(AName, '', ALineDistance, ALineAngle, ALineWidth, AMultiplier)
end;


{===============================================================================
                             Line patterns
===============================================================================}

constructor TsRawLinePattern.Create(AName: String;
  AElement1Length: Single; AElement1Count: Integer;
  ADistanceLength: Single; AUnit: TsChartLengthUnit);
begin
  inherited Create;
  Name := AName;
  Element1.Length := AElement1Length;
  Element1.Count := AElement1Count;
  Element2.Length := 0;
  Element2.Count := 0;
  DistanceLength := ADistanceLength;
  LengthUnit := AUnit;
end;

constructor TsRawLinePattern.Create(AName: String;
  AElement1Length: Single; AElement1Count: Integer;
  AElement2Length: Single; AElement2Count: Integer;
  ADistanceLength: Single; AUnit: TsChartLengthUnit);
begin
  inherited Create;
  Name := AName;
  Element1.Length := AElement1Length;
  Element1.Count := AElement1Count;
  Element2.Length := AElement2Length;
  Element2.Count := AElement2Count;
  DistanceLength := ADistanceLength;
  LengthUnit := AUnit;
end;

procedure TsRawLinePattern.CopyFrom(ASource: TsRawLinePattern);
begin
  Name := ASource.Name;
  Element1 := ASource.Element1;
  Element2 := ASource.Element2;
  DistanceLength := ASource.DistanceLength;
  LengthUnit := ASource.LengthUnit;
end;


{ TsRawLinePatternList }

procedure TsRawLinePatternList.AddBuiltinPatterns;
begin
  // solid line
  AddPattern(GetRawLinePatternName(clsSolid), 10, 1, 0, 0, 0, cluMillimeters);
  // no line
  AddPattern(GetRawLinePatternName(clsNoLine), 0, 0, 0, 0, 0, cluMillimeters);
  // fine dots
  AddPattern(GetRawLinePatternName(clsFineDot), 120, 1, 0, 0, 120, cluPercentage);
  // dotted
  AddPattern(GetRawLinePatternName(clsDot), 120, 1, 0, 0, 500, cluPercentage);
  // dashed  (- - - - )
  AddPattern(GetRawLinePatternName(clsDash), 800, 1, 0, 0, 600, cluPercentage);
  // dash-dot  (- . - . - )
  AddPattern(GetRawLinePatternName(clsDashDot), 800, 1, 120, 1, 600, cluPercentage);
  // long dash  (--  --  --)
  AddPattern(GetRawLinePatternName(clsLongDash), 2400, 1, 0, 0, 600, cluPercentage);
  // long dash-dot  (-- . -- . -- . )
  AddPattern(GetRawLinePatternName(clsLongDashDot), 2400, 1, 120, 1, 600, cluPercentage);
  // long dash-dot-dot  (-- . . -- . . )
  AddPattern(GetRawLinePatternName(clsLongDashDotDot), 2400, 1, 120, 2, 600, cluPercentage);
end;

function TsRawLinePatternList.AddOrReplace(APattern: TsRawLinePattern): Integer;
var
  idx: Integer;
begin
  idx := IndexOfName(APattern.Name);
  if idx = -1 then
    Result := Add(APattern)
  else
  begin
    Items[idx].CopyFrom(APattern);
//    Insert(idx, APattern);
    Result := idx;
  end;
end;

function TsRawLinePatternList.AddPattern(AName: String;
  AElement1Length: Single; AElement1Count: Integer;
  AElement2Length: Single; AElement2Count: Integer;
  ADistanceLength: Single; ALengthUnit: TsChartLengthUnit): Integer;
var
  patt: TsRawLinePattern;
begin
  patt := TsRawLinePattern.Create(AName,
    AElement1Length, AElement1Count,
    AElement2Length, AElement2Count,
    ADistanceLength, ALengthUnit
  );
//  patt.FExcelName := AExcelName;
  Result := AddOrReplace(patt);
end;

function TsRawLinePatternList.FindPatternIndex(
  AElement1Length: Single; AElement1Count: Integer;
  AElement2Length: Single; AElement2Count: Integer;
  ADistanceLength: Single; ALengthUnit: TsChartlengthUnit): Integer;
var
  i: Integer;
  patt: TsRawLinePattern;
begin
  for i := 0 to Count-1 do
  begin
    patt := Items[i];
    if SameValue(AElement1Length, patt.Element1.Length, 0.1) and
       (AElement1Count = patt.Element1.Count) and
       SameValue(AElement2Length, patt.Element2.Length, 0.1) and
       (AElement2Count = patt.Element2.Count) and
       SameValue(ADistanceLength, patt.DistanceLength, 0.1) and
       (ALengthUnit = patt.LengthUnit) then
    begin
      Result := i;
      exit;
    end;
  end;
  Result := -1;
end;


function TsRawLinePatternList.GetItem(AIndex: Integer): TsRawLinePattern;
begin
  Result := TsRawLinePattern(inherited Items[AIndex]);
end;

function TsRawLinePatternList.IndexOfName(AName: String): Integer;
begin
  for Result := 0 to Count-1 do
    if Items[Result].Name = AName then
      exit;
  Result := -1;
end;

procedure TsRawLinePatternList.SetItem(AIndex: Integer; AValue: TsRawLinePattern);
begin
  TsRawLinePattern(inherited Items[AIndex]).CopyFrom(AValue);
end;


{-------------------------------------------------------------------------------
                        global line pattern procedures
-------------------------------------------------------------------------------}
var
  RawLinePatterns: TsRawLinePatternList = nil;
  RawLinePatterns_ReferenceCounter: Integer = 0;

procedure CreateRawLinePatterns;
begin
  if RawLinePatterns_ReferenceCounter = 0 then
  begin
    RawLinePatterns := TsRawLinePatternList.Create;
    RawLinePatterns.AddBuiltinPatterns;
  end;
  inc(RawLinePatterns_ReferenceCounter);
end;

procedure DestroyRawLinePatterns;
begin
  dec(RawLinePatterns_ReferenceCounter);
  if RawLinePatterns_ReferenceCounter <= 0 then
  begin
    FreeAndNil(RawLinePatterns);
    RawLinePatterns_ReferenceCounter := 0;
  end;
end;

function GetRawLinePattern(APatternIndex: Integer): TsRawLinePattern;
begin
  if Assigned(RawLinePatterns) then
    Result := RawLinePatterns.Items[APatternIndex]
  else
    Result := nil;
end;

function GetRawLinePatternCount: Integer;
begin
  if Assigned(RawLinePatterns) then
    Result := RawLinePatterns.Count
  else
    Result := 0;
end;

{@@ Returns the index of the line pattern with the given pattern name. }
function GetRawLinePatternIndex(APatternName: String): Integer;
begin
  if Assigned(RawLinePatterns) then
    Result := RawLinePatterns.IndexOfName(APatternName)
  else
    Result := -1;
end;

{@@ Finds the index of the line pattern having the specified parameters.
  Returns -1 if not found. }
function GetRawLinePatternIndex(AElement1Length: Single; AElement1Count: Integer;
  AElement2Length: Single; AElement2Count: Integer;
  ADistanceLength: Single; AUnit: TsChartLengthUnit): Integer;
begin
  if Assigned(RawLinePatterns) then
    Result := RawLinePatterns.FindPatternIndex(
      AElement1Length, AElement1Count,
      AElement2Length, AElement2Count,
      ADistanceLength, AUnit
    )
  else
    Result := -1;
end;

function GetRawLinePatternName(APatternStyle: TsChartLinePatternStyle): String;
const
  PatternName: Array[TsChartLinePatternStyle] of string = (
    'SOLID', 'NO_LINE', 'FINE_DOT', 'DOT', 'DASH', 'DASH_DOT',
    'LONG_DASH', 'LONG_DASH_DOT', 'LONG_DASH_DOT_DOT', 'CUSTOM'
  );
begin
  Result := PatternName[APatternStyle];
end;

function RegisterRawLinePattern(AName: String; AElementLength: Single; AElementCount: Integer;
  ADistanceLength: Single; AUnit: TsChartLengthUnit): Integer;
begin
  Result := RegisterRawLinePattern(AName, AElementLength, AElementCount, 0, 0, ADistanceLength, AUnit);
end;

function RegisterRawLinePattern(AName: String;
  AElement1Length: Single; AElement1Count: Integer;
  AElement2Length: Single; AElement2Count: Integer;
  ADistanceLength: Single; AUnit: TsChartLengthUnit): Integer;
begin
  if not Assigned(RawLinePatterns) then
    CreateRawLinePatterns;
  Result := RawLinePatterns.AddPattern(AName, AElement1Length, AElement1Count, AElement2Length, AElement2Count, ADistanceLength, AUnit);
end;


finalization
  FreeAndNil(RawFillPatterns);
  FreeAndNil(RawLinePatterns);

end.

