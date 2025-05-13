unit fpsPatterns;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, Classes, Contnrs, Math,
  fpsTypes, fpsChart;

const
  FPS_GRAY06 = 'GRAY06';
  FPS_GRAY12 = 'GRAY12';
  FPS_GRAY20 = 'GRAY20';
  FPS_GRAY25 = 'GRAY25';
  FPS_GRAY30 = 'GRAY30';
  FPS_GRAY40 = 'GRAY40';
  FPS_GRAY50 = 'GRAY50';
  FPS_GRAY60 = 'GRAY60';
  FPS_GRAY70 = 'GRAY70';
  FPS_GRAY75 = 'GRAY75';
  FPS_GRAY80 = 'GRAY80';
  FPS_GRAY90 = 'GRAY90';

  FPS_HOR_THICK = 'HOR_THICK';
  FPS_VERT_THICK = 'VERT_THICK';
  FPS_DIAG_UP_THICK = 'DIAG_UP_THICK';
  FPS_DIAG_DOWN_THICK = 'DIAG_DOWN_THICK';
  FPS_HATCH_THICK = 'HATCH_THICK';
  FPS_CROSS_THICK = 'CROSS_THICK';

  FPS_HOR_THIN = 'HOR_THIN';
  FPS_VERT_THIN = 'VERT_THIN';
  FPS_DIAG_UP_THIN = 'DIAG_UP_THIN';
  FPS_DIAG_DOWN_THIN = 'DIAG_DOWN_THIN';
  FPS_HATCH_THIN = 'HATCH_THIN';
  FPS_CROSS_THIN = 'CROSS_THIN';

  FPS_HOR_NARROW = 'HOR_NARROW';
  FPS_VERT_NARROW = 'VERT_NARROW';
  FPS_DIAG_UP_NARROW = 'DIAG_UP_NARROW';
  FPS_DIAG_DOWN_NARROW = 'DIAG_DOWN_NARROW';
  FPS_HATCH_NARROW = 'HATCH_NARROW';
  FPS_CROSS_NARROW = 'CROSS_NARROW';

  FPS_HOR_DASH = 'HOR_DASH';
  FPS_VERT_DASH = 'VERT_DASH';
  FPS_DIAG_UP_DASH = 'DIAG_UP_DASH';
  FPS_DIAG_DOWN_DASH = 'DIAG_DOWN_DASH';
  FPS_CROSS_DOT = 'CROSS_DOT';
  FPS_HATCH_DOT = 'HATCH_DOT';

  FPS_BRICK_HOR = 'BRICK_HOR';
  FPS_BRICK_DIAG = 'BRICK_DIAG';
  FPS_CHECKERBOARD_LARGE = 'CHECKERBOARD_LARGE';
  FPS_CHECKERBOARD_SMALL = 'CHECKERBOARD_SMALL';
  FPS_CONFETTI_LARGE = 'CONFETTI_LARGE';
  FPS_CONFETTI_SMALL = 'CONFETTI_SMALL';
  FPS_DIAMOND = 'DIAMOND';
  FPS_DIVOT = 'DIVOT';
  FPS_PLAID = 'PLAID';
  FPS_SHINGLE = 'SHINGLE';
  FPS_SPHERE = 'SPHERE';
  FPS_TRELLIS = 'TRELLIS';
  FPS_WAVE = 'WAVE';
  FPS_WEAVE = 'WEAVE';
  FPS_ZIGZAG = 'ZIGZAG';

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
    function AddFillPattern(AName: String; ADotPattern: TsDotFillPattern;
      ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier): Integer;
    function AddDotFillPattern(AName: String; APattern: TsDotFillPattern): Integer; overload;
    function AddLineFillPattern(AName: String; ALineDistance, ALineAngle, ALineWidth: Single;
      AMultiplier: TsLineFillPatternMultiplier): Integer;
    function ClonePattern(AIndex: Integer; AName: String): Integer;
    function FindByName(AName: String): TsRawFillPattern;
    function FindLinePatternIndex(ALineDistance, ALineAngle, ALineWidth: Single;
      AMultiplier: TsLineFillPatternMultiplier): Integer;
    function IndexOfName(AName: String): Integer;
    property Items[AIndex: Integer]: TsRawFillPattern read GetItem write SetItem; default;
  end;


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
begin
  fpsGray06 := AddDotFillPattern(FPS_GRAY06,
    StringToDotPattern(GRAY06_PATTERN));
  fpsGray12 := AddDotFillPattern(FPS_GRAY12,
    StringToDotPattern(GRAY12_PATTERN));
  fpsGray20 := AddDotFillPattern(FPS_GRAY20,
    StringToDotPattern(GRAY20_PATTERN));
  fpsGray25 := AddDotFillPattern(FPS_GRAY25,
    StringToDotPattern(GRAY25_PATTERN));
  fpsGray30 := AddDotFillPattern(FPS_GRAY30,
    StringToDotPattern(GRAY30_PATTERN));
  fpsGray40 := AddDotFillPattern(FPS_GRAY40,
    StringToDotPattern(GRAY40_PATTERN));
  fpsGray50 := AddDotFillPattern(FPS_GRAY50,
    StringToDotPattern(GRAY50_PATTERN));
  fpsGray60 := AddDotFillPattern(FPS_GRAY60,
    StringToDotPattern(GRAY60_PATTERN));
  fpsGray75 := AddDotFillPattern(FPS_GRAY75,
    StringToDotPattern(GRAY75_PATTERN));
  fpsGray80 := AddDotFillPattern(FPS_GRAY80,
    StringToDotPattern(GRAY80_PATTERN));
  fpsGray90 := AddDotFillPattern(FPS_GRAY90,
    StringToDotPattern(GRAY90_PATTERN));

  fpsHorThick := AddFillPattern(FPS_HOR_THICK,
    StringToDotPattern(HOR_PATTERN_THICK),
    WIDE_DISTANCE, 0.0, THICK_LINE, lfpmSingle);
  fpsVertThick := AddFillPattern(FPS_VERT_THICK,
    StringToDotPattern(VERT_PATTERN_THICK),
    WIDE_DISTANCE, 90.0, THICK_LINE, lfpmSingle);
  fpsDiagUpThick := AddFillPattern(FPS_DIAG_UP_THICK,
    StringToDotPattern(DIAG_UP_PATTERN_THICK),
    WIDE_DISTANCE, 45.0, THICK_LINE, lfpmSingle);
  fpsDiagDownThick := AddFillPattern(FPS_DIAG_DOWN_THICK,
    StringToDotPattern(DIAG_DOWN_PATTERN_THICK),
    WIDE_DISTANCE, -45.0, THICK_LINE, lfpmSingle);
  fpsHatchThick := AddFillPattern(FPS_HATCH_THICK,
    StringToDotPattern(HATCH_PATTERN_THICK),
    WIDE_DISTANCE, 45.0, THICK_LINE, lfpmDouble);
  fpsCrossThick := AddFillPattern(FPS_CROSS_THICK,
    StringToDotPattern(CROSS_PATTERN_THICK),
    WIDE_DISTANCE, 0.0, THICK_LINE, lfpmDouble);

  fpsHorThin := AddFillPattern(FPS_HOR_THIN,
    StringToDotPattern(HOR_PATTERN_THIN),
    WIDE_DISTANCE, 0.0, THIN_LINE, lfpmSingle);
  fpsVertThin := AddFillPattern(FPS_VERT_THIN,
    StringToDotPattern(VERT_PATTERN_THIN),
    WIDE_DISTANCE, 90.0, THIN_LINE, lfpmSingle);
  fpsDiagUpThin := AddFillPattern(FPS_DIAG_UP_THIN,
    StringToDotPattern(DIAG_UP_PATTERN_THIN),
    WIDE_DISTANCE, 45.0, THIN_LINE, lfpmSingle);
  fpsDiagDownThin := AddFillPattern(FPS_DIAG_DOWN_THIN,
    StringToDotPattern(DIAG_DOWN_PATTERN_THIN),
    WIDE_DISTANCE, -45.0, THIN_LINE, lfpmSingle);
  fpsHatchThin := AddFillPattern(FPS_HATCH_THIN,
    StringToDotPattern(HATCH_PATTERN_THIN),
    WIDE_DISTANCE, 45.0, THIN_LINE, lfpmDouble);
  fpsCrossThin := AddFillPattern(FPS_CROSS_THIN,
    StringToDotPattern(CROSS_PATTERN_THIN),
    WIDE_DISTANCE, 0.0, THIN_LINE, lfpmDouble);

  fpsHorNarrow := AddFillPattern(FPS_HOR_NARROW,
    StringToDotPattern(HOR_PATTERN_NARROW),
    NARROW_DISTANCE, 0.0, THIN_LINE, lfpmSingle);
  fpsVertNarrow := AddFillPattern(FPS_VERT_NARROW,
    StringToDotPattern(VERT_PATTERN_NARROW),
    NARROW_DISTANCE, 90.0, THIN_LINE, lfpmSingle);
  fpsDiagUpNarrow := AddFillPattern(FPS_DIAG_UP_NARROW,
    StringToDotPattern(DIAG_UP_PATTERN_NARROW),
    NARROW_DISTANCE, 45.0, THIN_LINE, lfpmSingle);
  fpsDiagDownNarrow := AddFillPattern(FPS_DIAG_DOWN_NARROW,
    StringToDotPattern(DIAG_DOWN_PATTERN_NARROW),
    NARROW_DISTANCE, -45.0, THIN_LINE, lfpmSingle);
  fpsHatchNarrow := AddFillPattern(FPS_HATCH_NARROW,
    StringToDotPattern(HATCH_PATTERN_NARROW),
    NARROW_DISTANCE, 45.0, THIN_LINE, lfpmDouble);
  fpsCrossNarrow := AddFillPattern(FPS_CROSS_NARROW,
    StringToDotPattern(CROSS_PATTERN_NARROW),
    NARROW_DISTANCE, 0.0, THIN_LINE, lfpmDouble);

  fpsHorDash := AddDotFillPattern(FPS_HOR_DASH,
    StringToDotPattern(HOR_DASH_PATTERN));
  fpsVertDash := AddDotFillPattern(FPS_VERT_DASH,
    StringToDotPattern(VERT_DASH_PATTERN));
  fpsDiagUpDash := AddDotFillPattern(FPS_DIAG_UP_DASH,
    StringToDotPattern(DIAG_UP_DASH_PATTERN));
  fpsDiagDownDash := AddDotFillPattern(FPS_DIAG_DOWN_DASH,
    StringToDotPattern(DIAG_DOWN_DASH_PATTERN));
  fpsCrossDot := AddDotFillPattern(FPS_CROSS_DOT,
    StringToDotPattern(CROSS_PATTERN_DOT));
  fpsHatchDot := AddDotFillPattern(FPS_HATCH_DOT,
    StringToDotPattern(HATCH_PATTERN_DOT));

  fpsBrickHor := AddDotFillPattern(FPS_BRICK_HOR,
    StringToDotPattern(BRICK_HOR_PATTERN));
  fpsBrickHor := AddDotFillPattern(FPS_BRICK_DIAG,
    StringToDotPattern(BRICK_DIAG_PATTERN));
  fpsCheckerBoardLarge := AddDotFillPattern(FPS_CHECKERBOARD_LARGE,
    StringToDotPattern(CHECKERBOARD_LARGE_PATTERN));
  fpsCheckerBoardSmall := AddDotFillPattern(FPS_CHECKERBOARD_SMALL,
    StringToDotPattern(CHECKERBOARD_SMALL_PATTERN));
  fpsConfettiLarge := AddDotFillPattern(FPS_CONFETTI_LARGE,
    StringToDotPattern(CONFETTI_LARGE_PATTERN));
  fpsConfettiSmall := AddDotFillPattern(FPS_CONFETTI_SMALL,
    StringToDotPattern(CONFETTI_SMALL_PATTERN));
  fpsDiamond := AddDotFillPattern(FPS_DIAMOND,
    StringToDotPattern(DIAMOND_PATTERN));
  fpsDivot := AddDotFillPattern(FPS_DIVOT,
    StringToDotPattern(DIVOT_PATTERN));
  fpsPlaid := AddDotFillPattern(FPS_PLAID,
    StringToDotPattern(PLAID_PATTERN));
  fpsShingle := AddDotFillPattern(FPS_SHINGLE,
    StringToDotPattern(SHINGLE_PATTERN));
  fpsSphere := AddDotFillPattern(FPS_SPHERE,
    StringToDotPattern(SPHERE_PATTERN));
  fpsTrellis := AddDotFillPattern(FPS_TRELLIS,
    StringToDotPattern(TRELLIS_PATTERN));
  fpsWave := AddDotFillPattern(FPS_WAVE,
    StringToDotPattern(WAVE_PATTERN));
  fpsWeave := AddDotFillPattern(FPS_WEAVE,
    StringToDotPattern(WEAVE_PATTERN));
  fpsZigzag := AddDotFillPattern(FPS_ZIGZAG,
    StringToDotPattern(ZIGZAG_PATTERN));
end;

function TsRawFillPatternList.AddDotFillPattern(AName: String;
  APattern: TsDotFillPattern): Integer;
var
  patt: TsRawFillPattern;
begin
  patt := TsRawFillPattern.Create(AName, APattern);
  Result := AddOrReplace(patt);
end;

function TsRawFillPatternList.AddFillPattern(AName: String; ADotPattern: TsDotFillPattern;
  ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
var
  patt: TsRawFillPattern;
begin
  patt := TsRawFillPattern.Create(AName,
    ADotPattern,
    ALineDistance, ALineAngle, ALineWidth, AMultiplier
  );
  Result := AddOrReplace(patt);
end;

function TsRawFillPatternList.AddLineFillPattern(AName: String;
  ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
var
  patt: TsRawFillPattern;
begin
  patt := TsRawFillPattern.Create(AName, ALineDistance, ALineAngle, ALineWidth, AMultiplier);
  patt.DotPattern := FindSimilarDotPattern(ALineDistance, ALineAngle, ALineWidth, AMultiplier);
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
    Insert(idx, APattern);
    Result := idx;
  end;
end;

{ Adds a copy of the pattern at the given index and gives it a new name. }
function TsRawFillPatternList.ClonePattern(AIndex: Integer; AName: String): Integer;
var
  patt: TsRawFillPattern;
begin
  patt := Items[AIndex];
  Result := AddFillPattern(AName,
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
    if (ALineDistance >= WIDE_DISTANCE) or (ALineDistance <= -8) then
      similarIdx := IfThen(singleLine, fpsHorThin, fpsCrossThin)
    else
      similarIdx := IfThen(singleLine, fpsHorNarrow, fpsCrossNarrow);
  end else
  if SameValue(cosAngle, 0.0, 0.2) then  // about 90 +/- 12°
  begin
    if (ALineDistance >= WIDE_DISTANCE) or (ALineDistance <= -8) then
      similarIdx := IfThen(singleLine, fpsVertThin, fpsCrossThin)
    else
      similarIdx := IfThen(singleLine, fpsVertNarrow, fpsCrossNarrow);
  end else
  if ALineAngle > 0 then
  begin
    if (ALineDistance >= WIDE_DISTANCE) or (ALineDistance <= -8) then
      similarIdx := IfThen(singleLine, fpsDiagUpThin, fpsHatchThin)
    else
      similarIdx := IfThen(singleLine, fpsDiagUpNarrow, fpsHatchNarrow);
  end else
  begin
    if (ALineDistance >= WIDE_DISTANCE) or (ALineDistance <= -8) then
      similarIdx := IfThen(singleLine, fpsDiagDownThin, fpsHatchThin)
    else
      similarIdx := IfThen(singleLine, fpsDiagDownNarrow, fpsHatchNarrow);
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

end.

