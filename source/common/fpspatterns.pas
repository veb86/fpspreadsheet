unit fpsPatterns;

{$mode objfpc}{$H+}

interface

uses
  SysUtils, Classes, Contnrs, Math, Types, FPImage, FPCanvas,
  fpsTypes, fpsChart;

const
  FPS_GRAY06 = 'GRAY06';
  FPS_GRAY12 = 'GRAY12';
  FPS_GRAY25 = 'GRAY25';
  FPS_GRAY50 = 'GRAY50';
  FPS_GRAY75 = 'GRAY75';

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

  FPS_CROSS_DOT = 'CROSS_DOT';
  FPS_HATCH_DOT = 'HATCH_DOT';

  FPS_BRICK_HOR = 'BRICK_HOR';
  FPS_BRICK_DIAG = 'BRICK_DIAG';
  FPS_CHECKERBOARD_LARGE = 'CHECKERBOARD_LARGE';
  FPS_CHECKERBOARD_SMALL = 'CHECKERBOARD_SMALL';
  FPS_DIAMOND = 'DIAMOND';
  FPS_SHINGLE = 'SHINGLE';
  FPS_WAVE = 'WAVE';
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
  { TsFillPattern
    Combines the same visual pattern as a dot matrix and as vector strokes.
    Vector strokes are supported only by ODS }
  TsFillPattern = class
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
    procedure CopyFrom(ASource: TsFillPattern); virtual;
    property Name: String read FName;
    property DotPattern: TsDotFillPattern read FDotPattern write FDotPattern;
    property LinePattern: TsLineFillPattern read FLinePattern write SetLinePattern;
  end;

  TsFillPatternList = class(TFPObjectlist)
  private
    function GetItem(AIndex: Integer): TsFillPattern;
    procedure SetItem(AIndex: Integer; AValue: TsFillPattern);
  protected
    function AddOrReplace(APattern: TsFillPattern): Integer;
    function FindSimilarDotPattern(ALineDistance, ALineAngle, ALineWidth: Single;
      AMultiplier: TsLineFillPatternMultiplier): TsDotFillPattern;
  public
    function AddFillPattern(AName: String; ADotPattern: TsDotFillPattern;
      ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier): Integer;
    function AddDotFillPattern(AName: String; APattern: TsDotFillPattern): Integer; overload;
    function AddLineFillPattern(AName: String; ALineDistance, ALineAngle, ALineWidth: Single;
      AMultiplier: TsLineFillPatternMultiplier): Integer;
    function FindByName(AName: String): TsFillPattern;
    function IndexOfName(AName: String): Integer;
    property Items[AIndex: Integer]: TsFillPattern read GetItem write SetItem; default;
  end;

function StringToDotPattern(APattern: String): TsDotFillPattern;

procedure CreateFillPatterns(WithDefaultPatterns: Boolean);
function RegisterFillPattern(AName: String; ADotPattern: TsDotFillPattern): Integer;
function RegisterFillPattern(AName: String;
  ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier): Integer;
function RegisterFillPattern(AName: String; ADotPattern: TsDotFillPattern;
  ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier): Integer;

function ClonePattern(AIndex: Integer; AName: String): Integer;
function FindLineFillPatternIndex(ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
function GetFillPattern(AIndex: Integer): TsFillPattern;
function IndexOfFillPattern(AName: String): Integer;
function NumFillPatterns: Integer;

implementation

var
  FillPatternList: TsFillPatternList = nil;

const
  WIDE_DISTANCE = 3.0;
  NARROW_DISTANCE = 1.5;
  THIN_LINE = 0.1;
  THICK_LINE = 0.3;

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

{ TsFillPattern }

constructor TsFillPattern.Create(AName: String; ADotPattern: TsDotFillPattern);
begin
  inherited Create;
  FName := AName;
  FDotPattern := ADotPattern;
  FLinePattern := nil;  // no line pattern
end;

constructor TsFillPattern.Create(AName: String;
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

constructor TsFillPattern.Create(AName: String; ADotPattern: TsDotFillPattern;
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

destructor TsFillPattern.Destroy;
begin
  FName := '';
  FLinePattern.Free;
  inherited;
end;

procedure TsFillPattern.CopyFrom(ASource: TsFillPattern);
begin
  FName := ASource.Name;
  FDotPattern := ASource.DotPattern;
  SetLinePattern(ASource.LinePattern);
end;

procedure TsFillPattern.SetLinePattern(AValue: TsLineFillPattern);
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

{ TsFillPatternList }

function TsFillPatternList.AddDotFillPattern(AName: String;
  APattern: TsDotFillPattern): Integer;
var
  patt: TsFillPattern;
begin
  patt := TsFillPattern.Create(AName, APattern);
  Result := AddOrReplace(patt);
end;

function TsFillPatternList.AddFillPattern(AName: String; ADotPattern: TsDotFillPattern;
  ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
var
  patt: TsFillPattern;
begin
  patt := TsFillPattern.Create(AName,
    ADotPattern,
    ALineDistance, ALineAngle, ALineWidth, AMultiplier
  );
  Result := AddOrReplace(patt);
end;

function TsFillPatternList.AddLineFillPattern(AName: String;
  ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
var
  patt: TsFillPattern;
begin
  patt := TsFillPattern.Create(AName, ALineDistance, ALineAngle, ALineWidth, AMultiplier);
  patt.DotPattern := FindSimilarDotPattern(ALineDistance, ALineAngle, ALineWidth, AMultiplier);
  Result := AddOrReplace(patt);
end;

function TsFillPatternList.AddOrReplace(APattern: TsFillPattern): Integer;
var
  idx: Integer;
begin
  idx := IndexOfName(APattern.Name);
  if idx = -1 then
    Result := Add(APattern)
  else
  begin
    Items[idx].CopyFrom(APattern);
//    Items[idx].Free;
//    Delete(idx);
    Insert(idx, APattern);
    Result := idx;
  end;
end;

function TsFillPatternList.FindByName(AName: String): TsFillPattern;
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

function TsFillPatternList.FindSimilarDotPattern(ALineDistance, ALineAngle,
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

function TsFillPatternList.IndexOfName(AName: String): Integer;
begin
  for Result := 0 to Count-1 do
    if Items[Result].Name = AName then
      exit;
  Result := -1;
end;

function TsFillPatternList.GetItem(AIndex: Integer): TsFillPattern;
begin
  Result := TsFillPattern(inherited Items[AIndex]);
end;

procedure TsFillPatternList.SetItem(AIndex: Integer; AValue: TsFillPattern);
begin
  TsFillPattern(inherited Items[AIndex]).CopyFrom(AValue);
end;


{ Globals }

const
  GRAY75_PATTERN =
    ' xxx xxx'+
    'xxxxxxxx'+
    'xx xxx x'+
    'xxxxxxxx'+
    ' xxx xxx'+
    'xxxxxxxx'+
    'xx xxx x'+
    'xxxxxxxx';
  GRAY50_PATTERN =
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x'+
    'x x x x '+
    ' x x x x';
  GRAY25_PATTERN =
    'x   x   '+
    '  x   x '+
    'x   x   '+
    '  x   x '+
    'x   x   '+
    '  x   x '+
    'x   x   '+
    '  x   x ';
  GRAY12_PATTERN =
    'x   x   '+
    '        '+
    '  x   x '+
    '        '+
    'x   x   '+
    '        '+
    '  x   x '+
    '        ';
  GRAY06_PATTERN =
    'x       '+
    '        '+
    '    x   '+
    '        '+
    'x       '+
    '        '+
    '    x   '+
    '        ';
  HOR_PATTERN_THIN =
    'xxxxxxxx'+
    '        '+
    '        '+
    '        '+
    '        '+
    '        '+
    '        '+
    '        ';
  HOR_PATTERN_THICK =
    'xxxxxxxx'+
    'xxxxxxxx'+
    '        '+
    '        '+
    '        '+
    '        '+
    '        '+
    '        ';
  HOR_PATTERN_NARROW =
    'xxxxxxxx'+
    '        '+
    '        '+
    '        '+
    'xxxxxxxx'+
    '        '+
    '        '+
    '        ';
  VERT_PATTERN_THIN =
    'x       '+
    'x       '+
    'x       '+
    'x       '+
    'x       '+
    'x       '+
    'x       '+
    'x       ';
  VERT_PATTERN_THICK =
    'xx      '+
    'xx      '+
    'xx      '+
    'xx      '+
    'xx      '+
    'xx      '+
    'xx      '+
    'xx      ';
  VERT_PATTERN_NARROW =
    'x   x   '+
    'x   x   '+
    'x   x   '+
    'x   x   '+
    'x   x   '+
    'x   x   '+
    'x   x   '+
    'x   x   ';
  DIAG_DOWN_PATTERN_THIN =
    'x       '+
    ' x      '+
    '  x     '+
    '   x    '+
    '    x   '+
    '     x  '+
    '      x '+
    '       x';
  DIAG_DOWN_PATTERN_THICK =
    'xx      '+
    ' xx     '+
    '  xx    '+
    '   xx   '+
    '    xx  '+
    '     xx '+
    '      xx'+
    'x      x';
  DIAG_DOWN_PATTERN_NARROW =
    'x   x   '+
    ' x   x  '+
    '  x   x '+
    '   x   x'+
    'x   x   '+
    ' x   x  '+
    '  x   x '+
    '   x   x';
  DIAG_UP_PATTERN_THIN =
    '       x'+
    '      x '+
    '     x  '+
    '    x   '+
    '   x    '+
    '  x     '+
    ' x      '+
    'x       ';
  DIAG_UP_PATTERN_THICK =
    'x      x'+
    '      xx'+
    '     xx '+
    '    xx  '+
    '   xx   '+
    '  xx    '+
    ' xx     '+
    'xx      ';
  DIAG_UP_PATTERN_NARROW =
    '   x   x'+
    '  x   x '+
    ' x   x  '+
    'x   x   '+
    '   x   x'+
    '  x   x '+
    ' x   x  '+
    'x   x   ';
  HATCH_PATTERN_THIN =
    '       x'+
    'x     x '+
    ' x   x  '+
    '  x x   '+
    '   x    '+
    '  x x   '+
    ' x   x  '+
    'x     x ';
  HATCH_PATTERN_THICK =
    'x      x'+
    'xx    xx'+
    ' xx  xx '+
    '  xxxx  '+
    '   xx   '+
    '  xxxx  '+
    ' xx  xx '+
    'xx    xx';
  HATCH_PATTERN_NARROW =
    '   x   x'+
    'x x x x '+
    ' x   x  '+
    'x x x x '+
    '   x   x'+
    'x x x  x'+
    ' x   x  '+
    'x x x x ';
  HATCH_PATTERN_DOT =
    'x       '+
    '        '+
    '  x   x '+
    '        '+
    '    x   '+
    '        '+
    '  x   x '+
    '        ';
  CROSS_PATTERN_THIN =
    '   x    '+
    '   x    '+
    '   x    '+
    'xxxxxxxx'+
    '   x    '+
    '   x    '+
    '   x    '+
    '   x    ';
  CROSS_PATTERN_THICK =
    '   xx   '+
    '   xx   '+
    '   xx   '+
    'xxxxxxxx'+
    'xxxxxxxx'+
    '   xx   '+
    '   xx   '+
    '   xx   ';
  CROSS_PATTERN_NARROW =
    '  x   x '+
    'xxxxxxxx'+
    '  x   x '+
    '  x   x '+
    '  x   x '+
    'xxxxxxxx'+
    '  x   x '+
    '  x   x ';
  CROSS_PATTERN_DOT =
    'x x x x '+
    '        '+
    'x       '+
    '        '+
    'x       '+
    '        '+
    'x       '+
    '        ';
  CHECKERBOARD_LARGE_PATTERN =
    'xxxx    '+
    'xxxx    '+
    'xxxx    '+
    'xxxx    '+
    '    xxxx'+
    '    xxxx'+
    '    xxxx'+
    '    xxxx';
  CHECKERBOARD_SMALL_PATTERN =
    'xx  xx  '+
    'xx  xx  '+
    '  xx  xx'+
    '  xx  xx'+
    'xx  xx  '+
    'xx  xx  '+
    '  xx  xx'+
    '  xx  xx';
  DIAMOND_PATTERN =
    '   x    '+
    '  xxx   '+
    ' xxxxx  '+
    'xxxxxxx '+
    ' xxxxx  '+
    '  xxx   '+
    '   x    '+
    '        ';
  BRICK_HOR_PATTERN =
    'xxxxxxxx'+
    'x       '+
    'x       '+
    'x       '+
    'xxxxxxxx'+
    '    x   '+
    '    x   '+
    '    x   ';
  BRICK_DIAG_PATTERN =
    '       x'+
    '      x '+
    '     x  '+
    '    x   '+
    '   xx   '+
    '  x  x  '+
    ' x    x '+
    'x      x';
  SHINGLE_PATTERN =
    '      xx'+
    'x    x  '+
    ' x  x   '+
    '  xx    '+
    '    xx  '+
    '      x '+
    '       x'+
    '       x';
  WAVE_PATTERN =
    '        '+
    '   xx   '+
    '  x  x x'+
    'xx      '+
    '        '+
    '   xx   '+
    '  x  x x'+
    'xx      ';
  ZIGZAG_PATTERN =
    'x      x'+
    ' x    x '+
    '  x  x  '+
    '   xx   '+
    'x      x'+
    ' x    x '+
    '  x  x  '+
    '   xx   ';

procedure CreateFillPatterns(WithDefaultPatterns: Boolean);
begin
  if (FillPatternList <> nil) and (FillPatternList.Count > 0) and WithDefaultPatterns then
    exit;

  if FillPatternList = nil then
    FillPatternList := TsFillPatternList.Create;

  if WithDefaultPatterns then
  begin
    fpsGray75 := RegisterFillPattern(FPS_GRAY75, StringToDotPattern(GRAY75_PATTERN));
    fpsGray50 := RegisterFillPattern(FPS_GRAY50, StringToDotPattern(GRAY50_PATTERN));
    fpsGray25 := RegisterFillPattern(FPS_GRAY25, StringToDotPattern(GRAY25_PATTERN));
    fpsGray12 := RegisterFillPattern(FPS_GRAY12, StringToDotPattern(GRAY12_PATTERN));
    fpsGray06 := RegisterFillPattern(FPS_GRAY06, StringToDotPattern(GRAY06_PATTERN));

    fpsHorThick := RegisterFillPattern(FPS_HOR_THICK,
      StringToDotPattern(HOR_PATTERN_THICK),
      WIDE_DISTANCE, 0.0, THICK_LINE, lfpmSingle);
    fpsVertThick := RegisterFillPattern(FPS_VERT_THICK,
      StringToDotPattern(VERT_PATTERN_THICK),
      WIDE_DISTANCE, 90.0, THICK_LINE, lfpmSingle);
    fpsDiagUpThick := RegisterFillPattern(FPS_DIAG_UP_THICK,
      StringToDotPattern(DIAG_UP_PATTERN_THICK),
      WIDE_DISTANCE, 45.0, THICK_LINE, lfpmSingle);
    fpsDiagDownThick := RegisterFillPattern(FPS_DIAG_DOWN_THICK,
      StringToDotPattern(DIAG_DOWN_PATTERN_THICK),
      WIDE_DISTANCE, -45.0, THICK_LINE, lfpmSingle);
    fpsHatchThick := RegisterFillPattern(FPS_HATCH_THICK,
      StringToDotPattern(HATCH_PATTERN_THICK),
      WIDE_DISTANCE, 45.0, THICK_LINE, lfpmDouble);
    fpsCrossThick := RegisterFillPattern(FPS_CROSS_THICK,
      StringToDotPattern(CROSS_PATTERN_THICK),
      WIDE_DISTANCE, 0.0, THICK_LINE, lfpmDouble);

    fpsHorThin := RegisterFillPattern(FPS_HOR_THIN,
      StringToDotPattern(HOR_PATTERN_THIN),
      WIDE_DISTANCE, 0.0, THIN_LINE, lfpmSingle);
    fpsVertThin := RegisterFillPattern(FPS_VERT_THIN,
      StringToDotPattern(VERT_PATTERN_THIN),
      WIDE_DISTANCE, 90.0, THIN_LINE, lfpmSingle);
    fpsDiagUpThin := RegisterFillPattern(FPS_DIAG_UP_THIN,
      StringToDotPattern(DIAG_UP_PATTERN_THIN),
      WIDE_DISTANCE, 45.0, THIN_LINE, lfpmSingle);
    fpsDiagDownThin := RegisterFillPattern(FPS_DIAG_DOWN_THIN,
      StringToDotPattern(DIAG_DOWN_PATTERN_THIN),
      WIDE_DISTANCE, -45.0, THIN_LINE, lfpmSingle);
    fpsHatchThin := RegisterFillPattern(FPS_HATCH_THIN,
      StringToDotPattern(HATCH_PATTERN_THIN),
      WIDE_DISTANCE, 45.0, THIN_LINE, lfpmDouble);
    fpsCrossThin := RegisterFillPattern(FPS_CROSS_THIN,
      StringToDotPattern(CROSS_PATTERN_THIN),
      WIDE_DISTANCE, 0.0, THIN_LINE, lfpmDouble);

    fpsHorNarrow := RegisterFillPattern(FPS_HOR_NARROW,
      StringToDotPattern(HOR_PATTERN_NARROW),
      NARROW_DISTANCE, 0.0, THIN_LINE, lfpmSingle);
    fpsVertNarrow := RegisterFillPattern(FPS_VERT_NARROW,
      StringToDotPattern(VERT_PATTERN_NARROW),
      NARROW_DISTANCE, 90.0, THIN_LINE, lfpmSingle);
    fpsDiagUpNarrow := RegisterFillPattern(FPS_DIAG_UP_NARROW,
      StringToDotPattern(DIAG_UP_PATTERN_NARROW),
      NARROW_DISTANCE, 45.0, THIN_LINE, lfpmSingle);
    fpsDiagDownNarrow := RegisterFillPattern(FPS_DIAG_DOWN_NARROW,
      StringToDotPattern(DIAG_DOWN_PATTERN_NARROW),
      NARROW_DISTANCE, -45.0, THIN_LINE, lfpmSingle);
    fpsHatchNarrow := RegisterFillPattern(FPS_HATCH_NARROW,
      StringToDotPattern(HATCH_PATTERN_NARROW),
      NARROW_DISTANCE, 45.0, THIN_LINE, lfpmDouble);
    fpsCrossNarrow := RegisterFillPattern(FPS_CROSS_NARROW,
      StringToDotPattern(CROSS_PATTERN_NARROW),
      NARROW_DISTANCE, 0.0, THIN_LINE, lfpmDouble);

    fpsCrossDot := RegisterFillPattern(FPS_CROSS_DOT, StringToDotPattern(CROSS_PATTERN_DOT));
    fpsHatchDot := RegisterFillPattern(FPS_HATCH_DOT, StringToDotPattern(HATCH_PATTERN_DOT));

    fpsBrickHor := RegisterFillPattern(FPS_BRICK_HOR, StringToDotPattern(BRICK_HOR_PATTERN));
    fpsBrickHor := RegisterFillPattern(FPS_BRICK_DIAG, StringToDotPattern(BRICK_DIAG_PATTERN));
    fpsCheckerBoardLarge := RegisterFillPattern(FPS_CHECKERBOARD_LARGE, StringToDotPattern(CHECKERBOARD_LARGE_PATTERN));
    fpsCheckerBoardSmall := RegisterFillPattern(FPS_CHECKERBOARD_SMALL, StringToDotPattern(CHECKERBOARD_SMALL_PATTERN));
    fpsDiamond := RegisterFillPattern(FPS_DIAMOND, StringToDotPattern(DIAMOND_PATTERN));
    fpsShingle := RegisterFillPattern(FPS_SHINGLE, StringToDotPattern(SHINGLE_PATTERN));
    fpsWave := RegisterFillPattern(FPS_WAVE, StringToDotPattern(WAVE_PATTERN));
    fpsZigzag := RegisterFillPattern(FPS_ZIGZAG, StringToDotPattern(ZIGZAG_PATTERN));

  end;
end;

procedure EnsureFillPatternList;
begin
  CreateFillPatterns(false);
end;

function RegisterFillPattern(AName: String; ADotPattern: TsDotFillPattern;
  ALineDistance, ALineAngle, ALineWidth: Single; AMultiplier: TsLineFillPatternMultiplier): Integer;
begin
  EnsureFillPatternList;
  Result := FillPatternList.AddFillPattern(AName, ADotPattern,
    ALineDistance, ALineAngle, ALineWidth, AMultiplier);
end;

function RegisterFillPattern(AName: String; ADotPattern: TsDotFillPattern): Integer;
begin
  EnsureFillPatternList;
  Result := FillPatternList.AddDotFillPattern(AName, ADotPattern);
end;

function RegisterFillPattern(AName: String; ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
begin
  EnsureFillPatternList;
  Result := FillPatternList.AddLineFillPattern(AName, ALineDistance, ALineAngle, ALineWidth, AMultiplier);
end;

function ClonePattern(AIndex: Integer; AName: String): Integer;
var
  patt: TsFillPattern;
begin
  if Assigned(FillPatternList) then
  begin
    patt := GetFillPattern(AIndex);
    Result := RegisterFillPattern(AName,
      patt.DotPattern,
      patt.LinePattern.Distance, patt.LinePattern.Angle, patt.LinePattern.LineWidth, patt.LinePattern.Multiplier
    );
  end else
    Result := -1;
end;

function FindLineFillPatternIndex(ALineDistance, ALineAngle, ALineWidth: Single;
  AMultiplier: TsLineFillPatternMultiplier): Integer;
var
  i: Integer;
  patt: TsFillPattern;
begin
  if Assigned(FillPatternList) then
    for i := 0 to FillPatternList.Count-1 do
    begin
      patt := FillPatternList[i];
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

function GetFillPattern(AIndex: Integer): TsFillPattern;
begin
  if Assigned(FillPatternList) then
    Result := FillPatternList.Items[AIndex]
  else
    Result := nil;
end;

function IndexOfFillPattern(AName: String): Integer;
begin
  if Assigned(FillPatternList) then
    Result := FillPatternList.IndexOfName(AName)
  else
    Result := -1;
end;

function NumFillPatterns: Integer;
begin
  if Assigned(FillPatternList) then
    Result := FillPatternList.Count
  else
    Result := 0;
end;

finalization
  FreeAndNil(FillPatternList);

end.

