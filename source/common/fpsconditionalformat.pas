unit fpsConditionalFormat;

{$mode objfpc}{$H+}
{$WARN 6058 off : Call to subroutine "$1" marked as inline is not inlined}
interface

uses
  Classes, Contnrs, SysUtils, Variants, fpsTypes;

type
  TsCFRule = class
  public
    procedure Assign(ASource: TsCFRule); virtual; abstract;
  end;

  { Cell is... }
  TsCFCondition = (
    cfcEqual, cfcNotEqual,
    cfcGreaterThan, cfcLessThan, cfcGreaterEqual, cfcLessEqual,
    cfcBetween, cfcNotBetween,
    cfcAboveAverage, cfcBelowAverage, cfcAboveEqualAverage, cfcBelowEqualAverage,
    cfcTop, cfcBottom, cfcTopPercent, cfcBottomPercent,
    cfcDuplicate, cfcUnique,
    cfcBeginsWith, cfcEndsWith,
    cfcContainsText, cfcNotContainsText,
    cfcContainsErrors, cfcNotContainsErrors,
    cfcExpression
  );

  TsCFCellRule = class(TsCFRule)
  public
    Condition: TsCFCondition;
    Operand1: Variant;
    Operand2: Variant;
    FormatIndex: Integer;
    procedure Assign(ASource: TsCFRule); override;
  end;

  TsCFValueKind = (vkNone, vkMin, vkMax, vkPercent, vkPercentile, vkValue);

  { Color range }
  TsCFColorRangeRule = class(TsCFRule)
    StartValueKind: TsCFValueKind;
    CenterValueKind: TsCFValueKind;
    EndValueKind: TsCFValueKind;
    StartValue: Double;
    CenterValue: Double;
    EndValue: Double;
    StartColor: TsColor;
    CenterColor: TsColor;
    EndColor: TsColor;
    ThreeColors: Boolean;
    constructor Create;
    procedure Assign(ASource: TsCFRule); override;
    procedure SetupEnd(AColor: TsColor; AKind: TsCFValueKind; AValue: Double);
    procedure SetupCenter(AColor: TsColor; AKind: TsCFValueKind; AValue: Double);
    procedure SetupStart(AColor: TsColor; AKind: TsCFValueKind; AValue: Double);
  end;

  { DataBars }
  TsCFDatabarRule = class(TsCFRule)
    StartValueKind: TsCFValueKind;
    EndValueKind: TsCFValueKind;
    StartValue: Double;
    EndValue: Double;
    Color: TsColor;
    constructor Create;
    procedure Assign(ASource: TsCFRule); override;
  end;

  { Icon sets }
  TsCFIconSet = (
    is3Arrows, is3ArrowsGray, is3Flags,
    is3TrafficLights1,  // x14 in xlsx
    is3TrafficLights2, is3Signs, is3Symbols, is3Symbols2,
    is3Smilies, is3Stars, is3Triangles, is3ColorSmilies,  // need x14 in xlsx
    is4Arrows, is4ArrowsGray, is4RedToBlack, is4Rating,
    is4TrafficLights,   // not in ODS
    is5Arrows, is5ArrowsGray, is5Rating, is5Quarters,
    is5Boxes            // needs x14 in Excel
  );
  TsCFIconSetRule = class(TsCFRule)
  private
    FIconSet: TsCFIconSet;
    FReverse: Boolean;
    FShowValue: Boolean;
    FValueKinds: array of TsCFValuekind;
    FValues: array of double;
    function GetIconCount: Integer;
    function GetValueKinds(AIndex: Integer): TsCFValueKind;
    function GetValues(AIndex: Integer): Double;
    procedure SetIconSet(AValue: TsCFIconSet);
    procedure SetValueKinds(AIndex: Integer; AKind: TsCFValueKind);
    procedure SetValues(AIndex: Integer; AValue: Double);
  public
    constructor Create;
    procedure Assign(ASource: TsCFRule); override;
    property IconSet: TsCFIconSet read FIconSet write SetIconSet;
    property IconCount: Integer read GetIconCount;
    property Values[AIndex: Integer]: Double read GetValues write SetValues;
    property ValueKinds[AIndex: Integer]: TsCFValueKind read GetValueKinds write SetValueKinds;
    property Reverse: Boolean read FReverse write FReverse;
    property ShowValue: Boolean read FShowValue write FShowValue;
  end;

  { Rules }
  TsCFRules = class(TFPObjectList)
  private
    function GetItem(AIndex: Integer): TsCFRule;
    function GetPriority(AIndex: Integer): Integer;
    procedure SetItem(AIndex: Integer; const AValue: TsCFRule);
  public
    property Items[AIndex: Integer]: TsCFRule read GetItem write SetItem; default;
    property Priority[AIndex: Integer]: Integer read GetPriority;
  end;

  { Conditional format item }
  TsConditionalFormat = class
  private
    FWorksheet: TsBasicWorksheet;
    FCellRange: TsCellRange;
    FRules: TsCFRules;
    function GetRules(AIndex: Integer): TsCFRule;
    function GetRulesCount: Integer;
  public
    constructor Create(AWorksheet: TsBasicWorksheet; ACellRange: TsCellRange);
    destructor Destroy; override;

    property CellRange: TsCellRange read FCellRange;
    property Rules[AIndex: Integer]: TsCFRule read GetRules;
    property RulesCount: Integer read GetRulesCount;
    property Worksheet: TsBasicWorksheet read FWorksheet;
  end;

  TsConditionalFormatList = class(TFPObjectList)
  protected
    function AddRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      ARule: TsCFRule): Integer;
  public
    function AddCellRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      ACondition: TsCFCondition;   ACellFormatIndex: Integer): Integer; overload;
    function AddCellRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      ACondition: TsCFCondition; AParam: Variant; ACellFormatIndex: Integer): Integer; overload;
    function AddCellRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      ACondition: TsCFCondition; AParam1, AParam2: Variant; ACellFormatIndex: Integer): Integer; overload;

    function AddColorRangeRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      AStartColor, AEndColor: TsColor): Integer; overload;
    function AddColorRangeRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      AStartColor, ACenterColor, AEndColor: TsColor): Integer; overload;
    function AddColorRangeRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      AStartColor: TsColor; AStartKind: TsCFValueKind; AStartValue: Double;
      AEndColor: TsColor; AEndKind: TsCFValueKind; AEndValue: Double): Integer; overload;
    function AddColorRangeRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      AStartColor: TsColor; AStartKind: TsCFValueKind; AStartValue: Double;
      ACenterColor: TsColor; ACenterKind: TsCFValueKind; ACenterValue: Double;
      AEndColor: TsColor; AEndKind: TsCFValueKind; AEndValue: Double): Integer; overload;

    function AddDataBarRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      ABarColor: TsColor): Integer; overload;
    function AddDataBarRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      ABarColor: TsColor; AStartKind: TsCFValueKind; AStartValue: Double;
      AEndKind: TsCFValueKind; AEndValue: Double): Integer; overload;

    function AddIconSetRule(ASheet: TsBasicWorksheet; ARange: TsCellRange;
      AIconSet: TsCFIconSet; AHideValue: Boolean = false; AReverse: Boolean = false): Integer; overload;
    function AddIconSetRule(ASheet: TsBasicWorksheet; ARange: TsCellRange; AIconSet: TsCFIconSet;
      AValueKind1: TsCFValueKind; AValue1: Double;
      AValueKind2: TsCFValueKind; AValue2: Double;
      AHideValue: Boolean = false; AReverse: Boolean = false): Integer; overload;
    function AddIconSetRule(ASheet: TsBasicWorksheet; ARange: TsCellRange; AIconSet: TsCFIconSet;
      AValueKind1: TsCFValueKind; AValue1: Double;
      AValueKind2: TsCFValueKind; AValue2: Double;
      AValueKind3: TsCFValueKind; AValue3: Double;
      AHideValue: Boolean = false; AReverse: Boolean = false): Integer; overload;
    function AddIconSetRule(ASheet: TsBasicWorksheet; ARange: TsCellRange; AIconSet: TsCFIconSet;
      AValueKind1: TsCFValueKind; AValue1: Double;
      AValueKind2: TsCFValueKind; AValue2: Double;
      AValueKind3: TsCFValueKind; AValue3: Double;
      AValueKind4: TsCFValueKind; AValue4: Double;
      AHideValue: Boolean = false; AReverse: Boolean = false): Integer; overload;

    procedure Delete(AIndex: Integer);
    function Find(ASheet: TsBasicWorksheet; ARange: TsCellRange): Integer;
  end;

  function GetCFIconCount(AIconSet: TsCFIconSet): Integer;

implementation

uses
  Math, TypInfo,
  fpSpreadsheet;

function GetCFIconCount(AIconSet: TsCFIconSet): Integer;
var
  s: String;
begin
  s := GetEnumName(TypeInfo(TsCFIconSet), integer(AIconSet));
  Result := ord(s[3]) - ord('0');
end;

procedure TsCFCellRule.Assign(ASource: TsCFRule);
begin
  if ASource is TsCFCellRule then
  begin
    Condition := TsCFCellRule(ASource).Condition;
    Operand1 := TsCFCellRule(ASource).Operand1;
    Operand2 := TsCFCellRule(ASource).Operand2;
    FormatIndex := TsCFCellRule(ASource).FormatIndex;
  end else
    raise Exception.Create('Source cannot be assigned to TCVCellRule');
end;

constructor TsCFDataBarRule.Create;
begin
  inherited;
  StartValueKind := vkMin;
  EndValueKind := vkMax;
  Color := scBlue;
end;

procedure TsCFDataBarRule.Assign(ASource: TsCFRule);
begin
  if ASource is TsCFDataBarRule then
  begin
    //
  end else
    raise Exception.Create('Source cannot be assigned to TsCFDataBarRule');
end;

constructor TsCFColorRangeRule.Create;
begin
  inherited;
  ThreeColors := true;
  SetupStart(scRed, vkMin, 0.0);
  SetupCenter(scYellow, vkPercent, 50.0);
  SetupEnd(scBlue, vkMax, 0.0);
  EndValueKind := vkMax;
  EndValue := 0;
  EndColor := scBlue;
end;

procedure TsCFColorRangeRule.Assign(ASource: TsCFRule);
begin
  if ASource is TsCFColorRangeRule then
  begin
    ThreeColors := TsCFColorRangeRule(ASource).ThreeColors;
    StartValueKind := TsCFColorRangeRule(ASource).StartValueKind;
    CenterValueKind := TsCFColorRangeRule(ASource).CenterValueKind;
    EndValueKind := TsCFColorRangeRule(ASource).EndValueKind;
    StartValue := TsCFColorRangeRule(ASource).StartValue;
    CenterValue := TsCFColorRangeRule(ASource).CenterValue;
    EndValue := TsCFColorRangeRule(ASource).EndValue;
    StartColor := TsCFColorRangeRule(ASource).StartColor;
    CenterColor := TsCFColorRangeRule(ASource).CenterColor;
    EndColor := TsCFColorRangeRule(ASource).EndColor;
  end else
    raise Exception.Create('Source cannot be assigned to TsCFColorRangeRule');
end;

procedure TsCFColorRangeRule.SetupCenter(AColor: TsColor;
  AKind: TsCFValueKind; AValue: Double);
begin
  CenterValueKind := AKind;
  CenterValue := AValue;
  CenterColor := AColor;
end;

procedure TsCFColorRangeRule.SetupEnd(AColor: TsColor;
  AKind: TsCFValueKind; AValue: Double);
begin
  EndValueKind := AKind;
  EndValue := AValue;
  EndColor := AColor;
end;

procedure TsCFColorRangeRule.SetupStart(AColor: TsColor;
  AKind: TsCFValueKind; AValue: Double);
begin
  StartValueKind := AKind;
  StartValue := AValue;
  StartColor := AColor;
end;


{ TsCFIconSetRule }

constructor TsCFIconSetRule.Create;
begin
  FIconSet := is3Arrows;
  SetLength(FValues, 2);
  Setlength(FValueKinds, 2);
  FValues[0] := 33;
  FValues[1] := 66;
  FValueKinds[0] := vkPercent;
  FValueKinds[1] := vkPercent;
  FShowValue := true;
  FReverse := false;
end;

procedure TsCFIconSetRule.Assign(ASource: TsCFRule);
var
  i: Integer;
begin
  if ASource is TsCFIconSetRule then
  begin
    SetIconSet(TsCFIconSetRule(ASource).IconSet);
    for i := 0 to High(FValues) do
      FValues[i] := TsCFIconSetRule(ASource).Values[i];
    for i := 0 to High(FValueKinds) do
      FValueKinds[i] := TsCFIconSetRule(ASource).ValueKinds[i];
    FShowValue := TsCFIconSetRule(ASource).ShowValue;
    FReverse := TsCFIconSetRule(ASource).Reverse;
  end else
    raise Exception.Create('Source cannot be assigned to TsCFIconSetRule');
end;

function TsCFIconSetRule.GetIconCount: Integer;
begin
  Result := Length(FValues) + 1;
end;

function TsCFIconSetRule.GetValueKinds(AIndex: Integer): TsCFValueKind;
begin
  Result := FValueKinds[AIndex];
end;

function TsCFIconSetRule.GetValues(AIndex: Integer): Double;
begin
  Result := FValues[AIndex];
end;

procedure TsCFIconSetRule.SetIconSet(AValue: TsCFIconSet);
var
  s: String;
  i, n: Integer;
begin
  if AValue = FIconSet then exit;

  FIconSet := AValue;

  s := GetEnumName(TypeInfo(TsCFIconSet), integer(AValue));
  n := Ord(s[3]) - ord('0');
  SetLength(FValues, n - 1);
  for i := 0 to High(FValues) do
    FValues[i] := (i + 1) * 100 div n;
  SetLength(FValueKinds, n - 1);
  for i := 0 to High(FValueKinds) do
    FValueKinds[i] := vkPercent;

  //                     value index
  //    (min)       0         1         2         (max)
  //      |---------|---------|---------|----------|
  //         icon0     icon1     icon2     icon3
end;

procedure TsCFIconSetRule.SetValueKinds(AIndex: Integer; AKind: TsCFValueKind);
begin
  FValueKinds[AIndex] := AKind;
end;

procedure TsCFIconSetRule.SetValues(AIndex: Integer; AValue: Double);
begin
  FValues[AIndex] := AValue;
end;


{ TsCFRule }

function TsCFRules.GetItem(AIndex: Integer): TsCFRule;
begin
  Result := TsCFRule(inherited Items[AIndex]);
end;

function TsCFRules.GetPriority(AIndex: Integer): Integer;
begin
  Result := Count - AIndex;
end;

procedure TsCFRules.SetItem(AIndex: Integer; const AValue: TsCFRule);
var
  item: TsCFRule;
begin
  item := GetItem(AIndex);
  item.Assign(AValue);
  inherited Items[AIndex] := item;
end;


{ TsConditonalFormat }

constructor TsConditionalFormat.Create(AWorksheet: TsBasicWorksheet;
  ACellRange: TsCellRange);
begin
  inherited Create;
  FWorksheet := AWorksheet;
  FCellRange := ACellRange;
  FRules := TsCFRules.Create;
end;

destructor TsConditionalFormat.Destroy;
begin
  FRules.Free;
  inherited;
end;

function TsConditionalFormat.GetRules(AIndex: Integer): TsCFRule;
begin
  Result := FRules[AIndex];
end;

function TsConditionalFormat.GetRulesCount: Integer;
begin
  Result := FRules.Count;
end;


{ TsConditionalFormatList }

{@@ ----------------------------------------------------------------------------
  Adds a new conditional format to the list.
  The format is specified by the cell range to which it is applied and by
  the rule describing the format.
  The rules are grouped for the same cell ranges.
-------------------------------------------------------------------------------}
function TsConditionalFormatList.AddRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; ARule: TsCFRule): Integer;
var
  CF: TsConditionalFormat;
  idx: Integer;
begin
  idx := Find(ASheet, ARange);
  if idx = -1 then begin
    CF := TsConditionalFormat.Create(ASheet, ARange);
    idx := Add(CF);
  end else
    CF := TsConditionalFormat(Items[idx]);
  CF.FRules.Add(ARule);
  Result := idx;
end;

// TODO: Add pre-checks for compatibility of condition and operands

function TsConditionalFormatList.AddCellRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; ACondition: TsCFCondition;
  ACellFormatIndex: Integer): Integer;
var
  rule: TsCFCellRule;
begin
  rule := TsCFCellRule.Create;
  rule.Condition := ACondition;
  VarClear(rule.Operand1);
  VarClear(rule.Operand2);
  rule.FormatIndex := ACellFormatIndex;
  Result := AddRule(ASheet, ARange, rule);
end;

function TsConditionalFormatList.AddCellRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; ACondition: TsCFCondition; AParam: Variant;
  ACellFormatIndex: Integer): Integer;
var
  rule: TsCFCellRule;
begin
  rule := TsCFCellRule.Create;
  rule.Condition := ACondition;
  rule.Operand1 := AParam;
  VarClear(rule.Operand2);
  rule.FormatIndex := ACellFormatIndex;
  Result := AddRule(ASheet, ARange, rule);
end;

function TsConditionalFormatList.AddCellRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; ACondition: TsCFCondition; AParam1, AParam2: Variant;
  ACellFormatIndex: Integer): Integer;
var
  rule: TsCFCellRule;
begin
  rule := TsCFCellRule.Create;
  rule.Condition := ACondition;
  rule.Operand1 := AParam1;
  rule.Operand2 := AParam2;
  rule.FormatIndex := ACellFormatIndex;
  Result := AddRule(ASheet, ARange, rule);
end;

function TsConditionalFormatList.AddColorRangeRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; AStartColor, ACenterColor, AEndColor: TsColor): Integer;
var
  rule: TsCFColorRangeRule;
begin
  rule := TsCFColorRangeRule.Create;
  rule.StartColor := AStartColor;
  rule.CenterColor := ACenterColor;
  rule.EndColor := AEndColor;
  rule.ThreeColors := true;
  Result := AddRule(ASheet, ARange, rule);
end;

function TsConditionalFormatList.AddColorRangeRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; AStartColor, AEndColor: TsColor): Integer;
var
  rule: TsCFColorRangeRule;
begin
  rule := TsCFColorRangeRule.Create;
  rule.StartColor := AStartColor;
  rule.CenterColor := scNotDefined;
  rule.EndColor := AEndColor;
  rule.ThreeColors := false;
  Result := AddRule(ASheet, ARange, rule);
end;

function TsConditionalFormatList.AddColorRangeRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange;
  AStartColor: TsColor; AStartKind: TsCFValueKind; AStartValue: Double;
  AEndColor: TsColor; AEndKind: TsCFValueKind; AEndValue: Double): Integer;
var
  rule: TsCFColorRangeRule;
begin
  rule := TsCFColorRangeRule.Create;
  rule.SetupStart(AStartColor, AStartKind, AStartValue);
  rule.SetupCenter(scNotDefined, vkNone, NaN);
  rule.SetupEnd(AEndColor, AEndKind, AEndValue);
  rule.ThreeColors := false;
  Result := AddRule(ASheet, ARange, rule);
end;

function TsConditionalFormatList.AddColorRangeRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange;
  AStartColor: TsColor; AStartKind: TsCFValueKind; AStartValue: Double;
  ACenterColor: TsColor; ACenterKind: TsCFValueKind; ACenterValue: Double;
  AEndColor: TsColor; AEndKind: TsCFValueKind; AEndValue: Double): Integer;
var
  rule: TsCFColorRangeRule;
begin
  rule := TsCFColorRangeRule.Create;
  rule.SetupStart(AStartColor, AStartKind, AStartValue);
  rule.SetupCenter(ACenterColor, ACenterKind, ACenterValue);
  rule.SetupEnd(AEndColor, AEndKind, AEndValue);
  rule.ThreeColors := true;
  Result := AddRule(ASheet, ARange, rule);
end;

function TsConditionalFormatlist.AddDataBarRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; ABarColor: TsColor): Integer;
var
  rule: TsCFDataBarRule;
begin
  rule := TsCFDataBarRule.Create;
  rule.Color := ABarColor;
  Result := AddRule(ASheet, ARange, rule);
end;

function TsConditionalFormatlist.AddDataBarRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; ABarColor: TsColor; AStartKind: TsCFValuekind;
  AStartValue: Double; AEndKind: TsCFValueKind; AEndValue: Double): Integer;
var
  rule: TsCFDataBarRule;
begin
  rule := TsCFDataBarRule.Create;
  rule.Color:= ABarColor;
  rule.StartValueKind := AStartKind;
  rule.StartValue := AStartValue;
  rule.EndValueKind := AEndKind;
  rule.EndValue := AEndValue;
  Result := AddRule(ASheet, ARange, rule);
end;


function TsConditionalFormatList.AddIconSetRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; AIconSet: TsCFIconSet; AHideValue: Boolean = false;
  AReverse: Boolean = false): Integer;
var
  rule: TsCFIconSetRule;
  i, n: Integer;
begin
  rule := TsCFIconSetRule.Create;
  rule.IconSet := AIconset;
  n := rule.IconCount;
  for i := 0 to n - 2 do
  begin
    rule.ValueKinds[i] := vkPercent;
    rule.Values[i] := ((i+1) * 100) div n
  end;
  rule.ShowValue := not AHideValue;
  rule.Reverse := AReverse;
  Result := AddRule(ASheet, ARange, rule);
end;

{ IconSet conditional format for 3 icons, ie. 2 values }
function TsConditionalFormatList.AddIconSetRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; AIconSet: TsCFIconSet;
  AValueKind1: TsCFValueKind; AValue1: Double;
  AValueKind2: TsCFValueKind; AValue2: Double;
  AHideValue: Boolean = false; AReverse: Boolean = false): Integer;
var
  rule: TsCFIconSetRule;
  n: Integer;
begin
  rule := TsCFIconSetRule.Create;
  rule.IconSet := AIconset;
  n := rule.IconCount;
  if n <> 3 then begin
    rule.Free;
    Result := -1;
    exit;
  end;

  rule.ValueKinds[0] := AValueKind1;   rule.Values[0] := AValue1;
  rule.ValueKinds[1] := AValueKind2;   rule.Values[1] := AValue2;

  rule.ShowValue := not AHideValue;
  rule.Reverse := AReverse;

  Result := AddRule(ASheet, ARange, rule);
end;

{ IconSet conditional format for 4 icons, i.e. 3 values }
function TsConditionalFormatList.AddIconSetRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; AIconSet: TsCFIconSet;
  AValueKind1: TsCFValueKind; AValue1: Double;
  AValueKind2: TsCFValueKind; AValue2: Double;
  AValueKind3: TsCFValueKind; AValue3: Double;
  AHideValue: Boolean = false; AReverse: Boolean = false): Integer;
var
  rule: TsCFIconSetRule;
  n: Integer;
begin
  rule := TsCFIconSetRule.Create;
  rule.IconSet := AIconset;
  n := rule.IconCount;
  if n <> 4 then begin
    rule.Free;
    Result := -1;
    exit;
  end;

  rule.ValueKinds[0] := AValueKind1;   rule.Values[0] := AValue1;
  rule.ValueKinds[1] := AValueKind2;   rule.Values[1] := AValue2;
  rule.ValueKinds[2] := AValueKind3;   rule.Values[2] := AValue3;

  rule.ShowValue := not AHideValue;
  rule.Reverse := AReverse;

  Result := AddRule(ASheet, ARange, rule);
end;

{ Iconset conditional format for 5 icons, i.e. 4 values }
function TsConditionalFormatList.AddIconSetRule(ASheet: TsBasicWorksheet;
  ARange: TsCellRange; AIconSet: TsCFIconSet;
  AValueKind1: TsCFValueKind; AValue1: Double; AValueKind2: TsCFValueKind; AValue2: Double;
  AValueKind3: TsCFValueKind; AValue3: Double; AValueKind4: TsCFValueKind; AValue4: Double;
  AHideValue: Boolean = false;  AReverse: Boolean = false): Integer;
var
  rule: TsCFIconSetRule;
  n: Integer;
begin
  rule := TsCFIconSetRule.Create;
  rule.IconSet := AIconset;
  n := rule.IconCount;
  if n <> 5 then begin
    rule.Free;
    Result := -1;
    exit;
  end;

  rule.ValueKinds[0] := AValueKind1;   rule.Values[0] := AValue1;
  rule.ValueKinds[1] := AValueKind2;   rule.Values[1] := AValue2;
  rule.ValueKinds[2] := AValueKind3;   rule.Values[2] := AValue3;
  rule.ValueKinds[3] := AValueKind4;   rule.Values[3] := AValue4;

  rule.ShowValue := not AHideValue;
  rule.Reverse := AReverse;

  Result := AddRule(ASheet, ARange, rule);
end;


{@@ ----------------------------------------------------------------------------
  Deletes the conditional format at the given index from the list.
  Iterates also through all cell in the range of the CF and removess the
  format index from the cell's ConditionalFormatIndex array.
-------------------------------------------------------------------------------}
procedure TsConditionalFormatList.Delete(AIndex: Integer);
var
  CF: TsConditionalFormat;
  r, c: Cardinal;
  i: Integer;
  cell: PCell;
begin
  CF := TsConditionalFormat(Items[AIndex]);
  for r := CF.CellRange.Row1 to CF.CellRange.Row2 do
    for c := CF.CellRange.Col1 to CF.CellRange.Col2 do
    begin
      cell := TsWorksheet(CF.Worksheet).FindCell(r, c);
      if Assigned(cell) and (Length(cell^.ConditionalFormatIndex) > 0) then begin
        for i := AIndex+1 to High(cell^.ConditionalFormatIndex) do
          cell^.ConditionalFormatIndex[i-1] := cell^.ConditionalFormatIndex[i];
        SetLength(cell^.ConditionalFormatIndex, Length(cell^.ConditionalFormatIndex)-1);
      end;
    end;

  inherited Delete(AIndex);
end;


{@@ ----------------------------------------------------------------------------
  The conditional format list must be unique regarding cell ranges.
  This function searches all format item whether a given cell ranges is
  already listed.
-------------------------------------------------------------------------------}
function TsConditionalFormatList.Find(ASheet: TsBasicWorksheet;
  ARange: TsCellRange): Integer;
var
  i: Integer;
  CF: TsConditionalFormat;
  CFRange: TsCellRange;
begin
  for i := 0 to Count-1 do
  begin
    CF := TsConditionalFormat(Items[i]);
    if CF.Worksheet = ASheet then
    begin
      CFRange := CF.CellRange;
      if (CFRange.Row1 = ARange.Row1) and (CFRange.Row2 = ARange.Row2) and
         (CFRange.Col1 = ARange.Col1) and (CFRange.Col2 = ARange.Col2) then
      begin
        Result := i;
        exit;
      end;
    end;
  end;
  Result := -1;
end;

end.

