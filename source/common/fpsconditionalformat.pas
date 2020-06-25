unit fpsConditionalFormat;

{$mode objfpc}{$H+}

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
    cfcAboveAverage, cfcBelowAverage,
    cfcBeginsWidth, cfcEndsWith,
    cfcDuplicate, cfcUnique,
    cfcContainsText, cfcNotContaisText,
    cfcContainsErrors, cfcNotContainsErrors
  );

  {cellIs
   expression
   colorScale, dataBar, iconSet
   containsText, notContainsText, beginsWith, endsWith, containsBlanks, notContainsBlanks, containsErrors, notContainsErrors
   }

  TsCFCellRule = class(TsCFRule)
  public
    Condition: TsCFCondition;
    Operand1: Variant;
    Operand2: Variant;
    FormatIndex: Integer;
    procedure Assign(ASource: TsCFRule); override;
  end;

  { Color range }
  TsCFColorRangeValue = (crvMin, crvMax, crvPercentile);

  TsCFColorRangeRule = class(TsCFRule)
    StartValue: TsCFColorRangeValue;
    CenterValue: TsCFColorRangeValue;
    EndValue: TsCFColorRangeValue;
    StartValueParam: Double;
    CenterValueParam: Double;
    EndValueParam: Double;
    StartColor: TsColor;
    CenterColor: TsColor;
    EndColor: TsColor;
    procedure Assign(ASource: TsCFRule); override;
  end;

  { DataBars }
  TsCFDatabarRule = class(TsCFRule)
    procedure Assign(ASource: TsCFRule); override;
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
    FCellRange: TsCellRange;
    FRules: TsCFRules;
    function GetRules(AIndex: Integer): TsCFRule;
    function GetRulesCount: Integer;
  public
    constructor Create(ACellRange: TsCellRange);
    destructor Destroy; override;

    property CellRange: TsCellRange read FCellRange;
    property Rules[AIndex: Integer]: TsCFRule read GetRules;
    property RulesCount: Integer read GetRulesCount;
  end;

  TsConditionalFormatList = class(TFPObjectList)
  protected
    function AddRule(ARange: TsCellRange; ARule: TsCFRule): Integer;
  public
    function AddCellRule(ARange: TsCellRange; ACondition: TsCFCondition;
      ACellFormatIndex: Integer): Integer; overload;
    function AddCellRule(ARange: TsCellRange; ACondition: TsCFCondition;
      AParam: Variant; ACellFormatIndex: Integer): Integer; overload;
    function AddCellRule(ARange: TsCellRange; ACondition: TsCFCondition;
      AParam1, AParam2: Variant; ACellFormatIndex: Integer): Integer; overload;
    procedure AddColorRangeRule(ARange: TsCellRange);
    procedure AddDataBarRule(ARange: TsCellRange);
    function Find(ARange: TsCellRange): Integer;
  end;


implementation

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

procedure TsCFDataBarRule.Assign(ASource: TsCFRule);
begin
  if ASource is TsCFDataBarRule then
  begin
    //
  end else
    raise Exception.Create('Source cannot be assigned to TCVDataBarRule');
end;

procedure TsCFColorRangeRule.Assign(ASource: TsCFRule);
begin
  if ASource is TsCFColorRangeRule then
  begin
    StartValue := TsCFColorRangeRule(ASource).StartValue;
    CenterValue := TsCFColorRangeRule(ASource).CenterValue;
    EndValue := TsCFColorRangeRule(ASource).EndValue;
    StartValueParam := TsCFColorRangeRule(ASource).StartValueParam;
    CenterValueParam := TsCFColorRangeRule(ASource).CenterValueParam;
    EndValueParam := TsCFColorRangeRule(ASource).EndValueParam;
    StartColor := TsCFColorRangeRule(ASource).StartColor;
    CenterColor := TsCFColorRangeRule(ASource).CenterColor;
    EndColor := TsCFColorRangeRule(ASource).EndColor;
  end else
    raise Exception.Create('Source cannot be assigned to TCVDataBarRule');
end;


{ TCFRule }

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

constructor TsConditionalFormat.Create(ACellRange: TsCellRange);
begin
  inherited Create;
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
function TsConditionalFormatList.AddRule(ARange: TsCellRange;
  ARule: TsCFRule): Integer;
var
  CF: TsConditionalFormat;
  idx: Integer;
begin
  idx := Find(ARange);
  if idx = -1 then begin
    CF := TsConditionalFormat.Create(ARange);
    idx := Add(CF);
  end else
    CF := TsConditionalFormat(Items[idx]);
  CF.FRules.Add(ARule);
  Result := idx;
end;

// TODO: Add pre-checks for compatibility of condition and operands

function TsConditionalFormatList.AddCellRule(ARange: TsCellRange;
  ACondition: TsCFCondition; ACellFormatIndex: Integer): Integer;
var
  rule: TsCFCellRule;
begin
  rule := TsCFCellRule.Create;
  rule.Condition := ACondition;
  rule.Operand1 := varNull;
  rule.Operand2 := varNull;
  rule.FormatIndex := ACellFormatIndex;
  Result := AddRule(ARange, rule);
end;

function TsConditionalFormatList.AddCellRule(ARange: TsCellRange;
  ACondition: TsCFCondition; AParam: Variant; ACellFormatIndex: Integer): Integer;
var
  rule: TsCFCellRule;
begin
  rule := TsCFCellRule.Create;
  rule.Condition := ACondition;
  rule.Operand1 := AParam;
  rule.Operand2 := varNull;
  rule.FormatIndex := ACellFormatIndex;
  Result := AddRule(ARange, rule);
end;

function TsConditionalFormatList.AddCellRule(ARange: TsCellRange;
  ACondition: TsCFCondition; AParam1, AParam2: Variant;
  ACellFormatIndex: Integer): Integer;
var
  rule: TsCFCellRule;
begin
  rule := TsCFCellRule.Create;
  rule.Condition := ACondition;
  rule.Operand1 := AParam1;
  rule.Operand2 := AParam2;
  rule.FormatIndex := ACellFormatIndex;
  Result := AddRule(ARange, rule);
end;

procedure TsConditionalFormatList.AddColorRangeRule(ARange: TsCellRange);
begin
  raise EXception.Create('ColorRange not yet implemented.');
end;

procedure TsConditionalFormatlist.AddDataBarRule(ARange: TsCellRange);
begin
  raise Exception.Create('DataBars not yet implemented.');
end;


{@@ ----------------------------------------------------------------------------
  The conditional format list must be unique regarding cell ranges.
  This function searches all format item whether a given cell ranges is
  already listed.
-------------------------------------------------------------------------------}
function TsConditionalFormatList.Find(ARange: TsCellRange): Integer;
var
  i: Integer;
  CF: TsConditionalFormat;
  CFRange: TsCellRange;
begin
  for i := 0 to Count-1 do
  begin
    CF := TsConditionalFormat(Items[i]);
    CFRange := CF.CellRange;
    if (CFRange.Row1 = ARange.Row1) and (CFRange.Row2 = ARange.Row2) and
       (CFRange.Col1 = ARange.Col1) and (CFRange.Col2 = ARange.Col2) then
    begin
      Result := i;
      exit;
    end;
  end;
  Result := -1;
end;

end.

