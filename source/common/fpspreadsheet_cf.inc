{ Included by fpspreadsheet.pas }

{ Code for conditional formatting }

{==============================================================================}
{   TsWorksheet code for conditional formats                                   }
{==============================================================================}

procedure StoreCFIndexInCells(AWorksheet: TsWorksheet; AIndex: Integer;
  ARange: TsCellRange);
var
  r, c: Cardinal;
  n: Integer;
  cell: PCell;
begin
  for r := ARange.Row1 to ARange.Row2 do
    for c := ARange.Col1 to ARange.Col2 do
    begin
      cell := AWorksheet.GetCell(r, c);
      n := Length(cell^.ConditionalFormatIndex);
      SetLength(cell^.ConditionalFormatIndex, n+1);
      cell^.ConditionalFormatIndex[n] := AIndex;
    end;
end;


{@@ ----------------------------------------------------------------------------
  Creates a conditional format item for the cells given by ARange.
  The condition specified here must not require parameters, e.g. cfcEmpty
  The format is primarily stored in the worksheet's ConditionalFormats list,
  but the format index is also stored in the cell's ConditionalFormatIndex list.
  Returns the index of the conditional format in the workbook's CF list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteConditionalCellFormat(ARange: TsCellRange;
  ACondition: TsCFCondition; ACellFormatIndex: Integer): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddCellRule(Self, ARange,
    ACondition, ACellFormatIndex);
  StoreCFIndexInCells(self, Result, ARange);
end;


{@@ ----------------------------------------------------------------------------
  Creates a conditional format item for the cells given by ARange.
  The condition specified must require one parameter, e.g. cfcEqual,
  and the parameter must be specified as AParam.
  The format is primarily stored in the worksheet's ConditionalFormats list,
  but the format index is also stored in the cell's ConditionalFormatIndex list.
  Returns the index of the conditional format in the workbook's CF list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteConditionalCellFormat(ARange: TsCellRange;
  ACondition: TsCFCondition; AParam: Variant; ACellFormatIndex: Integer): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddCellRule(Self, ARange, ACondition,
    AParam, ACellFormatIndex);
  StoreCFIndexInCells(self, Result, ARange);
end;


{@@ ----------------------------------------------------------------------------
  Creates a conditional format item for the cells given by ARange.
  The condition specified must requored two parameters, e.g. cfcBetween,
  and the parameters must be specified as AParam1 and AParam2.
  The format is primarily stored in the worksheet's ConditionalFormats list,
  but the format index is also stored in the cell's ConditionalFormatIndex list.
  Returns the index of the conditional format in the workbook's CF list.
-------------------------------------------------------------------------------}
function TsWorksheet.WriteConditionalCellFormat(ARange: TsCellRange;
  ACondition: TsCFCondition; AParam1, AParam2: Variant;
  ACellFormatIndex: Integer): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddCellRule(Self, ARange, ACondition,
    AParam1, AParam2, ACellFormatIndex);
  StoreCFIndexInCells(self, Result, ARange);
end;


{@@ ----------------------------------------------------------------------------
  Writes the conditional format "color range"
-------------------------------------------------------------------------------}
function TsWorksheet.WriteColorRange(ARange: TsCellRange;
  AStartColor, AEndColor: TsColor): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddColorRangeRule(Self, ARange,
    AStartColor, AEndColor);
  StoreCFIndexInCells(Self, Result, ARange);
end;


function TsWorksheet.WriteColorRange(ARange: TsCellRange;
  AStartColor, ACenterColor, AEndColor: TsColor): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddColorRangeRule(Self, ARange,
    AStartColor, ACenterColor, AEndColor);
  StoreCFIndexInCells(Self, Result, ARange);
end;


function TsWorksheet.WriteColorRange(ARange: TsCellRange;
  AStartColor: TsColor; AStartKind: TsCFValueKind; AStartValue: Double;
  AEndColor: TsColor; AEndKind: TsCFValueKind; AEndValue: Double): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddColorRangeRule(Self, ARange,
    AStartColor, AStartKind, AStartValue,
    AEndColor, AEndKind, AEndValue);
  StoreCFIndexInCells(Self, Result, ARange);
end;


function TsWorksheet.WriteColorRange(ARange: TsCellRange;
  AStartColor: TsColor; AStartKind: TsCFValueKind; AStartValue: Double;
  ACenterColor: TsColor; ACenterKind: TsCFValueKind; ACenterValue: Double;
  AEndColor: TsColor; AEndKind: TsCFValueKind; AEndValue: Double): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddColorRangeRule(Self, ARange,
    AStartColor, AStartKind, AStartValue,
    ACenterColor, ACenterKind, ACenterValue,
    AEndColor, AEndKind, AEndValue);
  StoreCFIndexInCells(Self, Result, ARange);
end;


{@@ ----------------------------------------------------------------------------
  Writes the conditional format "data bars"
-------------------------------------------------------------------------------}
function TsWorksheet.WriteDataBars(ARange: TscellRange; ABarColor: TsColor): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddDataBarRule(Self, ARange, ABarColor);
  StoreCFIndexInCells(self, Result, ARange);
end;


function TsWorksheet.WriteDataBars(ARange: TscellRange; ABarColor: TsColor;
  AStartKind: TsCFValueKind; AStartValue: Double;
  AEndKind: TsCFValueKind; AEndValue: Double): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddDataBarRule(Self, ARange,
    ABarColor,
    AStartKind, AStartValue,
    AEndKind, AEndValue
  );
  StoreCFIndexInCells(self, Result, ARange);
end;


{@@ ----------------------------------------------------------------------------
  Writes the conditional format "icon set"
-------------------------------------------------------------------------------}

function TsWorksheet.WriteIconSet(ARange: TsCellRange; AIconSet: TsCFIconSet;
  AHideText: Boolean = false; AReverse: Boolean = false): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddIconSetRule(Self, ARange,
    AIconSet,
    AHideText, AReverse);
  StoreCFIndexInCells(self, Result, ARange);
end;

function TsWorksheet.WriteIconSet(ARange: TsCellRange; AIconSet: TsCFIconSet;
  AValueKind1: TsCFValueKind; AValue1: Double;
  AValueKind2: TsCFValueKind; AValue2: Double;
  AHideText: Boolean = false; AReverse: Boolean = false): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddIconSetRule(Self, ARange,
    AIconSet,
    AValueKind1, AValue1,
    AValueKind2, AValue2,
    AHideText, AReverse
  );
  if Result <> -1 then
    StoreCFIndexInCells(self, Result, ARange);
end;

function TsWorksheet.WriteIconSet(ARange: TsCellRange; AIconSet: TsCFIconSet;
  AValueKind1: TsCFValueKind; AValue1: Double;
  AValueKind2: TsCFValueKind; AValue2: Double;
  AValueKind3: TsCFValueKind; AValue3: Double;
  AHideText: Boolean = false; AReverse: Boolean = false): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddIconSetRule(Self, ARange,
    AIconSet,
    AValueKind1, AValue1,
    AValueKind2, AValue2,
    AValueKind3, AValue3,
    AHideText, AReverse
  );
  if Result <> -1 then
    StoreCFIndexInCells(self, Result, ARange);
end;

function TsWorksheet.WriteIconSet(ARange: TsCellRange; AIconSet: TsCFIconSet;
  AValueKind1: TsCFValueKind; AValue1: Double;
  AValueKind2: TsCFValueKind; AValue2: Double;
  AValueKind3: TsCFValueKind; AValue3: Double;
  AValueKind4: TsCFValueKind; AValue4: Double;
  AHideText: Boolean = false; AReverse: Boolean = false): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddIconSetRule(Self, ARange,
    AIconSet,
    AValueKind1, AValue1,
    AValueKind2, AValue2,
    AValueKind3, AValue3,
    AValueKind4, AValue4,
    AHideText, AReverse
  );
  if Result <> -1 then
    StoreCFIndexInCells(self, Result, ARange);
end;



{==============================================================================}
{   TsWorkbook code for conditional formats                                    }
{==============================================================================}

function TsWorkbook.GetConditionalFormat(AIndex: Integer): TsConditionalFormat;
begin
  Result := FConditionalFormatList[AIndex] as TsConditionalFormat;
end;


function TsWorkbook.GetNumConditionalFormats: Integer;
begin
  Result := FConditionalFormatList.Count;
end;

