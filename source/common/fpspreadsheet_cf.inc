{ Included by fpspreadsheet.pas }

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
  AStartColor: TsColor; AStartKind: TsCFColorRangeValueKind; AStartValue: Double;
  AEndColor: TsColor; AEndKind: TsCFColorRangeValueKind; AEndValue: Double): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddColorRangeRule(Self, ARange,
    AStartColor, AStartKind, AStartValue,
    AEndColor, AEndKind, AEndValue);
  StoreCFIndexInCells(Self, Result, ARange);
end;

function TsWorksheet.WriteColorRange(ARange: TsCellRange;
  AStartColor: TsColor; AStartKind: TsCFColorRangeValueKind; AStartValue: Double;
  ACenterColor: TsColor; ACenterKind: TsCFColorRangeValueKind; ACenterValue: Double;
  AEndColor: TsColor; AEndKind: TsCFColorRangeValueKind; AEndValue: Double): Integer;
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
function TsWorksheet.WriteDataBars(ARange: TscellRange): Integer;
begin
  Result := FWorkbook.FConditionalFormatList.AddDataBarRule(Self, ARange);
  StoreCFIndexInCells(self, Result, ARange);
end;

